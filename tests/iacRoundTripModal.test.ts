import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { build } from 'esbuild';
import type { ComponentProps } from 'react';
import type IaCRoundTripModal from '../src/components/IaCRoundTripModal';
import { buildIaCBaseline, buildStarterTemplate, compareDiagramToBaseline } from '../src/services/iacRoundTrip';

type ModalProps = ComponentProps<typeof IaCRoundTripModal>;
type RenderModal = (props: ModalProps, language: 'en' | 'ja') => string;
let renderer: Promise<RenderModal> | undefined;

function loadRenderer(): Promise<RenderModal> {
  renderer ??= (async () => {
    const bundle = await build({
      stdin: {
        contents: `
          import React from 'react';
          import { renderToStaticMarkup } from 'react-dom/server';
          import Modal from './src/components/IaCRoundTripModal';
          import { setLanguage } from './src/i18n/LanguageContext';
          export function render(props, language) {
            setLanguage(language);
            return renderToStaticMarkup(React.createElement(Modal, props));
          }
        `,
        resolveDir: process.cwd(), loader: 'tsx',
      },
      bundle: true, write: false, platform: 'node', format: 'cjs', packages: 'external',
      loader: { '.css': 'empty' }, logLevel: 'silent',
      plugins: [{
        name: 'iac-modal-language-boundary',
        setup(builder) {
          builder.onResolve({ filter: /\/i18n\/LanguageContext$/ }, () => ({
            path: 'language-context', namespace: 'language-boundary',
          }));
          builder.onLoad({ filter: /.*/, namespace: 'language-boundary' }, () => ({
            contents: `
              let language = 'en';
              export const setLanguage = value => { language = value; };
              export const useLanguage = () => ({ language, t: key => key });
            `,
            loader: 'js',
          }));
        },
      }],
    });
    const module: { exports: unknown } = { exports: {} };
    new Function('module', 'exports', 'require', bundle.outputFiles[0].text)(
      module, module.exports, createRequire(import.meta.url),
    );
    const exported = module.exports;
    assert.ok(typeof exported === 'object' && exported !== null && 'render' in exported);
    assert.ok(typeof exported.render === 'function');
    const render = exported.render;
    return (props: ModalProps, language: 'en' | 'ja') => {
      const markup: unknown = render(props, language);
      assert.ok(typeof markup === 'string');
      return markup;
    };
  })();
  return renderer;
}

function fixture(incomplete: boolean) {
  const baseline = buildIaCBaseline({
    format: 'bicep',
    importedAt: '2026-09-06T00:00:00.000Z',
    files: incomplete ? [{ name: 'partial.bicep', text: "module child './child.bicep' = { name: 'child' }" }] : [],
  });
  const comparison = compareDiagramToBaseline([], baseline);
  assert.ok(comparison);
  return {
    isOpen: true, onClose() {}, baseline, comparison, driftPlan: null,
    onImportDriftPlan() {}, onClearDriftPlan() {}, onDownloadStarter() {},
    bicepStarter: buildStarterTemplate([], 'bicep'),
    terraformStarter: buildStarterTemplate([], 'terraform'),
    diagramServiceCount: 0,
  } satisfies ModalProps;
}

for (const language of ['en', 'ja'] as const) {
  test(`partial IaC comparisons visibly qualify counts and empty states in ${language}`, async () => {
    const render = await loadRenderer();
    const html = render(fixture(true), language);
    const title = language === 'ja' ? '暫定比較：ソースの一部のみ解析' : 'Provisional comparison: partially parsed source';
    assert.ok(html.includes(title), 'the incomplete comparison needs an explicit localized notice');
    assert.ok(html.includes('role="note"'), 'the disclosure is available to assistive technology');
    assert.ok(html.indexOf('iac-provisional-notice') < html.indexOf('iac-roundtrip-summary'),
      'the qualification precedes the counts, rather than being buried in parsing notes');
    assert.ok(!html.includes('Every baseline resource has a current match.'));
    assert.ok(!html.includes('すべてのベースライン リソースに現在の一致があります。'));
    assert.ok(!html.includes('No new service nodes beyond the baseline.'));
    assert.ok(!html.includes('ベースラインを超える新しいサービス ノードはありません。'));
  });

  test(`complete IaC comparisons retain their existing claims in ${language}`, async () => {
    const render = await loadRenderer();
    const html = render(fixture(false), language);
    assert.ok(!html.includes('iac-provisional-notice'));
    assert.ok(html.includes(language === 'ja'
      ? 'すべてのベースライン リソースに現在の一致があります。'
      : 'Every baseline resource has a current match.'));
  });
}

test('legacy warnings and comparison-only incomplete status cannot be hidden by a false baseline flag', async () => {
  const render = await loadRenderer();
  const partial = fixture(true);
  const complete = fixture(false);
  for (const props of [
    { ...partial, baseline: { ...partial.baseline, incomplete: false }, comparison: { ...partial.comparison, incomplete: false } },
    { ...complete, comparison: { ...complete.comparison, incomplete: true } },
    { ...complete, comparison: { ...complete.comparison, warnings: ['The source could not be fully read.'] } },
  ]) {
    assert.ok(render(props, 'en').includes('iac-provisional-notice'), 'partial results must stay qualified');
  }
  assert.equal(render({ ...partial, isOpen: false }, 'en'), '', 'closed dialogs render no content');
});
