import test from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import type { ValidationFinding } from '../src/services/architectureValidator';
import { normalizeValidationFindingSource } from '../src/services/validationFindingSource';
import { updateValidationReview } from '../src/services/validationReview';

const finding: ValidationFinding = {
  severity: 'high', category: 'Identity', issue: 'Authentication is missing',
  recommendation: 'Configure authentication', resources: ['Web'],
};

test('the legacy AI source requested by the validator becomes canonical before history validation', () => {
  const original = { ...finding, source: 'ai-analysis', id: 'identity-1', resourceIds: ['web-1'] };
  const normalized = normalizeValidationFindingSource(original);
  assert.equal(normalized.source, 'ai');
  assert.equal(original.source, 'ai-analysis');
  assert.deepEqual(normalized, { ...original, source: 'ai' });
  const records = updateValidationReview([], {
    overallScore: 70, summary: 'Review', timestamp: '2026-01-01T00:00:00Z',
    pillars: [{ pillar: 'Security', score: 70, findings: [normalized] }], quickWins: [],
  });
  assert.equal(records[0].finding.source, 'ai');
});

test('known sources remain distinct and absent provenance is never guessed', () => {
  for (const source of ['ai', 'rule-based'] as const) {
    assert.equal(normalizeValidationFindingSource({ ...finding, source }).source, source);
  }
  assert.equal(normalizeValidationFindingSource(finding).source, undefined);
  for (const source of ['invented', null, 1, {}]) {
    assert.throws(() => normalizeValidationFindingSource({ ...finding, source }), /source/);
  }
});

let validatorPromise: Promise<{
  validateArchitecture: typeof import('../src/services/architectureValidator').validateArchitecture;
  formatValidationReport: typeof import('../src/services/architectureValidator').formatValidationReport;
  setProviderContent: (content: string) => void;
}> | undefined;

function loadValidator() {
  validatorPromise ??= (async () => {
    const boundaries: Record<string, string> = {
      apiHelper: `
        let content = '';
        export const setProviderContent = value => { content = value; };
        export const buildRequestBody = () => ({});
        export const callAzureOpenAIProxy = async () => ({ ok: true, data: {} });
        export const parseApiResponse = () => ({ content, promptTokens: 1, completionTokens: 2, totalTokens: 3 });
        export const createOpenAIProxyError = () => new Error('Provider error');
        export const getApiFormatLabel = () => 'test';
      `,
      aiModelRuntime: `
        export const resolveAIModelRuntime = () => ({
          displayName: 'Test provider', telemetryModel: 'test', deployment: 'test',
          apiFormat: 'chat-completions', maxCompletionTokens: 8000, isReasoning: false,
        });
      `,
      telemetryService: `export const trackAIModelUsage = () => {};`,
    };
    const result = await build({
      stdin: {
        contents: `
          export { validateArchitecture, formatValidationReport } from './src/services/architectureValidator';
          export { setProviderContent } from './src/services/apiHelper';
        `,
        resolveDir: process.cwd(), loader: 'ts',
      },
      bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
      define: { 'import.meta.env': '{}' },
      plugins: [{
        name: 'validation-provider-boundaries',
        setup(builder) {
          builder.onResolve({ filter: /\/(apiHelper|aiModelRuntime|telemetryService)$/ }, args => ({
            path: args.path.split('/').at(-1)!, namespace: 'boundary',
          }));
          builder.onLoad({ filter: /.*/, namespace: 'boundary' }, args => ({
            contents: boundaries[args.path], loader: 'js',
          }));
        },
      }],
    });
    const module = { exports: {} };
    new Function('module', 'exports', result.outputFiles[0].text)(module, module.exports);
    return module.exports as Awaited<NonNullable<typeof validatorPromise>>;
  })();
  return validatorPromise;
}

test('the provider path retains remote findings, model metadata and canonical provenance', async t => {
  t.mock.method(console, 'log', () => {});
  const validator = await loadValidator();
  for (const source of [undefined, 'ai-analysis', 'ai', 'rule-based']) {
    validator.setProviderContent(JSON.stringify({
      overallScore: 0, summary: 'Provider review',
      pillars: [{ pillar: 'Security', score: 120, findings: [{
        ...finding, source, id: 'identity-1', resourceIds: ['web-1'],
        evidence: ['No identity service is shown.'], remediation: ['Configure identity.'],
        referenceUrl: 'https://learn.microsoft.com/azure/well-architected/security/',
        applyAction: { type: 'add-service', label: 'Add identity', serviceType: 'Microsoft Entra ID' },
      }] }],
    }));
    const report = await validator.validateArchitecture([{ name: 'Web', type: 'App Service', category: 'compute' }], []);
    const normalized = report.pillars[0].findings[0];
    assert.equal(normalized.source, source === 'ai-analysis' ? 'ai' : source);
    assert.equal(normalized.id, 'identity-1');
    assert.deepEqual(normalized.resourceIds, ['web-1']);
    assert.deepEqual(normalized.resources, ['Web']);
    assert.deepEqual(normalized.evidence, ['No identity service is shown.']);
    assert.deepEqual(normalized.remediation, ['Configure identity.']);
    assert.equal(normalized.applyAction?.serviceType, 'Microsoft Entra ID');
    assert.equal(report.overallScore, 0);
    assert.equal(report.pillars[0].score, 100);
    assert.deepEqual(report.quickWins, []);
    assert.equal(report.modelUsed, 'Test provider');
    assert.equal(report.metrics?.totalTokens, 3);
    const history = updateValidationReview([], report);
    assert.deepEqual(history[0].finding.applyAction, normalized.applyAction);
    if (source === undefined) {
      assert.match(validator.formatValidationReport(report), /\*\*Source:\*\* Not specified/);
    }
  }
});

test('the provider path rejects unknown provenance instead of silently relabelling it AI', async t => {
  t.mock.method(console, 'log', () => {});
  t.mock.method(console, 'error', () => {});
  const validator = await loadValidator();
  for (const source of ['unsupported-provider', null, 7]) {
    validator.setProviderContent(JSON.stringify({
      overallScore: 70, summary: 'Review', quickWins: [{ ...finding, source }],
      pillars: [{ pillar: 'Security', score: 70, findings: [] }],
    }));
    await assert.rejects(validator.validateArchitecture([], []), /Invalid validation finding source/);
  }
});
