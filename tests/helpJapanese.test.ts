// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { helpJapanese } from '../src/i18n/helpJapanese';

// Proper nouns whose official Japanese form is intentionally identical to the
// English brand name. Keep this list as small as possible.
const IDENTICAL_ALLOWED = new Set<string>([
  'Azure Well-Architected Framework',
]);

// Hiragana, Katakana, or CJK ideographs.
const hasJapanese = (value: string): boolean => /[\u3040-\u30FF\u4E00-\u9FFF]/.test(value);

test('helpJapanese: every entry has a non-empty Japanese translation distinct from its English key', () => {
  const entries = Object.entries(helpJapanese);
  assert.ok(entries.length > 0, 'helpJapanese should not be empty');
  for (const [key, value] of entries) {
    assert.equal(typeof value, 'string', `value for "${key}" must be a string`);
    assert.ok(value.trim().length > 0, `value for "${key}" must not be empty`);
    if (!IDENTICAL_ALLOWED.has(key)) {
      assert.notEqual(value, key, `value for "${key}" must not be identical to the English key`);
      assert.ok(
        hasJapanese(value),
        `value for "${key}" must contain Japanese characters (got "${value}")`,
      );
    }
  }
});

// --- Guard: no GuidedHelpPanel string silently falls back to English ---

const panelSource = readFileSync(
  new URL('../src/components/GuidedHelpPanel.tsx', import.meta.url),
  'utf8',
);

const unescape = (raw: string): string => raw.replace(/\\(['"\\])/g, '$1');

// Extract every English source string that the panel routes through translate()
// and therefore needs a Japanese entry. The structured prompt is special-cased in
// the component (STRUCTURED_PROMPT_JA) and is intentionally excluded.
const collectRequiredKeys = (source: string): Set<string> => {
  const keys = new Set<string>();
  const quoted = /(['"])((?:\\.|(?!\1)[\s\S])*?)\1/g;

  // A. Inline translate('...') / translate("...") calls.
  const inline = /\btranslate\(\s*(['"])((?:\\.|(?!\1)[\s\S])*?)\1\s*\)/g;
  for (const m of source.matchAll(inline)) keys.add(unescape(m[2]));

  // B. Data-array content props: label / title / detail / body.
  const props = /\b(?:label|title|detail|body):\s*(['"])((?:\\.|(?!\1)[\s\S])*?)\1/g;
  for (const m of source.matchAll(props)) keys.add(unescape(m[2]));

  // C. steps: [...] arrays (Choose a Path).
  const stepsBlocks = /steps:\s*\[([^\]]*)\]/g;
  for (const block of source.matchAll(stepsBlocks)) {
    for (const m of block[1].matchAll(quoted)) keys.add(unescape(m[2]));
  }

  // D. EXAMPLE_PROMPTS array (Prompt Lab).
  const examples = /const EXAMPLE_PROMPTS\s*=\s*\[([\s\S]*?)\];/.exec(source);
  if (examples) {
    for (const m of examples[1].matchAll(quoted)) keys.add(unescape(m[2]));
  }

  // E. FeatureSection JSX props (double-quoted eyebrow / title / intro).
  const jsxProps = /\b(?:eyebrow|intro|title)="((?:[^"\\]|\\.)*)"/g;
  for (const m of source.matchAll(jsxProps)) keys.add(unescape(m[1]));

  keys.delete('');
  return keys;
};

test('GuidedHelpPanel: every rendered string has a Japanese translation (no English fallback)', () => {
  const required = collectRequiredKeys(panelSource);
  assert.ok(required.size > 40, `expected to extract many strings, got ${required.size}`);
  const missing = [...required].filter((key) => !(key in helpJapanese));
  assert.deepEqual(
    missing,
    [],
    `these GuidedHelpPanel strings are missing from helpJapanese and would render in English:\n${missing.join('\n')}`,
  );
});

test('GuidedHelpPanel: the structured prompt has a dedicated Japanese variant', () => {
  assert.match(panelSource, /const STRUCTURED_PROMPT\s*=/);
  assert.match(panelSource, /const STRUCTURED_PROMPT_JA\s*=/);
});
