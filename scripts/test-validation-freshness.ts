import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  getCurrentValidationScore,
  resolveValidationFreshness,
} from '../src/utils/validationFreshness';

assert.deepEqual(
  resolveValidationFreshness(true, true),
  { keepResult: true, needsRefresh: true },
  'Applying recommendations should retain the previous report and require revalidation.',
);

assert.deepEqual(
  resolveValidationFreshness(true, false),
  { keepResult: false, needsRefresh: false },
  'Ordinary generation should clear an obsolete validation report.',
);

assert.deepEqual(
  resolveValidationFreshness(false, true),
  { keepResult: false, needsRefresh: false },
  'There is no report to retain when recommendations are applied without prior results.',
);

assert.equal(getCurrentValidationScore(88, false), 88);
assert.equal(
  getCurrentValidationScore(88, true),
  undefined,
  'Stale validation scores must not be saved, exported, or treated as current progress.',
);

const appSource = readFileSync(new URL('../src/App.tsx', import.meta.url), 'utf8');
assert.match(
  appSource,
  /handleAIGenerate\(\s*improvedArchitecture,\s*bannerText,\s*true,\s*true,\s*undefined,\s*true,\s*true,\s*\)/s,
  'The recommendation flow must pass the final preserve-validation argument.',
);
assert.match(
  appSource,
  /validationScore=\{currentValidationScore \?\? null\}/,
  'The workflow rail must ignore stale validation scores.',
);
assert.match(
  appSource,
  /validationScore: currentValidationScore/,
  'Downloaded diagram data must ignore stale validation scores.',
);

console.log('Validation freshness tests passed.');