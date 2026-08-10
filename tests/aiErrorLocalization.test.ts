import test from 'node:test';
import assert from 'node:assert/strict';
import { exactJapanese } from '../src/i18n/LanguageContext.tsx';
import { PROXY_ERROR_MESSAGE_CODES, proxyErrorMessageForCode } from '../src/services/apiHelper.ts';
import { MODEL_JSON_ERROR_MESSAGES } from '../src/services/aiRetry.ts';

// This test is the guard-rail for HIGH 1: the AI error strings emitted by the
// service layer and the Japanese dictionary that localises them live in
// different files, so they can silently drift apart. Enumerating every
// producible English message here and asserting an exact Japanese entry exists
// makes any future divergence a failing test rather than mixed-language garbage
// on screen for Japanese users.

/**
 * Static AI-error messages thrown directly by the owned services (not routed
 * through `proxyErrorMessageForCode`). Kept here as literals on purpose: if one
 * of these strings changes without a matching dictionary update, this test
 * fails and points at the exact message.
 */
const STANDALONE_AI_ERROR_MESSAGES: readonly string[] = [
  // azureOpenAI.ts — empty architecture guard (HIGH 3)
  'The AI model returned an empty architecture (no services). Please try again or rephrase your request.',
  // azureOpenAI.ts — configuration + generic fallbacks (MEDIUM 7 / LOW 10)
  'No AI model is configured. Check the environment configuration or connect a custom AI endpoint.',
  'Failed to generate architecture. Please try again.',
  // azureOpenAI.ts — IaC import + image-analysis guards (Issue 4: no interpolation)
  'Failed to parse the template. Please try again.',
  'The selected model does not support image analysis. Choose a vision-capable model in AI settings.',
];

test('every proxy error code produces a message with an exact Japanese entry', () => {
  const missing: Array<{ code: string; message: string; vision: boolean }> = [];

  for (const code of PROXY_ERROR_MESSAGE_CODES) {
    // Both the plain and the vision variant are user-reachable — a diagram
    // import passes vision: true, so its message must be localised too.
    for (const vision of [false, true]) {
      const message = proxyErrorMessageForCode(code, { vision });
      if (exactJapanese[message] === undefined) {
        missing.push({ code, message, vision });
      }
    }
  }

  assert.deepEqual(
    missing,
    [],
    `Missing exact Japanese entries for:\n${missing
      .map((m) => `  [${m.code}${m.vision ? ' vision' : ''}] ${m.message}`)
      .join('\n')}`,
  );
});

test('every ModelJsonError message has an exact Japanese entry', () => {
  for (const message of Object.values(MODEL_JSON_ERROR_MESSAGES)) {
    assert.notEqual(
      exactJapanese[message],
      undefined,
      `Missing exact Japanese entry for ModelJsonError message: ${message}`,
    );
  }
});

test('standalone AI-error messages have exact Japanese entries', () => {
  for (const message of STANDALONE_AI_ERROR_MESSAGES) {
    assert.notEqual(
      exactJapanese[message],
      undefined,
      `Missing exact Japanese entry for: ${message}`,
    );
  }
});

test('Japanese entries are real translations, not English copies', () => {
  // A JA value identical to its English key is a copy-paste miss, not a
  // translation. Spot-check the AI-error surface we own.
  const sample = [
    ...Object.values(MODEL_JSON_ERROR_MESSAGES),
    ...STANDALONE_AI_ERROR_MESSAGES,
    proxyErrorMessageForCode('azure_openai_timeout'),
  ];
  for (const message of sample) {
    const ja = exactJapanese[message];
    assert.ok(ja, `no entry for: ${message}`);
    assert.notEqual(ja, message, `Japanese entry is an English copy for: ${message}`);
    assert.match(ja as string, /[\u3040-\u30ff\u4e00-\u9faf]/, `no kana/kanji in: ${message}`);
  }
});
