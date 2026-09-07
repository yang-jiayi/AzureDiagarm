import test from 'node:test';
import assert from 'node:assert/strict';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { exactJapanese, LanguageProvider, useLanguage, type Language } from '../src/i18n/LanguageContext.tsx';
import { PROXY_ERROR_MESSAGE_CODES, proxyErrorMessageForCode } from '../src/services/apiHelper.ts';
import { MODEL_JSON_ERROR_MESSAGES } from '../src/services/aiRetry.ts';
import { AIResponseValidationError } from '../src/services/aiResponseValidation.ts';
import { BYOAI_STORAGE_ERROR_MESSAGES } from '../src/stores/byoAISettingsStore.ts';

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
  'The AI model returned an invalid architecture. Try again or revise the request.',
  // azureOpenAI.ts — generic fallbacks (MEDIUM 7 / LOW 10)
  'Failed to generate architecture. Please try again.',
  // azureOpenAI.ts — IaC import (Issue 4: no interpolation)
  'Failed to parse the template. Please try again.',
  // Shared request admission errors also reach the generator's BOTH mode.
  'The AI concurrency budget is invalid. Check the server configuration.',
  'The AI concurrency budget could not be checked. Please try again.',
  'AI request capacity stayed busy. Try again when capacity is available.',
  'AI capacity was repeatedly claimed by other requests. Try this model again.',
  'The AI request timed out after 225 seconds. This timeout was not automatically retried. Try again later; failures with unknown usage may still count toward the application budget.',
];

const ASTRA_CONFIGURATION_MESSAGES = [
  'Only the managed GPT-6 Astra model can run in this application.',
  'Only the configured GPT-6 Astra deployment can run.',
  'The requested reasoning effort is not supported by GPT-6 Astra.',
  'GPT-6 Astra requests must use the Responses API.',
  'Bring-your-own AI is disabled. Only managed GPT-6 Astra can run.',
  'No deployment configured for GPT-6 Astra. Set VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA.',
  'Azure OpenAI is not configured. Please check your environment.',
  'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment.',
  'The configured GPT-6 Astra deployment rejected the image analysis request. Check the image and contact the application administrator if the problem persists.',
];

const ASTRA_SERVER_MESSAGES = [
  'Only GPT-6 Astra Responses requests are supported.',
  'Custom AI endpoints are not supported.',
  'Configure the explicit GPT-6 Astra deployment and its identical singleton allowlist.',
  'Only the configured GPT-6 Astra deployment is allowed.',
  'deployment and body.model must explicitly match the configured GPT-6 Astra alias without endpoint overrides.',
  'Use a Responses request body for the configured GPT-6 Astra deployment.',
  'The managed Azure OpenAI endpoint is not configured correctly.',
];

const BYO_CONNECTION_MESSAGES = [
  'Enter a connection name of 1–80 characters.',
  'Use an Azure resource HTTPS origin or https://api.openai.com, without credentials, paths, ports, query, or fragment.',
  'Enter a whole-number output limit from 1 to 32768 tokens.',
  'Enter a valid API key. Keys stay only in this browser tab memory.',
  'You can save up to 10 AI connection profiles.',
  'The selected AI connection is missing. Select another profile or managed GPT-6 Astra.',
  'Enter this profile’s API key and test the connection in this browser tab.',
  'Test this AI connection successfully before using it.',
  'This AI connection failed its test. Check its settings and key, then test again.',
  'Bring-your-own AI is disabled by the application administrator.',
  'Bring-your-own AI availability has not been confirmed by the application server.',
  'The selected AI connection has invalid settings. Edit the profile and test again.',
  'The selected AI connection changed before this request was sent. Review the connection and submit again.',
  'The AI connection changed during its test. Test the updated profile again.',
  'The connection test did not return a complete expected response. Check the model, API format, and output settings, then test again.',
  'The connection test timed out. Check the endpoint and model, then test again.',
  'The connection test failed. Check the connection settings and try again.',
  'The AI provider rejected this profile’s API key. Re-enter the key and test the connection again.',
  'The AI connection settings are invalid. Check the provider, endpoint, and model, then test again.',
  'The selected AI connection does not support images. Select a vision-capable connection.',
];

test('BYO status and validation errors have exact Japanese guidance without automatic fallback', () => {
  for (const message of [...BYO_CONNECTION_MESSAGES, ...Object.values(BYOAI_STORAGE_ERROR_MESSAGES)]) {
    const translated = exactJapanese[message];
    assert.ok(translated, `Missing Japanese BYO error: ${message}`);
    assert.match(translated, /[\u3040-\u30ff\u4e00-\u9faf]/);
    assert.doesNotMatch(translated, /自動で.*Astra|GPT-5/);
  }
});

test('Astra-only configuration errors have exact Japanese entries without model-switching advice', () => {
  for (const message of [...ASTRA_CONFIGURATION_MESSAGES, ...ASTRA_SERVER_MESSAGES]) {
    const translated = exactJapanese[message];
    assert.ok(translated, `Missing Japanese configuration error: ${message}`);
    assert.match(translated, /[\u3040-\u30ff\u4e00-\u9faf]/);
    assert.doesNotMatch(translated, /別のモデル|モデルを選択|カスタム.*接続|GPT-5/);
    if (message.includes('GPT-6 Astra')) assert.match(translated, /GPT-6 Astra/);
    if (message.includes('VITE_')) assert.match(translated, /VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA/);
  }
});

for (const language of ['en', 'ja'] as const) {
  test(`Astra-only ${language} UI and server policy errors preserve reasoning choices request IDs and historical names`, t => {
    const originalStorage = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
    Object.defineProperty(globalThis, 'localStorage', {
      configurable: true, value: { getItem: () => language },
    });
    t.after(() => {
      if (originalStorage) Object.defineProperty(globalThis, 'localStorage', originalStorage);
      else Reflect.deleteProperty(globalThis, 'localStorage');
    });
    const observed: Record<string, string> = {};
    function Probe() {
      const { t: text, translate } = useLanguage();
      observed.defaultReasoning = text('ai.defaultReasoning');
      observed.perFeature = text('ai.perFeatureReasoning');
      observed.inherit = text('ai.defaultReasoningChoice', { effort: text('High') });
      observed.setup = text('ai.astraNotConfigured');
      observed.unsupported = translate(`${ASTRA_CONFIGURATION_MESSAGES[1]} Request ID: astra-case-123`);
      observed.history = translate('GPT-5.6 Sol / retired-deployment-123');
      observed.byoFailure = translate(`${BYO_CONNECTION_MESSAGES[17]} Request ID: byo-key-rejected-123`);
      observed.byoRequestFailure = translate(`${proxyErrorMessageForCode('byo_request_failed', { status: 400 })} Request ID: byo-request-rejected-123`);
      ASTRA_SERVER_MESSAGES.forEach((message, index) => {
        observed[message] = translate(message);
        observed[`policy-${index}`] = translate(`${message} Request ID: astra-policy-${index}`);
      });
      return null;
    }
    renderToStaticMarkup(createElement(LanguageProvider, { children: createElement(Probe) }));
    const expected: Record<Language, [string, string, string]> = {
      en: ['Default reasoning', 'Per-feature reasoning', 'Default (High)'],
      ja: ['既定の推論強度', '機能別の推論強度', '既定値 (高)'],
    };
    assert.deepEqual([observed.defaultReasoning, observed.perFeature, observed.inherit], expected[language]);
    assert.match(observed.setup, /GPT-6 Astra/);
    assert.match(observed.setup, language === 'ja' ? /管理者/ : /administrator/);
    assert.equal(observed.unsupported, language === 'ja'
      ? `${exactJapanese[ASTRA_CONFIGURATION_MESSAGES[1]]} リクエスト ID: astra-case-123`
      : `${ASTRA_CONFIGURATION_MESSAGES[1]} Request ID: astra-case-123`);
    assert.equal(observed.history, 'GPT-5.6 Sol / retired-deployment-123');
    assert.equal(observed.byoFailure, language === 'ja'
      ? `${exactJapanese[BYO_CONNECTION_MESSAGES[17]]} リクエスト ID: byo-key-rejected-123`
      : `${BYO_CONNECTION_MESSAGES[17]} Request ID: byo-key-rejected-123`);
    assert.match(observed.byoRequestFailure, /byo-request-rejected-123/);
    assert.match(observed.byoRequestFailure, language === 'ja' ? /リクエスト.*失敗.*リクエスト ID:/ : /request failed.*Request ID:/);
    ASTRA_SERVER_MESSAGES.forEach((message, index) => {
      const translated = language === 'ja' ? exactJapanese[message] : message;
      assert.equal(observed[message], translated);
      assert.equal(observed[`policy-${index}`],
        `${translated} ${language === 'ja' ? 'リクエスト ID:' : 'Request ID:'} astra-policy-${index}`);
      assert.doesNotMatch(observed[message], /choose.*model|try using GPT-|別のモデル|モデルを選択|GPT-5/i);
      if (message.includes('Responses')) assert.match(observed[message], /Responses/);
      if (message.includes('body.model')) {
        assert.match(observed[message], /body\.model/);
        assert.match(observed[message], /deployment/);
      }
    });
  });
}

test('structured response errors keep diagnostics separate from their translated user message', () => {
  for (const detail of ['Blueprint parent cycle.', 'Duplicate resource IDs.', 'Reference stages are empty.']) {
    const error = new AIResponseValidationError(detail);
    assert.equal(error.detail, detail);
    assert.equal(error.retryable, false);
    assert.ok(STANDALONE_AI_ERROR_MESSAGES.includes(error.message));
    assert.match(exactJapanese[error.message], /無効な構造/);
    assert.ok(!error.message.includes(detail));
  }
});

test('every proxy error code produces a message with an exact Japanese entry', () => {
  const missing: Array<{ code: string; message: string; vision: boolean }> = [];

  for (const code of PROXY_ERROR_MESSAGE_CODES) {
    // Both the plain and the vision variant are user-reachable — a diagram
    // import passes vision: true, so its message must be localised too.
    for (const vision of [false, true]) {
      const message = proxyErrorMessageForCode(code, { vision });
      assert.doesNotMatch(message, /choose.*model|try using GPT-|select a different model/i,
        `Normalized ${code} must not suggest switching models`);
      if (exactJapanese[message] === undefined) {
        missing.push({ code, message, vision });
      } else {
        assert.doesNotMatch(exactJapanese[message], /別のモデル|モデルを選択|GPT-5/,
          `Normalized ${code} must not suggest switching models in Japanese`);
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
