// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AIModelConfigurationError } from '../stores/modelSettingsStore';
import { getBYOAISettings, validateBYOAIProfile } from '../stores/byoAISettingsStore';
import {
  beginBYOAIConnectionTest,
  finishBYOAIConnectionTest,
  isBYOAIConnectionTestCurrent,
  isValidBYOAIApiKey,
  readBYOAIConnectionSecret,
  type BYOAIConnectionError,
} from './byoAIConnectionSession';
import {
  buildRequestBody, callAzureOpenAIProxy, createOpenAIProxyError, OpenAIProxyError, parseApiResponse,
} from './apiHelper';
import { awaitWithAISignal, loadRuntimeConfig, runtimeConfigCancellationError } from './runtimeConfig';

export const BYO_AI_TEST_TIMEOUT_MS = 45_000;
export const BYO_AI_TEST_MAX_TOKENS = 4096;

function configurationError(code: string, message: string): AIModelConfigurationError {
  return new AIModelConfigurationError(code, message);
}

function completeTestResponse(data: any, apiFormat: 'responses' | 'chat-completions'): boolean {
  if (data?.error) return false;
  if (apiFormat === 'responses') {
    if (data?.status !== 'completed' || data.incomplete_details) return false;
    if (Array.isArray(data.output) && data.output.some((item: any) =>
      (item?.status !== undefined && item.status !== 'completed')
      || (Array.isArray(item?.content) && item.content.some((part: any) => part?.type === 'refusal')))) return false;
  } else {
    const choice = data?.choices?.[0];
    if (choice?.finish_reason !== 'stop' || choice.message?.refusal || choice.message?.tool_calls?.length) return false;
  }
  try {
    const content = JSON.parse(parseApiResponse(data, apiFormat).content);
    return content && typeof content === 'object' && !Array.isArray(content)
      && content.status === 'ok' && Object.keys(content).length === 1;
  } catch { return false; }
}

/**
 * Explicit, bounded smoke test, never a generation or an activation.
 * A key/configuration edit aborts the operation and revokes its verification token.
 */
export async function testBYOAIConnection(
  profileId: string,
  options: { signal?: AbortSignal } = {},
): Promise<{ profileId: string; verified: true; revision: number }> {
  if (options.signal?.aborted) throw runtimeConfigCancellationError();
  const profile = getBYOAISettings().profiles.find(item => item.id === profileId);
  if (!profile) throw configurationError('byo_profile_missing', 'The AI connection profile is missing. Select or create a profile.');
  const validation = validateBYOAIProfile(profile);
  if (!validation.valid || !validation.profile) {
    throw configurationError('byo_invalid_profile', 'The AI connection settings are invalid. Edit the profile and test again.');
  }
  const controller = new AbortController();
  const cancel = () => controller.abort();
  options.signal?.addEventListener('abort', cancel, { once: true });
  let timedOut = false;
  const timer = setTimeout(() => { timedOut = true; controller.abort(); }, BYO_AI_TEST_TIMEOUT_MS);
  const test = beginBYOAIConnectionTest(profileId, controller);
  const assertCurrent = () => {
    if (!isBYOAIConnectionTestCurrent(profileId, test.token, test.revision)) {
      throw configurationError('stale_ai_configuration', 'The AI connection changed during its test. Test the updated profile again.');
    }
    if (controller.signal.aborted) throw runtimeConfigCancellationError();
  };
  try {
    const apiKey = readBYOAIConnectionSecret(profileId, test.revision);
    if (!isValidBYOAIApiKey(apiKey)) throw configurationError('byo_key_required', 'Enter this profile’s API key before testing the connection.');
    // Force an actual server capability read, not the settings UI's cached flag.
    const capability = await loadRuntimeConfig(true, { signal: controller.signal });
    assertCurrent();
    if (capability.status !== 'ready') {
      throw configurationError('byo_availability_unknown', 'The application server could not confirm bring-your-own AI availability.');
    }
    if (!capability.bringYourOwnAI) {
      throw configurationError('byo_not_enabled', 'Bring-your-own AI is disabled by the application administrator.');
    }
    const request = buildRequestBody({
      deployment: profile.model,
      messages: [{ role: 'user', content: 'Connection test only. Reply with exactly this JSON object and nothing else: {"status":"ok"}' }],
      apiFormat: profile.apiFormat,
      maxTokens: Math.min(profile.maxCompletionTokens, BYO_AI_TEST_MAX_TOKENS),
      isReasoning: profile.isReasoning,
      reasoningEffort: profile.reasoningEffort,
    });
    const result = await awaitWithAISignal(callAzureOpenAIProxy({
      apiFormat: profile.apiFormat,
      deployment: profile.model,
      body: request,
      byo: { provider: profile.provider, endpoint: validation.profile.endpoint, apiKey },
      purpose: 'connection-test',
      signal: controller.signal,
    }), controller.signal);
    assertCurrent();
    if (!result.ok) throw createOpenAIProxyError(result);
    if (!completeTestResponse(result.data, profile.apiFormat)) {
      throw configurationError('byo_test_incomplete', 'The connection test did not return a complete expected response. Check the model, API format, and output settings, then test again.');
    }
    if (!finishBYOAIConnectionTest(profileId, test.token, test.revision, { verified: true })) {
      throw configurationError('stale_ai_configuration', 'The AI connection changed during its test. Test the updated profile again.');
    }
    return { profileId, verified: true, revision: test.revision };
  } catch (error) {
    const failure = options.signal?.aborted ? runtimeConfigCancellationError()
      : timedOut ? configurationError('byo_test_timeout', 'The connection test timed out. Check the endpoint and model, then test again.')
      : !isBYOAIConnectionTestCurrent(profileId, test.token, test.revision)
        ? configurationError('stale_ai_configuration', 'The AI connection changed during its test. Test the updated profile again.')
        : error instanceof OpenAIProxyError || error instanceof AIModelConfigurationError ? error
          : configurationError('byo_test_failed', 'The connection test failed. Check the connection settings and try again.');
    const details: BYOAIConnectionError | undefined = failure.name === 'AbortError' ? undefined : {
      code: 'code' in failure ? String(failure.code) : 'byo_test_failed',
      source: 'source' in failure ? String(failure.source) : 'client',
      message: failure.message,
      ...(failure instanceof OpenAIProxyError ? { status: failure.status, requestId: failure.requestId } : {}),
    };
    finishBYOAIConnectionTest(profileId, test.token, test.revision, { verified: false, error: details });
    throw failure;
  } finally {
    clearTimeout(timer);
    options.signal?.removeEventListener('abort', cancel);
  }
}
