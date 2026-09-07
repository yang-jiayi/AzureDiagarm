import test from 'node:test';
import assert from 'node:assert/strict';
import type { BYOAIProfile } from '../src/stores/byoAISettingsStore';
import {
  BYO_STORAGE_KEY, capabilityResponse, deferred, jsonResponse, loadBYOClient, settle, testResponse, verifiedProfile,
} from './fixtures/byoClientHarness';

const keyA = 'sk-offline-secret-profile-A';
const keyB = 'sk-offline-secret-profile-B';
const publicKeys = [
  'apiFormat', 'endpoint', 'id', 'isReasoning', 'maxCompletionTokens', 'model',
  'name', 'provider', 'reasoningEffort', 'supportsVision',
].sort();
const profile = (id = 'profile-a'): BYOAIProfile => ({
  id, name: `Connection ${id}`, provider: 'azure-openai', endpoint: 'https://contoso.openai.azure.com',
  model: 'customer-production-thinking', apiFormat: 'responses',
  reasoningEffort: 'max', isReasoning: true, supportsVision: true, maxCompletionTokens: 32768,
});

function successfulFetch(url: unknown, options?: RequestInit): Promise<Response> {
  assert.ok(url === '/api/runtime-config' || url === '/api/openai', 'tests must use only mocked application routes');
  return Promise.resolve(url === '/api/runtime-config' ? capabilityResponse()
    : testResponse(JSON.parse(String(options?.body)).apiFormat));
}

for (const version of [1, 2]) {
  for (const enabled of [true, false]) {
    test(`v${version} migration preserves enabled=${enabled}, public fields, and history but removes all secret/verification fields`, async () => {
      const history = JSON.stringify({ model: 'historical-deployment', apiKey: 'history-is-not-ours', diagram: { nodes: [] } });
      const entries = new Map([['azure-diagrams-history', history], ['review-history', history]]);
      const { client } = await loadBYOClient({
        version, enabled, ...profile(), name: 'Saved public label',
        apiKey: keyA, token: keyB, credentials: { apiKey: keyA }, verification: { valid: true },
        nested: { secret: keyB }, verified: true, capabilityMode: 'manual',
      }, undefined, entries);
      const saved = client.getBYOAISettings();
      assert.equal(saved.profiles.length, 1);
      const migrated = saved.profiles[0];
      assert.equal(saved.activeProfileId, enabled ? migrated.id : null);
      assert.equal(migrated.model, profile().model);
      assert.equal(migrated.reasoningEffort, 'max');
      assert.equal(migrated.maxCompletionTokens, 32768);
      assert.deepEqual(Object.keys(migrated).sort(), publicKeys);
      assert.equal(client.getBYOAIConnectionState(migrated.id).status, 'key-required');
      const persisted = entries.get(BYO_STORAGE_KEY)!;
      assert.doesNotMatch(persisted, /sk-offline|apiKey|token|credentials|verification|verified|nested|capabilityMode/);
      assert.equal(JSON.parse(persisted).version, 3);
      assert.equal(entries.get('azure-diagrams-history'), history);
      assert.equal(entries.get('review-history'), history);
      assert.equal(client.isAnyAIModelConfigured(), !enabled);
    });
  }
}

test('legacy OpenAI connection migrates to its fixed official origin and requires a fresh key/test', async () => {
  const { client } = await loadBYOClient({
    version: 1, enabled: true, provider: 'openai', endpoint: 'https://arbitrary.invalid/path?apiKey=secret',
    model: 'ft:gpt-customer:org:actual-name', apiFormat: 'chat-completions',
    reasoningEffort: 'minimal', isReasoning: false, supportsVision: false,
  });
  assert.deepEqual(client.getBYOAISettings().profiles[0], {
    id: 'byo-migrated', name: 'Migrated AI connection', provider: 'openai', endpoint: 'https://api.openai.com',
    model: 'ft:gpt-customer:org:actual-name', apiFormat: 'chat-completions',
    reasoningEffort: 'minimal', isReasoning: false, supportsVision: false, maxCompletionTokens: 32000,
  });
  assert.equal(client.getBYOAIConnectionState('byo-migrated').status, 'key-required');
  assert.throws(() => client.resolveAIModelRuntime('architectureGeneration'), { code: 'byo_availability_unknown' });
});

test('all public inputs, returned profiles, nested settings arrays and states are detached allowlisted snapshots', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const { client, entries } = await loadBYOClient();
  const input = { ...profile(), apiKey: keyA, nested: { password: keyB }, verified: true };
  const saved = client.upsertBYOAIProfile(input);
  input.name = 'Mutated input';
  input.model = 'mutated-model';
  saved.name = 'Mutated result';
  const settings = client.getBYOAISettings();
  settings.profiles[0].endpoint = 'https://attacker.invalid';
  settings.profiles.push(profile('rogue'));
  settings.activeProfileId = 'rogue';
  assert.equal(client.getBYOAISettings().profiles[0].name, profile().name);
  assert.equal(client.getBYOAISettings().profiles.length, 1);
  assert.equal(client.getBYOAISettings().activeProfileId, null);
  client.setBYOAIApiKey(saved.id, keyA);
  await client.testBYOAIConnection(saved.id);
  const state = client.getBYOAIConnectionState(saved.id);
  state.verified = false;
  state.status = 'failed';
  assert.equal(client.getBYOAIConnectionState(saved.id).verified, true);
  assert.doesNotMatch(entries.get(BYO_STORAGE_KEY)!, /apiKey|nested|password|verified|sk-offline/);
  for (const name of ['getBYOAIApiKey', 'saveBYOAIConfiguration', 'readBYOAIConnectionSecret']) {
    assert.equal(name in client, false, 'public store/service barrel must not expose credential getters');
  }
});

test('v3 rehydration strips nested credentials from every profile and never recovers keys/verification', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const { client, entries } = await loadBYOClient({
    version: 3, activeProfileId: 'profile-b', apiKey: keyA,
    profiles: [
      { ...profile(), apiKey: keyA, connection: { key: keyB }, verified: true },
      { ...profile('profile-b'), apiKey: keyB, token: keyA, status: 'verified' },
    ],
  });
  assert.equal(client.getBYOAISettings().activeProfileId, 'profile-b');
  assert.equal(client.getBYOAIConnectionState('profile-a').status, 'key-required');
  assert.equal(client.getBYOAIConnectionState('profile-b').status, 'key-required');
  client.setBYOAIApiKey('profile-b', keyB);
  await client.testBYOAIConnection('profile-b');
  client.selectBYOAIProfile('profile-b');
  const old = client.captureRuntimeModelOverride('validation');
  client.reloadBYOAISettings();
  assert.equal(client.getBYOAISettings().activeProfileId, 'profile-b');
  assert.equal(client.getBYOAIConnectionState('profile-b').status, 'key-required');
  assert.throws(() => client.resolveAIModelRuntime('validation', old), { code: 'stale_ai_configuration' });
  assert.doesNotMatch(entries.get(BYO_STORAGE_KEY)!, /apiKey|connection|token|verified|status|sk-offline/);
});

test('a second tab sees public selection but never another tab’s key or verified state', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const first = await loadBYOClient();
  const saved = await verifiedProfile(first.client, profile(), keyA);
  first.client.selectBYOAIProfile(saved.id);
  const second = await loadBYOClient(undefined, undefined, first.entries);
  assert.equal(second.client.getBYOAISettings().activeProfileId, saved.id);
  assert.equal(second.client.getBYOAIConnectionState(saved.id).hasApiKey, false);
  assert.equal(second.client.getBYOAIConnectionState(saved.id).verified, false);
  assert.equal(first.client.getBYOAIConnectionState(saved.id).verified, true);
});

test('ten named profiles are bounded without evicting connections; updates remain possible', async () => {
  const { client } = await loadBYOClient();
  for (let i = 0; i < 10; i++) client.upsertBYOAIProfile(profile(`p-${i}`));
  assert.throws(() => client.upsertBYOAIProfile(profile('p-10')), { code: 'byo_profile_limit' });
  client.upsertBYOAIProfile({ ...profile('p-0'), name: 'Renamed' });
  assert.equal(client.getBYOAISettings().profiles.length, 10);
  assert.equal(client.getBYOAISettings().profiles[0].name, 'Renamed');
  const loaded = await loadBYOClient({
    version: 3, activeProfileId: 'p-10',
    profiles: Array.from({ length: 12 }, (_, i) => ({ ...profile(`p-${i}`), secret: keyA })),
  });
  assert.equal(loaded.client.getBYOAISettings().profiles.length, 10);
  assert.equal(loaded.client.getEffectiveAIModelInfo('validation').code, 'byo_profile_missing');
});

test('malformed or unknown-provider selected profiles never become an implicit managed connection', async () => {
  for (const raw of [
    { version: 3, profiles: [], activeProfileId: 'removed' },
    { version: 3, profiles: [{ ...profile(), provider: 'anthropic' }], activeProfileId: 'profile-a' },
    { version: 3, profiles: [{ ...profile(), apiFormat: 'anthropic-messages' }], activeProfileId: 'profile-a' },
    { version: 3, profiles: [], activeProfileId: { token: keyA } },
  ]) {
    const { client } = await loadBYOClient(raw);
    assert.equal(client.getEffectiveAIModelInfo('blueprint').source, 'bring-your-own');
    assert.equal(client.isAnyAIModelConfigured(), false);
    assert.throws(() => client.resolveAIModelRuntime('validation'), { code: 'byo_profile_missing' });
    client.selectBYOAIProfile(null);
    assert.equal(client.isAnyAIModelConfigured(), true);
  }
  const entries = new Map([[BYO_STORAGE_KEY, 'not-json']]);
  const { client } = await loadBYOClient(undefined, undefined, entries);
  assert.equal(client.getEffectiveAIModelInfo('validation').code, 'byo_profile_missing');
  assert.equal(JSON.parse(entries.get(BYO_STORAGE_KEY)!).activeProfileId, 'missing-byo-profile');
  for (const value of [null, true, 'corrupt string', []]) {
    const loaded = await loadBYOClient(value);
    assert.equal(loaded.client.getEffectiveAIModelInfo('validation').code, 'byo_profile_missing');
  }
});

for (const endpoint of [
  'http://contoso.openai.azure.com', 'https://localhost', 'https://127.0.0.1', 'https://[::1]',
  'https://api.openai.com.evil.example', 'https://evil.example', 'https://openai.azure.com',
  'https://nested.contoso.openai.azure.com', 'https://contoso.openai.azure.com:443',
  'https://contoso.openai.azure.com:444', 'https://contoso.openai.azure.com/openai/v1',
  'https://user:password@contoso.openai.azure.com', 'https://@contoso.openai.azure.com',
  'https://contoso.openai.azure.com?', 'https://contoso.openai.azure.com#',
  'https://contoso.openai.azure.com?apiKey=value', 'https://contoso.openai.azure.com\\@evil.example',
]) {
  test(`rejects untrusted or non-origin Azure endpoint: ${endpoint}`, async () => {
    const { client } = await loadBYOClient();
    assert.equal(client.validateBYOAIProfile({ ...profile(), endpoint }).valid, false);
    assert.throws(() => client.upsertBYOAIProfile({ ...profile(), endpoint }), { code: 'byo_invalid_profile' });
  });
}

test('official OpenAI is fixed; valid Azure cloud resource origins canonicalize without provider guessing', async () => {
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile({ ...profile(), endpoint: 'HTTPS://Contoso.OPENAI.AZURE.COM/' });
  assert.equal(saved.endpoint, 'https://contoso.openai.azure.com');
  for (const endpoint of [
    'https://contoso.openai.azure.us', 'https://contoso.openai.azure.cn',
    'https://contoso.cognitiveservices.azure.com', 'https://contoso.services.ai.azure.com',
  ]) assert.equal(client.validateBYOAIProfile({ ...profile(), endpoint }).valid, true);
  assert.equal(client.validateBYOAIProfile({ ...profile(), provider: 'openai', endpoint: 'https://api.openai.com/' }).valid, true);
  for (const endpoint of ['', 'https://custom.openai.azure.com', 'https://example.org', 'https://api.openai.com/v1']) {
    assert.equal(client.validateBYOAIProfile({ ...profile(), provider: 'openai', endpoint }).valid, false);
  }
});

for (const edit of [
  { maxCompletionTokens: 0 }, { maxCompletionTokens: -1 }, { maxCompletionTokens: 32769 },
  { maxCompletionTokens: NaN }, { maxCompletionTokens: Infinity }, { maxCompletionTokens: 1.5 },
  { model: 'bad model' }, { model: '../model' }, { model: '' }, { name: '' }, { name: 'x'.repeat(81) },
  { provider: 'anthropic' }, { apiFormat: 'anthropic-messages' }, { reasoningEffort: 'maximum' },
]) {
  test(`invalid public fields fail explicitly rather than silently changing quality: ${JSON.stringify(edit)}`, async () => {
    const { client } = await loadBYOClient();
    assert.equal(client.validateBYOAIProfile({ ...profile(), ...edit }).valid, false);
  });
}

test('keys and verification are isolated per profile; testing does not select and activation requires verification', async t => {
  const sent: any[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    const envelope = JSON.parse(String(options?.body));
    sent.push(envelope);
    return testResponse(envelope.apiFormat);
  });
  const { client } = await loadBYOClient();
  const a = client.upsertBYOAIProfile(profile());
  const b = client.upsertBYOAIProfile({ ...profile('profile-b'), provider: 'openai', endpoint: 'https://api.openai.com', model: 'ft:custom:model' });
  assert.throws(() => client.selectBYOAIProfile(a.id));
  client.setBYOAIApiKey(a.id, keyA);
  client.setBYOAIApiKey(b.id, keyB);
  assert.throws(() => client.selectBYOAIProfile(a.id));
  const result = await client.testBYOAIConnection(a.id);
  assert.deepEqual(Object.keys(result).sort(), ['profileId', 'revision', 'verified']);
  assert.equal(client.getBYOAISettings().activeProfileId, null);
  assert.equal(client.getBYOAIConnectionState(a.id).verified, true);
  assert.equal(client.getBYOAIConnectionState(b.id).verified, false);
  client.selectBYOAIProfile(a.id);
  assert.throws(() => client.selectBYOAIProfile(b.id), { code: 'byo_unverified' });
  assert.equal(client.getBYOAISettings().activeProfileId, a.id);
  await client.testBYOAIConnection(b.id);
  assert.deepEqual(sent.map(envelope => envelope.byo.apiKey), [keyA, keyB]);
  client.setBYOAIApiKey(b.id, '');
  assert.equal(client.getBYOAIConnectionState(a.id).verified, true);
  assert.equal(client.getBYOAIConnectionState(b.id).status, 'key-required');
  assert.equal(client.getEffectiveAIModelInfo('validation').ready, true);
  assert.doesNotMatch(JSON.stringify(client.getEffectiveAIModelInfo('validation')), /sk-offline|apiKey/);
});

for (const edit of [
  { model: 'different-deployment' }, { endpoint: 'https://another.openai.azure.com' },
  { provider: 'openai' as const, endpoint: 'https://api.openai.com' }, { apiFormat: 'chat-completions' as const },
  { reasoningEffort: 'none' as const }, { isReasoning: false }, { supportsVision: false }, { maxCompletionTokens: 32000 },
]) {
  test(`request-affecting edits invalidate verification and queued snapshots: ${JSON.stringify(edit)}`, async t => {
    t.mock.method(globalThis, 'fetch', successfulFetch);
    const { client } = await loadBYOClient();
    const saved = await verifiedProfile(client, profile(), keyA);
    client.selectBYOAIProfile(saved.id);
    const captured = client.captureRuntimeModelOverride('architectureGeneration');
    client.upsertBYOAIProfile({ ...saved, ...edit });
    assert.equal(client.getBYOAIConnectionState(saved.id).status, 'unverified');
    assert.equal(client.getBYOAISettings().activeProfileId, saved.id);
    assert.throws(() => client.resolveAIModelRuntime('architectureGeneration', captured), { code: 'stale_ai_configuration' });
    assert.throws(() => client.resolveAIModelRuntime('architectureGeneration'), { code: 'byo_unverified' });
  });
}

test('friendly rename and unchanged trimmed key preserve verification and captured request identity', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const { client } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), keyA);
  client.selectBYOAIProfile(saved.id);
  const captured = client.captureRuntimeModelOverride('architectureGeneration');
  const state = client.getBYOAIConnectionState(saved.id);
  client.upsertBYOAIProfile({ ...saved, name: 'A friendlier label' });
  client.setBYOAIApiKey(saved.id, `  ${keyA}  `);
  assert.deepEqual(client.getBYOAIConnectionState(saved.id), state);
  assert.match(client.getEffectiveAIModelInfo('architectureGeneration').displayName, /A friendlier label/);
  assert.equal(client.resolveAIModelRuntime('architectureGeneration', captured).displayName, captured.connection!.displayName);
});

test('draft edit invalidation blocks activation and queued work without saving drafts or clearing unchanged keys', async t => {
  const keys: string[] = [];
  t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    const envelope = JSON.parse(String(options?.body));
    keys.push(envelope.byo.apiKey);
    return testResponse(envelope.apiFormat);
  });
  const { client, entries } = await loadBYOClient();
  const a = await verifiedProfile(client, profile(), keyA);
  const b = await verifiedProfile(client, profile('profile-b'), keyB);
  client.selectBYOAIProfile(a.id);
  const captured = client.captureRuntimeModelOverride('architectureGeneration');
  const publicSettings = client.getBYOAISettings();
  const persisted = entries.get(BYO_STORAGE_KEY);
  const otherState = client.getBYOAIConnectionState(b.id);

  client.invalidateBYOAIProfile(a.id);

  assert.equal(client.getBYOAIConnectionState(a.id).status, 'unverified');
  assert.equal(client.getBYOAIConnectionState(a.id).hasApiKey, true);
  assert.deepEqual(client.getBYOAIConnectionState(b.id), otherState);
  assert.deepEqual(client.getBYOAISettings(), publicSettings);
  assert.equal(entries.get(BYO_STORAGE_KEY), persisted);
  assert.throws(() => client.selectBYOAIProfile(a.id), { code: 'byo_unverified' });
  assert.throws(() => client.resolveAIModelRuntime('architectureGeneration', captured), { code: 'stale_ai_configuration' });
  assert.throws(() => client.setBYOAIApiKey(a.id, 'sk-'), { code: 'byo_invalid_api_key' });
  assert.equal(client.getBYOAIConnectionState(a.id).verified, false);
  assert.equal(client.getBYOAIConnectionState(a.id).hasApiKey, true);
  await client.testBYOAIConnection(a.id);
  assert.deepEqual(keys, [keyA, keyB, keyA], 'retesting reuses the unchanged key; no public getter is necessary');
  assert.throws(() => client.invalidateBYOAIProfile('missing-profile'), { code: 'byo_profile_missing' });
});

test('deleting the selected profile leaves an actionable missing-profile state until an explicit switch', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const { client, entries } = await loadBYOClient();
  const saved = await verifiedProfile(client, profile(), keyA);
  client.selectBYOAIProfile(saved.id);
  client.removeBYOAIProfile(saved.id);
  assert.equal(client.getBYOAISettings().activeProfileId, saved.id);
  assert.equal(JSON.parse(entries.get(BYO_STORAGE_KEY)!).activeProfileId, saved.id);
  assert.equal(client.getEffectiveAIModelInfo('validation').code, 'byo_profile_missing');
  assert.throws(() => client.resolveAIModelRuntime('architectureGeneration'), { code: 'byo_profile_missing' });
  client.selectBYOAIProfile(null);
  assert.equal(client.resolveAIModelRuntime('architectureGeneration').source, 'managed');
});

for (const change of ['key', 'profile', 'draft', 'delete', 'reload']) {
  test(`late connection-test success cannot verify after a ${change} change`, async t => {
    const pending = deferred<Response>();
    let signal: AbortSignal | undefined;
    t.mock.method(globalThis, 'fetch', async (url: unknown, options?: RequestInit) => {
      if (url === '/api/runtime-config') return capabilityResponse();
      signal = options?.signal ?? undefined;
      return pending.promise;
    });
    const { client } = await loadBYOClient();
    const saved = client.upsertBYOAIProfile(profile());
    client.setBYOAIApiKey(saved.id, keyA);
    const request = client.testBYOAIConnection(saved.id);
    const rejection = assert.rejects(request, { code: 'stale_ai_configuration' });
    await settle();
    assert.equal(client.getBYOAIConnectionState(saved.id).status, 'testing');
    if (change === 'key') client.setBYOAIApiKey(saved.id, keyB);
    if (change === 'profile') client.upsertBYOAIProfile({ ...saved, model: 'updated-deployment' });
    if (change === 'draft') client.invalidateBYOAIProfile(saved.id);
    if (change === 'delete') client.removeBYOAIProfile(saved.id);
    if (change === 'reload') client.reloadBYOAISettings();
    await rejection;
    assert.equal(signal?.aborted, true);
    pending.resolve(testResponse());
    await settle();
    assert.equal(client.getBYOAIConnectionState(saved.id).verified, false);
  });
}

test('overlapping tests revoke the older result without clearing the newer verification', async t => {
  const first = deferred<Response>();
  const second = deferred<Response>();
  let calls = 0;
  t.mock.method(globalThis, 'fetch', async (url: unknown) => {
    if (url === '/api/runtime-config') return capabilityResponse();
    return ++calls === 1 ? first.promise : second.promise;
  });
  const { client } = await loadBYOClient();
  const saved = client.upsertBYOAIProfile(profile());
  client.setBYOAIApiKey(saved.id, keyA);
  const older = client.testBYOAIConnection(saved.id);
  const oldRejected = assert.rejects(older, { code: 'stale_ai_configuration' });
  await settle();
  const newer = client.testBYOAIConnection(saved.id);
  await settle();
  second.resolve(testResponse());
  await newer;
  first.resolve(testResponse());
  await oldRejected;
  assert.equal(client.getBYOAIConnectionState(saved.id).verified, true);
});

for (const payload of [
  {}, { status: 'completed', output_text: '' }, { status: 'completed', output_text: '{}' },
  { status: 'incomplete', output_text: '{"status":"ok"}' },
  { status: 'completed', output_text: '{"status":"ok"}', incomplete_details: { reason: 'max_output_tokens' } },
  { status: 'completed', output_text: '{"status":"ok"}', output: [{ type: 'message', status: 'incomplete' }] },
  { status: 'completed', output_text: '{"status":"ok"}', output: [{ type: 'message', content: [{ type: 'refusal' }] }] },
  { status: 'completed', output_text: '{"status":"ok"}', error: { message: keyA } },
  { choices: [{ finish_reason: 'length', message: { content: '{"status":"ok"}' } }] },
]) {
  test(`HTTP 200 is not verification without a meaningful complete response: ${JSON.stringify(payload)}`, async t => {
    t.mock.method(globalThis, 'fetch', async (url: unknown) => url === '/api/runtime-config' ? capabilityResponse() : jsonResponse(payload));
    const { client } = await loadBYOClient();
    const saved = client.upsertBYOAIProfile(profile());
    client.setBYOAIApiKey(saved.id, keyA);
    await assert.rejects(client.testBYOAIConnection(saved.id), { code: 'byo_test_incomplete' });
    assert.equal(client.getBYOAIConnectionState(saved.id).verified, false);
    assert.equal(client.getBYOAIConnectionState(saved.id).status, 'failed');
    assert.doesNotMatch(JSON.stringify(client.getBYOAIConnectionState(saved.id)), /sk-offline/);
  });
}

for (const operation of ['select-byo', 'select-managed', 'create', 'rename', 'edit-request', 'remove'] as const) {
  test(`storage transaction: ${operation} rolls back completely when required persistence fails`, async t => {
    t.mock.method(globalThis, 'fetch', successfulFetch);
    const { client, entries } = await loadBYOClient();
    const saved = await verifiedProfile(client, profile(), keyA);
    if (operation !== 'select-byo') client.selectBYOAIProfile(saved.id);
    const before = client.getBYOAISettings();
    const beforeSession = client.getBYOAIConnectionState(saved.id);
    const persisted = entries.get(BYO_STORAGE_KEY);
    const captured = client.captureRuntimeModelOverride('architectureGeneration');
    const failedWrite = t.mock.method(entries, 'set', () => {
      throw new DOMException('Sensitive storage details: sk-storage-red-secret', 'QuotaExceededError');
    });
    const mutate = () => {
      if (operation === 'select-byo') client.selectBYOAIProfile(saved.id);
      if (operation === 'select-managed') client.selectBYOAIProfile(null);
      if (operation === 'create') client.upsertBYOAIProfile(profile('another-profile'));
      if (operation === 'rename') client.upsertBYOAIProfile({ ...saved, name: 'New label' });
      if (operation === 'edit-request') client.upsertBYOAIProfile({ ...saved, model: 'new-model' });
      if (operation === 'remove') client.removeBYOAIProfile(saved.id);
    };
    assert.throws(mutate, (error: any) => {
      assert.equal(error.name, 'BYOAIStorageError');
      assert.equal(error.code, 'byo_settings_save_failed');
      assert.equal(error.source, 'client');
      assert.doesNotMatch(`${error.message} ${JSON.stringify(error)}`, /Sensitive storage|sk-storage|QuotaExceeded/);
      return true;
    });
    assert.deepEqual(client.getBYOAISettings(), before);
    assert.deepEqual(client.getBYOAIConnectionState(saved.id), beforeSession);
    assert.equal(entries.get(BYO_STORAGE_KEY), persisted);
    assert.equal(client.resolveAIModelRuntime('architectureGeneration', captured).connection, captured.connection,
      'failed saves must not revoke or replace the previously selected capture');
    failedWrite.mock.restore();
    client.reloadBYOAISettings();
    assert.equal(client.getBYOAISettings().activeProfileId, before.activeProfileId);
  });
}

test('storage transaction: failed migration rewrite remains blocked after key entry and a successful connection test', async t => {
  t.mock.method(globalThis, 'fetch', successfulFetch);
  const entries = new Map([[BYO_STORAGE_KEY, JSON.stringify({ version: 2, enabled: true, ...profile(), apiKey: keyA })]]);
  const original = entries.get(BYO_STORAGE_KEY);
  const failedWrite = t.mock.method(entries, 'set', () => {
    throw new DOMException('Sensitive migration storage details', 'QuotaExceededError');
  });
  const { client } = await loadBYOClient(undefined, undefined, entries);
  const saved = client.getBYOAISettings().profiles[0];
  assert.ok(saved, 'safe public migration data should remain available for explicit recovery');
  client.setBYOAIApiKey(saved.id, keyA);
  await client.testBYOAIConnection(saved.id);
  assert.equal(client.getEffectiveAIModelInfo('architectureGeneration').ready, false);
  assert.equal(client.getEffectiveAIModelInfo('architectureGeneration').code, 'byo_settings_save_failed');
  assert.throws(() => client.selectBYOAIProfile(saved.id), { name: 'BYOAIStorageError', code: 'byo_settings_save_failed' });
  assert.equal(entries.get(BYO_STORAGE_KEY), original);
  failedWrite.mock.restore();
  client.selectBYOAIProfile(saved.id);
  assert.equal(client.getEffectiveAIModelInfo('architectureGeneration').ready, true);
  assert.equal(JSON.parse(entries.get(BYO_STORAGE_KEY)!).version, 3);
  assert.doesNotMatch(entries.get(BYO_STORAGE_KEY)!, /apiKey|sk-offline/);
});

test('storage transaction: unreadable preferences expose a sanitized load failure and block implicit managed routing', async t => {
  const entries = new Map<string, string>();
  const get = entries.get.bind(entries);
  t.mock.method(entries, 'get', (name: string) => {
    if (name === BYO_STORAGE_KEY) throw new Error('Sensitive read details: sk-storage-red-secret');
    return get(name);
  });
  const { client } = await loadBYOClient(undefined, undefined, entries);
  const effective = client.getEffectiveAIModelInfo('architectureGeneration');
  assert.equal(effective.ready, false);
  assert.equal(effective.code, 'byo_settings_read_failed');
  assert.throws(() => client.captureRuntimeModelOverride('architectureGeneration'), { code: 'byo_settings_read_failed' });
  assert.doesNotMatch(JSON.stringify(client.getBYOAISettings()), /Sensitive read|sk-storage/);
});

for (const version of [2, 3]) {
  for (const endpoint of [
    'https://username:sk-url-red-secret@contoso.openai.azure.com',
    'https://contoso.openai.azure.com?api-key=sk-url-red-secret',
    'https://contoso.openai.azure.com#sk-url-red-secret',
  ]) {
    test(`unsafe migration endpoint v${version} is blanked in snapshots and rewritten storage`, async t => {
      const logs: string[] = [];
      t.mock.method(console, 'warn', (...args: unknown[]) => { logs.push(args.map(String).join(' ')); });
      const imported = { ...profile(), endpoint };
      const { client, entries } = await loadBYOClient(version === 2
        ? { version, enabled: true, ...imported }
        : { version, activeProfileId: imported.id, profiles: [imported] });
      const migrated = client.getBYOAISettings();
      assert.equal(migrated.profiles[0].endpoint, '');
      assert.equal(migrated.activeProfileId, migrated.profiles[0].id);
      assert.equal(client.isAnyAIModelConfigured(), false);
      assert.doesNotMatch(JSON.stringify(migrated), /sk-url-red-secret|username|api-key=/);
      assert.doesNotMatch(entries.get(BYO_STORAGE_KEY)!, /sk-url-red-secret|username|api-key=/);
      assert.doesNotMatch(logs.join('\n'), /sk-url-red-secret|username|api-key=/);
    });
  }
}
