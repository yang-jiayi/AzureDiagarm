// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const {
  deploymentConfig, createOriginGuard, assertTrustedAuth,
} = require('./deployment-security');
const {
  managedModelsConfiguration, runtimeManagedModelsConfiguration, byoEndpointsEnabled,
} = require('./astra-policy');
const read = (...parts) => fs.readFileSync(path.join(__dirname, '..', ...parts), 'utf8').replace(/\r\n/g, '\n');

const openAiId = '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/ai/providers/Microsoft.CognitiveServices/accounts/openai';
const foundryId = '/subscriptions/22222222-2222-2222-2222-222222222222/resourceGroups/ai/providers/Microsoft.CognitiveServices/accounts/foundry';
const publicEnv = {
  APP_DEPLOYMENT_MODE: 'public',
  ACCESS_CONTROL_ENABLED: 'true',
  EASY_AUTH_ENABLED: 'true',
  ACCESS_ADMIN_EMAIL: 'admin@example.com',
  AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: '/access-store',
  PUBLIC_URL: 'https://app.example.com',
  FRONT_DOOR_ID: '11111111-1111-1111-1111-111111111111',
  AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com/',
  AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'sol-prod',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'sol-prod',
  AZURE_OPENAI_RESOURCE_ID: openAiId,
  AZURE_FOUNDRY_ENDPOINT: 'https://example.services.ai.azure.com/',
  AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'opus-prod',
  AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'opus-prod',
  AZURE_FOUNDRY_RESOURCE_ID: foundryId,
  AZURE_COSMOS_ENDPOINT: 'https://cosmos.example.com',
};

test('public deployment fails closed for every required control and never uses memory', () => {
  assert.equal(deploymentConfig(publicEnv).mode, 'public');
  for (const key of Object.keys(publicEnv).filter(key => (
    key !== 'APP_DEPLOYMENT_MODE'
    && key !== 'AZURE_OPENAI_RESOURCE_ID'
    && key !== 'AZURE_FOUNDRY_RESOURCE_ID'
  ))) {
    assert.throws(() => deploymentConfig({ ...publicEnv, [key]: '' }), /require|safe to start|managed|model/i, key);
  }
  assert.throws(() => deploymentConfig({ ...publicEnv, AI_BUDGET_STORE: 'memory' }), /local-only/);
  assert.throws(() => deploymentConfig({ ...publicEnv, PUBLIC_URL: 'https://user:pass@app.example.com/?secret=1#token' }));
  assert.throws(() => deploymentConfig({ ...publicEnv, AI_DAILY_TOKEN_BUDGET: '-1' }));
  assert.throws(() => deploymentConfig({ ...publicEnv, AZURE_IMPORT_ENABLED: 'true' }));
});

test('local development works without cloud; production requires explicit mode', () => {
  assert.equal(deploymentConfig({}).store, 'memory');
  assert.equal(deploymentConfig({ APP_DEPLOYMENT_MODE: 'local', NODE_ENV: 'production' }).mode, 'local');
  assert.throws(() => deploymentConfig({ NODE_ENV: 'production' }), /no implicit local fallback/);
});

test('production requires both approved managed models; BYO cannot replace either', () => {
  const config = deploymentConfig({
    ...publicEnv,
    AZURE_TABLES_ENDPOINT: 'https://storage.table.core.windows.net',
  });
  assert.equal(config.store, 'table');
  assert.equal(config.managedModels.configured, true);
  assert.equal(config.managedModels.sol.model, 'gpt-6.1-sol');
  assert.equal(config.managedModels.opus.model, 'claude-opus-5-5');
  for (const invalid of [
    { AZURE_OPENAI_ENDPOINT: '', AZURE_OPENAI_DEPLOYMENT_GPT61SOL: '', AZURE_OPENAI_ALLOWED_DEPLOYMENTS: '' },
    { AZURE_FOUNDRY_ENDPOINT: '', AZURE_FOUNDRY_DEPLOYMENT_OPUS55: '', AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: '' },
    { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'sol-prod,other' },
    { AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'other' },
    { AZURE_FOUNDRY_API_KEY: 'orphan-managed-key' },
  ]) {
    assert.throws(() => deploymentConfig({
      ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true', ...invalid,
    }), /managed|model|singleton|GPT-6\.1|Claude/i);
  }
  assert.equal(config.store, 'table');
  assert.throws(() => deploymentConfig({ ...publicEnv, AZURE_IMPORT_ENABLED: 'TRUE' }), /must be false/);
  assert.throws(() => deploymentConfig({ ...publicEnv, AI_BUDGET_STORE: 'memory' }), /local-only/);
});

test('public BYO-only mode is allowed only when both managed tuples are entirely absent', () => {
  const byoOnly = { ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true' };
  for (const name of [
    'AZURE_OPENAI_ENDPOINT',
    'AZURE_OPENAI_DEPLOYMENT_GPT61SOL',
    'AZURE_OPENAI_ALLOWED_DEPLOYMENTS',
    'AZURE_OPENAI_RESOURCE_ID',
    'AZURE_OPENAI_API_KEY',
    'AZURE_FOUNDRY_ENDPOINT',
    'AZURE_FOUNDRY_DEPLOYMENT_OPUS55',
    'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS',
    'AZURE_FOUNDRY_RESOURCE_ID',
    'AZURE_FOUNDRY_API_KEY',
  ]) delete byoOnly[name];
  const config = deploymentConfig(byoOnly);
  assert.equal(config.mode, 'public');
  assert.equal(config.allowByoAIEndpoints, true);
  assert.equal(config.managedModels.configured, false);
  for (const partial of [
    { AZURE_OPENAI_ENDPOINT: publicEnv.AZURE_OPENAI_ENDPOINT },
    { AZURE_OPENAI_DEPLOYMENT_GPT61SOL: publicEnv.AZURE_OPENAI_DEPLOYMENT_GPT61SOL },
    { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: publicEnv.AZURE_OPENAI_ALLOWED_DEPLOYMENTS },
    { AZURE_FOUNDRY_ENDPOINT: publicEnv.AZURE_FOUNDRY_ENDPOINT },
    { AZURE_FOUNDRY_DEPLOYMENT_OPUS55: publicEnv.AZURE_FOUNDRY_DEPLOYMENT_OPUS55 },
    { AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: publicEnv.AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS },
    { AZURE_FOUNDRY_API_KEY: 'forbidden-managed-foundry-key' },
  ]) {
    assert.throws(() => deploymentConfig({ ...byoOnly, ...partial }), /managed|singleton|Claude/i);
  }
});

test('runtime exposes only the two managed models and BYO remains an independent explicit feature', () => {
  const models = managedModelsConfiguration(publicEnv);
  assert.deepEqual(models, {
    sol: {
      model: 'gpt-6.1-sol',
      apiFormat: 'responses',
      configured: true,
      deployment: 'sol-prod',
      endpoint: 'https://example.openai.azure.com/',
    },
    opus: {
      model: 'claude-opus-5-5',
      apiFormat: 'anthropic-messages',
      configured: true,
      deployment: 'opus-prod',
      endpoint: 'https://example.services.ai.azure.com/',
    },
    configured: true,
  });
  const runtime = runtimeManagedModelsConfiguration(models, true);
  assert.equal(runtime.features.bringYourOwnAI, true);
  assert.equal(runtime.ai.defaultModel, 'gpt-6.1-sol');
  assert.equal(runtime.ai.configured, true);
  assert.equal(runtime.ai.models.length, 2);
  assert.match(JSON.stringify(runtime.ai.models), /gpt-6\.1-sol/);
  assert.match(JSON.stringify(runtime.ai.models), /claude-opus-5-5/);
  for (const flag of [undefined, false, 'true', 'false', 'TRUE', null, 1]) {
    assert.equal(runtimeManagedModelsConfiguration(models, flag).features.bringYourOwnAI, false);
  }
  for (const flag of [undefined, '', 'false']) {
    assert.equal(byoEndpointsEnabled({ ALLOW_BYO_AI_ENDPOINTS: flag }), false);
  }
  for (const flag of ['TRUE', '1', 'yes', ' true ', true]) {
    assert.throws(() => deploymentConfig({ ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: flag }), /ALLOW_BYO_AI_ENDPOINTS/);
  }
});

test('managed deployment identity verifier validates both ARM accounts and model versions', async () => {
  const {
    verifyManagedModels, APPROVED_GPT61_SOL_VERSION, APPROVED_CLAUDE_OPUS55_VERSION,
  } = await import('../scripts/verify-managed-models.mjs');
  const env = {
    ...publicEnv,
    AZURE_OPENAI_DEPLOYMENT_GPT61SOL: 'openai-production-alias',
    AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'openai-production-alias',
    AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'foundry-production-alias',
    AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'foundry-production-alias',
  };
  const openAiAccount = {
    id: openAiId, kind: 'OpenAI', properties: { endpoint: env.AZURE_OPENAI_ENDPOINT },
  };
  const foundryAccount = {
    id: foundryId, kind: 'AIServices', properties: { endpoint: env.AZURE_FOUNDRY_ENDPOINT },
  };
  const openAiDeployment = {
    id: `${openAiId}/deployments/openai-production-alias`,
    properties: {
      provisioningState: 'Succeeded',
      model: { format: 'OpenAI', name: 'gpt-6.1-sol', version: APPROVED_GPT61_SOL_VERSION },
    },
  };
  const foundryDeployment = {
    id: `${foundryId}/deployments/foundry-production-alias`,
    properties: {
      provisioningState: 'Succeeded',
      model: { format: 'Anthropic', name: 'claude-opus-5-5', version: APPROVED_CLAUDE_OPUS55_VERSION },
    },
  };
  const fixtureFor = url => {
    if (url.includes(openAiId) && url.includes('/deployments/')) return openAiDeployment;
    if (url.includes(foundryId) && url.includes('/deployments/')) return foundryDeployment;
    if (url.includes(openAiId)) return openAiAccount;
    if (url.includes(foundryId)) return foundryAccount;
    throw new Error('Unexpected ARM URL');
  };
  const calls = [];
  verifyManagedModels(env, url => {
    calls.push(url);
    return structuredClone(fixtureFor(url));
  });
  assert.equal(calls.length, 4);
  assert.ok(calls.every(url => url.includes('api-version=2025-06-01')));
  assert.ok(calls.some(url => url.endsWith('/deployments/openai-production-alias?api-version=2025-06-01')));
  assert.ok(calls.some(url => url.endsWith('/deployments/foundry-production-alias?api-version=2025-06-01')));

  const invalidCases = [
    [openAiAccount, value => { value.kind = 'AIServices'; }],
    [openAiDeployment, value => { value.properties.model.name = 'gpt-6-astra'; }],
    [openAiDeployment, value => { value.properties.model.version = '2026-09-28'; }],
    [openAiDeployment, value => { value.properties.model.format = 'Anthropic'; }],
    [openAiDeployment, value => { value.properties.provisioningState = 'Failed'; }],
    [foundryAccount, value => { value.kind = 'OpenAI'; }],
    [foundryDeployment, value => { value.properties.model.name = 'claude-opus-5'; }],
    [foundryDeployment, value => { value.properties.model.version = '1'; }],
    [foundryDeployment, value => { value.properties.model.format = 'OpenAI'; }],
    [foundryDeployment, value => { value.id = `${foundryId}/deployments/other`; }],
  ];
  for (const [original, mutate] of invalidCases) {
    const changed = structuredClone(original);
    mutate(changed);
    assert.throws(() => verifyManagedModels(env, url => (
      url.includes('/deployments/')
        ? url.includes(openAiId)
          ? original === openAiDeployment ? changed : openAiDeployment
          : original === foundryDeployment ? changed : foundryDeployment
        : url.includes(openAiId)
          ? original === openAiAccount ? changed : openAiAccount
          : original === foundryAccount ? changed : foundryAccount
    )), /GPT-6\.1 Sol|Claude Opus 5\.5/);
  }
  assert.throws(() => verifyManagedModels(env, url => {
    if (url.includes(openAiId) && !url.includes('/deployments/')) {
      return { ...openAiAccount, properties: { endpoint: 'https://other.openai.azure.com/' } };
    }
    return fixtureFor(url);
  }), /endpoint does not belong/);
  assert.throws(() => verifyManagedModels(env, url => {
    if (url.includes(foundryId) && !url.includes('/deployments/')) {
      return { ...foundryAccount, properties: { endpoint: 'https://other.services.ai.azure.com/' } };
    }
    return fixtureFor(url);
  }), /endpoint does not belong/);
  assert.throws(() => verifyManagedModels({ ...env, AZURE_FOUNDRY_RESOURCE_ID: '' }, () => {
    throw new Error('must not read ARM');
  }), /AZURE_FOUNDRY_RESOURCE_ID/);
  assert.throws(() => verifyManagedModels(env, () => { throw new Error('ARM unavailable'); }), /ARM unavailable/);
});

test('deployment hash includes both models and rollback restores one coherent prior revision', () => {
  const workflow = read('.github', 'workflows', 'azurediagarm-sync-deploy.yml');
  const dockerfile = read('Dockerfile');
  const script = read('scripts', 'deploy_aca.sh');
  const hashKeys = workflow.match(/keys=\(([\s\S]*?)\n\s*\)/)?.[1];
  for (const key of [
    'CONFIG_SCHEMA',
    'AZURE_OPENAI_ENDPOINT', 'AZURE_OPENAI_DEPLOYMENT_GPT61SOL', 'AZURE_OPENAI_RESOURCE_ID',
    'AZURE_FOUNDRY_ENDPOINT', 'AZURE_FOUNDRY_DEPLOYMENT_OPUS55', 'AZURE_FOUNDRY_RESOURCE_ID',
    'ALLOW_BYO_AI_ENDPOINTS',
  ]) assert.ok(hashKeys?.includes(key), `changing ${key} must invalidate deployment drift hash`);
  for (const name of [
    'VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL',
    'VITE_AZURE_FOUNDRY_ENDPOINT',
    'VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55',
  ]) {
    assert.ok(dockerfile.includes(`ARG ${name}`));
    assert.ok(script.includes(name));
    assert.ok(workflow.includes(name));
  }
  assert.match(workflow, /node scripts\/verify-managed-models\.mjs/);
  assert.ok(workflow.includes('"AZURE_OPENAI_ALLOWED_DEPLOYMENTS=${{ vars.AZURE_OPENAI_DEPLOYMENT_GPT61SOL }}'));
  assert.ok(workflow.includes('"AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS=${{ vars.AZURE_FOUNDRY_DEPLOYMENT_OPUS55 }}'));
  const rollback = workflow.slice(workflow.indexOf('      - name: Roll back a failed production revision'));
  assert.match(rollback, /az containerapp revision copy/);
  assert.doesNotMatch(rollback, /--set-env-vars|--remove-env-vars|rollback_ai_tuple/);
  assert.match(rollback, /DESIRED_CONFIG_HASH remain one consistent/);
});

test('retired-environment cleanup removes Astra aliases but leaves shared managed-model settings intact', () => {
  const { execFileSync } = require('node:child_process');
  const names = execFileSync(process.execPath, [
    path.join(__dirname, '..', 'scripts', 'retired-ai-environment.mjs'),
  ], { encoding: 'utf8' }).trim().split(/\r?\n/);
  for (const name of [
    'AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA',
    'VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA',
    'AZURE_FOUNDRY_API_KEY',
    'AZURE_OPENAI_DEPLOYMENT_GPT56SOL',
  ]) assert.ok(names.includes(name), name);
  for (const name of [
    'AZURE_OPENAI_ALLOWED_DEPLOYMENTS',
    'AZURE_FOUNDRY_ENDPOINT',
    'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS',
    'AZURE_FOUNDRY_RESOURCE_ID',
    'ALLOW_BYO_AI_ENDPOINTS',
  ]) assert.ok(!names.includes(name), name);
});

test('mutations require exact public origin and local browser callers must use loopback', () => {
  const check = (mode, origin, method = 'POST') => {
    let code = 0;
    createOriginGuard({ mode, origin: publicEnv.PUBLIC_URL })(
      { method, get: () => origin }, { status(status) { code = status; return this; }, json() {} },
      () => { code = 204; },
    );
    return code;
  };
  assert.equal(check('public', publicEnv.PUBLIC_URL), 204);
  assert.equal(check('public', ''), 403);
  assert.equal(check('public', 'https://evil.example.com'), 403);
  assert.equal(check('public', '', 'GET'), 204);
  assert.equal(check('local', 'http://localhost:5173'), 204);
  assert.equal(check('local', ''), 204);
  assert.equal(check('local', 'https://evil.example.com'), 403);
});

test('deployment auth assertions reject disabled auth, wrong tenants/audiences and excluded APIs', () => {
  const client = '11111111-1111-1111-1111-111111111111';
  const auth = {
    platform: { enabled: true }, httpSettings: { requireHttps: true },
    globalValidation: { unauthenticatedClientAction: 'RedirectToLoginPage', redirectToProvider: 'azureactivedirectory', excludedPaths: ['/healthz'] },
    identityProviders: { azureActiveDirectory: {
      registration: { clientId: client, openIdIssuer: `https://login.microsoftonline.com/${client}/v2.0` },
      validation: { allowedAudiences: [client] },
    } },
  };
  assert.doesNotThrow(() => assertTrustedAuth(auth));
  for (const mutate of [
    value => { value.platform.enabled = false; },
    value => { value.globalValidation.excludedPaths.push('/api/*'); },
    value => { value.identityProviders.azureActiveDirectory.registration.openIdIssuer = 'https://login.microsoftonline.com/common/v2.0'; },
    value => { value.identityProviders.azureActiveDirectory.validation.allowedAudiences = []; },
  ]) {
    const altered = structuredClone(auth);
    mutate(altered);
    assert.throws(() => assertTrustedAuth(altered), /Easy Auth/);
  }
});
