// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const { deploymentConfig, createOriginGuard, assertTrustedAuth } = require('./deployment-security');
const { managedModelsConfiguration, runtimeManagedModelsConfiguration, byoEndpointsEnabled } = require('./astra-policy');
const read = (...parts) => fs.readFileSync(path.join(__dirname, '..', ...parts), 'utf8').replace(/\r\n/g, '\n');
const openAiId = '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/ai/providers/Microsoft.CognitiveServices/accounts/openai';
const publicEnv = {
  APP_DEPLOYMENT_MODE: 'public',
  ACCESS_CONTROL_ENABLED: 'true',
  EASY_AUTH_ENABLED: 'true',
  ACCESS_ADMIN_EMAIL: 'admin@example.com',
  AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: '/access-store',
  PUBLIC_URL: 'https://app.example.com',
  FRONT_DOOR_ID: '11111111-1111-1111-1111-111111111111',
  AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com/',
  AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: 'astra-prod',
  AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'sol-prod',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-prod,sol-prod',
  AZURE_OPENAI_RESOURCE_ID: openAiId,
  AZURE_COSMOS_ENDPOINT: 'https://cosmos.example.com',
};

test('public deployment fails closed for every required control and never uses memory', () => {
  assert.equal(deploymentConfig(publicEnv).mode, 'public');
  for (const key of Object.keys(publicEnv).filter(key => key !== 'APP_DEPLOYMENT_MODE' && key !== 'AZURE_OPENAI_RESOURCE_ID')) {
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

test('production requires both approved models and rejects partial or retired managed settings', () => {
  const config = deploymentConfig({ ...publicEnv, AZURE_TABLES_ENDPOINT: 'https://storage.table.core.windows.net' });
  assert.equal(config.store, 'table');
  assert.equal(config.managedModels.configured, true);
  assert.equal(config.managedModels.astra.model, 'gpt-6.2-astra');
  assert.equal(config.managedModels.sol.model, 'gpt-6.2-sol');
  for (const invalid of [
    { AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: '' },
    { AZURE_OPENAI_DEPLOYMENT_GPT62SOL: '' },
    { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-prod' },
    { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-prod,sol-prod,extra' },
    { AZURE_OPENAI_DEPLOYMENT_GPT62SOL: 'astra-prod', AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'astra-prod,astra-prod' },
    { AZURE_FOUNDRY_DEPLOYMENT_OPUS55: 'retired-opus' },
    { AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'retired-astra' },
  ]) assert.throws(() => deploymentConfig({ ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true', ...invalid }));
  assert.throws(() => deploymentConfig({ ...publicEnv, AZURE_IMPORT_ENABLED: 'TRUE' }), /must be false/);
  assert.throws(() => deploymentConfig({ ...publicEnv, AI_BUDGET_STORE: 'memory' }), /local-only/);
});

test('explicit BYO-only public startup remains supported without managed settings', () => {
  const byoOnly = {
    ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true',
    AZURE_OPENAI_ENDPOINT: '', AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: '',
    AZURE_OPENAI_DEPLOYMENT_GPT62SOL: '', AZURE_OPENAI_ALLOWED_DEPLOYMENTS: '',
    AZURE_OPENAI_RESOURCE_ID: '',
  };
  const config = deploymentConfig(byoOnly);
  assert.equal(config.mode, 'public');
  assert.equal(config.allowByoAIEndpoints, true);
  assert.equal(config.managedModels.configured, false);
  assert.throws(() => deploymentConfig({ ...byoOnly, AZURE_OPENAI_ENDPOINT: publicEnv.AZURE_OPENAI_ENDPOINT }));
  assert.throws(() => deploymentConfig({ ...byoOnly, AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA: 'astra-prod' }));
});

test('runtime exposes only GPT-6.2 Astra and Sol with Sol default and independent BYO policy', () => {
  const models = managedModelsConfiguration(publicEnv);
  assert.deepEqual(models, {
    astra: {
      model: 'gpt-6.2-astra', apiFormat: 'responses', configured: true,
      deployment: 'astra-prod', endpoint: 'https://example.openai.azure.com/',
    },
    sol: {
      model: 'gpt-6.2-sol', apiFormat: 'responses', configured: true,
      deployment: 'sol-prod', endpoint: 'https://example.openai.azure.com/',
    },
    configured: true,
  });
  const runtime = runtimeManagedModelsConfiguration(models, true);
  assert.equal(runtime.features.bringYourOwnAI, true);
  assert.equal(runtime.ai.defaultModel, 'gpt-6.2-sol');
  assert.deepEqual(runtime.ai.models.map(item => item.model), ['gpt-6.2-astra', 'gpt-6.2-sol']);
  for (const flag of [undefined, false, 'true', 'false', 'TRUE', null, 1]) {
    assert.equal(runtimeManagedModelsConfiguration(models, flag).features.bringYourOwnAI, false);
  }
  for (const flag of [undefined, '', 'false']) assert.equal(byoEndpointsEnabled({ ALLOW_BYO_AI_ENDPOINTS: flag }), false);
  for (const flag of ['TRUE', '1', 'yes', ' true ', true]) {
    assert.throws(() => deploymentConfig({ ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: flag }), /ALLOW_BYO_AI_ENDPOINTS/);
  }
});

test('ARM verifier checks exact OpenAI account identity and both GPT-6.2 deployment versions', async () => {
  const { verifyManagedModels, APPROVED_GPT62_VERSION } = await import('../scripts/verify-managed-models.mjs');
  const env = { ...publicEnv };
  const account = { id: openAiId, kind: 'OpenAI', properties: { endpoint: env.AZURE_OPENAI_ENDPOINT } };
  const deployments = {
    'astra-prod': {
      id: `${openAiId}/deployments/astra-prod`,
      properties: { provisioningState: 'Succeeded', model: { format: 'OpenAI', name: 'gpt-6.2-astra', version: APPROVED_GPT62_VERSION } },
    },
    'sol-prod': {
      id: `${openAiId}/deployments/sol-prod`,
      properties: { provisioningState: 'Succeeded', model: { format: 'OpenAI', name: 'gpt-6.2-sol', version: APPROVED_GPT62_VERSION } },
    },
  };
  const fixtureFor = url => url.includes('/deployments/')
    ? deployments[url.split('/deployments/')[1].split('?')[0]]
    : account;
  const calls = [];
  verifyManagedModels(env, url => { calls.push(url); return structuredClone(fixtureFor(url)); });
  assert.equal(calls.length, 4);
  assert.ok(calls.every(url => url.includes('api-version=2025-06-01')));
  for (const [id, mutate] of [
    [openAiId, item => { item.kind = 'AIServices'; }],
    [`${openAiId}/deployments/astra-prod`, item => { item.properties.model.version = 'wrong'; }],
    [`${openAiId}/deployments/astra-prod`, item => { item.properties.model.name = 'gpt-6-astra'; }],
    [`${openAiId}/deployments/sol-prod`, item => { item.properties.provisioningState = 'Failed'; }],
    [`${openAiId}/deployments/sol-prod`, item => { item.properties.model.format = 'Anthropic'; }],
  ]) {
    assert.throws(() => verifyManagedModels(env, url => {
      const current = url.includes('/deployments/') ? `${openAiId}/deployments/${url.split('/deployments/')[1].split('?')[0]}` : openAiId;
      const fixture = structuredClone(fixtureFor(url));
      if (current === id) mutate(fixture);
      return fixture;
    }));
  }
  assert.throws(() => verifyManagedModels({ ...env, AZURE_OPENAI_RESOURCE_ID: '' }, () => {
    throw new Error('must not read ARM');
  }), /AZURE_OPENAI_RESOURCE_ID/);
});

test('deployment hash, image build and rollback track both managed model aliases', () => {
  const workflow = read('.github', 'workflows', 'azurediagarm-sync-deploy.yml');
  const dockerfile = read('Dockerfile');
  const script = read('scripts', 'deploy_aca.sh');
  const hashKeys = workflow.match(/keys=\(([\s\S]*?)\n\s*\)/)?.[1];
  for (const key of [
    'CONFIG_SCHEMA', 'AZURE_OPENAI_ENDPOINT', 'AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA',
    'AZURE_OPENAI_DEPLOYMENT_GPT62SOL', 'AZURE_OPENAI_RESOURCE_ID', 'ALLOW_BYO_AI_ENDPOINTS',
  ]) assert.ok(hashKeys?.includes(key), `changing ${key} must invalidate deployment drift hash`);
  for (const name of ['VITE_AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA', 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT62SOL']) {
    assert.ok(dockerfile.includes(`ARG ${name}`));
    assert.ok(script.includes(name));
    assert.ok(workflow.includes(name));
  }
  assert.match(workflow, /node scripts\/verify-managed-models\.mjs/);
  assert.match(workflow, /AZURE_OPENAI_ALLOWED_DEPLOYMENTS=\$AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA,\$AZURE_OPENAI_DEPLOYMENT_GPT62SOL/);
  const rollback = workflow.slice(workflow.indexOf('      - name: Roll back a failed production revision'));
  assert.match(rollback, /az containerapp revision copy/);
  assert.doesNotMatch(rollback, /--set-env-vars|--remove-env-vars|rollback_ai_tuple/);
});

test('retired-environment cleanup removes former Astra, Sol and Foundry aliases', () => {
  const { execFileSync } = require('node:child_process');
  const names = execFileSync(process.execPath, [
    path.join(__dirname, '..', 'scripts', 'retired-ai-environment.mjs'),
  ], { encoding: 'utf8' }).trim().split(/\r?\n/);
  for (const name of [
    'AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA', 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA',
    'AZURE_OPENAI_DEPLOYMENT_GPT61SOL', 'AZURE_FOUNDRY_API_KEY', 'AZURE_FOUNDRY_ENDPOINT',
    'AZURE_OPENAI_DEPLOYMENT_GPT56SOL',
  ]) assert.ok(names.includes(name), name);
  for (const name of [
    'AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA', 'AZURE_OPENAI_DEPLOYMENT_GPT62SOL',
    'AZURE_OPENAI_ALLOWED_DEPLOYMENTS', 'ALLOW_BYO_AI_ENDPOINTS',
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
