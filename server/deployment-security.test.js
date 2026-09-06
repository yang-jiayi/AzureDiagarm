// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { deploymentConfig, createOriginGuard, assertTrustedAuth } = require('./deployment-security');

const publicEnv = {
  APP_DEPLOYMENT_MODE: 'public', ACCESS_CONTROL_ENABLED: 'true', EASY_AUTH_ENABLED: 'true',
  ACCESS_ADMIN_EMAIL: 'admin@example.com', AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: '/access-store',
  PUBLIC_URL: 'https://app.example.com', FRONT_DOOR_ID: '11111111-1111-1111-1111-111111111111',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'approved', AZURE_COSMOS_ENDPOINT: 'https://cosmos.example.com',
};
test('public deployment fails closed for every required control and never uses memory', () => {
  assert.equal(deploymentConfig(publicEnv).mode, 'public');
  for (const key of Object.keys(publicEnv).filter(key => key !== 'APP_DEPLOYMENT_MODE')) {
    assert.throws(() => deploymentConfig({ ...publicEnv, [key]: '' }), /require|safe to start/i, key);
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
test('public configuration preserves Foundry-only deployments and uses existing shared Tables by default', () => {
  const config = deploymentConfig({
    ...publicEnv, AZURE_OPENAI_ALLOWED_DEPLOYMENTS: '', AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'approved-claude',
    AZURE_TABLES_ENDPOINT: 'https://storage.table.core.windows.net',
  });
  assert.equal(config.store, 'table');
  assert.equal(config.legacyRetentionEnabled, false);
  assert.equal(deploymentConfig({ ...publicEnv, FEEDBACK_LEGACY_RETENTION_ENABLED: 'true' }).legacyRetentionEnabled, true);
  assert.throws(() => deploymentConfig({ ...publicEnv, AZURE_IMPORT_ENABLED: 'TRUE' }), /must be false/);
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

test('Astra build and runtime wiring uses its own deployment variable and preserves rollback models', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const read = (...parts) => fs.readFileSync(path.join(__dirname, '..', ...parts), 'utf8').replace(/\r\n/g, '\n');
  const key = 'AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA';
  const reference = '${{ vars.' + key + ' }}';
  const workflow = read('.github', 'workflows', 'azurediagarm-sync-deploy.yml');
  const dockerfile = read('Dockerfile');
  const main = read('infra', 'main.bicep');
  const parameters = JSON.parse(read('infra', 'main.parameters.json')).parameters;
  const script = read('scripts', 'deploy_aca.sh');

  assert.ok(dockerfile.includes(`ARG VITE_${key}\n`));
  assert.ok(dockerfile.includes(`ENV VITE_${key}=$VITE_${key}`));
  assert.ok(workflow.includes(`VITE_${key}=${reference}`));
  assert.ok(workflow.includes(`${key}: ${reference}`));
  const hashKeys = workflow.match(/keys=\(([\s\S]*?)\n\s*\)/)?.[1];
  assert.ok(hashKeys?.includes(key), 'changing Astra configuration must invalidate deployment drift hash');
  const allowlists = workflow.split('\n').filter(line => /AZURE_OPENAI_ALLOWED_DEPLOYMENTS[:=]/.test(line));
  assert.equal(allowlists.length, 3, 'validate, compare, and deploy must share the managed-model allowlist');
  for (const allowlist of allowlists) {
    for (const name of [key, 'AZURE_OPENAI_DEPLOYMENT_GPT56SOL', 'AZURE_OPENAI_DEPLOYMENT_GPT56TERRA', 'AZURE_OPENAI_DEPLOYMENT_GPT56LUNA']) {
      assert.ok(allowlist.includes(name), `${name} missing from ${allowlist.trim()}`);
    }
  }
  assert.equal(parameters.openAiDeploymentGpt6Astra.value, '${' + key + '}');
  assert.match(main, /param openAiDeploymentGpt6Astra string = ''/);
  assert.match(main, /var openAiAllowedDeployments = join\(\[[\s\S]*?openAiDeploymentGpt6Astra[\s\S]*?\], ','\)/);
  assert.ok(main.includes(`output ${key} string = openAiDeploymentGpt6Astra`));
  assert.ok(script.includes(`VITE_${key}`));
  assert.match(script, /OPENAI_DEPLOYMENTS\+=\("\$\{!var\}"\)/);
  assert.match(workflow, /Capture rollback baseline/);
  assert.match(workflow, /az containerapp revision copy/);
});
