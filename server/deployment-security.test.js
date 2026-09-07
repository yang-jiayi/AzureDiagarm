// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const { deploymentConfig, createOriginGuard, assertTrustedAuth } = require('./deployment-security');
const read = (...parts) => fs.readFileSync(path.join(__dirname, '..', ...parts), 'utf8').replace(/\r\n/g, '\n');

const publicEnv = {
  APP_DEPLOYMENT_MODE: 'public', ACCESS_CONTROL_ENABLED: 'true', EASY_AUTH_ENABLED: 'true',
  ACCESS_ADMIN_EMAIL: 'admin@example.com', AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: '/access-store',
  PUBLIC_URL: 'https://app.example.com', FRONT_DOOR_ID: '11111111-1111-1111-1111-111111111111',
  AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com/',
  AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'approved-astra',
  AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'approved-astra', AZURE_COSMOS_ENDPOINT: 'https://cosmos.example.com',
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
test('public Astra configuration uses existing shared Tables by default and rejects legacy-only AI', () => {
  const config = deploymentConfig({
    ...publicEnv,
    AZURE_TABLES_ENDPOINT: 'https://storage.table.core.windows.net',
  });
  test('explicit BYO-only public startup retains every authentication, origin, access-list and shared-budget control', () => {
    const byoOnly = {
      ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true',
      AZURE_OPENAI_ENDPOINT: '', AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: '', AZURE_OPENAI_ALLOWED_DEPLOYMENTS: '',
    };
    const config = deploymentConfig(byoOnly);
    assert.equal(config.astra.configured, false);
    assert.equal(config.allowByoAIEndpoints, true);
    assert.equal(config.mode, 'public');
    assert.equal(config.store, 'cosmos');
    for (const key of ['ACCESS_CONTROL_ENABLED', 'EASY_AUTH_ENABLED', 'ACCESS_ADMIN_EMAIL',
      'AZURE_ACCESS_KEY_VAULT_RESOURCE_ID', 'PUBLIC_URL', 'FRONT_DOOR_ID', 'AZURE_COSMOS_ENDPOINT']) {
      assert.throws(() => deploymentConfig({ ...byoOnly, [key]: '' }), /require|safe to start/i, key);
    }
    assert.throws(() => deploymentConfig({ ...byoOnly, AI_BUDGET_STORE: 'memory' }), /local-only/);
    assert.throws(() => deploymentConfig({ ...byoOnly, ALLOW_BYO_AI_ENDPOINTS: 'false' }), /safe to start/);
    for (const incomplete of [
      { AZURE_OPENAI_ENDPOINT: publicEnv.AZURE_OPENAI_ENDPOINT },
      { AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'alias' },
      { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'legacy' },
      { AZURE_OPENAI_ALLOWED_DEPLOYMENTS: ', ' },
      { AZURE_OPENAI_API_KEY: 'orphan-managed-key' },
      { AZURE_OPENAI_RESOURCE_ID: '/orphan-managed-account' },
    ]) assert.throws(() => deploymentConfig({ ...byoOnly, ...incomplete }), /Astra/);
  });
  assert.equal(config.store, 'table');
  assert.equal(config.legacyRetentionEnabled, false);
  assert.equal(deploymentConfig({ ...publicEnv, FEEDBACK_LEGACY_RETENTION_ENABLED: 'true' }).legacyRetentionEnabled, true);
  assert.throws(() => deploymentConfig({ ...publicEnv, AZURE_IMPORT_ENABLED: 'TRUE' }), /must be false/);
  assert.throws(() => deploymentConfig({
    ...publicEnv, AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: '',
    AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'gpt-5.6-sol', AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'approved-claude',
  }), /Astra/);
  assert.throws(() => deploymentConfig({
    ...publicEnv, AZURE_OPENAI_ALLOWED_DEPLOYMENTS: 'approved-astra,gpt-5.6-sol',
  }), /singleton/);
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

test('Astra is the sole managed model; explicit BYO opt-in participates in runtime, preflight, drift hashes and reference infrastructure', () => {
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
  assert.ok(hashKeys?.includes('ALLOW_BYO_AI_ENDPOINTS'), 'changing BYO policy must invalidate deployment drift hash');
  const allowlists = workflow.split('\n').filter(line => /AZURE_OPENAI_ALLOWED_DEPLOYMENTS[:=]/.test(line)
    && !line.includes('rollback_ai_tuple'));
  assert.equal(allowlists.length, 3, 'validate, compare, and deploy must share the managed-model allowlist');
  for (const allowlist of allowlists) {
    assert.ok(allowlist.includes(key));
    assert.doesNotMatch(allowlist, /,/);
  }
  assert.equal(parameters.openAiDeploymentGpt6Astra.value, '${' + key + '}');
  assert.match(main, /param openAiDeploymentGpt6Astra string = ''/);
  assert.match(main, /azureOpenAiDeploymentGpt6Astra: openAiDeploymentGpt6Astra/);
  assert.ok(main.includes(`output ${key} string = openAiDeploymentGpt6Astra`));
  assert.ok(script.includes(`VITE_${key}`));
  assert.match(script, /verify-astra-deployment\.mjs/);
  for (const active of [workflow, dockerfile, main, read('infra', 'resources.bicep'),
    read('infra', 'main.parameters.json'), script, read('.env.example'),
    read('scripts', 'dev-all.sh'), read('scripts', 'start-token-server.sh')]) {
    assert.doesNotMatch(active, /(?:VITE_)?AZURE_OPENAI_DEPLOYMENT_(?:GPT5|DEEPSEEK|GROK|KIMI|MISTRAL)|azureFoundryEndpoint/);
  }
  assert.match(main, /param allowByoAIEndpoints bool = false/);
  assert.equal(parameters.allowByoAIEndpoints.value, false);
  assert.match(read('infra', 'resources.bicep'), /name: 'ALLOW_BYO_AI_ENDPOINTS', value: string\(allowByoAIEndpoints\)/);
  assert.match(workflow, /ALLOW_BYO_AI_ENDPOINTS: \$\{\{ vars\.ALLOW_BYO_AI_ENDPOINTS \|\| 'false' \}\}/);
  assert.match(script, /export ALLOW_BYO_AI_ENDPOINTS="\$\{ALLOW_BYO_AI_ENDPOINTS:-false\}"/);
  assert.match(workflow, /node scripts\/retired-ai-environment\.mjs/);
  assert.match(workflow, /Verify actual GPT-6 Astra deployment identity/);
  const modelCheck = 'node scripts/verify-astra-deployment.mjs';
  assert.equal(workflow.split(modelCheck).length - 1, 2);
  assert.ok(workflow.indexOf(modelCheck) < workflow.indexOf('- name: Build and push application image'));
  assert.match(workflow, /node scripts\/verify-astra-deployment\.mjs\s+node scripts\/verify-deployment-source\.mjs\s+# Freeze[^\n]+\s+printf '%s\\0' "\$AZURE_OPENAI_ENDPOINT" "\$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA" "\$AZURE_OPENAI_RESOURCE_ID" "\$ALLOW_BYO_AI_ENDPOINTS" \\\s+> azurediagarm-verified-astra\.env\s+touch azurediagarm-deployment-started/);
  assert.match(workflow, /"AZURE_OPENAI_ALLOWED_DEPLOYMENTS=\$\{rollback_ai_tuple\[1\]\}"/);
  assert.match(workflow, /"ALLOW_BYO_AI_ENDPOINTS=\$\{rollback_ai_tuple\[3\]\}"/);
  assert.match(dockerfile, /COPY server\/astra-policy\.js \.\//);
  const identityHelper = read('scripts', 'verify-astra-deployment.mjs');
  assert.match(identityHelper, /'rest', '--method', 'get'/);
  assert.doesNotMatch(identityHelper, /account deployment (?:create|delete)|role assignment create/);
  assert.match(workflow, /Capture rollback baseline/);
  assert.match(workflow, /az containerapp revision copy/);
});

test('runtime BYO capability requires explicit opt-in and never revives a managed Foundry provider', () => {
  const { astraConfiguration, runtimeAstraConfiguration, byoEndpointsEnabled } = require('./astra-policy');
  const stale = {
    ...publicEnv, ALLOW_BYO_AI_ENDPOINTS: 'true',
    AZURE_FOUNDRY_ENDPOINT: 'https://old.services.ai.azure.com/',
    AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'claude-opus-5',
  };
  const configured = deploymentConfig(stale);
  const runtime = runtimeAstraConfiguration(configured.astra, configured.allowByoAIEndpoints);
  assert.deepEqual(runtime, {
    features: { bringYourOwnAI: true },
    ai: { model: 'gpt-6-astra', apiFormat: 'responses', deployment: 'approved-astra', configured: true },
  });
  assert.equal(deploymentConfig(stale).mode, 'public');
  for (const flag of [undefined, false, 'true', 'false', 'TRUE', null, 1]) {
    assert.equal(runtimeAstraConfiguration(configured.astra, flag).features.bringYourOwnAI, false);
  }
  for (const flag of [undefined, '', 'false']) assert.equal(byoEndpointsEnabled({ ALLOW_BYO_AI_ENDPOINTS: flag }), false);
  for (const flag of ['TRUE', '1', 'yes', ' true ', true]) {
    assert.throws(() => deploymentConfig({ ...stale, ALLOW_BYO_AI_ENDPOINTS: flag }), /ALLOW_BYO_AI_ENDPOINTS/);
  }
  assert.throws(() => astraConfiguration({ ...publicEnv, AZURE_OPENAI_ENDPOINT: 'https://api.openai.com' }), /Azure OpenAI HTTPS origin/);
});

test('predeployment verification checks actual model identity and account ownership, never alias spelling', async () => {
  const { verifyAstraDeployment, APPROVED_ASTRA_MODEL_VERSION } = await import('../scripts/verify-astra-deployment.mjs');
  assert.equal(
    read('infra', 'gpt6-astra.bicep').match(/\bversion:\s*'([^']+)'/)?.[1],
    APPROVED_ASTRA_MODEL_VERSION,
  );
  const id = '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/example/providers/Microsoft.CognitiveServices/accounts/example';
  const alias = 'architecture-production';
  const env = {
    AZURE_OPENAI_RESOURCE_ID: id, AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com/',
    AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: alias,
  };
  const account = { id, properties: { endpoint: env.AZURE_OPENAI_ENDPOINT } };
  const deployment = {
    id: `${id}/deployments/${alias}`,
    properties: {
      provisioningState: 'Succeeded',
      model: { format: 'OpenAI', name: 'gpt-6-astra', version: '2026-09-03' },
    },
  };
  const calls = [];
  verifyAstraDeployment(env, url => {
    calls.push(url);
    return structuredClone(url.includes('/deployments/') ? deployment : account);
  });
  assert.equal(calls.length, 2);
  assert.ok(calls.every(url => url.startsWith(`https://management.azure.com${id}`)));
  for (const mutate of [
    value => { value.properties.model.name = 'gpt-5.6-sol'; },
    value => { value.properties.model.name = 'gpt-5.6-astra'; },
    value => { delete value.properties.model.version; },
    value => { value.properties.model.version = '2026-09-04'; },
    value => { value.properties.model.version = ''; },
    value => { value.properties.model.version = null; },
    value => { value.properties.model.version = 20260903; },
    value => { value.properties.model.format = 'Anthropic'; },
    value => { value.properties.provisioningState = 'Failed'; },
    value => { value.id = `${id}/deployments/different`; },
  ]) {
    const altered = structuredClone(deployment);
    mutate(altered);
    assert.throws(() => verifyAstraDeployment(env, url => url.includes('/deployments/') ? altered : account), /gpt-6-astra/);
  }
  assert.throws(() => verifyAstraDeployment(env, () => ({ ...account, properties: { endpoint: 'https://other.openai.azure.com/' } })), /does not belong/);
  assert.throws(() => verifyAstraDeployment({ ...env, AZURE_OPENAI_RESOURCE_ID: '' }, () => { throw new Error('must not read'); }), /full AZURE_OPENAI_RESOURCE_ID/);
  assert.throws(() => verifyAstraDeployment(env, () => { throw new Error('ARM unavailable'); }), /ARM unavailable/);
});

test('retired runtime environment cleanup includes every removed model and provider knob', () => {
  const { execFileSync } = require('node:child_process');
  const path = require('node:path');
  const names = execFileSync(process.execPath, [path.join(__dirname, '..', 'scripts', 'retired-ai-environment.mjs')], { encoding: 'utf8' }).trim().split(/\r?\n/);
  for (const name of [
    'AZURE_FOUNDRY_ENDPOINT', 'AZURE_FOUNDRY_API_KEY',
    'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS', 'AZURE_OPENAI_DEPLOYMENT_GPT56SOL',
    'AZURE_OPENAI_DEPLOYMENT_GPT56TERRA', 'AZURE_OPENAI_DEPLOYMENT_GPT56LUNA',
    'VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL',
  ]) assert.ok(names.includes(name), name);
  assert.ok(!names.includes('AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA'));
  assert.ok(!names.includes('AZURE_OPENAI_ALLOWED_DEPLOYMENTS'));
  assert.ok(!names.includes('ALLOW_BYO_AI_ENDPOINTS'));
  const rollbackNames = execFileSync(process.execPath, [path.join(__dirname, '..', 'scripts', 'retired-ai-environment.mjs'), '--rollback'], { encoding: 'utf8' }).trim().split(/\r?\n/);
  assert.deepEqual(rollbackNames, names);
});
