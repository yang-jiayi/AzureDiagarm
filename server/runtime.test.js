// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { spawn } = require('node:child_process');
const path = require('node:path');

test('real local server starts without cloud calls and serves budget with no credentials', async (t) => {
  const child = spawn(process.execPath, [path.join(__dirname, 'token-server.js')], {
    // Deliberately do not inherit Azure credentials or any deployment settings.
    env: {
      PATH: process.env.PATH, SystemRoot: process.env.SystemRoot, APP_DEPLOYMENT_MODE: 'local', TOKEN_SERVER_PORT: '0',
      ALLOW_BYO_AI_ENDPOINTS: 'true', AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS: 'retired',
      AZURE_FOUNDRY_ENDPOINT: 'https://retired.services.ai.azure.com/',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  t.after(() => child.kill());
  const port = await new Promise((resolve, reject) => {
    let output = '';
    let errors = '';
    child.stderr.on('data', chunk => { errors += chunk.toString(); });
    const timeout = setTimeout(() => reject(new Error(`Local server did not start. ${errors}`)), 120_000);
    child.stdout.on('data', chunk => {
      output += chunk.toString();
      const match = output.match(/Listening on 127\.0\.0\.1:(\d+)/);
      if (match) { clearTimeout(timeout); resolve(Number(match[1])); }
    });
    child.once('exit', code => { clearTimeout(timeout); reject(new Error(`Local server exited: ${code}. ${errors}`)); });
  });
  const base = `http://127.0.0.1:${port}`;
  const response = await fetch(`${base}/api/ai/budget`);
  assert.equal(response.status, 200);
  assert.equal(response.headers.get('Cache-Control'), 'no-store');
  const budget = await response.json();
  assert.equal(budget.mode, 'local');
  assert.equal(budget.available, true);
  assert.equal(budget.remainingTokens, 250000);
  const health = await fetch(`${base}/healthz`);
  assert.equal(health.status, 200);
  assert.equal(await health.text(), 'ok\n');
  const readiness = await fetch(`${base}/readyz`);
  assert.equal(readiness.status, 200);
  assert.equal(await readiness.text(), 'ready\n');
  const runtime = await (await fetch(`${base}/api/runtime-config`)).json();
  assert.equal(runtime.features.bringYourOwnAI, true);
  assert.deepEqual(runtime.ai, { model: 'gpt-6-astra', apiFormat: 'responses', deployment: null, configured: false });
  const unconfigured = await fetch(`${base}/api/openai`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ apiFormat: 'responses', deployment: 'gpt-6-astra', body: { model: 'gpt-6-astra', input: 'test' } }),
  });
  assert.equal(unconfigured.status, 503);
  assert.equal((await unconfigured.json()).error.code, 'astra_not_configured');
  const policy = await (await fetch(`${base}/api/feedback/policy`)).json();
  assert.equal(policy.archiveEnabled, false);
  assert.equal(policy.contactEnabled, false);
  assert.equal(policy.legacyRetentionEnabled, false);
  assert.equal(policy.retentionDays, 30);
  assert.equal((await fetch(`${base}/api/access/check`)).status, 204);
  const denied = await fetch(`${base}/api/openai`, {
    method: 'POST', headers: { Origin: 'https://attacker.example', 'Content-Type': 'application/json' },
    body: '{}',
  });
  assert.equal(denied.status, 403);
});
