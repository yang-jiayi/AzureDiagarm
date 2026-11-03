// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const path = require('node:path');
const { test } = require('node:test');

test('token server releases upstream error bodies, keeps statuses, and does not log secrets', async (t) => {
  const serverPath = path.join(__dirname, 'token-server.js');
  const fixture = `
    const { DefaultAzureCredential } = require('@azure/identity');
    DefaultAzureCredential.prototype.getToken = async () => ({ token: 'credential-canary' });
    const calls = Object.create(null);
    global.fetch = async (input) => {
      const requestPath = new URL(input).pathname;
      const count = calls[requestPath] = (calls[requestPath] || 0) + 1;
      let status = requestPath.endsWith('/sts/v1.0/issueToken') ? 401
        : requestPath === '/api/mcp' ? 503
          : requestPath.endsWith('/cognitiveservices/avatar/relay/token/v1') ? 429
            : requestPath === '/subscriptions' ? 403
              : requestPath.endsWith('/resourcegroups') ? 404
                : requestPath.includes('/Microsoft.ResourceGraph/resources') ? 418 : 500;
      let body = 'response-body-canary credential-canary exception-canary';
      if (requestPath.endsWith('/sts/v1.0/issueToken') && count > 1) {
        status = 200;
        body = 'synthetic-speech-token';
      } else if (requestPath === '/subscriptions' && count > 1) {
        status = 200;
        body = JSON.stringify({ value: [{ subscriptionId: '00000000-0000-0000-0000-000000000000', displayName: 'Synthetic' }] });
      }
      const cancellationFails = requestPath.endsWith('/sts/v1.0/issueToken') && count === 1;
      return new Response(new ReadableStream({
        start(controller) {
          controller.enqueue(new TextEncoder().encode(body));
          if (status >= 200 && status < 300) controller.close();
        },
        cancel() {
          process.stdout.write('CANCEL:' + requestPath + '\\n');
          if (cancellationFails) throw new Error('cancel-canary');
        },
      }), { status, headers: { 'Content-Type': 'application/json' } });
    };
    console.error = (...args) => process.stdout.write('LOG:' + args.map(String).join(' ') + '\\n');
    require(${JSON.stringify(serverPath)});
  `;
  const child = spawn(process.execPath, ['-e', fixture], {
    cwd: __dirname,
    env: {
      PATH: process.env.PATH,
      SystemRoot: process.env.SystemRoot,
      APP_DEPLOYMENT_MODE: 'local',
      TOKEN_SERVER_PORT: '0',
      AZURE_SPEECH_REGION: 'westus2',
      AZURE_SPEECH_RESOURCE_ID:
        '/subscriptions/00000000-0000-0000-0000-000000000000/resourceGroups/test/providers/Microsoft.CognitiveServices/accounts/mock-speech',
      AZURE_IMPORT_ENABLED: 'true',
    },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  t.after(() => child.kill());

  let output = '';
  child.stdout.on('data', chunk => { output += chunk.toString(); });
  child.stderr.on('data', chunk => { output += chunk.toString(); });
  const port = await new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error(`Mock server did not start. ${output}`)), 30_000);
    child.stdout.on('data', () => {
      const match = output.match(/Listening on 127\.0\.0\.1:(\d+)/);
      if (match) {
        clearTimeout(timeout);
        resolve(Number(match[1]));
      }
    });
    child.once('error', error => {
      clearTimeout(timeout);
      reject(error);
    });
    child.once('exit', code => {
      clearTimeout(timeout);
      reject(new Error(`Mock server exited: ${code}. ${output}`));
    });
  });
  const base = `http://127.0.0.1:${port}`;
  const request = async (route, init) => {
    const response = await fetch(`${base}${route}`, init);
    return { status: response.status, text: await response.text() };
  };
  const subscriptionId = '00000000-0000-0000-0000-000000000000';

  assert.equal((await request('/api/speech-token')).status, 502);
  const docs = await request('/api/docs-search', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ query: 'synthetic query' }),
  });
  assert.equal(docs.status, 200);
  assert.deepEqual(JSON.parse(docs.text), { results: [], error: 'docs search failed' });
  assert.equal((await request('/api/ice-token')).status, 502);
  assert.equal((await request('/api/azure/subscriptions')).status, 502);
  assert.equal((await request(`/api/azure/resource-groups?subscriptionId=${subscriptionId}`)).status, 502);
  assert.equal((await request('/api/azure/resource-graph', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ subscriptionId, resourceGroup: 'synthetic-rg' }),
  })).status, 502);

  const speechSuccess = await request('/api/speech-token');
  assert.equal(speechSuccess.status, 200);
  assert.equal(JSON.parse(speechSuccess.text).token, 'synthetic-speech-token');
  const subscriptionSuccess = await request('/api/azure/subscriptions');
  assert.equal(subscriptionSuccess.status, 200);
  assert.deepEqual(JSON.parse(subscriptionSuccess.text), {
    subscriptions: [{ subscriptionId, displayName: 'Synthetic' }],
  });

  const expectedCancelledPaths = [
    '/sts/v1.0/issueToken',
    '/api/mcp',
    '/cognitiveservices/avatar/relay/token/v1',
    '/subscriptions',
    `/subscriptions/${subscriptionId}/resourcegroups`,
    '/providers/Microsoft.ResourceGraph/resources',
  ];
  for (const cancelledPath of expectedCancelledPaths) {
    assert.ok(output.includes(`CANCEL:${cancelledPath}`), `${cancelledPath} error body was not cancelled`);
  }
  assert.match(output, /upstream_status=401/);
  assert.match(output, /upstream_status=503/);
  assert.match(output, /upstream request failed with status 429/);
  assert.match(output, /subscriptions request failed with status 403/);
  assert.match(output, /resource-groups request failed with status 404/);
  assert.match(output, /resource-graph request failed with status 418/);
  assert.doesNotMatch(output, /response-body-canary|credential-canary|exception-canary|cancel-canary|synthetic-speech-token/);
});
