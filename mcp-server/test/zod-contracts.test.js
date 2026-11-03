import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createServer } from 'node:net';
import { once } from 'node:events';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

async function findAvailablePort() {
  const server = createServer();
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const { port } = server.address();
  await new Promise((resolve, reject) => {
    server.close(error => (error ? reject(error) : resolve()));
  });
  return port;
}

async function waitForHealth(url, child) {
  const deadline = Date.now() + 10_000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
      throw new Error(`MCP server exited before becoming ready (code ${child.exitCode})`);
    }
    try {
      const response = await fetch(`${url}/healthz`);
      if (response.ok) return;
    } catch {
      // The server may not have bound its port yet.
    }
    await new Promise(resolve => setTimeout(resolve, 50));
  }
  throw new Error('MCP server did not become ready in time');
}

test('Zod v4 preserves MCP tool input required, optional, default, and cap contracts', {
  timeout: 30_000,
}, async () => {
  const port = await findAvailablePort();
  const url = `http://127.0.0.1:${port}/mcp`;
  const child = spawn(
    process.execPath,
    [fileURLToPath(new URL('../dist/index.js', import.meta.url)), '--http'],
    {
      env: {
        ...process.env,
        MCP_HTTP_HOST: '127.0.0.1',
        MCP_HTTP_PORT: String(port),
        MCP_HTTP_STATELESS: 'true',
      },
      stdio: ['ignore', 'ignore', 'pipe'],
    },
  );
  const transport = new StreamableHTTPClientTransport(new URL(url));
  const client = new Client({ name: 'zod-contract-test', version: '1.0.0' });

  try {
    await waitForHealth(url.replace(/\/mcp$/, ''), child);
    await client.connect(transport);

    const validArchitecture = await client.callTool({
      name: 'validate_architecture',
      arguments: {
        services: [{ name: 'api', type: 'App Service' }],
      },
    });
    assert.notEqual(validArchitecture.isError, true);

    const missingRequiredServices = await client.callTool({
      name: 'validate_architecture',
      arguments: { connections: [] },
    });
    assert.equal(missingRequiredServices.isError, true);

    const overCapArchitecture = await client.callTool({
      name: 'validate_architecture',
      arguments: {
        services: Array.from({ length: 251 }, (_, index) => ({
          name: `Service ${index}`,
          type: 'App Service',
        })),
      },
    });
    assert.equal(overCapArchitecture.isError, true);

    const defaultedEstimate = await client.callTool({
      name: 'estimate_costs',
      arguments: {
        services: [{ name: 'database', type: 'SQL Database' }],
      },
    });
    assert.notEqual(defaultedEstimate.isError, true);
    assert.equal(defaultedEstimate.structuredContent.region, 'eastus2');
    assert.equal(defaultedEstimate.structuredContent.term, 'payg');
    assert.equal(defaultedEstimate.structuredContent.estimates[0].tier, 'standard');
    assert.equal(defaultedEstimate.structuredContent.estimates[0].quantity, 1);
  } finally {
    await transport.close().catch(() => {});
    if (child.exitCode === null) {
      const exited = once(child, 'exit');
      child.kill();
      await exited;
    }
  }
});
