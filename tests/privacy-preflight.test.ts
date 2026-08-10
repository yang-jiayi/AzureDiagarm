import assert from 'node:assert/strict';
import test from 'node:test';
import {
  anonymizeDiagramPayload,
  detectSensitiveData,
} from '../src/utils/privacyPreflight';

test('privacy preflight detects credentials, resource IDs, and internal values', () => {
  const payload = {
    nodes: [{
      id: 'api',
      data: {
        label: 'api.internal',
        description: 'api_key="super~secret-value"',
        resourceId: '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/prod-rg/providers/Microsoft.Web/sites/prod-app',
      },
    }],
    edges: [],
    titleBlockData: { author: 'owner@contoso.com' },
  };

  const kinds = new Set(detectSensitiveData(payload).map(finding => finding.kind));
  const previews = detectSensitiveData(payload).map(finding => finding.preview).join(' ');
  assert.equal(kinds.has('credential'), true);
  assert.equal(kinds.has('resource-id'), true);
  assert.equal(kinds.has('email'), true);
  assert.equal(kinds.has('internal-host'), true);
  assert.equal(previews.includes('super~secret-value'), false);
  assert.equal(previews.includes('owner@contoso.com'), false);
});

test('anonymization removes sensitive values without changing topology', () => {
  const payload = {
    nodes: [{
      id: 'api',
      position: { x: 10, y: 20 },
      data: {
        label: 'api.internal',
        description: 'password=topsecret123',
      },
    }],
    edges: [{ id: 'edge', source: 'api', target: 'api' }],
    architecturePrompt: 'Contact owner@contoso.com from 10.2.3.4',
  };
  const anonymized = anonymizeDiagramPayload(payload);
  const serialized = JSON.stringify(anonymized);

  assert.equal(anonymized.nodes[0].id, 'api');
  assert.equal(anonymized.edges[0].id, 'edge');
  assert.equal(serialized.includes('topsecret123'), false);
  assert.equal(serialized.includes('owner@contoso.com'), false);
  assert.equal(serialized.includes('10.2.3.4'), false);
  assert.deepEqual(detectSensitiveData(anonymized), []);
});

test('generic Azure service labels are not flagged', () => {
  const findings = detectSensitiveData({
    nodes: [{ id: 'kv', data: { label: 'Key Vault', serviceName: 'Key Vault' } }],
    edges: [],
  });
  assert.deepEqual(findings, []);
});

test('anonymization remaps sensitive graph identifiers without breaking references', () => {
  const payload = {
    nodes: [
      {
        id: 'client_secret=node~alpha1234',
        position: { x: 0, y: 0 },
        data: { label: 'API' },
      },
      {
        id: 'password=node!beta5678',
        parentNode: 'client_secret=node~alpha1234',
        position: { x: 40, y: 40 },
        data: { label: 'Worker' },
      },
    ],
    edges: [{
      id: 'api_key=edge~secret9012',
      source: 'client_secret=node~alpha1234',
      target: 'password=node!beta5678',
    }],
  };

  const anonymized = anonymizeDiagramPayload(payload);
  const [parent, child] = anonymized.nodes;
  const [edge] = anonymized.edges;

  assert.notEqual(parent.id, payload.nodes[0].id);
  assert.notEqual(child.id, payload.nodes[1].id);
  assert.notEqual(parent.id, child.id);
  assert.equal(child.parentNode, parent.id);
  assert.equal(edge.source, parent.id);
  assert.equal(edge.target, child.id);
  assert.notEqual(edge.id, payload.edges[0].id);
  assert.deepEqual(detectSensitiveData(anonymized), []);
});

test('detectSensitiveData returns findings deduplicated by id (single-pass Set)', () => {
  const payload = {
    nodes: [
      { id: 'a', data: { note: 'password: hunter2secret', host: 'app.internal' } },
      { id: 'b', data: { contact: 'owner@contoso.com', ip: '10.1.2.3' } },
    ],
    edges: [],
  };

  const findings = detectSensitiveData(payload);
  const ids = findings.map(finding => finding.id);

  // Every id is unique — the Set-based dedup left no duplicates.
  assert.equal(new Set(ids).size, ids.length);
  const kinds = new Set(findings.map(finding => finding.kind));
  assert.equal(kinds.has('credential'), true);
  assert.equal(kinds.has('internal-host'), true);
  assert.equal(kinds.has('email'), true);
  assert.equal(kinds.has('private-address'), true);
});

test('detectSensitiveData scans large payloads without quadratic blowup', () => {
  // Mostly-benign strings force a full recursive walk (the 100-finding cap is
  // never reached), so this exercises the per-string regex path. With the
  // compiled RegExps hoisted to module scope this scales linearly; the previous
  // per-call `new RegExp` recompilation would be dramatically slower.
  const nodes: Array<{ id: string; data: Record<string, string> }> = [];
  for (let index = 0; index < 20000; index += 1) {
    nodes.push({ id: `n${index}`, data: { label: `Service ${index}`, note: `benign note ${index}` } });
  }
  nodes.push({ id: 'secret', data: { note: 'api_key="abcdefgh12345678"' } });

  const start = performance.now();
  const findings = detectSensitiveData({ nodes, edges: [] });
  const elapsedMs = performance.now() - start;

  assert.equal(findings.some(finding => finding.kind === 'credential'), true);
  const ids = findings.map(finding => finding.id);
  assert.equal(new Set(ids).size, ids.length);
  assert.ok(elapsedMs < 1500, `detectSensitiveData took ${elapsedMs.toFixed(0)}ms — expected linear scan`);
});
