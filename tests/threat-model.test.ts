import assert from 'node:assert/strict';
import test from 'node:test';
import { analyzeThreatModel } from '../src/utils/threatModel';

test('threat model classifies exposure, data, and controls', () => {
  const markers = analyzeThreatModel([
    { id: 'front', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Azure Front Door' } },
    { id: 'sql', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Azure SQL Database' } },
    { id: 'vault', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Key Vault' } },
  ], []);

  assert.deepEqual(markers.map(marker => [marker.nodeId, marker.kind, marker.level]), [
    ['front', 'internet', 'high'],
    ['sql', 'data', 'medium'],
    ['vault', 'secrets', 'control'],
  ]);
});
