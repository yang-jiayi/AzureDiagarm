// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * A cost the canvas is hiding must not reappear in the exported file.
 *
 * The PNG and SVG captures were always right about this, because they
 * photograph the DOM and the DOM has no badge on it. The native exporters read
 * `data.pricing` off the node instead, so they printed the figure regardless —
 * on the deck, the Visio sheet and the HTML. That is the one number a user
 * takes off the screen deliberately before showing a diagram to a customer.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node } from 'reactflow';

import { nodesForExport } from '../src/utils/nodesForExport.ts';

function priced(id: string, extra: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: {
      label: id,
      category: 'compute',
      pricing: { estimatedCost: 128.4, tier: 'P1v3', region: 'japaneast', quantity: 1 },
      ...extra,
    },
  } as Node;
}

test('the cost survives when the canvas is showing it', () => {
  const [node] = nodesForExport([priced('a')], true);
  assert.ok((node.data as Record<string, unknown>).pricing);
});

test('turning cost badges off takes the figure out of the export', () => {
  const [node] = nodesForExport([priced('a')], false);
  assert.equal((node.data as Record<string, unknown>).pricing, undefined);
});

test('presentation styling takes the figure out even with badges enabled', () => {
  const [node] = nodesForExport([priced('a', { stylePreset: 'presentation' })], true);
  assert.equal((node.data as Record<string, unknown>).pricing, undefined);
});

test('the preset is read per node, not for the whole diagram', () => {
  const out = nodesForExport(
    [priced('a', { stylePreset: 'presentation' }), priced('b', { stylePreset: 'detailed' })],
    true,
  );
  assert.equal((out[0].data as Record<string, unknown>).pricing, undefined);
  assert.ok((out[1].data as Record<string, unknown>).pricing);
});

test('nothing else on the node is disturbed', () => {
  const [node] = nodesForExport([priced('a', { tags: ['PCI'], serviceName: 'App Service' })], false);
  const data = node.data as Record<string, unknown>;
  assert.equal(data.label, 'a');
  assert.equal(data.category, 'compute');
  assert.equal(data.serviceName, 'App Service');
  assert.deepEqual(data.tags, ['PCI']);
});

test('the caller list is not mutated', () => {
  const input = [priced('a')];
  nodesForExport(input, false);
  assert.ok((input[0].data as Record<string, unknown>).pricing, 'the canvas node lost its pricing');
});

test('a node with no pricing is passed through untouched', () => {
  const plain = { id: 'x', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'x' } } as Node;
  const [node] = nodesForExport([plain], false);
  assert.equal(node, plain);
});

test('a node with no data at all does not throw', () => {
  const bare = { id: 'x', type: 'azureNode', position: { x: 0, y: 0 } } as unknown as Node;
  assert.doesNotThrow(() => nodesForExport([bare], false));
});
