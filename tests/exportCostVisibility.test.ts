// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Two rules that are easy to re-break because the code that enforces them sits
 * a long way from the code that would violate them:
 *
 *  1. A cost figure the user has taken off the screen must not appear in a
 *     file they hand to a customer. `nodesForExport` is the only thing that
 *     enforces this, and it works by removing `data.pricing` — so any exporter
 *     that reads the *live* nodes instead silently opts out.
 *  2. Every CSV leaving this app opens with a UTF-8 BOM, or Excel renders
 *     Japanese service names as mojibake. The standalone download went through
 *     `csvBlob`; the copies inside the cost ZIP did not.
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node } from 'reactflow';

import { nodesForExport } from '../src/utils/nodesForExport.ts';
import { csvText, csvBlob, UTF8_BOM } from '../src/utils/csvBlob.ts';
import { buildWorkflowMarkdown } from '../src/services/workflowNarrativeExporter.ts';

function priced(id: string, extra: Record<string, unknown> = {}): Node {
  return {
    id,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: {
      label: id,
      serviceName: 'App Service',
      category: 'compute',
      pricing: { tier: 'Standard', skuName: 'S1', region: 'japaneast', monthlyCost: 73 },
      ...extra,
    },
  } as unknown as Node;
}

const TITLE = {
  architectureName: 'Contoso Platform',
  author: 'Swarm Data SE, Jiayi Yang',
  version: '1.0',
  date: '2026-08-15',
};

function markdownFor(nodes: Node[]): string {
  return buildWorkflowMarkdown({
    title: TITLE,
    nodes,
    edges: [],
    workflow: [],
    validationScore: null,
    // Passed unconditionally by the caller, exactly as in App.tsx — the gate
    // has to come from the nodes, not from this number.
    totalMonthlyCost: 4210,
    pricingMode: 'payg',
    region: 'japaneast',
  });
}

// ─── The narrative must not print a total the canvas is hiding ───────────────

test('the workflow narrative prints the cost when the canvas is showing it', () => {
  const md = markdownFor(nodesForExport([priced('web')], true));
  assert.match(md, /## Estimated Monthly Cost/);
  assert.match(md, /4,210/);
});

test('the workflow narrative drops the cost when the cost badges are off', () => {
  const md = markdownFor(nodesForExport([priced('web')], false));
  assert.doesNotMatch(md, /Estimated Monthly Cost/);
  assert.doesNotMatch(md, /4,210/);
  // The service itself must still be described — only the money goes.
  assert.match(md, /web/);
});

test('the workflow narrative drops the cost under the presentation preset', () => {
  const nodes = nodesForExport([priced('web', { stylePreset: 'presentation' })], true);
  const md = markdownFor(nodes);
  assert.doesNotMatch(md, /Estimated Monthly Cost/);
  assert.doesNotMatch(md, /4,210/);
});

test('one presentation tile does not hide a priced neighbour cost', () => {
  const nodes = nodesForExport(
    [priced('web', { stylePreset: 'presentation' }), priced('db')],
    true,
  );
  const md = markdownFor(nodes);
  assert.match(md, /## Estimated Monthly Cost/, 'db still carries pricing');
});

// ─── Every CSV opens with a BOM, wherever it is delivered ────────────────────

test('a CSV bound for a ZIP entry carries the same BOM as the download', async () => {
  const raw = 'サービス,月額\nApp Service,73';
  const inZip = csvText(raw);

  assert.ok(inZip.startsWith(UTF8_BOM), 'the ZIP entry opens with a BOM');
  assert.equal(inZip.slice(1), raw, 'and is otherwise byte-identical');

  // The standalone download and the ZIP entry must be the same file. Compare
  // bytes, not `Blob.text()` — UTF-8 decode strips a leading BOM by spec, so
  // reading the blob back as text would hide the very thing under test.
  const downloaded = new Uint8Array(await csvBlob(raw).arrayBuffer());
  const bytes = new TextEncoder().encode(inZip);
  assert.deepEqual(Array.from(downloaded), Array.from(bytes));
  assert.deepEqual(Array.from(bytes.slice(0, 3)), [0xef, 0xbb, 0xbf]);
});

test('BOM prefixing is idempotent', () => {
  assert.equal(csvText(csvText('a,b')), csvText('a,b'));
});
