import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { exportToDrawio } from '../src/services/drawioExporter.ts';

function countOccurrences(haystack: string, needle: string): number {
  return haystack.split(needle).length - 1;
}

// Balanced-tag + escaped-ampersand well-formedness check (same shape as the
// VSDX suite's helper) so it correctly ignores self-closing tags.
function assertWellFormed(xml: string, label: string): void {
  const stack: string[] = [];
  const tagPattern = /<\/?([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = tagPattern.exec(xml)) !== null) {
    const text = xml.slice(lastIndex, match.index);
    assert.equal(/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(text), false,
      `${label}: unescaped ampersand near ${JSON.stringify(text.slice(0, 60))}`);
    lastIndex = tagPattern.lastIndex;
    const raw = match[0];
    const name = match[1];
    if (raw.startsWith('<?') || raw.startsWith('<!')) continue;
    if (raw.startsWith('</')) {
      assert.equal(stack.pop(), name, `${label}: mismatched closing tag </${name}>`);
    } else if (!match[3]) {
      stack.push(name);
    }
  }
  assert.deepEqual(stack, [], `${label}: unclosed tags ${stack.join(', ')}`);
}

test('draw.io export places nested groups at their resolved absolute position (fixes 8 & 12)', async () => {
  const nodes = [
    { id: 'outer', type: 'groupNode', position: { x: 1000, y: 500 }, style: { width: 800, height: 600 }, data: { label: 'Subscription' } },
    { id: 'inner', type: 'groupNode', position: { x: 40, y: 60 }, parentNode: 'inner-parent', style: { width: 400, height: 300 }, data: { label: 'VNet', customColor: { border: '#dc2626' } } },
    { id: 'svc', type: 'azureNode', position: { x: 20, y: 30 }, parentNode: 'inner', width: 150, height: 75, data: { label: 'VM', sku: 'D2s_v3', region: 'eastus' } },
  ] as unknown as Node[];
  // Fix the parent chain: inner→outer, svc→inner.
  (nodes[1] as { parentNode?: string }).parentNode = 'outer';

  const edges: Edge[] = [];
  const xml = await exportToDrawio(nodes, edges, 'Nested');

  // Nested group lands at outer + relative offset, not its bare relative offset.
  assert.match(xml, /x="1040" y="560"/, 'inner zone uses absolute coordinates');
  assert.match(xml, /x="1060" y="590"/, 'the child service follows its zone');
  // The zone's own colour is honoured (fix 6).
  assert.match(xml, /strokeColor=#dc2626/, 'custom zone border colour is applied');
  // Metadata sub-line rides along on the service value (fix 10).
  assert.match(xml, /VM&#xa;D2s_v3 · eastus/, 'service carries a SKU · region sub-line');

  // Page bounds come from the resolved boxes, comfortably past the outer zone.
  const page = xml.match(/pageWidth="(\d+)" pageHeight="(\d+)"/);
  assert.ok(page);
  assert.ok(Number(page![1]) >= 1800, 'page width covers the resolved content');
  assert.ok(Number(page![2]) >= 1100, 'page height covers the resolved content');
});

test('draw.io export colours a security edge red and keeps XML balanced (fixes 4 & 11)', async () => {
  const nodes = [
    { id: 'a', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Firewall', category: 'security' } },
    { id: 'b', type: 'azureNode', position: { x: 400, y: 0 }, data: { label: 'Gateway', category: 'networking' } },
  ] as unknown as Node[];
  const edges = [
    { id: 'sec', source: 'a', target: 'b', data: { connectionType: 'security', label: 'mTLS' } },
  ] as unknown as Edge[];

  const xml = await exportToDrawio(nodes, edges, 'Security');

  assert.match(xml, /strokeColor=#dc2626;/, 'the security connector is red');
  assert.match(xml, /dashed=1/, 'the security connector is dashed');
  assert.match(xml, /value="mTLS"/, 'the edge label is emitted');
  // The two service tiles use the shared category palette, not one grey default.
  assert.ok(countOccurrences(xml, 'fillColor=') >= 3, 'groups and services are all filled');
  assertWellFormed(xml, 'security.drawio');
});

test('draw.io export renders self-loops and parallel edges distinctly (fix 7)', async () => {
  const nodes = [
    { id: 'a', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Queue' } },
    { id: 'b', type: 'azureNode', position: { x: 500, y: 300 }, data: { label: 'Worker' } },
  ] as unknown as Node[];
  const edges = [
    { id: 'loop', source: 'a', target: 'a', data: { label: 'retry' } },
    { id: 'p1', source: 'a', target: 'b' },
    { id: 'p2', source: 'a', target: 'b' },
  ] as unknown as Edge[];

  const xml = await exportToDrawio(nodes, edges, 'Loops');
  // Self-loop and the 2nd parallel edge both add explicit waypoints.
  assert.ok(countOccurrences(xml, '<Array as="points">') >= 2, 'de-collision waypoints are emitted');
  assertWellFormed(xml, 'loops.drawio');
});
