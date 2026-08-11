import test from 'node:test';
import assert from 'node:assert/strict';
import { sequenceWorkflowSvg } from '../src/utils/sequenceWorkflow.ts';

/**
 * The animated SVG schedules each step at `(slot - 1) * STEP_DUR` inside a loop
 * that is exactly `workflow.length * STEP_DUR` long. Scheduling by the step's
 * declared number instead of its position would push a gap-numbered step past
 * the end of the loop, where it can never play. These tests pin that invariant
 * and the shared step->edge mapping the exports also depend on.
 */

const NODES = [
  { id: 'a', position: { x: 0, y: 0 }, width: 160, height: 113 },
  { id: 'b', position: { x: 400, y: 0 }, width: 160, height: 113 },
  { id: 'c', position: { x: 800, y: 0 }, width: 160, height: 113 },
];

const EDGES = [
  { id: 'e1', source: 'a', target: 'b' },
  { id: 'e2', source: 'b', target: 'c' },
];

/** Minimal capture of what ReactFlow renders: one group + path per edge. */
const SVG = `<svg viewBox="0 0 1200 400">
<g data-testid="rf__edge-e1"><path class="react-flow__edge-path" d="M160,56 L400,56" stroke-width="2px"/></g>
<g data-testid="rf__edge-e2"><path class="react-flow__edge-path" d="M560,56 L800,56" stroke-width="2px"/></g>
</svg>`;

/** Every `begin`-like offset the animation schedules, in seconds. */
function scheduledOffsets(svg: string, total: number): number[] {
  return [...svg.matchAll(/keyTimes="0;([\d.]+);/g)]
    .map((match) => Number(match[1]) * total)
    .filter((value) => value > 0);
}

test('a gap in the workflow numbering still animates inside the loop', () => {
  const stepDurSec = 3;
  // Steps 1 and 5 with only two entries: TOTAL is 6s, so a naive
  // (5 - 1) * 3 = 12s offset would fall outside the loop entirely.
  const svg = sequenceWorkflowSvg(SVG, {
    nodes: NODES,
    edges: EDGES,
    stepDurSec,
    workflow: [
      { step: 1, description: 'Client calls the front end', services: ['a', 'b'] },
      { step: 5, description: 'The front end queries the database', services: ['b', 'c'] },
    ],
  });

  const total = 2 * stepDurSec;
  const offsets = scheduledOffsets(svg, total);
  assert.ok(offsets.length > 0, 'the animation scheduled something');
  for (const offset of offsets) {
    assert.ok(offset < total, `offset ${offset}s must fall inside the ${total}s loop`);
  }
  // Both edges have to animate; neither may be dropped by the mapping.
  const animatedPaths = new Set(
    [...svg.matchAll(/<mpath[^>]*href="#([^"]+)"/g)].map((match) => match[1]),
  );
  assert.equal(animatedPaths.size, 2, 'both edges carry a flowing marker');
});

test('out-of-order workflow entries animate in step order, not array order', () => {
  const svg = sequenceWorkflowSvg(SVG, {
    nodes: NODES,
    edges: EDGES,
    stepDurSec: 3,
    workflow: [
      { step: 2, description: 'Second', services: ['b', 'c'] },
      { step: 1, description: 'First', services: ['a', 'b'] },
    ],
  });
  // Captions are emitted in sorted order, so step 1 must be titled first.
  const titles = [...svg.matchAll(/Step (\d+) of (\d+)/g)].map((m) => m[1]);
  assert.deepEqual(titles, ['1', '2']);
});

test('an empty workflow leaves the SVG untouched', () => {
  assert.equal(sequenceWorkflowSvg(SVG, { nodes: NODES, edges: EDGES, workflow: [] }), SVG);
});
