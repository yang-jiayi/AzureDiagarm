import test from 'node:test';
import assert from 'node:assert/strict';
import { planBothRun, type PendingRetry } from '../src/utils/bothModeRetry.ts';

const BRIEF = 'A retail platform with an API tier and a SQL database.';

function pending(overrides: Partial<PendingRetry> = {}): PendingRetry {
  return {
    missing: 'blueprint',
    brief: BRIEF,
    prompt: `enriched: ${BRIEF}`,
    manifest: { components: [], zones: [] } as unknown as PendingRetry['manifest'],
    canvasApplied: true,
    ...overrides,
  };
}

test('a normal run generates both deliverables and builds a fresh manifest', () => {
  const plan = planBothRun(null, BRIEF);
  assert.equal(plan.retry, null);
  assert.equal(plan.runTopology, true);
  assert.equal(plan.runBlueprint, true);
  assert.equal(plan.reuseManifest, false);
});

test('retrying a failed blueprint leaves the applied topology untouched', () => {
  const plan = planBothRun(pending({ missing: 'blueprint' }), BRIEF);
  assert.ok(plan.retry, 'the retry is honoured');
  assert.equal(plan.runBlueprint, true, 'the missing deliverable is regenerated');
  assert.equal(plan.runTopology, false, 'the canvas is not overwritten or billed again');
  assert.equal(plan.reuseManifest, true, 'the manifest pre-pass is not paid for twice');
});

test('retrying a failed topology leaves the stashed blueprint untouched', () => {
  const plan = planBothRun(pending({ missing: 'topology' }), BRIEF);
  assert.equal(plan.runTopology, true);
  assert.equal(plan.runBlueprint, false);
});

test('editing the brief cancels the retry and regenerates everything', () => {
  // Otherwise the user would silently get only half of a changed request.
  const plan = planBothRun(pending(), `${BRIEF} Add a cache.`);
  assert.equal(plan.retry, null);
  assert.equal(plan.runTopology, true);
  assert.equal(plan.runBlueprint, true);
  assert.equal(plan.reuseManifest, false, 'a changed brief invalidates the cached manifest');
});

test('whitespace-level brief edits still count as a change', () => {
  assert.equal(planBothRun(pending(), `${BRIEF} `).retry, null);
});

test('the retry carries the original prompt, not one rebuilt from the canvas', () => {
  const plan = planBothRun(pending(), BRIEF);
  // Rebuilding would produce a "MODIFY EXISTING ARCHITECTURE" instruction
  // because the topology from the first attempt is already on the canvas.
  assert.equal(plan.retry?.prompt, `enriched: ${BRIEF}`);
  assert.equal(plan.retry?.canvasApplied, true);
});
