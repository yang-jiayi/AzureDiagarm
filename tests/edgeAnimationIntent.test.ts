import test from 'node:test';
import assert from 'node:assert/strict';
import { edgeAnimationIntent, resolveFlowAnimated } from '../src/utils/edgePresentation.ts';

/**
 * A live edge's `flowAnimated` is the product of two independent preferences:
 * the user's intent for that edge, and the app-wide "animate connections"
 * switch. Only the product is serialised into a saved diagram, and the switch
 * lives in localStorage rather than in the file -- so the switch at save time
 * and at open time legitimately differ. A restore therefore has to read the
 * intent, never the product.
 */

test('a restore reads the intent, not the resolved state it was last saved at', () => {
  // Saved while the global switch was OFF: every edge carries flowAnimated:false.
  const saved = { baseFlowAnimated: true, flowAnimated: false };
  assert.equal(edgeAnimationIntent(saved, true), true);
  assert.equal(resolveFlowAnimated(saved, true, true), true);
});

test('a paused edge stays paused when the global switch is on', () => {
  const paused = { baseFlowAnimated: false, flowAnimated: false };
  assert.equal(resolveFlowAnimated(paused, true, true), false);
});

test('the stored product is read only for files written before the intent existed', () => {
  assert.equal(edgeAnimationIntent({ flowAnimated: false }, true), false);
  assert.equal(edgeAnimationIntent({ flowAnimated: true }, false), true);
});

test('an edge carrying neither field falls back to its connection type default', () => {
  assert.equal(edgeAnimationIntent({}, true), true);
  assert.equal(edgeAnimationIntent({}, false), false);
  assert.equal(edgeAnimationIntent(undefined, false), false);
});

test('nothing animates while the global switch is off', () => {
  assert.equal(resolveFlowAnimated({ baseFlowAnimated: true }, true, false), false);
});

/**
 * The full save/reopen matrix. Before the fix exactly one row diverged: an edge
 * whose intent was "animate", saved while the switch was off, came back
 * permanently static because its own `flowAnimated:false` masked its
 * `baseFlowAnimated:true`. PPTX maps animated to dashed, so that divergence
 * reached the exports too.
 */
test('the reopened edge agrees with the authoring intent in every combination', () => {
  const divergent: string[] = [];
  for (const intent of [false, true]) {
    for (const switchAtSave of [false, true]) {
      for (const switchAtOpen of [false, true]) {
        const stored = {
          baseFlowAnimated: intent,
          flowAnimated: switchAtSave && intent,
        };
        const reopened = resolveFlowAnimated(stored, true, switchAtOpen);
        const expected = switchAtOpen && intent;
        if (reopened !== expected) {
          divergent.push(
            `intent=${intent} saveSwitch=${switchAtSave} openSwitch=${switchAtOpen}`
            + ` -> ${reopened}, expected ${expected}`,
          );
        }
      }
    }
  }
  assert.deepEqual(divergent, []);
});
