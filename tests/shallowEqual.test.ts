import assert from 'node:assert/strict';
import test from 'node:test';

import { shallowArrayEqual, shallowEqual } from '../src/utils/shallowEqual';

test('shallowEqual compares one level of own keys', () => {
  assert.equal(shallowEqual({ a: 1, b: 'x' }, { a: 1, b: 'x' }), true);
  assert.equal(shallowEqual({ a: 1 }, { a: 2 }), false);

  // Key order must not matter — JSON.stringify got this wrong and forced a
  // re-render whenever an object was rebuilt with a different insertion order.
  assert.equal(shallowEqual({ a: 1, b: 2 }, { b: 2, a: 1 }), true);

  // An explicitly undefined field is a different shape from a missing one;
  // JSON.stringify collapses both to the same output and misses the change.
  assert.equal(shallowEqual({ a: 1, b: undefined }, { a: 1 }), false);

  assert.equal(shallowEqual({ a: 1 }, { a: 1, b: 2 }), false);
  assert.equal(shallowEqual(undefined, undefined), true);
  assert.equal(shallowEqual(undefined, {}), false);
  assert.equal(shallowEqual(null, {}), false);

  // The real runtime case is pricing being REMOVED from a node, which puts the
  // nullish value on the right. Without the right-hand guard this throws a
  // TypeError from inside a React.memo comparator and takes the canvas down.
  assert.equal(shallowEqual({ a: 1 }, undefined), false);
  assert.equal(shallowEqual({ a: 1 }, null), false);

  // Same key count, no shared keys: only the per-key ownership check rejects
  // this, and both values are undefined so a value comparison would pass.
  assert.equal(shallowEqual({ a: undefined }, { b: undefined }), false);

  assert.equal(shallowEqual({ a: NaN }, { a: NaN }), true, 'Object.is treats NaN as equal');
  // JSON.stringify serialises both as "0" and would wrongly skip the render.
  assert.equal(shallowEqual({ a: -0 }, { a: 0 }), false);
});

test('shallowEqual does not recurse into nested objects', () => {
  const nested = { deep: 1 };
  assert.equal(shallowEqual({ nested }, { nested }), true);
  assert.equal(shallowEqual({ nested: { deep: 1 } }, { nested: { deep: 1 } }), false);
});

test('shallowArrayEqual compares elements without allocating', () => {
  assert.equal(shallowArrayEqual(['a', 'b'], ['a', 'b']), true);
  assert.equal(shallowArrayEqual(['a', 'b'], ['b', 'a']), false);
  assert.equal(shallowArrayEqual(['a'], ['a', 'b']), false);
  assert.equal(shallowArrayEqual(undefined, undefined), true);
  assert.equal(shallowArrayEqual(undefined, []), false);
  // Tags being removed from a node puts the nullish value on the right.
  assert.equal(shallowArrayEqual(['a'], undefined), false);
  assert.equal(shallowArrayEqual([], []), true);
  // Not an array on either side: without the Array.isArray guard, a plain
  // object would fall through to a length comparison of undefined === undefined.
  assert.equal(shallowArrayEqual({ 0: 'a', length: 1 }, ['a']), false);
  assert.equal(shallowArrayEqual([NaN], [NaN]), true);
  assert.equal(shallowArrayEqual([-0], [0]), false);
});
