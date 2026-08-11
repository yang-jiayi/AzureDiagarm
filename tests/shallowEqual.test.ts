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
  assert.equal(shallowEqual({ a: NaN }, { a: NaN }), true, 'Object.is treats NaN as equal');
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
  assert.equal(shallowArrayEqual([], []), true);
});
