/*
 * Allocation-free equality helpers for React.memo comparators.
 *
 * The node comparators run for every node on every React Flow store update
 * (drag, pan, selection, zoom). `JSON.stringify(a) === JSON.stringify(b)` was
 * costing two full serialisations per node per frame on large diagrams, and it
 * is also wrong in two ways: it is sensitive to key insertion order, so a
 * rebuilt object with identical values compares unequal, and it collapses an
 * explicitly `undefined` field into a missing one, so some real changes compare
 * equal.
 */

/** One-level object comparison. Handles null/undefined and differing key sets. */
export function shallowEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (
    typeof left !== 'object' || left === null
    || typeof right !== 'object' || right === null
  ) return false;

  const leftKeys = Object.keys(left as Record<string, unknown>);
  const rightKeys = Object.keys(right as Record<string, unknown>);
  if (leftKeys.length !== rightKeys.length) return false;

  for (const key of leftKeys) {
    if (!Object.prototype.hasOwnProperty.call(right, key)) return false;
    if (!Object.is(
      (left as Record<string, unknown>)[key],
      (right as Record<string, unknown>)[key],
    )) return false;
  }
  return true;
}

/** Element-wise comparison for the small primitive arrays carried on node data. */
export function shallowArrayEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (!Array.isArray(left) || !Array.isArray(right)) return false;
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (!Object.is(left[index], right[index])) return false;
  }
  return true;
}
