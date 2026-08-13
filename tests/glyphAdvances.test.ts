import test from 'node:test';
import assert from 'node:assert/strict';
import { widestGlyphIn, estimateTextWidthIn } from '../src/services/pptxExporter.ts';

/**
 * Real advances for Yu Gothic UI, in em, measured from the installed font with
 * GDI+ (`Graphics.MeasureString`, `GenericTypographic`, 20 repeats of a glyph
 * at 100pt so the fitting error divides out).
 *
 * This table is the only width number in the repo that does not come from the
 * exporter's own model. It exists because the audit's `auditTextWidthIn` is a
 * character-for-character copy of `estimateTextWidthIn`: a gate that shares its
 * estimator with the thing it gates cannot observe the estimator being wrong,
 * and for the whole life of the file it did not — both gave every non-CJK
 * character 0.54 em, which is Segoe UI's average LOWERCASE advance. Used as a
 * maximum that understates `@` by 77%, and the "does one letter fit?" guard
 * passed a box that holds no capital at all: 39 chips 0.181in wide with 31
 * characters stacked one per line, 2.55in of smear on the first slide.
 */
const MEASURED_EM: Record<string, number> = {
  '@': 0.955, W: 0.934, M: 0.898, m: 0.861, '%': 0.818, '&': 0.800,
  O: 0.754, Q: 0.754, N: 0.748, w: 0.723, H: 0.710, U: 0.687, G: 0.686,
  '+': 0.684, '=': 0.684, '<': 0.684, '>': 0.684, '~': 0.684, '^': 0.684,
  D: 0.701, A: 0.645, V: 0.621, C: 0.619, R: 0.598, '#': 0.591, X: 0.590,
  d: 0.589, g: 0.589, q: 0.589, o: 0.586, b: 0.588, p: 0.588, K: 0.580,
  B: 0.573, Z: 0.570, P: 0.560, h: 0.566, n: 0.566, u: 0.566, Y: 0.553,
  a: 0.509, S: 0.531, T: 0.524, E: 0.506, F: 0.488, y: 0.484, v: 0.479,
  x: 0.459, z: 0.452, L: 0.471, c: 0.462, k: 0.497, e: 0.523,
  '0': 0.539, '5': 0.539, '9': 0.539, $: 0.539, '\\': 0.539,
  '?': 0.448, '_': 0.415, s: 0.424, '*': 0.417, '-': 0.400, '/': 0.390,
  '"': 0.392, r: 0.348, t: 0.339, f: 0.313, '(': 0.302, ')': 0.302,
  '[': 0.302, ']': 0.302, '{': 0.302, '}': 0.302, '!': 0.284,
  '`': 0.268, I: 0.266, i: 0.242, j: 0.242, l: 0.242, '|': 0.239,
  "'": 0.230, '.': 0.217, ',': 0.217, ':': 0.217, ';': 0.217,
  // Full-width: one em by construction, and both estimators agree.
  '注': 1.000, 'あ': 0.816, 'Ａ': 1.000,
};

test('widestGlyphIn never reports a glyph narrower than the font actually draws', () => {
  const under: string[] = [];
  for (const [glyph, em] of Object.entries(MEASURED_EM)) {
    // 100pt so a rounding hair cannot decide the comparison.
    const modelled = widestGlyphIn(glyph, 100) * 72 / 100;
    if (modelled < em - 0.001) under.push(`${glyph}: model ${modelled.toFixed(3)} < measured ${em.toFixed(3)}`);
  }
  assert.deepEqual(under, [], `the sizer would let these glyphs into a box too narrow to hold them:\n${under.join('\n')}`);
});

test('widestGlyphIn reports the widest glyph of a run, not its average', () => {
  // "Managed identity authentication" is the exact string that shipped 39
  // one-glyph-per-line chips. Its widest letter is M.
  const run = widestGlyphIn('Managed identity authentication', 6);
  assert.ok(run >= (MEASURED_EM.M * 6) / 72 - 0.0005, `widest of the run was ${run.toFixed(4)}in, M needs ${((MEASURED_EM.M * 6) / 72).toFixed(4)}in`);
  // The average estimator is what it must NOT be: 0.54 em at 6pt is 0.045in.
  assert.ok(run > estimateTextWidthIn('M', 6), 'the widest glyph must exceed the average-advance estimate for a capital');
});

test('widestGlyphIn ignores whitespace and is zero for an empty run', () => {
  assert.equal(widestGlyphIn('', 9), 0);
  assert.equal(widestGlyphIn('   ', 9), 0);
});
