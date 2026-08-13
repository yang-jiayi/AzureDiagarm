import test from 'node:test';
import assert from 'node:assert/strict';
import { widestGlyphIn, estimateTextWidthIn } from '../src/services/pptxExporter.ts';
import { hasMeasuredAdvance } from '../src/services/diagramExportGeometry.ts';

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
  // Beyond printable ASCII, where both estimators used to fall through to a
  // flat 0.54. The ellipsis is the one that mattered most: `fitLabelToLines`
  // appends it at every truncation point, so it is drawn 249 times across the
  // audit corpus and was charged 26% under its real width every time.
  '\u2026': 0.733, '\u2192': 1.000, '\u2190': 1.000, '\u2014': 1.000,
  '\u00d7': 0.684, '\u2013': 0.500, '\u00b7': 0.217, '\u2019': 0.229,
};

/**
 * Whitespace, measured the same way. It is kept apart from the table above
 * because `widestGlyphIn` reports the widest INK in a run and must ignore it,
 * while every width and wrap decision must charge for it: `wrapOneLine`
 * accumulates `used += w` across the runs of a line, so a space priced at zero
 * was free in the middle of a label. Charged nothing, the model bought two
 * lines for "step 19" in a 0.220in column that really needs three, and the
 * third line was drawn outside the chip.
 */
const MEASURED_WS_EM: Record<string, number> = {
  ' ': 0.274, '\u00a0': 0.274,
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
  // The average estimator is what it must NOT be: `widestGlyphIn` has to
  // report the run's widest letter, not the width its characters average to.
  // Comparing against `estimateTextWidthIn('M')` used to express that while the
  // estimator charged a flat 0.54 em for every character; now that it measures
  // real advances, `estimateTextWidthIn('M')` IS M's advance and the comparison
  // is a tautology. Compare against the run's mean advance instead, which is
  // the quantity the rule exists to reject.
  const mean = estimateTextWidthIn('Managed identity authentication', 6) / 'Managed identity authentication'.length;
  assert.ok(run > mean, `the widest glyph (${run.toFixed(4)}in) must exceed the run's mean advance (${mean.toFixed(4)}in)`);
});

test('widestGlyphIn ignores whitespace and is zero for an empty run', () => {
  assert.equal(widestGlyphIn('', 9), 0);
  assert.equal(widestGlyphIn('   ', 9), 0);
});

test('a space is charged the width it draws, not nothing', () => {
  for (const [glyph, em] of Object.entries(MEASURED_WS_EM)) {
    const modelled = estimateTextWidthIn(glyph, 100) * 72 / 100;
    assert.ok(
      modelled >= em - 0.001,
      `U+${glyph.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')} is modelled at `
      + `${modelled.toFixed(3)} em against a measured ${em.toFixed(3)}`,
    );
  }
  // The property the corpus failed on: the same letters, one more word break.
  assert.ok(
    estimateTextWidthIn('step 19', 9) > estimateTextWidthIn('step19', 9),
    'a run with an interior space must be wider than the same run without it',
  );
});

/**
 * The oracle this file could not previously provide.
 *
 * Every other test here compares the exporter's table against a measured one,
 * and the audit compares the exporter against its own copy - but neither can
 * see a character that is MISSING FROM BOTH, and both were missing the same
 * ones. A shared blind spot is not a disagreement: the ellipsis, the arrows and
 * the em dash all fell through to a flat average on both sides, so both agreed,
 * and both were wrong by up to 46%.
 *
 * So this asserts COVERAGE rather than agreement. It fails on the commit that
 * first draws a character nobody has measured, whether or not the two models
 * still match, and it is the only test here that would have caught the bug
 * before the deck shipped it.
 */
test('every character the exporters draw has a measured advance', () => {
  // Punctuation and symbols the product itself emits: the ellipsis appended by
  // every truncation, the arrows and dashes authors put in connector labels,
  // the separators the meta line is built from, and the quotes a pasted name
  // arrives with.
  const drawn = ' !"#$%&\'()*+,-./0123456789:;<=>?@'
    + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_`abcdefghijklmnopqrstuvwxyz{|}~'
    + '\u00a0\u00b7\u00d7\u2013\u2014\u2018\u2019\u201c\u201d\u2022\u2026'
    + '\u2190\u2192\u2194\u21d2\u2212'
    + '\u6ce8\u3042\uff21';
  const missing = [...drawn].filter((character) => !hasMeasuredAdvance(character));
  assert.deepEqual(
    missing,
    [],
    'these characters have no measured advance, so every width and wrap that touches them is a '
    + `guess: ${missing.map((c) => `U+${c.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`).join(', ')}`,
  );
});
