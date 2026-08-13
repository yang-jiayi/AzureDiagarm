import test from 'node:test';
import assert from 'node:assert/strict';
import { widestGlyphIn, estimateTextWidthIn } from '../src/services/pptxExporter.ts';
import { hasMeasuredAdvance, advanceTier, drawableInColumn, advanceWidthIn, widestGlyphUpperIn, widestGlyphIn as widestGlyphLowerIn } from '../src/services/diagramExportGeometry.ts';

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

/**
 * Latin-1 Supplement and Latin Extended-A, measured the same way.
 *
 * These fell through to the 1 em unknown-character fallback, which the model
 * defended as harmless over-reservation. It is not: `widestGlyphIn` gates
 * whether a tile is named at all, so an over-charge there DELETES the name -
 * and `Reseau prive securise` spelled with its accents drew nothing on a tile
 * that drew the unaccented spelling in full.
 */
const MEASURED_LATIN_EM: Record<string, number> = {
  '\u00e9': 0.523, '\u00e7': 0.462, '\u00e3': 0.509, '\u00c9': 0.506,
  '\u00f1': 0.566, '\u00fc': 0.566, '\u00b0': 0.377, '\u0131': 0.266,
  '\u00df': 0.544, '\u0142': 0.265, '\u0105': 0.554, '\u017e': 0.452,
  '\u0153': 0.928, '\u00e6': 0.832, '\u00d8': 0.754, '\u016f': 0.580,
};

test('Latin-1 and Latin Extended-A are measured, not guessed at a full em', () => {
  for (const [character, em] of Object.entries(MEASURED_LATIN_EM)) {
    assert.equal(hasMeasuredAdvance(character), true, `${character} has no measured advance`);
    const got = estimateTextWidthIn(character, 72);
    assert.ok(
      Math.abs(got - em) < 0.02,
      `${character} measures ${em} em but the model charges ${got.toFixed(3)}`,
    );
  }
});

test('an accent does not make a name unnameable', () => {
  // Same string, twice, differing only in four accents. `e-acute` is NARROWER
  // than the `R` that really is the widest glyph, so the two must agree
  // exactly - a divergence here is the deck and the sheet naming different
  // services in one diagram.
  const accented = 'R\u00e9seau priv\u00e9 s\u00e9curis\u00e9';
  const plain = 'Reseau prive securise';
  assert.equal(widestGlyphIn(accented, 7), widestGlyphIn(plain, 7));
});

test('a wide glyph in an ordinary name does not withhold it', () => {
  // `drawableInColumn` asks two questions: can the widest glyph set at all,
  // and does the column hold two TYPICAL characters. The old single question,
  // `column >= 2 * widest`, let one `m` at 0.861 em speak for a string whose
  // mean is 0.55 and delete the whole name from a column setting 2.8
  // characters a line.
  const name = 'Cami\u00f3n log\u00edstica an\u00e1lisis';
  assert.equal(drawableInColumn(name, 7, 0.1508), true);
  // A column that cannot set the widest glyph still refuses.
  assert.equal(drawableInColumn(name, 7, 0.05), false);
  // So does one too narrow for two typical characters: the mean is 0.043in,
  // so anything under 0.086in spells the name one letter per line.
  assert.equal(drawableInColumn(name, 7, 0.085), false);
});

test('an emoji cluster is charged once, not once per code point', () => {
  const one = advanceWidthIn('\u{1f680}', 72);
  // A flag is two regional indicators drawn as a single glyph.
  assert.ok(Math.abs(advanceWidthIn('\u{1f1ef}\u{1f1f5}', 72) - one) < 1e-9);
  // So is a skin-toned thumb.
  assert.ok(Math.abs(advanceWidthIn('\u{1f44d}\u{1f3fd}', 72) - one) < 1e-9);
  // And a four-person family, which is four astral code points and three
  // joiners: 5.44 em charged against one glyph drawn.
  const family = '\u{1f468}\u200d\u{1f469}\u200d\u{1f467}\u200d\u{1f466}';
  assert.ok(Math.abs(advanceWidthIn(family, 72) - one) < 1e-9);
});

/**
 * Round 58: the same independent-measurement discipline, extended past Latin.
 *
 * These numbers come from WPF `GlyphTypeface` - `CharacterToGlyphMap` says
 * exactly which code points the font FILE holds, so there is no fallback
 * ambiguity, and `AdvanceWidths[glyph]` is the exact em advance. That is a
 * different API from the GDI+ `MeasureString` used above and from anything the
 * exporter does; where the two APIs could both be asked they agreed to 0.0004
 * em over 317 pre-existing entries, which is why a disagreement is worth
 * chasing. All four disagreements found this round were real defects.
 *
 * Where the reviewer's GDI+ reading of the same character is 5% lower, it is
 * because `MeasureString` was resolving the character through a SUBSTITUTE
 * font. These are the advances in Yu Gothic UI's own file, which is the font
 * the exporter names, so they are the ones that govern.
 */
const MEASURED_SCRIPT_EM: Record<string, number> = {
  '\u0430': 0.5088, '\u043d': 0.5771, '\u0435': 0.5229, '\u0441': 0.4619,
  '\u03b1': 0.6143, '\u03b4': 0.5840, '\u03c9': 0.8081,
  '\u021b': 0.3510, '\u0219': 0.5070,
};

test('the measured tables cover the scripts the fallback used to guess at 1 em', () => {
  for (const [glyph, em] of Object.entries(MEASURED_SCRIPT_EM)) {
    assert.equal(
      hasMeasuredAdvance(glyph),
      true,
      `U+${glyph.codePointAt(0)!.toString(16)} is still guessed`,
    );
    const got = advanceWidthIn(glyph, 72);
    assert.ok(
      Math.abs(got - em) <= 0.006,
      `U+${glyph.codePointAt(0)!.toString(16)}: model ${got.toFixed(4)} vs font ${em}`,
    );
  }
});

test('an unmeasured astral code point is not certified as measured', () => {
  // The oracle that exists to catch guesses returned true for EVERY astral code
  // point, so it could not fire on the one range the file documents as guessed.
  assert.equal(hasMeasuredAdvance('\u{1d400}'), false);
  assert.equal(hasMeasuredAdvance('\u{20000}'), true);
  assert.equal(hasMeasuredAdvance('\u{1f600}'), true);
});

test('an astral CJK ideograph is one em, not the emoji width', () => {
  assert.ok(Math.abs(advanceWidthIn('\u{20000}', 72) - 1) <= 0.001);
});

test('two adjacent flags are charged as two clusters', () => {
  // The join test looked only at the immediately preceding code point, so a
  // second flag joined the first and a pair was charged for one glyph.
  const one = advanceWidthIn('\u{1F1FA}\u{1F1F8}', 72);
  const two = advanceWidthIn('\u{1F1FA}\u{1F1F8}\u{1F1EC}\u{1F1E7}', 72);
  assert.ok(Math.abs(two - one * 2) <= 0.002, `${two} vs ${one * 2}`);
});

test('the three hand-set advances cross-measurement corrected', () => {
  // Found by re-measuring every new entry through the OTHER API: a bullet was
  // charged 14% under, which is the direction that paints outside the shape.
  assert.ok(Math.abs(advanceWidthIn('\u2022', 72) - 0.406) <= 0.004);
  assert.ok(Math.abs(advanceWidthIn('\u201c', 72) - 0.377) <= 0.004);
  assert.ok(Math.abs(advanceWidthIn('\u201d', 72) - 0.377) <= 0.004);
});

test('every Unicode space is charged its own width, not the plain space width', () => {
  // Every space in Unicode took the plain-space advance regardless, so an em
  // space and an IDEOGRAPHIC space - ordinary punctuation in the Japanese
  // service names this app draws - were charged 265% under.
  assert.ok(Math.abs(advanceWidthIn('\u2003', 72) - 1) <= 0.002, 'em space');
  assert.ok(Math.abs(advanceWidthIn('\u3000', 72) - 1) <= 0.002, 'ideographic space');
  assert.ok(Math.abs(advanceWidthIn('\u2002', 72) - 0.5) <= 0.002, 'en space');
  assert.ok(Math.abs(advanceWidthIn('\u2009', 72) - 0.2) <= 0.002, 'thin space');
  assert.ok(Math.abs(advanceWidthIn('\u200a', 72) - 0.1) <= 0.002, 'hair space');
});

test('both clauses of the column test take the same bound on an unknown glyph', () => {
  // A wholly untabled script argued both sides of its own case: the widest
  // glyph was taken at the lower bound and the mean at the upper, and the upper
  // bound bound. The name was refused, and on a Visio sheet a refused name is
  // nowhere on the page.
  const unknown = '\u{10A00}\u{10A01}\u{10A02}\u{10A03}';
  assert.equal(drawableInColumn(unknown, 7, 0.05), true);
});

test('the sizing bound and the drawing bound are different functions', () => {
  // One caller sizes a chip to hold the widest glyph and needs the UPPER bound;
  // the rest ask whether a name can be drawn and need the LOWER one. Sharing a
  // function made one of them wrong whichever way it was written.
  const unknown = '\u{10A00}';
  assert.ok(widestGlyphLowerIn(unknown, 7) < widestGlyphUpperIn(unknown, 7));
});

/**
 * ASK-58-A: a measurement has a TIER, and the message has to say which.
 *
 * "Measured" and "measured in a font that is not the label font" are both
 * numbers off a real `hmtx`, but they are different claims. Yu Gothic UI has
 * no Thai, no Hebrew and no precomposed Vietnamese at all, so for those a
 * substitute face is the only measurement that exists - calling them untabled
 * would send someone hunting for a table entry nobody can ever write, and
 * calling them tier 1 would claim a precision the font file does not have.
 */
test('an advance reports whether it came from the label font or a substitute', () => {
  for (const glyph of ['A', 'e', '9', '\u6ce8', '\u3042', '\u00e9', '\u0430']) {
    assert.equal(advanceTier(glyph), 'label', `${glyph} should be measured in the label font`);
  }
  // Scripts Yu Gothic UI does not contain, and emoji, which Segoe UI Emoji draws.
  for (const glyph of ['\u0e01', '\u05d0', '\u1ea1', '\u{1f600}']) {
    assert.equal(advanceTier(glyph), 'substitute',
      `${glyph} is drawn by a substitute face, so its tier must say so`);
  }
  // Genuinely untabled: mathematical alphanumerics are in no face here.
  assert.equal(advanceTier('\u{1d400}'), 'none');
  // The tier must never disagree with the boolean every call site gates on.
  for (const glyph of ['A', '\u6ce8', '\u0e01', '\u{1f600}', '\u{1d400}']) {
    assert.equal(advanceTier(glyph) !== 'none', hasMeasuredAdvance(glyph),
      `the tier and the boolean disagree about ${glyph}`);
  }
});
