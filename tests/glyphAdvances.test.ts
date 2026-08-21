import test from 'node:test';
import assert from 'node:assert/strict';
import { widestGlyphIn, estimateTextWidthIn } from '../src/services/pptxExporter.ts';
import { hasMeasuredAdvance, advanceTier, drawableInColumn, advanceWidthIn, widestGlyphUpperIn, widestGlyphIn as widestGlyphLowerIn, hasMeasuredCluster, fitLabelToWidth, singleLineName } from '../src/services/diagramExportGeometry.ts';

/**
 * Real advances for Arial, in em, measured by laying each character out in
 * Chromium one at a time with kerning off.
 *
 * This table is one of only two width numbers in the repo that do not come from
 * the exporter's own model. It exists because the audit's `auditTextWidthIn` is
 * a character-for-character copy of `estimateTextWidthIn`: a gate that shares
 * its estimator with the thing it gates cannot observe the estimator being
 * wrong, and for the whole life of the file it did not — both gave every
 * non-CJK character 0.54 em, an average LOWERCASE advance. Used as a maximum
 * that understates `@` by 88%, and the "does one letter fit?" guard passed a box
 * that holds no capital at all: 39 chips 0.181in wide with 31 characters stacked
 * one per line, 2.55in of smear on the first slide.
 *
 * A BROWSER is asked, deliberately, because the exporter's own tables are read
 * from the font file's `hmtx`. Generating both from one measurement would make
 * this file agree with the exporter by construction. Kerning is switched off and
 * characters are measured in isolation because the model prices one glyph at a
 * time and never charges for a pair: measuring `"f"` repeated turns it into `ﬀ`
 * ligatures and reports 0.261 em for a glyph that advances 0.278.
 */
const MEASURED_EM: Record<string, number> = {
  '@': 1.015, W: 0.944, '%': 0.889, M: 0.833, m: 0.833,
  O: 0.778, Q: 0.778, G: 0.778,
  N: 0.722, w: 0.722, H: 0.722, U: 0.722, D: 0.722, C: 0.722, R: 0.722,
  '&': 0.667, A: 0.667, V: 0.667, X: 0.667, K: 0.667, B: 0.667,
  P: 0.667, Y: 0.667, S: 0.667, E: 0.667,
  Z: 0.611, T: 0.611, F: 0.611,
  '+': 0.584, '=': 0.584, '<': 0.584, '>': 0.584, '~': 0.584,
  '#': 0.556, d: 0.556, g: 0.556, q: 0.556, o: 0.556, b: 0.556, p: 0.556,
  h: 0.556, n: 0.556, u: 0.556, a: 0.556, e: 0.556, L: 0.556,
  '0': 0.556, '5': 0.556, '9': 0.556, $: 0.556, '?': 0.556, '_': 0.556,
  y: 0.500, v: 0.500, x: 0.500, z: 0.500, c: 0.500, k: 0.500, s: 0.500,
  '^': 0.469, '*': 0.389, '"': 0.355,
  '{': 0.334, '}': 0.334,
  '-': 0.333, r: 0.333, '(': 0.333, ')': 0.333, '`': 0.333,
  '\\': 0.278, '/': 0.278, t: 0.278, f: 0.278, '[': 0.278, ']': 0.278,
  '!': 0.278, I: 0.278, '.': 0.278, ',': 0.278, ':': 0.278, ';': 0.278,
  '|': 0.260, i: 0.222, j: 0.222, l: 0.222, "'": 0.191,
  // Full-width and kana, which the East Asian face draws. `注` and `Ａ` really
  // are one em; `あ` is not, and is listed at its true width so the full-em RULE
  // the model applies is being tested as an upper bound rather than assumed.
  '注': 1.000, 'あ': 0.816, 'Ａ': 1.000,
  // Beyond printable ASCII, where both estimators used to fall through to a
  // flat 0.54. The ellipsis is the one that mattered most: `fitLabelToLines`
  // appends it at every truncation point, so it is drawn 249 times across the
  // audit corpus and was charged 26% under its real width every time — and this
  // font draws it wider still, at a full em.
  '\u2026': 1.000, '\u2192': 1.000, '\u2190': 1.000, '\u2014': 1.000,
  '\u00d7': 0.584, '\u2013': 0.556, '\u00b7': 0.333, '\u2019': 0.222,
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
  ' ': 0.278, '\u00a0': 0.278,
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
  '\u00e9': 0.556, '\u00e7': 0.500, '\u00e3': 0.556, '\u00c9': 0.667,
  '\u00f1': 0.556, '\u00fc': 0.556, '\u00b0': 0.400, '\u0131': 0.278,
  '\u00df': 0.611, '\u0142': 0.222, '\u0105': 0.556, '\u017e': 0.500,
  '\u0153': 0.944, '\u00e6': 0.889, '\u00d8': 0.778, '\u016f': 0.556,
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
 * Measured by Chromium, one character at a time with kerning off - a different
 * engine and a different API from the font-file `hmtx` read the exporter's own
 * tables come from, so a disagreement here is worth chasing. Over the 1,380
 * code points the two methods can both be asked about they differ on four that
 * the label font itself draws; every other divergence is a character neither
 * font contains, where the question is which face substitutes.
 *
 * Hebrew and Arabic used to sit in this file's substituted list and no longer
 * do: the label font contains them, so the advance is the font's own and the
 * uncertainty is gone rather than merely restated.
 */
const MEASURED_SCRIPT_EM: Record<string, number> = {
  '\u0430': 0.5562, '\u043d': 0.5522, '\u0435': 0.5562, '\u0441': 0.5000,
  '\u03b1': 0.5781, '\u03b4': 0.5566, '\u03c9': 0.7808,
  '\u021b': 0.2778, '\u0219': 0.5000, '\u05d0': 0.5630,
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
  // The numbers moved with the label font; that they are still read from the
  // file rather than set by hand is the part that matters.
  assert.ok(Math.abs(advanceWidthIn('\u2022', 72) - 0.35) <= 0.004);
  assert.ok(Math.abs(advanceWidthIn('\u201c', 72) - 0.333) <= 0.004);
  assert.ok(Math.abs(advanceWidthIn('\u201d', 72) - 0.333) <= 0.004);
});

test('every Unicode space is charged its own width, not the plain space width', () => {
  // Every space in Unicode took the plain-space advance regardless, so an em
  // space and an IDEOGRAPHIC space - ordinary punctuation in the Japanese
  // service names this app draws - were charged 260% under.
  assert.ok(Math.abs(advanceWidthIn('\u2003', 72) - 1) <= 0.002, 'em space');
  assert.ok(Math.abs(advanceWidthIn('\u3000', 72) - 1) <= 0.002, 'ideographic space');
  assert.ok(Math.abs(advanceWidthIn('\u2002', 72) - 0.5) <= 0.002, 'en space');
  assert.ok(Math.abs(advanceWidthIn('\u2009', 72) - 0.2) <= 0.002, 'thin space');
  assert.ok(Math.abs(advanceWidthIn('\u200a', 72) - 0.083) <= 0.002, 'hair space');
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
 * numbers off a real `hmtx`, but they are different claims. The label font has
 * no Thai and no emoji, so for those a substitute face is the only measurement
 * that exists - calling them untabled would send someone hunting for a table
 * entry nobody can ever write, and calling them tier 1 would claim a precision
 * the font file does not have. Moving to Arial shrank this tier: 490 code
 * points that needed a substitute now come from the label font itself.
 */
test('an advance reports whether it came from the label font or a substitute', () => {
  for (const glyph of ['A', 'e', '9', '\u6ce8', '\u3042', '\u00e9', '\u0430']) {
    assert.equal(advanceTier(glyph), 'label', `${glyph} should be measured in the label font`);
  }
  // Scripts the label font does not contain, and emoji, which Segoe UI Emoji
  // draws. Hebrew and Vietnamese have LEFT this list: Arial contains both, so
  // they are now priced from the label font's own file.
  for (const glyph of ['\u0e01', '\u{1f600}']) {
    assert.equal(advanceTier(glyph), 'substitute',
      `${glyph} is drawn by a substitute face, so its tier must say so`);
  }
  for (const glyph of ['\u05d0', '\u1ea1']) {
    assert.equal(advanceTier(glyph), 'label',
      `the label font contains ${glyph}, so it is no longer a substitute guess`);
  }
  // Genuinely untabled: mathematical alphanumerics are in no face here.
  assert.equal(advanceTier('\u{1d400}'), 'none');
  // The tier must never disagree with the boolean every call site gates on.
  for (const glyph of ['A', '\u6ce8', '\u0e01', '\u{1f600}', '\u{1d400}']) {
    assert.equal(advanceTier(glyph) !== 'none', hasMeasuredAdvance(glyph),
      `the tier and the boolean disagree about ${glyph}`);
  }
});

/**
 * Round 63: VARIATION SELECTOR-16 PROMOTES a cluster, it does not vanish.
 *
 * The selector was charged zero and the base was charged its own advance, so
 * the model priced an emoji at its TEXT width. A heart is 1.000 em as a
 * dingbat and 1.373 as an emoji - 27% under - and an ASCII base is far worse,
 * because a keycap is a digit plus the selector plus a zero-width combining
 * enclosure: 0.556 em charged against 1.373 drawn, 60% under. Under-charging
 * is the direction that paints a line out past its box.
 */
const EMOJI_EM = 1.373;

test('a variation selector promotes its base to the emoji advance', () => {
  for (const cluster of ['\u2764\ufe0f', '\u2601\ufe0f', '\u2699\ufe0f', '\u26a0\ufe0f']) {
    assert.ok(
      Math.abs(advanceWidthIn(cluster, 72) - EMOJI_EM) < 1e-9,
      `${cluster} is drawn as an emoji but the model charges `
      + `${advanceWidthIn(cluster, 72).toFixed(3)} em, not ${EMOJI_EM}`,
    );
  }
});

test('promotion reaches an ASCII base, so a keycap is one emoji wide', () => {
  // The enclosing keycap U+20E3 is already a zero-width combining mark in the
  // table, so this needs no rule of its own ONCE the digit is promoted. Zeroing
  // the selector alone left the digit at its own advance.
  for (const cluster of ['1\ufe0f\u20e3', '9\ufe0f\u20e3', '#\ufe0f\u20e3', '*\ufe0f\u20e3']) {
    assert.ok(
      Math.abs(advanceWidthIn(cluster, 72) - EMOJI_EM) < 1e-9,
      `keycap ${cluster} charges ${advanceWidthIn(cluster, 72).toFixed(3)} em, not ${EMOJI_EM}`,
    );
  }
  // And the bare digit must NOT be promoted.
  assert.ok(Math.abs(advanceWidthIn('1', 72) - 0.556) < 1e-9);
});

test('a joiner welds on what follows it whatever plane it lives in', () => {
  // The staff of aesculapius in a health worker is U+2695, in the BMP. The old
  // joiner rule only absorbed astral code points, so this cluster was charged
  // the man AND the staff - 2.373 em against 1.373 drawn, 73% over, and an
  // over-charge is what withholds a name entirely.
  for (const cluster of [
    '\u{1f468}\u200d\u2695\ufe0f',
    '\u{1f469}\u200d\u2696\ufe0f',
    '\u{1f9d1}\u200d\u2708\ufe0f',
  ]) {
    assert.ok(
      Math.abs(advanceWidthIn(cluster, 72) - EMOJI_EM) < 1e-9,
      `${cluster} charges ${advanceWidthIn(cluster, 72).toFixed(3)} em against one glyph drawn`,
    );
  }
});

test('a variation selector does not weld on the character after it', () => {
  // A selector restyles what came BEFORE it. Absorbing what follows would make
  // every character after an emoji free, which silently shortens a whole name.
  const promoted = advanceWidthIn('\u2764\ufe0f', 72);
  assert.ok(
    Math.abs(advanceWidthIn('\u2764\ufe0fA', 72) - (promoted + advanceWidthIn('A', 72))) < 1e-9,
    'the letter after a variation selector must still be charged',
  );
});

test('the widest-glyph pair measures the cluster, not the base inside it', () => {
  // `widestGlyphIn` decides whether a column is wide enough to draw a run in.
  // Reading the base code point of a keycap reported 0.539 em for something the
  // width model draws at 1.373, sizing the column at 39% of what the line needs
  // - the same defect as measuring `box.label` and drawing something else.
  for (const cluster of ['1\ufe0f\u20e3', '\u2764\ufe0f', '\u{1f468}\u200d\u2695\ufe0f']) {
    for (const [name, measure] of [
      ['lower', widestGlyphLowerIn], ['upper', widestGlyphUpperIn],
    ] as const) {
      assert.ok(
        Math.abs(measure(cluster, 72) - EMOJI_EM) < 1e-9,
        `${name} bound reports ${measure(cluster, 72).toFixed(3)} for a glyph drawn at ${EMOJI_EM}`,
      );
    }
    // And the column test agrees with both: a column narrower than the glyph
    // cannot set it.
    assert.equal(drawableInColumn(cluster, 72, EMOJI_EM - 0.01), false);
    assert.equal(drawableInColumn(cluster, 72, 2 * EMOJI_EM), true);
  }
});

test('a promoted cluster counts as measured, not as an unmeasurable guess', () => {
  // The coverage oracle asked per code point, so it called a promoted heart
  // unmeasured forever: no table entry for U+2764 can ever read 1.373, because
  // that is not the code point's width. A rule that can never be satisfied is a
  // rule nobody can act on - and it drove `widestGlyphIn` to the zero lower
  // bound, which sizes the column at nothing.
  assert.equal(hasMeasuredAdvance('\u2764'), false, 'the bare dingbat is honestly untabled');
  assert.equal(hasMeasuredCluster('\u2764\ufe0f'), true);
  assert.equal(hasMeasuredCluster('1\ufe0f\u20e3'), true);
  assert.equal(hasMeasuredCluster('\u{1f468}\u200d\u2695\ufe0f'), true);
  // It must still be able to say no, or it is not an oracle.
  assert.equal(hasMeasuredCluster('\u{1d400}'), false);
  assert.equal(hasMeasuredCluster(''), false);
});

test('a joiner welds only what could belong to an emoji cluster', () => {
  // Removing the astral demand from the joiner clause was right for a health
  // worker and wrong for everything else. U+200D and U+2060 are ordinary text:
  // ZWJ forms Indic conjuncts and U+2060 WORD JOINER is the standard invisible
  // no-break character documentation tooling emits as &NoBreak;. The clause
  // absorbed the next code point WHATEVER IT WAS at zero width, and because the
  // joiner is itself absorbed the error compounds.
  const plain = advanceWidthIn('Contoso Platform', 10);
  const joined = advanceWidthIn('Contoso\u2060Platform', 10);
  assert.ok(
    joined >= plain - advanceWidthIn(' ', 10) - 1e-9,
    `a word joiner cost ${(plain - joined).toFixed(4)}in of a ${plain.toFixed(4)}in name`,
  );
  // Two Ms are two Ms however many joiners sit between them.
  const mm = advanceWidthIn('MM', 10);
  assert.ok(Math.abs(advanceWidthIn('M\u2060M', 10) - mm) < 1e-9, 'a word joiner halved "MM"');
  assert.ok(Math.abs(advanceWidthIn('M\u200dM', 10) - mm) < 1e-9, 'a ZWJ halved "MM"');
  // Sixteen Ms measured as one glyph, 93.8% under, before this.
  const sixteen = 'M'.repeat(16);
  const welded = [...sixteen].join('\u2060');
  assert.ok(
    Math.abs(advanceWidthIn(welded, 10) - advanceWidthIn(sixteen, 10)) < 1e-9,
    `sixteen joined Ms measure ${advanceWidthIn(welded, 10).toFixed(4)}in `
    + `against ${advanceWidthIn(sixteen, 10).toFixed(4)}in drawn`,
  );
  // And a Devanagari conjunct is text, not an emoji.
  const deva = '\u092a\u094d\u0930\u094b\u0921\u0915\u094d\u0936\u0928';
  const conjunct = '\u092a\u094d\u200d\u0930\u094b\u0921\u0915\u094d\u200d\u0936\u0928';
  assert.ok(
    Math.abs(advanceWidthIn(conjunct, 10) - advanceWidthIn(deva, 10)) < 1e-9,
    'two ZWJ conjuncts shrank a Devanagari name by a fifth',
  );
});

test('a variation selector promotes only a base the emoji font can draw', () => {
  // `promoted` fired for any non-CJK, non-space base, not only the code points
  // with a defined emoji variation sequence. A letter came out 204% over, and
  // an over-charge is what deletes a name.
  for (const base of ['z', 'A', '-', '@', '~']) {
    assert.ok(
      Math.abs(advanceWidthIn(`${base}\ufe0f`, 72) - advanceWidthIn(base, 72)) < 1e-9,
      `"${base}" with a variation selector charges `
      + `${advanceWidthIn(`${base}\ufe0f`, 72).toFixed(3)} against ${advanceWidthIn(base, 72).toFixed(3)} drawn`,
    );
    assert.equal(hasMeasuredCluster(`${base}\ufe0f`), true, 'a plain letter is tabled');
  }
  // The genuine emoji-variation bases still promote.
  for (const base of ['\u2764', '\u2601', '\u2699', '1']) {
    assert.ok(Math.abs(advanceWidthIn(`${base}\ufe0f`, 72) - EMOJI_EM) < 1e-9);
  }
});

test('the same visible name costs the same composed and decomposed', () => {
  // A combining mark carries no advance: it is drawn over the letter before it.
  // Both width models summed a cluster's code points, so the decomposed
  // spelling of a name that draws identically was billed up to 16.2% wider -
  // and because BOTH models made the same mistake, the gate's divergence rule
  // agreed with the exporter and stayed silent. NFD is not exotic: it is what
  // macOS filenames, Finder paths and many clipboards hand over.
  const names = [
    'Réseau privé partagé',
    'Passerelle sécurisée données clés partagées',
    'Ingestión de datos años atrás',
    'Contoso Zahlungsverkehr Prüfung München',
  ];
  for (const name of names) {
    const composed = advanceWidthIn(name.normalize('NFC'), 72);
    const decomposed = advanceWidthIn(name.normalize('NFD'), 72);
    assert.ok(
      Math.abs(composed - decomposed) < 1e-9,
      `"${name}" costs ${decomposed.toFixed(4)} decomposed against ${composed.toFixed(4)} composed`,
    );
  }
});

test('a spacing combining mark still costs its width', () => {
  // Mn and Me only, never Mc. A Devanagari matra is drawn BESIDE its base and
  // does take width; charging it nothing would understate the name, which is
  // the direction that paints outside the tile.
  const bare = '\u0915';
  const withMatra = '\u0915\u093e';
  assert.ok(
    advanceWidthIn(withMatra, 72) > advanceWidthIn(bare, 72),
    'a spacing matra was charged nothing',
  );
});

test('a byte order mark is not whitespace', () => {
  // U+FEFF is a member of JavaScript's own \s class, so a whitespace-first
  // ordering charged an invisible character a full space. Both models had the
  // identical ordering, so again neither could see it.
  assert.ok(
    Math.abs(advanceWidthIn('ab\ufeffcd', 72) - advanceWidthIn('abcd', 72)) < 1e-9,
    'a byte order mark bought width',
  );
});

test('a subdivision flag is one glyph, not seven', () => {
  // A tag sequence carries no joiner, so every clause that keeps a sequence
  // together missed it and all six TAG code points were charged the astral
  // face on top of the base flag: +600% for one drawn glyph.
  const flag = '\u{1f3f4}\u{e0067}\u{e0062}\u{e0073}\u{e0063}\u{e0074}\u{e007f}';
  const plain = '\u{1f3f4}';
  assert.ok(
    Math.abs(advanceWidthIn(flag, 72) - advanceWidthIn(plain, 72)) < 1e-9,
    `the Scottish flag costs ${advanceWidthIn(flag, 72).toFixed(4)} against `
    + `${advanceWidthIn(plain, 72).toFixed(4)} for the black flag it draws as`,
  );
});

test('truncation never strands an accent on the ellipsis', () => {
  // Over-charging does not merely mis-measure a name, it CUTS it, and a cut
  // between a letter and its accent does not lose the accent - it MOVES it
  // onto whatever now precedes it. The deck printed an acute stacked on the
  // ellipsis and nothing could see it, because every code point was still
  // there and every one of them had been priced.
  const marks = /[\p{Mn}\p{Me}]/u;
  const names = [
    'Réplication accélérée déléguée américaine réservée',
    'Réseau privé sécurisé européen partagé',
    'Contrôleur délégué prééminent',
    'Ingestión de datos años atrás en la región',
  ].map((name) => name.normalize('NFD'));
  for (const name of names) {
    for (let widthIn = 0.2; widthIn <= 3; widthIn += 0.01) {
      const drawn = fitLabelToWidth(name, widthIn, 9 / 72);
      const points = [...drawn];
      for (const [i, point] of points.entries()) {
        if (!marks.test(point)) continue;
        const before = i === 0 ? '' : points[i - 1];
        assert.ok(
          before !== '' && before !== '\u2026' && !/\s/.test(before),
          `at ${widthIn.toFixed(2)}in the deck draws ${JSON.stringify(drawn)}`,
        );
      }
    }
  }
});


// ---------------------------------------------------------------------------
// Canonical spelling. Round 66: the same visible name cost different widths
// depending on which normalisation form it was typed in, because a combining
// mark is priced at nothing and a precomposed letter is priced from the table.
// The fix is to compose at every entry point, so these are the counterfactual
// half of the proof - a property no single exported file can show, because a
// file only ever contains one of the two spellings.
// ---------------------------------------------------------------------------

test('a name costs the same however it was typed', () => {
  const names = [
    'Sisli sube sebeke sunucusu'.replace(/S/g, '\u015e').replace(/s(?=[iue])/g, '\u015f'),
    'Ba\u011flant\u0131 a\u011f ge\u00e7idi da\u011f\u0131t\u0131m',
    'D\u1ecbch v\u1ee5 l\u01b0u tr\u1eef \u0111\u1ed1i t\u01b0\u1ee3ng',
    'C\u1ed5ng k\u1ebft n\u1ed1i ri\u00eang t\u01b0',
  ];
  for (const name of names) {
    const nfc = singleLineName(name.normalize('NFC'));
    const nfd = singleLineName(name.normalize('NFD'));
    assert.equal(nfd, nfc, 'the two spellings must reach the drawing as one string');
    assert.equal(
      advanceWidthIn(nfd, 72).toFixed(4),
      advanceWidthIn(nfc, 72).toFixed(4),
      `${JSON.stringify(name)} is priced differently in the two forms`,
    );
  }
});

test('a name is cut at the same place however it was typed', () => {
  const name = 'Ba\u011flant\u0131 a\u011f ge\u00e7idi da\u011f\u0131t\u0131m';
  for (let widthIn = 0.2; widthIn <= 2.0; widthIn += 0.05) {
    const nfc = fitLabelToWidth(singleLineName(name.normalize('NFC')), widthIn, 9 / 72);
    const nfd = fitLabelToWidth(singleLineName(name.normalize('NFD')), widthIn, 9 / 72);
    assert.equal(nfd, nfc, `at ${widthIn.toFixed(2)}in the two spellings are cut differently`);
  }
});

test('the bare horn vowels are measured, and cost what their toned forms cost', () => {
  // Latin Extended-B, the one gap in the dump. Read off the toned forms in the
  // same table rather than measured afresh: a tone mark adds no advance, so
  // `u-horn` and `u-horn-acute` are the same width by construction, and taking
  // the number from anywhere else would make them differ.
  const pairs: Array<[string, string]> = [
    ['\u01a0', '\u1eda'], ['\u01a1', '\u1edb'],
    ['\u01af', '\u1ee8'], ['\u01b0', '\u1ee9'],
  ];
  for (const [bare, toned] of pairs) {
    assert.ok(hasMeasuredAdvance(bare), `U+${bare.codePointAt(0)!.toString(16)} has no measured advance`);
    assert.equal(
      advanceWidthIn(bare, 72).toFixed(4),
      advanceWidthIn(toned.normalize('NFC'), 72).toFixed(4),
      `U+${bare.codePointAt(0)!.toString(16)} is not priced like its toned form`,
    );
  }
});

test('a mark with no precomposed form still costs nothing', () => {
  // Yoruba: `e-dot-below` exists as one code point, `e-dot-below-grave` does
  // not, so NFC leaves the grave standing. This is the only way a combining
  // mark reaches the exporter now, and it is the case the zero-width rule
  // still has to hold for.
  const base = '\u1eb9';
  const marked = '\u1eb9\u0300';
  assert.equal(marked.normalize('NFC'), marked, 'the fixture must not compose away');
  assert.equal(
    advanceWidthIn(marked, 72).toFixed(4),
    advanceWidthIn(base, 72).toFixed(4),
    'a residual combining mark was charged an advance',
  );
});
