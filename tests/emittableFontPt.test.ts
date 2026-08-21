import test from 'node:test';
import assert from 'node:assert/strict';
import { emittableFontPt } from '../src/services/diagramExportGeometry.ts';
import { wrappedLineCount } from '../src/services/pptxExporter.ts';

/**
 * The exporter may not measure a font size the file cannot express.
 *
 * OOXML stores a run's size in `<a:rPr sz="...">` in HUNDREDTHS of a point and
 * pptxgenjs writes `round(pt * 100)`, so a size derived from a continuous
 * quantity — tile height, a shrink loop, a badge diameter — is not the size
 * PowerPoint draws. Every line count, block height and icon budget in the deck
 * is measured at the derived value, so the whole tile is laid out for type it
 * will never see.
 *
 * The corpus only catches this class when a real string happens to land within
 * a hundredth of a point of its column, which is why it took 186 scenarios and
 * a font change to surface once. These tests pin the invariant directly.
 */

const COLUMN_IN = 0.9903959;
const DERIVED_PT = 8.665766198453074;
const LABEL = 'Onyx configuration store';

test('emittable sizes land exactly on the hundredth-of-a-point grid the file stores', () => {
  for (const pt of [6, 7.0001, 8.665766198453074, 9.999999, 10.005, 13.3333333, 24.567891]) {
    const emitted = emittableFontPt(pt);
    const stored = Math.round(emitted * 100);
    assert.equal(stored / 100, emitted, `${pt}pt does not survive the file's own rounding`);
    assert.ok(emitted <= pt, `${pt}pt was rounded UP to ${emitted}pt`);
    assert.ok(pt - emitted < 0.01, `${pt}pt lost ${pt - emitted}pt, more than one grid step`);
  }
});

test('emittable sizes are idempotent and monotonic', () => {
  assert.equal(emittableFontPt(emittableFontPt(DERIVED_PT)), emittableFontPt(DERIVED_PT));
  assert.ok(emittableFontPt(8.66) <= emittableFontPt(8.67));
  assert.equal(emittableFontPt(7), 7);
  assert.equal(emittableFontPt(Number.NaN), 0);
  assert.equal(emittableFontPt(Number.POSITIVE_INFINITY), 0);
});

test('the size measured is the size drawn: the tile that wrapped a third line', () => {
  // The exact corpus failure. Measured at the derived size the label sets on
  // two lines with 0.0003in to spare, so the tile reserved a two-line band and
  // sized its icon around it; at the 8.67pt the file actually carried the same
  // line is 0.0001in too wide, wraps to three, and the third is painted 0.163in
  // below the tile.
  const derivedLines = wrappedLineCount(LABEL, COLUMN_IN, DERIVED_PT);
  const roundedToNearest = Math.round(DERIVED_PT * 100) / 100;
  assert.equal(roundedToNearest, 8.67);
  assert.ok(
    wrappedLineCount(LABEL, COLUMN_IN, roundedToNearest) > derivedLines,
    'the regression case no longer distinguishes rounding directions',
  );

  const emitted = emittableFontPt(DERIVED_PT);
  assert.equal(emitted, 8.66);
  assert.ok(
    wrappedLineCount(LABEL, COLUMN_IN, emitted) <= derivedLines,
    'the emitted size wraps to more lines than the size the tile was measured at',
  );
});

test('rounding down never costs a line anywhere on the grid', () => {
  // The other half of the contract: a size the file can express must never
  // measure WORSE than the continuous one it came from, at any column. Rounding
  // to nearest fails this for half of all inputs by construction.
  const labels = [
    'Onyx configuration store',
    'Azure Kubernetes Service',
    'Payments reconciliation function app',
    '仮想ネットワークゲートウェイ',
  ];
  for (const label of labels) {
    for (let column = 0.3; column <= 1.6; column += 0.1) {
      for (let pt = 7; pt < 14; pt += 0.137) {
        const emitted = emittableFontPt(pt);
        assert.ok(
          wrappedLineCount(label, column, emitted) <= wrappedLineCount(label, column, pt),
          `"${label}" at ${pt}pt in ${column.toFixed(1)}in wraps wider once emitted as ${emitted}pt`,
        );
      }
    }
  }
});
