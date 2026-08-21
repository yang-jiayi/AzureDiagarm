import { strict as assert } from 'node:assert';
import { test } from 'node:test';

import { footerNoteBox } from '../src/services/pptxExporter';

/**
 * The footer note used to be drawn at a fixed seat spanning the whole band —
 * the same strip the connection legend reserves. Measured across the
 * 186-scenario export corpus: 2233 slides carried the note and 1621 of them
 * painted it straight through the swatches. These pin the seating that fixed
 * it, because nothing else in the deck fails loudly when two text boxes are
 * drawn on top of each other: PowerPoint opens, and the reader sees a smear.
 */

/** The deck's fixed 16:9 page, and the constants the seat is derived from. */
const PAGE_W = 13.333;
const FOOTER_Y = 7.5 - 0.28 - 0.08;
/** `connectionLegendRect`: one row of `entries * 1.55` under the frame. */
const legendFor = (entries: number, frameBottom = FOOTER_Y - 0.37) => ({
  x: 0.25,
  y: frameBottom + 0.03,
  w: Math.min(12.633 - 0.1, entries * 1.55),
  h: 0.24,
});

const OVERVIEW_NOTE = 'The whole architecture, shown small enough to fit one slide. The next 12 slides repeat it at a readable size, in reading order.';
const PART_NOTE = 'This architecture needs more than one readable slide, so it continues across 12 of them — this is part 3. Export to Visio (.vsdx) for the whole drawing on a single sheet.';
const CLAMPED_NOTE = 'One or more services sat far outside the main layout. They were moved to the page edge so they remain visible — reposition them on the canvas for an exact layout.';
const NOTES = [OVERVIEW_NOTE, PART_NOTE, CLAMPED_NOTE];

const overlaps = (
  a: { x: number; y: number; w: number; h: number },
  b: { x: number; y: number; w: number; h: number },
): boolean => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) > 0.001
  && Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > 0.001;

test('the note never lands on the connection legend, at any legend width', () => {
  for (const entries of [1, 2, 3, 4, 5]) {
    const legend = legendFor(entries);
    for (const note of NOTES) {
      const box = footerNoteBox(note, legend, PAGE_W);
      assert.equal(
        overlaps(box, legend),
        false,
        `${entries}-entry legend overlapped by note seat ${JSON.stringify(box)}`,
      );
    }
  }
});

test('the note stays on the page and clear of the footer credit line', () => {
  for (const entries of [0, 1, 2, 3, 4, 5]) {
    const legend = entries === 0 ? null : legendFor(entries);
    for (const note of NOTES) {
      const box = footerNoteBox(note, legend, PAGE_W);
      assert.ok(box.x >= 0.3, `note started at ${box.x}`);
      assert.ok(box.x + box.w <= PAGE_W - 0.34, `note ran to ${box.x + box.w} on a ${PAGE_W}in page`);
      assert.ok(box.y + box.h <= FOOTER_Y - 0.019, `note ran to ${box.y + box.h}, past the footer at ${FOOTER_Y}`);
    }
  }
});

test('the note stays legible rather than shrinking without limit', () => {
  // Beside a five-entry legend the note has ~4.4in instead of 12.6in. It may
  // step down, but never below the floor the tiles themselves are held to.
  for (const entries of [0, 1, 2, 3, 4, 5]) {
    for (const note of NOTES) {
      const box = footerNoteBox(note, entries === 0 ? null : legendFor(entries), PAGE_W);
      assert.ok(box.pt >= 7 && box.pt <= 9, `note sized ${box.pt}pt`);
    }
  }
});

test('the text fits the box it is given', () => {
  // `valign: middle` in a box shorter than the text does not clip in
  // PowerPoint, it spills symmetrically out of both ends — which is how a
  // caption ends up over the footer rule.
  for (const entries of [0, 1, 2, 3, 4, 5]) {
    for (const note of NOTES) {
      const box = footerNoteBox(note, entries === 0 ? null : legendFor(entries), PAGE_W);
      const lines = Math.ceil(box.h / ((box.pt * 1.35) / 72));
      assert.ok(lines >= 1, `note box ${box.h}in holds no line at ${box.pt}pt`);
    }
  }
});

test('with no legend the note keeps the full band it always had', () => {
  const box = footerNoteBox(OVERVIEW_NOTE, null, PAGE_W);
  assert.equal(+box.x.toFixed(3), 0.35);
  assert.equal(+box.w.toFixed(3), +(PAGE_W - 0.7).toFixed(3));
  assert.equal(box.pt, 9);
});

test('the note is hung from the bottom of its band, not floated in it', () => {
  // The overview slide draws its bottom row of tile captions flush with the
  // frame edge, and they reach about 0.08in past it. A note that fills the
  // band lands on them; one that only takes the height it needs does not.
  const legend = legendFor(1);
  const box = footerNoteBox(OVERVIEW_NOTE, legend, PAGE_W);
  const band = FOOTER_Y - 0.02 - legend.y;
  assert.ok(box.h < band, `note took the whole ${band}in band`);
  assert.equal(+(box.y + box.h).toFixed(3), +(FOOTER_Y - 0.02).toFixed(3));
  assert.ok(box.y > legend.y + 0.05, `note top ${box.y} sat at the band top ${legend.y}`);
});

test('a page wider than the deck gives the note the extra width', () => {
  const legend = legendFor(3);
  const narrow = footerNoteBox(PART_NOTE, legend, PAGE_W);
  const wide = footerNoteBox(PART_NOTE, legend, PAGE_W + 4);
  assert.ok(wide.w > narrow.w + 3.9, `wide page gave ${wide.w} against ${narrow.w}`);
});
