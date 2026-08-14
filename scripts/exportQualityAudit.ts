/**
 * Objective quality audit for the native-shape exports (PPTX + VSDX).
 *
 * Office formats cannot be rendered head-less here, so quality is measured
 * from the emitted shape XML: every shape's geometry, text and font size are
 * parsed back out and scored against legibility rules that mirror what a human
 * sees when they open the deck.
 *
 * Run: npx tsx scripts/exportQualityAudit.ts
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import path from 'node:path';
import { Worker } from 'node:worker_threads';
import JSZip from 'jszip';

// The index row a service gets when its tile is drawn but carries no caption.
// Shared with the exporter by value; both sides must name the same string or the
// mark rules stop recognising the row and start reporting it as a lost mark.
const UNLABELLED_ROW = '(drawn unlabelled)';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx, buildArchitectureDeckPptx, calloutPlanFor } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';
import { buildVsdxPackage, calloutMagnificationFor } from '../src/services/visioVsdxExporter.ts';
import { WRAP_TRIGGER_RATIO } from '../src/utils/serpentineWrap.ts';

import { narrateEdgeCallouts, readEdgeLabel, CATEGORY_STYLES, singleLineName, advanceWidthIn } from '../src/services/diagramExportGeometry.ts';
import { readStepNumber } from '../src/utils/workflowStepMapping.ts';
import { stripXmlForbidden } from '../src/utils/xmlText.ts';

const OUT = path.join(process.cwd(), 'tmp-export-audit');
const EMU_PER_INCH = 914400;
/** Layout pixels per inch — the scale both exporters draw the canvas at. */
const PX_PER_IN = 96;
/** The standard 16:9 slide both decks start from, before any page growth. */
const BASE_SLIDE_W_IN = 13.333;
const BASE_SLIDE_H_IN = 7.5;
/**
 * Real per-character advances for Yu Gothic UI, in em, printable ASCII 33–126.
 *
 * Measured from the installed font with GDI+ (`Graphics.MeasureString`,
 * `GenericTypographic`, 20 repeats of each glyph at 100pt to divide out the
 * fitting error). This is the only number in this file that does not come from
 * the exporter's own model, and it exists because a shared estimator makes a
 * width error structurally unobservable: `auditTextWidthIn` is a
 * character-for-character copy of `estimateTextWidthIn`, so the two agree about
 * every mistake either of them makes.
 *
 * What that hid: both give every non-CJK character 0.54 em, Segoe UI's average
 * *lowercase* advance. Used as a maximum it understates `@` (0.955), `W`
 * (0.934) and `M` (0.898) by up to 77%, so the "does one letter fit?" guard
 * passed boxes that hold no capital at all.
 */
const YU_GOTHIC_ADVANCE_EM = [
  0.284, 0.392, 0.591, 0.539, 0.818, 0.800, 0.230, 0.302, 0.302, 0.417,
  0.684, 0.217, 0.400, 0.217, 0.390, 0.539, 0.539, 0.539, 0.539, 0.539,
  0.539, 0.539, 0.539, 0.539, 0.539, 0.217, 0.217, 0.684, 0.684, 0.684,
  0.448, 0.955, 0.645, 0.573, 0.619, 0.701, 0.506, 0.488, 0.686, 0.710,
  0.266, 0.357, 0.580, 0.471, 0.898, 0.748, 0.754, 0.560, 0.754, 0.598,
  0.531, 0.524, 0.687, 0.621, 0.934, 0.590, 0.553, 0.570, 0.302, 0.539,
  0.302, 0.684, 0.415, 0.268, 0.509, 0.588, 0.462, 0.589, 0.523, 0.313,
  0.589, 0.566, 0.242, 0.242, 0.497, 0.242, 0.861, 0.566, 0.586, 0.588,
  0.589, 0.348, 0.424, 0.339, 0.566, 0.479, 0.723, 0.459, 0.484, 0.452,
  0.302, 0.239, 0.302, 0.684,
];

/**
 * Advances for the non-ASCII characters the exporters actually emit, in em,
 * measured the same way as the table above.
 *
 * The old `: 0.54` fallback was a LOWER bound for anything unusual, which is
 * the wrong direction for a rule that asks whether text fits. The ellipsis is
 * the one that mattered: `fitLabelToLines` appends it at every truncation
 * point, 249 times across the corpus, and it advances 0.733 em rather than
 * 0.54.
 */
const YU_GOTHIC_EXTRA_EM: Record<string, number> = {
  '\u00a0': 0.274,
  '\u00b7': 0.217,
  '\u00d7': 0.684,
  '\u2013': 0.5,
  '\u2014': 1,
  '\u2018': 0.229,
  '\u2019': 0.229,
  '\u201c': 0.377,
  '\u201d': 0.377,
  '\u2026': 0.733,
  '\u2190': 1,
  '\u2192': 1,
  '\u2194': 1,
  '\u21d2': 1,
  '\u2212': 0.684,
  '\u2022': 0.406,
};

/**
 * A space advances 0.274 em. Zero is only correct for the whitespace that ENDS
 * a line, which `auditLineWidths` discounts explicitly at the point a renderer
 * discounts it.
 */
const AUDIT_SPACE_EM = 0.274;

/**
 * Latin-1 Supplement and Latin Extended-A, U+00A1 to U+017F, measured from the
 * installed font with GDI+ - the gate's own copy, taken from its own run.
 *
 * These fell through to the 1 em unknown fallback, which is 91% over for `é`.
 * That made the gate agree with the exporter that a 0.190in column could not
 * hold `Réseau privé sécurisé`, when the widest glyph in that string is an `R`
 * at 0.598 em and the column held it comfortably. A shared blind spot reads
 * exactly like a passing test.
 */
const YU_GOTHIC_LATIN_EM = [
  0.284, 0.539, 0.539, 0.556, 0.539, 0.239, 0.448, 0.414, 0.89, 0.392,
  0.506, 0.684, 0, 0.89, 0.415, 0.377, 0.684, 0.366, 0.366, 0.282,
  0.577, 0.458, 0.217, 0.205, 0.351, 0.431, 0.506, 0.906, 0.931, 0.952,
  0.448, 0.645, 0.645, 0.645, 0.645, 0.645, 0.645, 0.86, 0.619, 0.506,
  0.506, 0.506, 0.506, 0.266, 0.266, 0.266, 0.266, 0.701, 0.748, 0.754,
  0.754, 0.754, 0.754, 0.754, 0.684, 0.754, 0.687, 0.687, 0.687, 0.687,
  0.553, 0.56, 0.544, 0.509, 0.509, 0.509, 0.509, 0.509, 0.509, 0.832,
  0.462, 0.523, 0.523, 0.523, 0.523, 0.242, 0.242, 0.242, 0.242, 0.559,
  0.566, 0.586, 0.586, 0.586, 0.586, 0.586, 0.684, 0.586, 0.566, 0.566,
  0.566, 0.566, 0.484, 0.588, 0.484, 0.65, 0.554, 0.65, 0.554, 0.65,
  0.554, 0.683, 0.519, 0.683, 0.519, 0.683, 0.519, 0.683, 0.519, 0.744,
  0.672, 0.744, 0.578, 0.618, 0.555, 0.618, 0.555, 0.618, 0.555, 0.618,
  0.555, 0.618, 0.555, 0.728, 0.547, 0.728, 0.547, 0.728, 0.547, 0.728,
  0.547, 0.746, 0.58, 0.746, 0.582, 0.298, 0.266, 0.298, 0.266, 0.297,
  0.266, 0.297, 0.266, 0.298, 0.266, 0.676, 0.542, 0.4, 0.296, 0.691,
  0.531, 0.531, 0.596, 0.266, 0.596, 0.266, 0.596, 0.358, 0.596, 0.354,
  0.596, 0.265, 0.749, 0.58, 0.749, 0.58, 0.749, 0.58, 0.641, 0.749,
  0.58, 0.732, 0.559, 0.732, 0.559, 0.732, 0.559, 0.931, 0.928, 0.671,
  0.366, 0.671, 0.366, 0.671, 0.366, 0.628, 0.507, 0.628, 0.507, 0.628,
  0.507, 0.531, 0.424, 0.626, 0.351, 0.626, 0.406, 0.626, 0.35, 0.739,
  0.58, 0.739, 0.58, 0.739, 0.58, 0.739, 0.58, 0.739, 0.58, 0.74,
  0.58, 0.93, 0.774, 0.65, 0.489, 0.553, 0.612, 0.455, 0.612, 0.455,
  0.57, 0.452, 0.316,
];

/**
 * Everything the label font draws above U+017F, and the scripts it does not
 * draw at all - the gate's own copy, from the gate's own run.
 *
 * The round-56 table stopped at U+017F, so Cyrillic, Greek, IPA, combining
 * marks and the punctuation blocks were still charged a flat one em by BOTH
 * sides. Both agreeing to guess is the failure mode this file exists to avoid:
 * a Cyrillic name measured twice its ink, the exporter refused to draw it, and
 * the gate confirmed the refusal.
 *
 * Independence here is in the measurement, not in the typing. The first table
 * is re-measured through GDI+ string measurement, a different API from the
 * font metric table the exporter reads, at 200px over 20 repeats: of 533 code
 * points the two APIs disagree about 4, none by more than 0.06 em, and those
 * four are recorded as defects rather than smoothed over. The second table
 * cannot be independent in the same way - the question there is WHICH font the
 * renderer substitutes, not what that font measures - so it is deliberately
 * the same source, and its uncertainty is stated where it is defined.
 */
const YU_GOTHIC_WIDE_EM: ReadonlyArray<readonly [number, readonly number[]]> = [
  [0x192, [0.539, 0.728]],
  [0x1c2, [0.434]],
  [0x1cd, [0.65, 0.554, 0.298, 0.266, 0.732, 0.559, 0.739, 0.58, 0.739, 0.58, 0.739, 0.58, 0.739, 0.58, 0.739, 0.58]],
  [0x1f5, [0.547]],
  [0x1f8, [0.749, 0.58, 0.65, 0.554, 1.012, 0.88, 0.732, 0.559]],
  [0x218, [0.628, 0.507, 0.626, 0.351]],
  [0x237, [0.296]],
  [0x250, [0.554, 0.582, 0.587, 0.576, 0.519, 0.519, 0.577, 0.577, 0.555, 0.555, 0.748, 0.51, 0.51]],
  [0x25e, [0.57, 0.3, 0.577, 0.577, 0.556, 0.498, 0.488, 0.58, 0.578, 0.578, 0.299]],
  [0x26a, [0.265]],
  [0x26c, [0.376, 0.265, 0.605, 0.861, 0.861, 0.861, 0.58, 0.58, 0.583, 0.559, 0.768]],
  [0x278, [0.604, 0.366, 0.366, 0.366]],
  [0x27d, [0.366, 0.353]],
  [0x280, [0.521, 0.521, 0.507, 0.264, 0.3]],
  [0x288, [0.351, 0.598, 0.559, 0.551, 0.489, 0.774, 0.489, 0.493, 0.455, 0.506, 0.522]],
  [0x294, [0.502, 0.502]],
  [0x298, [0.732, 0.546]],
  [0x29c, [0.586, 0.332]],
  [0x29f, [0.447]],
  [0x2a1, [0.502, 0.502]],
  [0x2a4, [0.925]],
  [0x2b0, [0.427]],
  [0x2b2, [0.227]],
  [0x2b7, [0.539]],
  [0x2bb, [0.259, 0.259]],
  [0x2c1, [0.347]],
  [0x2c6, [0.371, 0.5, 0.213]],
  [0x2cc, [0.213]],
  [0x2d0, [0.284, 0.284]],
  [0x2d8, [0.5, 0.5, 0.5, 0.5, 0.337, 0.5, 0.267]],
  [0x2e0, [0.345, 0.211]],
  [0x2e5, [0.401, 0.401, 0.401, 0.401, 0.401]],
  [0x300, [0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]],
  [0x30a, [0.5, 0.5, 0.5]],
  [0x30f, [0.5]],
  [0x318, [0, 0, 0]],
  [0x31c, [0, 0, 0, 0, 0]],
  [0x324, [0, 0]],
  [0x327, [0.5, 0.5, 0, 0]],
  [0x32c, [0]],
  [0x32f, [0, 0]],
  [0x332, [0.5]],
  [0x334, [0]],
  [0x339, [0, 0, 0, 0, 0]],
  [0x361, [0]],
  [0x384, [0.273, 0.273, 0.645]],
  [0x388, [0.57, 0.774, 0.378]],
  [0x38c, [0.801]],
  [0x38e, [0.667, 0.826, 0.267, 0.645, 0.573, 0.472, 0.644, 0.506, 0.57, 0.71, 0.754, 0.266, 0.58, 0.629, 0.898, 0.748, 0.51, 0.754, 0.713, 0.56]],
  [0x3a3, [0.516, 0.524, 0.553, 0.754, 0.59, 0.776, 0.755, 0.266, 0.553, 0.614, 0.438, 0.575, 0.267, 0.553, 0.614, 0.548, 0.519, 0.584, 0.438, 0.442, 0.575, 0.586, 0.267, 0.524, 0.498, 0.577, 0.526, 0.445, 0.586, 0.627, 0.586, 0.46, 0.575, 0.487, 0.553, 0.699, 0.538, 0.75, 0.808, 0.267, 0.553, 0.586, 0.553, 0.808]],
  [0x3d0, [0.548, 0.591]],
  [0x3d5, [0.638]],
  [0x3db, [0.455]],
  [0x401, [0.506, 0.7, 0.472, 0.617, 0.531, 0.266, 0.266, 0.357, 0.981, 0.983, 0.723, 0.58, 0.749, 0.567, 0.709, 0.645, 0.572, 0.573, 0.472, 0.693, 0.506, 0.867, 0.54, 0.749, 0.749, 0.58, 0.673, 0.898, 0.71, 0.754, 0.713, 0.56, 0.619, 0.524, 0.567, 0.727, 0.59, 0.742, 0.661, 0.949, 0.98, 0.706, 0.783, 0.576, 0.616, 1.019, 0.591, 0.509, 0.579, 0.53, 0.383, 0.547, 0.523, 0.746, 0.446, 0.581, 0.581, 0.497, 0.527, 0.702, 0.577, 0.586, 0.577, 0.588, 0.462, 0.41, 0.484, 0.686, 0.459, 0.6, 0.565, 0.8, 0.824, 0.591, 0.71, 0.504, 0.462, 0.813, 0.503]],
  [0x451, [0.523, 0.577, 0.383, 0.462, 0.424, 0.242, 0.242, 0.242, 0.79, 0.807, 0.567, 0.497]],
  [0x45e, [0.484, 0.577]],
  [0x490, [0.469, 0.391]],
  [0x9f2, [0.54, 0.596]],
  [0xe3f, [0.679]],
  [0x17db, [0.549]],
  [0x1e3e, [0.919, 0.861]],
  [0x1e80, [0.93, 0.774, 0.93, 0.774, 0.93, 0.774]],
  [0x1ebc, [0.618, 0.555]],
  [0x1ef2, [0.65, 0.489]],
  [0x1f70, [0.582, 0.582, 0.51, 0.51]],
  [0x2002, [0.5]],
  [0x2011, [0.428, 0.5, 0.5]],
  [0x2018, [0.229, 0.229, 0.229, 0.259, 0.377, 0.377, 0.377, 0.424, 0.375, 0.375, 0.406]],
  [0x2026, [0.733]],
  [0x2030, [1.21]],
  [0x2039, [0.316, 0.316]],
  [0x203d, [0.527, 0.5, 0.477]],
  [0x2044, [0.076]],
  [0x2070, [0.361]],
  [0x2074, [0.361, 0.361, 0.361, 0.361, 0.361, 0.361]],
  [0x207f, [0.427, 0.361, 0.361, 0.361, 0.361, 0.361, 0.361, 0.361, 0.361, 0.361, 0.361]],
  [0x20a0, [0.683, 0.683, 0.683, 0.619, 0.578]],
  [0x20a6, [0.777, 0.99, 0.895, 0.93, 0.763, 0.578, 0.539]],
  [0x20ae, [0.626, 0.972, 0.683, 0.74]],
  [0x20b8, [0.631, 0.589, 0.615]],
  [0x20e3, [0]],
  [0x2116, [1.122]],
  [0x2122, [0.773]],
  [0x2126, [0.756, 0.756]],
  [0x212e, [0.832]],
  [0x2153, [0.798, 0.798, 0.798]],
  [0x215b, [0.798, 0.798, 0.798, 0.798]],
  [0x2206, [0.766]],
  [0x2209, [0.722]],
  [0x220f, [0.788]],
  [0x2219, [0.259]],
  [0x2225, [0.722, 0.722]],
  [0x223c, [0.704]],
  [0x2245, [0.722]],
  [0x2248, [0.722]],
  [0x2262, [0.722]],
  [0x2264, [0.722, 0.722]],
  [0x2276, [0.722, 0.722]],
  [0x2284, [0.722, 0.722]],
  [0x228a, [0.722, 0.722]],
  [0x229e, [0.814]],
  [0x22da, [0.722, 0.722]],
  [0x2305, [0.722, 0.722]],
  [0x2318, [0.924]],
  [0x2329, [0.5, 0.5]],
  [0x25ca, [0.632]],
  [0x2e40, [0.5]],
  [0xa7b5, [0.583]],
  [0xab53, [0.524]],
  [0xfb00, [0.622, 0.538, 0.538, 0.803, 0.803]],
  [0xff61, [0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]],
  [0xffe8, [0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5]],
];

const YU_GOTHIC_FALLBACK_EM: ReadonlyArray<readonly [number, readonly number[]]> = [
  // Latin Extended-B's two horn vowels, the one Latin gap in the dump. Read
  // off this table's own toned forms at 0x1ee8 and 0x1eda rather than measured
  // afresh, because a tone mark adds no advance and `ư` must not be priced
  // differently from `ứ`. See the twin note in `diagramExportGeometry.ts`.
  //
  // Devanagari is the sharp edge of what this table does NOT cover. Only the
  // digits at 0x966 are here; every consonant falls back to a flat 1.0 em -
  // and so does the exporter's table, from its own identical gap. So on an
  // Indic name the divergence rule is not merely silent, the two models AGREE
  // ON A GUESS, which is the one failure this file's whole two-model design
  // was meant to make impossible. Only the coverage rule stands between that
  // and a whole script, and the coverage rule reports rather than blocks.
  [0x1a0, [0.763, 0.597]],
  [0x1af, [0.708, 0.588]],
  [0x591, [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.4, 0, 0.202, 0, 0, 0.217, 0, 0, 0.399, 0]],
  [0x5d0, [0.637, 0.57, 0.439, 0.481, 0.678, 0.268, 0.337, 0.674, 0.681, 0.268, 0.559, 0.545, 0.551, 0.694, 0.674, 0.268, 0.399, 0.676, 0.605, 0.611, 0.631, 0.565, 0.585, 0.664, 0.559, 0.785, 0.726]],
  [0x5ef, [0.542, 0.522, 0.522, 0.522, 0.229, 0.377]],
  [0x600, [1.227, 0.832, 0.531, 1.522, 1.55, 0, 0.668, 0.668, 0.918, 0.751, 0.918, 0.585, 0.251, 0.334, 1.085, 0.585, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.251, 0.917, 0.61, 0.501, 0.501, 0.751, 0.417, 0.251, 0.251, 0.501, 0.251, 0.751, 0.251, 0.918, 0.501, 0.918, 0.918, 0.585, 0.585, 0.585, 0.501, 0.501, 0.417, 0.417, 1.169, 1.169, 1.336, 1.336, 0.835, 0.835, 0.501, 0.501, 0.918, 0.918, 0.751, 0.751, 0.751, 0.167]],
  [0x642, [0.751, 0.835, 0.668, 0.585, 0.668, 0.501, 0.501, 0.751, 0.751, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.417, 0.251, 0.501, 0.585, 0.417, 0.501, 0.501, 0.585, 0.585, 0.501, 0.585, 0.251, 0.251, 0.389, 0.918, 0.751, 0, 0.251, 0.251, 0.251, 0.251, 0.417, 0.501, 0.501, 0.751, 0.918, 0.918, 0.918, 0.918, 0.918, 0.918, 0.918, 0.918, 0.585, 0.585, 0.585, 0.585, 0.585, 0.585, 0.585, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.417, 0.417, 0.417, 0.417, 0.501, 0.417, 0.417, 0.417, 0.417, 1.169, 1.169, 1.169, 1.336, 1.336, 0.835, 0.501]],
  [0x6a7, [0.751, 0.751, 0.918, 1.085, 0.918, 0.835, 0.835, 0.835, 0.918, 0.918, 0.918, 0.918, 0.918, 0.918, 0.668, 0.668, 0.668, 0.668, 0.668, 0.668, 0.668, 0.668, 0.668, 0.751, 0.543, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.501, 0.751, 0.835, 0.751, 0.501, 0.751, 0.751, 0.918, 0.918, 0.251, 0.501, 0, 0, 0, 0, 0, 0, 0, 1.123, 1.085, 0, 0, 0, 0, 0, 0, 0.251, 0.417, 0, 0, 0.6, 0, 0, 0, 0, 0.501, 0.417, 0.417, 0.251, 0.501, 0.585, 0.501, 0.585, 0.501, 0.585, 0.585, 0.501, 1.169, 1.336, 0.501, 0.417, 0.585, 0.751]],
  [0x966, [0.519, 0.461, 0.547, 0.524, 0.588, 0.61, 0.613, 0.656, 0.58, 0.566]],
  [0xe01, [0.574, 0.603, 0.615, 0.575, 0.575, 0.619, 0.388, 0.509, 0.619, 0.601, 0.608, 0.831, 0.825, 0.598, 0.598, 0.487, 0.679, 0.849, 0.892, 0.575, 0.575, 0.574, 0.59, 0.51, 0.594, 0.569, 0.569, 0.598, 0.598, 0.667, 0.667, 0.598, 0.561, 0.524, 0.429, 0.574, 0.539, 0.598, 0.455, 0.576, 0.59, 0.539, 0.589, 0.669, 0.528, 0.538, 0.505, 0.342, 0, 0.455, 0.455, 0, 0, 0, 0, 0, 0, 0]],
  [0xe40, [0.266, 0.483, 0.409, 0.391, 0.399, 0.455, 0.444, 0, 0, 0, 0, 0, 0, 0, 0, 0.541, 0.71, 0.768, 0.783, 0.745, 0.725, 0.725, 0.645, 0.881, 0.765, 0.827, 0.646, 0.89]],
  [0x1e00, [0.645, 0.509, 0.573, 0.588, 0.573, 0.588, 0.573, 0.588, 0.619, 0.462, 0.701, 0.589, 0.701, 0.589, 0.701, 0.589, 0.701, 0.589, 0.701, 0.589, 0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.488, 0.313, 0.686, 0.589, 0.71, 0.566, 0.71, 0.566, 0.71, 0.566, 0.71, 0.566, 0.71, 0.566, 0.266, 0.242, 0.266, 0.242, 0.58, 0.497, 0.58, 0.497, 0.58, 0.497, 0.471, 0.242, 0.471, 0.242, 0.471, 0.242, 0.471, 0.242]],
  [0x1e40, [0.898, 0.861, 0.898, 0.861, 0.748, 0.566, 0.748, 0.566, 0.748, 0.566, 0.748, 0.566, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.56, 0.588, 0.56, 0.588, 0.598, 0.348, 0.598, 0.348, 0.598, 0.348, 0.598, 0.348, 0.531, 0.424, 0.531, 0.424, 0.531, 0.424, 0.531, 0.424, 0.531, 0.424, 0.524, 0.339, 0.524, 0.339, 0.524, 0.339, 0.524, 0.339, 0.687, 0.566, 0.687, 0.566, 0.687, 0.566, 0.687, 0.566, 0.687, 0.566, 0.621, 0.479, 0.621, 0.479]],
  [0x1e86, [0.934, 0.723, 0.934, 0.723, 0.59, 0.459, 0.59, 0.459, 0.553, 0.484, 0.57, 0.452, 0.57, 0.452, 0.57, 0.452, 0.566, 0.339, 0.723, 0.484, 0.509, 0.241, 0.313, 0.313, 0.62, 0.584, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.645, 0.509, 0.506, 0.523, 0.506, 0.523]],
  [0x1ebe, [0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.506, 0.523, 0.266, 0.242, 0.266, 0.242, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.754, 0.586, 0.763, 0.597, 0.763, 0.597, 0.763, 0.597, 0.763, 0.597, 0.763, 0.597, 0.687, 0.566, 0.687, 0.566, 0.708, 0.588, 0.708, 0.588, 0.708, 0.588, 0.708, 0.588, 0.708, 0.588]],
  [0x1ef4, [0.553, 0.484, 0.553, 0.484, 0.553, 0.484, 0.646, 0.408, 0.6, 0.516, 0.579, 0.484]],
];

/**
 * XML entities, undone.
 *
 * Every rule that harvests text back out of an emitted file reads it escaped:
 * a slide carrying `Backup &amp; Recovery` is correct OOXML for `Backup &
 * Recovery`, but a comparison against the authored label sees two different
 * strings and reports that the deck dropped the name. The measurement rules
 * are worse than the comparison ones - `&amp;` measures 2.975 em against `&`'s
 * 0.800, so a correct text block is declared 3.7x short of its own ink.
 *
 * The direction is always over-measure, so this class of bug produces false
 * RED rather than misses. On a project where the gate arbitrates every round,
 * a gate that cries defect is as expensive as one that stays quiet.
 */
function unescapeXml(text: string): string {
  return text
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_m, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/** Both tables above, flattened once into a lookup. */
/**
 * The Unicode spaces, in em. See the exporter's copy: the generator's
 * "skip anything already one em" filter is right for a glyph and wrong for
 * whitespace, whose fallback is the plain space width, so it dropped U+2003 and
 * U+3000 and left them charged 0.274 against a true 1.0.
 */
const YU_GOTHIC_SPACE_WIDE_EM: ReadonlyArray<readonly [number, number]> = [
  [0x2000, 0.5], [0x2001, 1], [0x2002, 0.5], [0x2003, 1],
  [0x2004, 1 / 3], [0x2005, 0.25], [0x2006, 1 / 6],
  [0x2007, YU_GOTHIC_ADVANCE_EM['0'.charCodeAt(0) - 33]],
  [0x2008, YU_GOTHIC_ADVANCE_EM['.'.charCodeAt(0) - 33]],
  [0x2009, 0.2], [0x200a, 0.1], [0x202f, 0.2], [0x205f, 4 / 18], [0x3000, 1],
];

const AUDIT_MEASURED_EM: ReadonlyMap<number, number> = (() => {
  const table = new Map<number, number>();
  for (const source of [YU_GOTHIC_WIDE_EM, YU_GOTHIC_FALLBACK_EM]) {
    for (const [start, values] of source) {
      values.forEach((value, offset) => table.set(start + offset, value));
    }
  }
  for (const [code, value] of YU_GOTHIC_SPACE_WIDE_EM) table.set(code, value);
  return table;
})();

/** Astral ideographs are a full em; the emoji blocks are what the font draws. */
const AUDIT_ASTRAL_CJK_MIN = 0x20000;
const AUDIT_ASTRAL_CJK_MAX = 0x3ffff;
const AUDIT_EMOJI_RE = /[\u{1f000}-\u{1faff}]/u;

/** Emoji advance, in em - the widest glyph in the substituted face. */
const AUDIT_EMOJI_EM = 1.373;

/** A joiner welds on what follows; a variation selector only restyles what came before. */
const AUDIT_JOINER_RE = /[\u200d\u2060]/;

/** VARIATION SELECTOR-16: draw the character before me as an emoji. */
const AUDIT_VS16 = '\ufe0f';

/** The BMP blocks the emoji font draws from. */
const AUDIT_EMOJI_BMP_RE = /[\u203c\u2049\u2122\u2139\u2194-\u21aa\u231a\u231b\u2328\u23cf\u23e9-\u23f3\u23f8-\u23fa\u24c2\u25aa\u25ab\u25b6\u25c0\u25fb-\u25fe\u2600-\u27bf\u2934\u2935\u2b00-\u2bff\u3030\u303d\u3297\u3299]/;

/** Unicode's own grapheme breaker, which is not this file's opinion about text. */
const AUDIT_GRAPHEMES = new Intl.Segmenter('en', { granularity: 'grapheme' });

/**
 * Drawn ON the code point before it, so it adds nothing to the run's advance.
 *
 * Asked of Unicode's general category rather than of a list, because the point
 * of this side of the gate is to reach an answer the exporter's tables were not
 * consulted about. Mn and Me are non-spacing and enclosing; a spacing mark (Mc)
 * is drawn beside its base and is charged. TAG characters are invisible by
 * definition and terminate a flag sequence.
 */
const AUDIT_COMBINING_RE = /[\p{Mn}\p{Me}]|[\u{E0020}-\u{E007F}]/u;

/** Can this code point be part of an emoji cluster at all? */
function auditEmojiCapable(character: string): boolean {
  const code = character.codePointAt(0) ?? 0;
  return code >= 0x10000 || AUDIT_EMOJI_BMP_RE.test(character);
}

/**
 * `text` split into rendered clusters, independently of the exporter's walker.
 *
 * A DIFFERENT ALGORITHM, not a second transcription of the same one. This used
 * to be the exporter's forward walk restated statement for statement - the same
 * `joined` disjunction, the same `promoted` predicate, the same absorb in the
 * same position - with only the advance tables differing. Two transcriptions of
 * one walk agree on every input INCLUDING every input on which the walk is
 * wrong, which is precisely the case a gate exists to catch, so the oracle
 * could not have disagreed about the joiner and variation-selector defects even
 * in principle. A gate rule and the code it checks must not be the same
 * algorithm twice.
 *
 * So the clustering is delegated to `Intl.Segmenter`, which is ICU's UAX #29
 * implementation and knows nothing about this repository: it breaks emoji ZWJ
 * sequences, regional-indicator pairs, skin-tone modifiers, Indic conjuncts and
 * combining sequences on the standard's terms. What is left here is only the
 * width question - is this grapheme drawn from the emoji face, and if not what
 * do its code points measure - which is the part the tables answer.
 */
function auditClusters(text: string): Array<{ text: string; em: number; measured: boolean }> {
  const clusters: Array<{ text: string; em: number; measured: boolean }> = [];
  for (const { segment } of AUDIT_GRAPHEMES.segment(text)) {
    const points = [...segment];
    const base = points[0] ?? '';
    // One glyph from the emoji face, whatever it took to spell it: a selector
    // over a base the face can draw, a keycap, a skin tone, a flag pair, or a
    // joined sequence. A joiner over a base the face CANNOT draw is text - a
    // Devanagari conjunct, or the word joiner documentation tooling emits - and
    // falls through to be measured code point by code point.
    const emoji = (points.includes(AUDIT_VS16) && (auditEmojiCapable(base) || /[0-9#*]/.test(base)))
      || points.some((p) => /[\u{1f3fb}-\u{1f3ff}]/u.test(p))
      || (points.length > 1
        && /[\u{1f1e6}-\u{1f1ff}]/u.test(base)
        && /[\u{1f1e6}-\u{1f1ff}]/u.test(points[1]))
      || (points.some((p) => AUDIT_JOINER_RE.test(p)) && auditEmojiCapable(base));
    if (emoji) {
      clusters.push({ text: segment, em: AUDIT_EMOJI_EM, measured: true });
      continue;
    }
    let em = 0;
    let measured = true;
    for (const point of points) {
      // A COMBINING MARK IS NOT A GLYPH. ICU was already being asked where the
      // clusters begin and then ignored about what they cost, so the same
      // visible name spelled in NFD was billed for every accent as though it
      // stood on its own: "Réseau privé partagé" measured 16.2% wider
      // decomposed than composed. Both models made the identical mistake, which
      // is why the divergence rule could not see it - a second implementation
      // that shares the first's pricing model is not a second opinion.
      if (AUDIT_COMBINING_RE.test(point)) continue;
      em += measuredAdvanceEm(point);
      if (!hasAuditAdvance(point)) measured = false;
    }
    clusters.push({ text: segment, em, measured });
  }
  return clusters;
}

/**
 * True when `character` has a measured advance rather than the fallback.
 *
 * The coverage rule is the gate's only honest oracle: it reports the characters
 * the width model is guessing at. It used to answer true for EVERYTHING above
 * the BMP, so the one rule written to catch a guess could never fire on the one
 * range documented as guessed - it certified it instead.
 */
function hasAuditAdvance(character: string): boolean {
  if (/\s/.test(character)) return true;
  if (/[\u200b-\u200f\u2060\ufe00-\ufe0f\ufeff]/.test(character)) return true;
  if (/[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]/.test(character)) return true;
  if (YU_GOTHIC_EXTRA_EM[character] !== undefined) return true;
  const code = character.codePointAt(0) ?? 0;
  if (code >= AUDIT_ASTRAL_CJK_MIN && code <= AUDIT_ASTRAL_CJK_MAX) return true;
  if (AUDIT_EMOJI_RE.test(character)) return true;
  if (code >= 0x10000) return false;
  if (code >= 0xa1 && code <= 0x17f) return true;
  if (AUDIT_MEASURED_EM.has(code)) return true;
  return code >= 33 && code <= 126;
}

/**
 * Measured advance of one character, in em. CJK is a full em by construction;
 * an unknown character is charged a full em, an UPPER bound, because a rule
 * that guesses low reports that text fits when it does not.
 *
 * Astral characters that are not ideographs are drawn by a substituted colour
 * font rather than by the label font, so no measurement of the label font can
 * settle them. 1.373 em is the widest glyph in the font Windows substitutes,
 * charged on purpose: a rule may safely over-reserve, never under-reserve.
 */
function measuredAdvanceEm(character: string): number {
  const code = character.codePointAt(0) ?? 0;
  // ZERO WIDTH IS TESTED FIRST, and the order is the whole point. U+FEFF is a
  // member of JavaScript's own \s class, so asking "is it whitespace?" first
  // answered yes for a character that has no advance at all and charged a byte
  // order mark a full space. A BOM is what a UTF-8 file pasted into a name
  // brings with it, it is invisible on the slide, and it was buying width.
  if (/[\u200b-\u200f\u2060\ufe00-\ufe0f\ufeff]/.test(character)) return 0;
  if (/\s/.test(character)) return AUDIT_MEASURED_EM.get(code) ?? AUDIT_SPACE_EM;
  if (/[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]/.test(character)) return 1;
  const extra = YU_GOTHIC_EXTRA_EM[character];
  if (extra !== undefined) return extra;
  if (code >= AUDIT_ASTRAL_CJK_MIN && code <= AUDIT_ASTRAL_CJK_MAX) return 1;
  if (code >= 0x10000) return 1.373;
  if (code >= 0xa1 && code <= 0x17f) return YU_GOTHIC_LATIN_EM[code - 0xa1];
  const measured = AUDIT_MEASURED_EM.get(code);
  if (measured !== undefined) return measured;
  return code >= 33 && code <= 126 ? YU_GOTHIC_ADVANCE_EM[code - 33] : 1;
}

/** The width of the whitespace `text` ends with, in inches. */
function measuredTrailingWsIn(text: string, fontSizePt: number): number {
  const trimmed = text.replace(/\s+$/, '');
  if (trimmed.length === text.length) return 0;
  let em = 0;
  for (const character of text.slice(trimmed.length)) em += measuredAdvanceEm(character);
  return (em * fontSizePt) / 72;
}

/**
 * Widest character in `text`, in inches, from the measured table.
 *
 * An unmeasured character is charged NOTHING here. This is the one measurement
 * in the gate that decides whether the exporter was entitled to withhold a
 * name, and a rule that guesses high there agrees with the exporter that a
 * column was too narrow when it was not. There is no positive lower bound to
 * use: a combining mark advances 0 and U+2044 advances 0.076, so the 0.205
 * this carried was a bound over the sample rather than over the repertoire.
 */
function measuredWidestGlyphIn(text: string, fontSizePt: number): number {
  let widest = 0;
  for (const cluster of auditClusters(text)) {
    // Whitespace is not a glyph: it advances, but a column that holds only a
    // space holds no ink.
    if (/^\s*$/.test(cluster.text)) continue;
    widest = Math.max(widest, cluster.measured ? cluster.em : 0);
  }
  return (widest * fontSizePt) / 72;
}

/**
 * Whether `text` is worth drawing in a column `columnIn` wide - the gate's own
 * copy of the exporter's two-clause test.
 *
 * Both clauses take the LOWER bound on an unmeasured character. They used to
 * disagree, the widest-glyph clause charging nothing and the mean clause a full
 * em, so a name in an untabled script was refused by a column more than twice
 * wide enough for it.
 */
function measuredDrawableInColumn(text: string, fontSizePt: number, columnIn: number): boolean {
  const glyphs = auditClusters(text).filter((cluster) => !/^\s*$/.test(cluster.text));
  if (glyphs.length === 0) return false;
  if (columnIn < measuredWidestGlyphIn(text, fontSizePt)) return false;
  let em = 0;
  for (const glyph of glyphs) em += glyph.measured ? glyph.em : 0;
  return columnIn >= (2 * em * fontSizePt) / 72 / glyphs.length;
}

/** Measured width of a run, in inches, from the real advance table. */
function measuredTextWidthIn(text: string, fontSizePt: number): number {
  let em = 0;
  for (const cluster of auditClusters(text)) em += cluster.em;
  return (em * fontSizePt) / 72;
}

/**
 * Lines a run takes in a column, wrapped the way a renderer wraps it, using
 * measured advances rather than the exporter's flat 0.54 em.
 *
 * `auditWrappedLines` shares its width model with `estimateTextWidthIn`, so it
 * agrees with the exporter about every string the exporter mis-measures — and
 * 0.54 em/char under-measures "HTTPS" (real: 0.710/0.524/0.524/0.560/0.531) by
 * a fifth, which is a whole extra line in a narrow column. This is the string
 * -width half of the independent oracle; `measuredWidestGlyphIn` was the
 * glyph half.
 */
function measuredWrappedLines(text: string, widthIn: number, fontSizePt: number): number {
  if (!text) return 1;
  const box = Math.max(0.001, widthIn);
  let lines = 0;
  for (const paragraph of text.split(/\r\n|\r|\n/)) {
    const runs = paragraph.split(/(?<=[\s\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6])/);
    let rows = 1;
    let used = 0;
    for (const run of runs) {
      const width = measuredTextWidthIn(run, fontSizePt);
      // Fit on visible ink, advance by the full width: a renderer lets the
      // run-final spaces hang past the column.
      const visible = width - measuredTrailingWsIn(run, fontSizePt);
      if (used > 0 && used + visible > box) { rows += 1; used = 0; }
      if (width > box) {
        // Break inside the run, one character at a time, because that is what
        // the renderer does with a word wider than its column.
        let row = 0;
        for (const character of run) {
          const advance = (measuredAdvanceEm(character) * fontSizePt) / 72;
          if (row > 0 && row + advance > box) { rows += 1; row = 0; }
          row += advance;
        }
        used = row;
        continue;
      }
      used += width;
    }
    lines += Math.max(1, rows);
  }
  return Math.max(1, lines);
}

/**
 * How wide a run of text is at a point size, measured here rather than imported.
 *
 * The exporter has its own copy of this arithmetic, and a rule that shares the
 * exporter's estimator cannot catch the exporter mis-estimating — it would agree
 * with the bug.
 *
 * That independence was bought with a flat 0.54 em per non-CJK character, and
 * the price turned out to be accuracy: 0.54 is Yu Gothic UI's average LOWERCASE
 * advance, so it over-states `i` (0.242 measured) by more than a factor of two,
 * under-states `W` (0.934) by 42%, and charges a full 0.54 for a space that
 * advances 0. Rules built on it reported 770 phantom overlaps and 142 phantom
 * short labels the moment the exporters started measuring for real.
 *
 * It now measures with `measuredAdvanceEm`, the per-glyph table above. That
 * keeps the independence where it actually matters: the table is hard-coded
 * HERE rather than imported from `diagramExportGeometry`, so if the shared
 * table drifts these rules disagree; and the wrap below is still a separate
 * implementation from the exporter's, which is how the mid-word `ceil(w / box)`
 * under-count was caught.
 */
/** Whitespace of any shape, reduced to single spaces, for text comparisons. */
function collapseWs(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

function auditTextWidthIn(text: string, fontSizePt: number): number {
  return measuredTextWidthIn(text, fontSizePt);
}
/**
 * How many lines a run takes in a box, wrapping the way a renderer wraps.
 *
 * `ceil(width / box)` is a lower bound, not a count: it assumes text can break
 * anywhere. Latin prose breaks between words and abandons the rest of a line
 * whenever the next word will not fit, so a name made of several long tokens
 * takes more lines than the ratio predicts — which is precisely the case where
 * a table measured onto the page prints below it. Written out here rather than
 * imported, for the same reason as the estimator above.
 *
 * Counting is a view of the walk below, not a second walk. Two implementations
 * of one wrap — however carefully kept in step — is the arrangement that has
 * hidden every wrapping defect this audit has ever missed.
 */
function auditWrappedLines(text: string, boxIn: number, fontSizePt: number): number {
  if (!text) return 1;
  return auditLineWidths(text, boxIn, fontSizePt).length;
}

/**
 * The narrowest column the audit will pretend a box has.
 *
 * One number, shared by the wrap and the column measurement, because two
 * different floors let a real defect through by cancelling out: at 0.1 the
 * wrap counted two lines for a caption PowerPoint stacks into four, and at
 * 0.05 the column said the ink did not fit — the containment rule saw a
 * caption that fits perfectly and the physics rule saw a violation, and the
 * only reason the class was visible at all was that the two disagreed.
 *
 * Small enough to be under any real column, because a box narrower than a
 * glyph is exactly the case being measured: PowerPoint does not clip, it puts
 * one character on each line and paints the rest below the box.
 */
const MIN_TEXT_COLUMN_IN = 0.01;

/**
 * The ink width of each line the wrap produces.
 *
 * A box's width is what it *may* use; the widest of these is what it *does*
 * use, and the difference is the padding that decides whether two overlapping
 * boxes are actually two overlapping sentences.
 *
 * Hard breaks are lines: both renderers start a new paragraph at a newline, and
 * splitting on whitespace alone only ends a run at one.
 */
function auditLineWidths(text: string, boxIn: number, fontSizePt: number): number[] {
  const box = Math.max(MIN_TEXT_COLUMN_IN, boxIn);
  const widths: number[] = [];
  for (const paragraph of String(text ?? '').split(/\r\n|\r|\n/)) {
    const tokens = paragraph.split(/(?<=[\s\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6])/);
    let used = 0;
    for (const token of tokens) {
      const w = auditTextWidthIn(token, fontSizePt);
      // Fit on visible ink, advance by the full width: run-final spaces hang
      // past the column rather than wrapping the line.
      const visible = w - measuredTrailingWsIn(token, fontSizePt);
      if (used > 0 && used + visible > box) {
        widths.push(used);
        used = 0;
      }
      if (w > box) {
        // A single run wider than the box breaks inside itself, one CHARACTER
        // at a time. `ceil(w / box)` assumes a word packs exactly a boxful per
        // line, which is only true if a break may fall part-way through a
        // glyph; breaks fall between glyphs, so the ratio is a lower bound.
        let row = 0;
        for (const character of token) {
          const advance = (measuredAdvanceEm(character) * fontSizePt) / 72;
          if (row > 0 && row + advance > box) { widths.push(row); row = 0; }
          row += advance;
        }
        used = row;
        continue;
      }
      used += w;
    }
    widths.push(used);
  }
  return widths;
}
/**
 * Chrome the exporter adds around the drawing, in inches.
 *
 * Not a guess: `visioVsdxExporter.ts` pads the sheet by `PAGE_PADDING_IN` on
 * every side, so the slack is exactly 1.2in on both axes for any drawing large
 * enough that the 11x8.5in minimums do not bind — measured at 1.20 on a
 * 31x3in chain and a 12x15in estate alike. A looser figure hides real
 * inflation: at 4in a mis-packed one- or two-lane stray strip fits underneath
 * the allowance and the sheet-size invariant never fires.
 */
const PAGE_CHROME_SLACK_IN = 1.5;

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const PIXEL_PNG_BYTES = Uint8Array.from(
  Buffer.from(PIXEL_PNG.slice(PIXEL_PNG.indexOf(',') + 1), 'base64'),
);

interface Shape {
  name: string;
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  /** Drawn text, one entry per `<a:p>`. Evidence, not a calculation. */
  paragraphs: string[];
  fontSize: number | null;
  /** Vertical anchor of the text body: 't', 'ctr', or 'b'. */
  anchor: string | null;
  /** Shape fill as RRGGBB, or null when the shape declares no fill. */
  fill: string | null;
  /** Fill opacity, 0..1. A translucent chip shows what is underneath it. */
  fillAlpha: number;
  /** Every drawn text run, with the colour and size it is drawn at. */
  runs: { color: string | null; sizePt: number; bold: boolean; text: string }[];
  /**
   * Left plus right text inset in inches, read from `<a:bodyPr>` rather than
   * assumed. PowerPoint's default is 0.1in a side, but a shape that sets
   * `margin: 0` gets its full width as a column — assuming the default there
   * shrinks the modelled column by 0.2in and hands back a line count that is
   * too low on exactly the shapes that need the most care.
   */
  insetX: number;
  /** Top plus bottom text inset in inches, read from `<a:bodyPr>`. */
  insetY: number;
  /** True when the body declares `wrap="none"`: one line however wide. */
  wrapNone: boolean;
  /**
   * Paragraph alignment, read from `<a:pPr algn="…">`. Ink is not always in the
   * middle of its box: a left-aligned caption in a band twice its width paints
   * only the left half, so measuring a collision against a centred rect both
   * misses hits on the left and invents them on the right.
   */
  align: string | null;
  /**
   * Declared line spacing as a multiple, or null when the shape states none.
   * Read so a box sized at the spacing it asked for is measured at it.
   */
  lineSpacing: number | null;
  path?: { x: number; y: number }[];
}

/** sRGB relative luminance, per WCAG 2.1. */
function luminance(hex: string): number {
  const v = [0, 2, 4].map((i) => {
    const c = parseInt(hex.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
}

/** WCAG contrast ratio between two RRGGBB colours, 1..21. */
function contrastRatio(a: string, b: string): number {
  const la = luminance(a);
  const lb = luminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/** What a translucent fill actually looks like over what is behind it. */
function blend(fg: string, bg: string, alpha: number): string {
  if (alpha >= 1) return fg;
  const mix = (i: number) => {
    const f = parseInt(fg.slice(i, i + 2), 16);
    const b = parseInt(bg.slice(i, i + 2), 16);
    return Math.round(f * alpha + b * (1 - alpha))
      .toString(16)
      .padStart(2, '0');
  };
  return `${mix(0)}${mix(2)}${mix(4)}`;
}

/** Approximate rendered text width in inches. CJK glyphs are full-width. */
/**
 * Width of a run in inches, from the measured advances.
 *
 * This was a fifth copy of the flat 0.54 em model, with a CJK range that did
 * not even match the other four (`\u3000-\u9fff\uff00-\uffef` against
 * `\u2e80-\u9fff\uac00-\ud7af...`). It charged 0.54 em for every space, which
 * advances 0, so a six-word sub-line was measured 3.2 em too wide and the rule
 * reported an overflow that is not there.
 */
function textWidthIn(text: string, fontSizePt: number): number {
  return measuredTextWidthIn(text, fontSizePt);
}

function parseShapes(xml: string): Shape[] {
  const shapes: Shape[] = [];
  const spRe = /<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g;
  let m: RegExpExecArray | null;
  while ((m = spRe.exec(xml))) {
    const body = m[2];
    const name = /name="([^"]*)"/.exec(body)?.[1] ?? '';
    const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(body);
    const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(body);
    if (!off || !ext) continue;
    const texts = [...body.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => unescapeXml(t[1]));
    // Per paragraph, joined with the break the renderer actually draws. Joining
    // every `<a:t>` with nothing at all made a shape carrying five paragraphs
    // indistinguishable from one carrying a single line, so no rule in this
    // file could see a hard break — the exact blind spot that let a sixteen-row
    // table measure 5.83in and draw 10.33in.
    const paragraphs = [...body.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
      .map((p) => [...p[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => unescapeXml(t[1])).join(''));
    const sz = /sz="(\d+)"/.exec(body);
    const x = +off[1] / EMU_PER_INCH;
    const y = +off[2] / EMU_PER_INCH;
    const w = +ext[1] / EMU_PER_INCH;
    const h = +ext[2] / EMU_PER_INCH;
    // The line a connector actually draws, not the box that contains it. An
    // L-shaped hop's bounding box covers the whole corner, so measuring a chip
    // against the box calls it "on" an arrow that runs nowhere near it.
    let path: { x: number; y: number }[] | undefined;
    const pts = [...body.matchAll(/<a:pt x="(-?\d+)" y="(-?\d+)"\s*\/>/g)];
    if (/<a:custGeom>/.test(body) && pts.length >= 2) {
      path = pts.map((pt) => ({ x: x + +pt[1] / EMU_PER_INCH, y: y + +pt[2] / EMU_PER_INCH }));
    } else if (/prst="line"/.test(body)) {
      const flipH = /flipH="1"/.test(body);
      const flipV = /flipV="1"/.test(body);
      path = [
        { x: flipH ? x + w : x, y: flipV ? y + h : y },
        { x: flipH ? x : x + w, y: flipV ? y : y + h },
      ];
    }
    // Colour, so a rule can ask whether the text is actually readable against
    // what is drawn behind it. The fill lives in spPr before <a:ln>, which has
    // a solidFill of its own.
    const txIdx = body.indexOf('<p:txBody>');
    const spPr = txIdx >= 0 ? body.slice(0, txIdx) : body;
    const beforeLn = spPr.split('<a:ln')[0];
    const fillMatch = /<a:solidFill>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"\s*(?:\/>|>([\s\S]*?)<\/a:srgbClr>)/.exec(beforeLn);
    const fill = /<a:noFill\/>/.test(beforeLn) ? null : (fillMatch?.[1]?.toLowerCase() ?? null);
    const alphaMatch = fillMatch?.[2] ? /<a:alpha val="(\d+)"\/>/.exec(fillMatch[2]) : null;
    const fillAlpha = alphaMatch ? +alphaMatch[1] / 100000 : 1;
    // Read, not assumed. PowerPoint's defaults are 0.1in left/right and 0.05in
    // top/bottom, but any shape may override them and the tile name does.
    const bodyPr = /<a:bodyPr[^>]*>/.exec(body)?.[0] ?? '';
    const inset = (attr: string, dflt: number): number => {
      const v = new RegExp(`\\b${attr}="(-?\\d+)"`).exec(bodyPr)?.[1];
      return v === undefined ? dflt : +v / EMU_PER_INCH;
    };
    const insetX = inset('lIns', 0.1) + inset('rIns', 0.1);
    const insetY = inset('tIns', 0.05) + inset('bIns', 0.05);
    const runs = [...body.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((r) => {
      const rb = r[1];
      const rpr = /<a:rPr[^>]*>([\s\S]*?)<\/a:rPr>/.exec(rb);
      return {
        color: /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(rpr?.[1] ?? '')?.[1]?.toLowerCase() ?? null,
        sizePt: (+(/<a:rPr[^>]*\bsz="(\d+)"/.exec(rb)?.[1] ?? 0) || 1800) / 100,
        bold: /<a:rPr[^>]*\bb="1"/.test(rb),
        text: unescapeXml(/<a:t>([\s\S]*?)<\/a:t>/.exec(rb)?.[1] ?? '').trim(),
      };
    });
    shapes.push({
      name,
      x,
      y,
      w,
      h,
      text: (paragraphs.length > 0 ? paragraphs.join('\n') : texts.join(''))
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
      paragraphs: paragraphs.map((p) => p.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')),
      fontSize: sz ? +sz[1] / 100 : null,
      anchor: /<a:bodyPr[^>]*\banchor="([^"]+)"/.exec(body)?.[1] ?? null,
      fill,
      fillAlpha,
      runs,
      insetX,
      insetY,
      wrapNone: /\bwrap="none"/.test(bodyPr),
      align: /<a:pPr[^>]*\balgn="([^"]+)"/.exec(body)?.[1] ?? null,
      lineSpacing: (() => {
        const pct = /<a:lnSpc>\s*<a:spcPct val="(\d+)"\s*\/>/.exec(body)?.[1];
        return pct === undefined ? null : +pct / 100000;
      })(),
      path,
    });
  }
  return shapes;
}

/**
 * Where the words in a text shape are actually drawn.
 *
 * A text box is laid out to the room available, not to its contents: a service
 * caption on an icon-less tile is given nearly the whole tile, and the name is
 * then centred inside it on one or two lines. Anything asking "is the name
 * covered" has to ask about the lines, because asking about the box is asking
 * about the tile.
 */
/**
 * Whether every word on the slide is actually readable against what is drawn
 * behind it, to WCAG 2.1 AA. Nothing had ever measured this: the audit only
 * ever built the light deck, and no rule looked at colour at all.
 */
/**
 * The 7pt floor, applied to the chips that sit on the arrows.
 *
 * A connector chip is the only text on the slide that says *why* two services
 * are joined, and it was the one piece of text the legibility rules never
 * measured: the tile rule filters on `service-label-`, so a chip drawn at 6.39pt
 * beside a 7pt tile name passed the audit clean. It is not exempt for being
 * secondary — an arrow whose reason cannot be read is an arrow the reader has
 * to guess about.
 */
function connectorLabelFontIssues(shapes: Shape[], prefix: string): string[] {
  const chips = shapes.filter(
    (s) => s.name.startsWith('connector-label-') && s.text.trim() !== '' && s.fontSize !== null,
  );
  const under = chips.filter((s) => (s.fontSize ?? 99) < 7);
  if (under.length === 0) return [];
  const worst = under.reduce((a, b) => ((a.fontSize ?? 99) <= (b.fontSize ?? 99) ? a : b));
  return [
    `${prefix}${under.length} connector label(s) drawn below the 7pt legibility floor, smallest ${worst.fontSize}pt: "${worst.text.slice(0, 40)}"`,
  ];
}

function contrastIssues(shapes: Shape[], slideBg: string): string[] {
  const issues: string[] = [];
  shapes.forEach((shape, idx) => {
    const readable = shape.runs.filter((r) => r.text !== '' && r.color);
    if (readable.length === 0) return;
    // The backdrop is everything already drawn under these glyphs, composited
    // in order — a caption on a translucent zone inside a tile is read against
    // the result, not against any one of them.
    let backdrop = slideBg;
    for (let i = 0; i < idx; i++) {
      const under = shapes[i];
      if (!under.fill) continue;
      if (
        under.x <= shape.x + 0.02 &&
        under.y <= shape.y + 0.02 &&
        under.x + under.w >= shape.x + shape.w - 0.02 &&
        under.y + under.h >= shape.y + shape.h - 0.02
      ) {
        backdrop = blend(under.fill, backdrop, under.fillAlpha);
      }
    }
    if (shape.fill) backdrop = blend(shape.fill, backdrop, shape.fillAlpha);
    let worst: { ratio: number; need: number; run: (typeof readable)[number] } | null = null;
    for (const run of readable) {
      const ratio = contrastRatio(run.color!, backdrop);
      const large = run.sizePt >= 18 || (run.bold && run.sizePt >= 14);
      const need = large ? 3 : 4.5;
      if (ratio < need && (!worst || ratio < worst.ratio)) worst = { ratio, need, run };
    }
    if (worst) {
      const sample = worst.run.text.length > 28 ? `${worst.run.text.slice(0, 28)}…` : worst.run.text;
      issues.push(
        `"${sample}" in ${shape.name || 'shape'} is #${worst.run.color} on #${backdrop} — ` +
          `contrast ${worst.ratio.toFixed(2)}:1, below the ${worst.need}:1 WCAG AA bar at ${worst.run.sizePt}pt`,
      );
    }
  });
  return issues;
}

/**
 * The height one wrapped line occupies, in inches.
 *
 * 1.35 em is a deliberate margin over the ~1.2 em a renderer actually uses,
 * and every box the exporter sizes for itself is checked against it. But a
 * shape may *state* its line spacing — `<a:lnSpc><a:spcPct>` — and two do: the
 * tile name and the zone caption both ask for 0.9. Measuring a box that was
 * sized at 0.9 against a 1.35 model reports every one of them as 10% over
 * before a single word is wrong, which is not a defect, it is the model
 * disagreeing with the file. Where the file says what it wants, believe it.
 */
function linePitchIn(shape: Shape, pt: number): number {
  return (pt * (shape.lineSpacing === null ? 1.35 : 1.2 * shape.lineSpacing)) / 72;
}

/**
 * The width a wrapped line actually has: the shape less the insets it declares,
 * not less a guess at PowerPoint's defaults.
 */
function textColumnIn(shape: Shape): number {
  return Math.max(MIN_TEXT_COLUMN_IN, shape.w - shape.insetX);
}

function drawnTextRect(shape: Shape, singleLine = false): { x: number; y: number; w: number; h: number } | null {
  const text = shape.text.trim();
  if (text === '' || !shape.fontSize) return null;
  // A `wrap="none"` run is drawn on one line whatever its width, and a centred
  // one that outgrows its box overflows equally on both sides rather than
  // wrapping. Modelling it as wrapped would claim rows above it that hold
  // nothing and report chips that never touched a glyph.
  //
  // The wrapped count is a wrap, not a ratio. Dividing total ink by the column
  // is the break-anywhere assumption — it says how many lines the ink would
  // need if words could split at any character, which is a lower bound and
  // never the answer. It is also the *same* arithmetic the tile exporter used,
  // so the rule and the code it checks were wrong together: a tile name that
  // drew five lines of a three-line box reported three here and no collision
  // with anything.
  const lines = singleLine ? 1 : auditWrappedLines(text, textColumnIn(shape), shape.fontSize);
  // The real pitch, not the margin: this rect is used for collisions, where
  // over-stating the ink invents overlaps. A declared spacing still counts.
  const h = Math.min(shape.h, (lines * shape.fontSize * 1.22 * (shape.lineSpacing ?? 1)) / 72);
  const w = lines > 1
    ? shape.w
    : (singleLine ? textWidthIn(text, shape.fontSize) : Math.min(shape.w, textWidthIn(text, shape.fontSize)));
  const x = shape.align === 'l'
    ? shape.x + shape.insetX / 2
    : shape.align === 'r'
      ? shape.x + shape.w - shape.insetX / 2 - w
      : shape.x + (shape.w - w) / 2;
  const y = shape.anchor === 't'
    ? shape.y
    : shape.anchor === 'b'
      ? shape.y + shape.h - h
      : shape.y + (shape.h - h) / 2;
  return { x, y, w, h };
}

/** Distance from a point to the nearest edge of a shape, zero when inside it. */
function edgeGap(box: { x: number; y: number; w: number; h: number }, at: { x: number; y: number }): number {
  return Math.hypot(
    at.x - Math.max(box.x, Math.min(at.x, box.x + box.w)),
    at.y - Math.max(box.y, Math.min(at.y, box.y + box.h)),
  );
}

/** Distance from a point to a connector's drawn path, falling back to its box. */
function pathGap(shape: Shape, at: { x: number; y: number }): number {
  if (!shape.path || shape.path.length < 2) return edgeGap(shape, at);
  let best = Number.POSITIVE_INFINITY;
  for (let i = 1; i < shape.path.length; i += 1) {
    const a = shape.path[i - 1];
    const b = shape.path[i];
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const len = dx * dx + dy * dy;
    const t = len > 0 ? Math.max(0, Math.min(1, ((at.x - a.x) * dx + (at.y - a.y) * dy) / len)) : 0;
    best = Math.min(best, Math.hypot(at.x - (a.x + t * dx), at.y - (a.y + t * dy)));
  }
  return best;
}
/** Overlapping area of two rectangles. Only the geometry is read, so a bare
 * rect — an ink box, a panel, a table row — is as good an argument as a shape. */
function overlapArea(
  a: { x: number; y: number; w: number; h: number },
  b: { x: number; y: number; w: number; h: number },
): number {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 ? w * h : 0;
}

interface Scenario {
  id: string;
  nodes: Node[];
  edges: Edge[];
  /**
   * Set when the nodes came out of the real layout engine. The strip rule below
   * only applies to those: a hand-placed strip is the user's own canvas and an
   * exporter that silently refolded it would no longer match what they drew.
   */
  fromLayoutEngine?: boolean;
  /**
   * Export against the dark palette. Every check had only ever run against the
   * light theme, so no colour the dark deck uses had ever been measured.
   */
  dark?: boolean;
  /**
   * The header triple, when the scenario is about it.
   *
   * Every run priced the same 16-character name, 5-character author and
   * 10-character date, hardcoded at the two call sites - so the cover title,
   * the section headers and the diagram slide header were outside every
   * rule's field of view on every scenario, by construction. The name is free
   * text the user types and nothing on the path caps it.
   */
  title?: string;
  author?: string;
}

/**
 * Every XML part in an OPC package has to be XML.
 *
 * Not a layout rule — a "does the file open at all" rule, and the only one of
 * those in this file that no amount of measuring geometry would ever catch.
 * The forbidden code points cannot be escaped, so an exporter that faithfully
 * passes a label through produces a package Word, PowerPoint and Visio all
 * refuse, while the export itself reports success. Cheap enough to run over
 * every part of every scenario: it is one regex per string already in memory.
 *
 * The test is written out here rather than imported from `xmlText`, on purpose.
 * A gate that asks the code under test whether the code under test is correct
 * is not a gate: the first version of this rule called `hasXmlForbidden`, and
 * neutering `stripXmlForbidden` disabled the strip and the detector in one
 * edit, so the mutation came back green.
 */
const AUDIT_FORBIDDEN = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

function xmlWellFormednessIssues(parts: Array<{ path: string; text: string }>, prefix: string): string[] {
  const issues: string[] = [];
  for (const part of parts) {
    const hit = AUDIT_FORBIDDEN.exec(part.text);
    if (!hit) continue;
    const code = hit[0].codePointAt(hit[0].length - 1) ?? 0;
    issues.push(
      `${prefix}${part.path} carries U+${code.toString(16).toUpperCase().padStart(4, '0')}, which XML 1.0 forbids and no escaping can encode — the package will not open`,
    );
  }
  return issues;
}

function svc(id: string, label: string, x: number, y: number, parent?: string, icon = true, category?: string): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 75,
    ...(parent ? { parentNode: parent } : {}),
    data: {
      label,
      serviceName: label,
      ...(category ? { category } : {}),
      ...(icon
        ? { iconPath: '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Virtual-Machine.svg' }
        : {}),
    },
  } as Node;
}

/**
 * A DR region drawn far east of the primary, with a sovereignty band overlapping
 * only its western half — the shape an Architecture Center multi-region diagram
 * takes when the author annotates data residency across part of a region.
 *
 * The band declares no members. Reading membership geometrically instead lets it
 * claim the two services it happens to cross, which are then packed as part of
 * the band while the virtual network that owns them is packed somewhere else:
 * the finished sheet shows two services standing outside the network they are
 * inside, which is a false statement about the architecture, not a cosmetic
 * slip.
 */
/**
 * A dense core with a secondary region parked far from it, and one service the
 * region's box happens to cover without owning.
 *
 * Two jobs. Closing empty bands is now the first thing either exporter does, so
 * every fixture that separated its outliers with blank canvas on both axes
 * stopped reaching the parking code that trims and packs — several hundred
 * lines of it were being carried untested. Compaction is far stronger than it
 * looks: the previous version of this fixture separated its region by 1030px,
 * which survives compaction, and still came out at 39.69x10.99in against a
 * 55.10in gate — comfortably inside the fit, so the parking code never ran and
 * the commit message that claimed it did was wrong. Reaching it needs a
 * drawing whose *compacted* span still overflows, so the ten hops below are
 * spaced 1400px apart: each gap is under the 1536px bar, so compaction keeps
 * every one of them, and ten of them in series is a sheet no page can hold.
 *
 * And `probe` is what tells declared membership from geometric. It sits inside
 * the secondary region's rectangle but belongs to no zone — an annotation band
 * drawn across something it does not own, which is exactly what a compliance
 * or residency boundary looks like. Reading membership from the drawing rather
 * than from the author's own `parentNode` puts it in a different cluster from
 * the region, and clusters are packed into separate slots, so the finished
 * sheet shows the boundary in one place and the service that was standing in
 * it in another.
 */
function zoneStrayScenario(): Scenario {
  const names = [
    'Azure Front Door', 'Application Gateway', 'Azure App Service', 'Azure Functions',
    'Azure SQL Database', 'Azure Cosmos DB', 'Azure Key Vault', 'Azure Monitor',
    'Azure Service Bus', 'Azure Cache for Redis', 'Azure Blob Storage', 'Microsoft Entra ID',
  ];
  const nodes: Node[] = [
    ...Array.from({ length: 60 }, (_, i) => svc(
      `g-${i}`,
      names[i % names.length],
      (i % 10) * 250,
      Math.floor(i / 10) * 200,
    )),
    ...Array.from({ length: 10 }, (_, i) => svc(
      `h-${i}`,
      names[(i + 3) % names.length],
      2400 + 1400 * i,
      400,
    )),
    grp('dr', 'Secondary region', 17800, 0, 620, 300),
    svc('dr-gw', 'Azure VPN Gateway', 30, 40, 'dr'),
    svc('dr-db', 'Azure SQL Database', 330, 40, 'dr'),
    svc('dr-store', 'Azure Blob Storage', 30, 180, 'dr'),
    svc('probe', 'Azure Policy', 18130, 180),
  ];
  const edges: Edge[] = [
    { id: 'c1', source: 'g-0', target: 'g-1', label: 'Routes' } as Edge,
    { id: 'c2', source: 'g-1', target: 'g-2', label: 'Balances' } as Edge,
    { id: 'c3', source: 'g-2', target: 'h-0', label: 'Queries' } as Edge,
    ...Array.from({ length: 9 }, (_, i) => (
      { id: `hop-${i}`, source: `h-${i}`, target: `h-${i + 1}`, label: 'Forwards' } as Edge
    )),
    { id: 'c4', source: 'h-9', target: 'dr-db', label: 'Replicates' } as Edge,
    { id: 'c5', source: 'dr-gw', target: 'dr-db', label: 'Connects' } as Edge,
    { id: 'c6', source: 'dr-db', target: 'dr-store', label: 'Archives' } as Edge,
  ];
  return { id: 'pipeline-region', nodes, edges };
}

/**
 * Two regions far apart, with one rectangle drawn around both of them.
 *
 * A subscription frame, a tenant boundary, an "Azure" box — the most ordinary
 * annotation in the Architecture Center, and it spans every empty band in the
 * drawing. Judging emptiness by every rectangle therefore found no void at all,
 * so a sheet that is nine tenths blank was exported at full size, and the gate
 * meant to catch that was blinded by the same rectangle.
 */
function boundaryVoidScenario(): Scenario {
  const nodes: Node[] = [
    grp('azure', 'Azure', -80, -80, 7060, 1000),
    ...Array.from({ length: 6 }, (_, i) => svc(
      `east-${i}`,
      ['Azure Front Door', 'Azure App Service', 'Azure SQL Database'][i % 3],
      80 + (i % 3) * 200,
      80 + Math.floor(i / 3) * 180,
      'azure',
    )),
    ...Array.from({ length: 6 }, (_, i) => svc(
      `west-${i}`,
      ['Azure Traffic Manager', 'Azure Functions', 'Azure Cosmos DB'][i % 3],
      6080 + (i % 3) * 200,
      80 + Math.floor(i / 3) * 180,
      'azure',
    )),
  ];
  const edges: Edge[] = [
    { id: 'b1', source: 'east-0', target: 'east-1', label: 'Routes' } as Edge,
    { id: 'b2', source: 'east-2', target: 'west-2', label: 'Replicates' } as Edge,
    { id: 'b3', source: 'west-0', target: 'west-1', label: 'Serves' } as Edge,
  ];
  return { id: 'boundary-void', nodes, edges };
}

/**
 * Subnets stacked one above another inside a virtual network — the shape of
 * every hub-and-spoke and every N-tier drawing the Architecture Center
 * publishes.
 *
 * The band immediately above a zone is clear of service tiles and belongs to
 * the zone above it. Scoring title placement against tiles alone therefore
 * moved a title out of its own box and into its neighbour's, so the drawing
 * asserted "Data subnet" was part of the application tier.
 */
function stackedSubnetsScenario(): Scenario {
  const nodes: Node[] = [];
  const tiers = [
    ['web', 'Web subnet', 'Azure Application Gateway', 'Azure Front Door'],
    ['app', 'Application subnet', 'Azure App Service', 'Azure Functions'],
    ['data', 'Data subnet', 'Azure SQL Database', 'Azure Cosmos DB'],
  ];
  tiers.forEach(([id, label, first, second], tier) => {
    // Drawn tight around their contents and stacked close, the way a real
    // subnet stack is: there is no clear band inside the box for a title, and
    // the only clear band near it belongs to the subnet above.
    nodes.push(grp(id, label, 0, tier * 118, 620, 95));
    nodes.push(svc(`${id}-a`, first, 40, 10, id));
    nodes.push(svc(`${id}-b`, second, 320, 10, id));
  });
  const edges: Edge[] = [
    { id: 's1', source: 'web-a', target: 'app-a', label: 'Forwards' } as Edge,
    { id: 's2', source: 'app-a', target: 'data-a', label: 'Queries' } as Edge,
    { id: 's3', source: 'app-b', target: 'data-b', label: 'Reads' } as Edge,
  ];
  return { id: 'stacked-subnets', nodes, edges };
}

/**
 * The same stack with the boxes actually full.
 *
 * `stacked-subnets` leaves a quarter of each row free, which is enough for a
 * half-width band to find clear space — so it passes for a reason that has
 * nothing to do with the rule being right. Fill the row and the fixed-share
 * candidates run out: three tiles across 620px cover 54% of every band on
 * offer, four across 640px cover 69%, and the audit fails a title at 25%. That
 * is not a placement that scores badly, it is no legal placement at all, and a
 * subnet drawn full is the ordinary case rather than the corner one.
 */
function tightSubnetsScenario(): Scenario {
  const nodes: Node[] = [];
  const rows: Array<[string, string, number, number]> = [
    ['web', 'Web subnet', 3, 620],
    ['app', 'Application subnet', 4, 640],
    ['data', 'Data subnet', 4, 780],
  ];
  const names = [
    'Azure Application Gateway', 'Azure Front Door', 'Azure App Service', 'Azure Functions',
    'Azure SQL Database', 'Azure Cosmos DB', 'Azure Key Vault', 'Azure Monitor',
  ];
  rows.forEach(([id, label, count, width], tier) => {
    nodes.push(grp(id, label, 0, tier * 118, width, 95));
    const pitch = (width - 20) / count;
    for (let i = 0; i < count; i += 1) {
      nodes.push(svc(`${id}-${i}`, names[(tier * 3 + i) % names.length], 10 + i * pitch, 10, id));
    }
  });
  const edges: Edge[] = [
    { id: 't1', source: 'web-0', target: 'app-0', label: 'Forwards' } as Edge,
    { id: 't2', source: 'app-0', target: 'data-0', label: 'Queries' } as Edge,
    { id: 't3', source: 'app-1', target: 'data-1', label: 'Reads' } as Edge,
  ];
  return { id: 'tight-subnets', nodes, edges };
}

/**
 * `tight-subnets` plus one edge the author numbered with a STRING.
 *
 * Models emit `"2"` about as often as `2`, and the Load-diagram path never
 * validates the field, so a saved or hand-edited file carries it verbatim into
 * the exporter. The promotion allocator read the field directly instead of
 * through `readStepNumber`, so it did not see the `"2"`: it started counting at
 * zero, handed the promoted chip the number 2 as well, and then found a row
 * already at step 2 and dropped the wording it exists to save. Two badges read
 * "2" and one arrow's label was deleted.
 *
 * The fixture also pins the second half of the same predicate mismatch — an
 * unlabelled edge can never promote, so it must not consume a number and push
 * the ones the reader sees out into the teens.
 */
function stringStepPromotionScenario(): Scenario {
  const base = tightSubnetsScenario();
  return {
    id: 'string-step-promotion',
    nodes: base.nodes,
    edges: [
      ...base.edges,
      {
        id: 'a0',
        source: 'web-1',
        target: 'app-2',
        label: 'Authenticates',
        data: { stepNumber: '2', stepDescription: 'Validates the caller token' },
      } as unknown as Edge,
    ],
  };
}

/**
 * The same deck with a crowd of plain, wordless connectors alongside.
 *
 * Promotion hands a stranded label the next free number. The numbers it may
 * hand out are bounded by what the workflow list can explain — the highest
 * authored step plus one per labelled arrow — but the allocator used to queue
 * every unnumbered edge, wordless ones included. Those can never become a row,
 * so each one silently pushed the real promotion further out: three authored
 * steps and twelve plain arrows numbered the one promoted hop 16, and the
 * reader was shown a list reading 1, 2, 3, 16.
 *
 * `string-step-promotion` cannot see this, because there the extra numbers
 * happen to land in a gap and collide with nothing.
 */
/**
 * The same promotion, with the wording in `data.label` instead of `label`.
 *
 * `readEdgeLabel` exists because either field can be the sole carrier — the
 * editor keeps them in step but the load path does not enforce it, so a saved
 * or model-authored file reaches the exporter with only the nested one. The
 * exporter reads both; the gate used to read only the top-level field, so it
 * called a correctly promoted badge spurious and skipped the very edge the
 * promotion-ceiling rule was written to watch.
 */
function dataLabelPromotionScenario(): Scenario {
  const base = tightSubnetsScenario();
  return {
    id: 'data-label-promotion',
    nodes: base.nodes,
    edges: base.edges.map((e) => (e.id === 't1'
      ? ({ id: 't1', source: 'web-0', target: 'app-0', data: { label: 'Forwards' } } as unknown as Edge)
      : e)),
  };
}

/**
 * Tiles with an authored width of a few pixels, alongside ordinary ones.
 *
 * The corpus could not reach this shape by scaling: across 9,607 tiles the
 * narrowest reading-slide tile is 1.1076in, and a 20,000px zone still leaves
 * the page scale above 0.82. It takes an AUTHORED width — which File → Load
 * hands over without complaint, because `validateRestoredNodes` checks `id`,
 * `position`, `data`, `type`, `parentNode`, tags, pricing and colour, and
 * never looks at `width` or `height`.
 *
 * That gap hid a synchronous infinite loop in the tile-name fitter for any
 * tile under 0.113in: no error, no watchdog — the event loop is blocked, so
 * the tab freezes and force-quit is the only exit.
 *
 * Deliberately keeps three ordinary tiles beside the hairlines so the deck
 * still has real geometry to measure rather than degenerating into a page of
 * slivers that every other rule skips.
 */
/**
 * Connector labels made of the characters the width table used to guess at,
 * and of ordinary words separated by ordinary spaces.
 *
 * Both are the same defect seen from two sides. A space was charged 0 em, so
 * every interior gap on a line was free and a chip that really wraps to two
 * lines was measured as one - `"step 19"` in a 0.220in column is the whole bug
 * in seven characters. And anything outside printable ASCII fell through to a
 * flat 0.54, so an arrow measured 46% narrow and the ellipsis the exporter
 * appends at every truncation point measured 26% narrow.
 *
 * Neither the exporter nor this audit could see either one, because both sides
 * were incomplete in exactly the same places.
 */
function probeArrowScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const labels = [
    'Ingest \u2192 Transform \u2192 Serve',
    'Extract \u2014 Load \u2014 Report',
    'Ingest \u2192 Serve',
    'validates the request and issues a token',
    'step 19',
    'writes the order document to Cosmos DB',
  ];
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  labels.forEach((label, i) => {
    nodes.push({
      id: `pa${i}a`,
      type: 'azureNode',
      position: { x: 0, y: i * 230 },
      width: 110,
      height: 55,
      data: { label: 'Azure Data Factory pipeline', serviceName: 'Data Factory', category: 'analytics', iconPath: icon },
    } as unknown as Node);
    nodes.push({
      id: `pa${i}b`,
      type: 'azureNode',
      position: { x: 450, y: i * 230 },
      width: 110,
      height: 55,
      data: { label: 'Azure Synapse Analytics workspace', serviceName: 'Synapse', category: 'analytics', iconPath: icon },
    } as unknown as Node);
    edges.push({ id: `pae${i}`, source: `pa${i}a`, target: `pa${i}b`, label } as Edge);
  });
  // A single unbroken token, several columns wide.
  //
  // This is the case that separates character-by-character packing from
  // `ceil(width / column)`: the ratio assumes a word packs exactly a columnful
  // per line, which is only true if a break may fall part-way through a glyph.
  // Breaks fall between glyphs, so it is a lower bound and never a count - and
  // an Azure resource id is exactly the kind of name that has no space in it
  // for forty characters. The mutation that put the ratio back survived all 96
  // files of the corpus until this fixture existed.
  ['contoso-prod-eastus-vnet-gateway-connection', 'OrderManagementWorkflowOrchestrator']
    .forEach((label, i) => {
      nodes.push({
        id: `pa-token${i}`,
        type: 'azureNode',
        position: { x: 900 + i * 200, y: 0 },
        width: 46,
        height: 210,
        data: { label, serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
      } as unknown as Node);
    });
  return { id: 'probe-arrow', nodes, edges };
}

/**
 * The same tile, twice, spelled once with accents and once without.
 *
 * Latin-1 and Latin Extended-A had no measured advance and took the 1 em
 * unknown-character fallback, which is 91% over for `e-acute` and 276% over
 * for a dotless `i`. The fallback was defended as harmless over-reservation:
 * charge too much and the type shrinks or buys a line it did not need. That
 * reasoning is false for exactly one caller. `widestGlyphIn` decides whether a
 * tile is wide enough to be worth naming at all, so an over-charge there does
 * not shrink anything - it DELETES the name, and a 0.190in column that really
 * holds this string was told its widest glyph was a full em when the true
 * widest is an `R` at 0.598.
 *
 * The pair is the whole point. Both tiles are 0.250 x 1.563in and differ only
 * in four accents, so any rule that reports one drawn and the other missing is
 * describing the width model and nothing else. Every European language a
 * customer might name a resource group in lives in this range.
 */
function probeAccentScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const names = [
    'R\u00e9seau priv\u00e9 s\u00e9curis\u00e9',
    'Reseau prive securise',
    'Cami\u00f3n log\u00edstica an\u00e1lisis',
    'Zar\u0105dzanie sieci\u0105 wirtualn\u0105',
  ];
  const nodes: Node[] = names.map((label, i) => ({
    id: `acc${i}`,
    type: 'azureNode',
    position: { x: (i % 2) * 120, y: Math.floor(i / 2) * 260 },
    width: i < 2 ? 24 : 26,
    height: 150,
    data: { label, serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
  } as unknown as Node));
  nodes.push({
    id: 'acchub',
    type: 'azureNode',
    position: { x: 420, y: 130 },
    width: 150,
    height: 75,
    data: { label: 'Hub virtual network', serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
  } as unknown as Node);
  const edges: Edge[] = names.map((_, i) => (
    { id: `acce${i}`, source: `acc${i}`, target: 'acchub', label: 'peers' } as Edge
  ));
  // Sizes at which the 20% over-charge changes the DRAWN STRING rather than
  // the draw/skip decision. At 80x55 and 140x30 the measured table sets
  // `Reseau prive securise partage` whole and the fallback cuts it to
  // `Reseau prive secu...partage` - the tile had the room all along, and the
  // ellipsis is the only thing the reader ever sees of the difference.
  [[80, 55], [90, 55], [140, 30], [50, 75]].forEach(([w, h], i) => {
    nodes.push({
      id: `acc-cut${i}`,
      type: 'azureNode',
      position: { x: 700 + (i % 2) * 340, y: 420 + Math.floor(i / 2) * 220 },
      width: w,
      height: h,
      data: {
        label: 'R\u00e9seau priv\u00e9 s\u00e9curis\u00e9 partag\u00e9',
        serviceName: 'Virtual Network',
        category: 'networking',
        iconPath: icon,
      },
    } as unknown as Node);
  });
  return { id: 'probe-accent', nodes, edges };
}

/**
 * Ordinary tiles whose names contain an ampersand.
 *
 * `Backup & Recovery` is written into a slide as `Backup &amp; Recovery`,
 * which is correct OOXML and the only legal way to write it. Every rule that
 * read text back out of an emitted file compared and MEASURED that escaped
 * form: the comparison saw two different strings and reported that the deck
 * had dropped a name it had drawn perfectly, and the measurement charged
 * `&amp;` 2.975 em against `&`'s 0.800, declaring a correct text block 3.7x
 * short of its own ink.
 *
 * The direction is always over-measure, so this is a false-RED class rather
 * than a miss - which on a gate that arbitrates every round costs exactly as
 * much as a miss. The third name is the control: same length, same structure,
 * no ampersand.
 */
function probeAmpScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const names = [
    'Backup & Recovery & Archive vault',
    'R&D analytics & reporting workspace',
    'Ordinary control plane workspace',
    'Identity <core> "primary" workspace',
  ];
  const nodes: Node[] = names.map((label, i) => ({
    id: `amp${i}`,
    type: 'azureNode',
    position: { x: (i % 2) * 320, y: Math.floor(i / 2) * 200 },
    width: 150,
    height: 75,
    data: {
      label,
      serviceName: 'Recovery Services',
      category: 'management',
      iconPath: icon,
      sku: 'GRS',
      region: 'japaneast',
    },
  } as unknown as Node));
  const edges: Edge[] = [
    { id: 'ampe0', source: 'amp0', target: 'amp1', label: 'replicates' } as Edge,
    { id: 'ampe1', source: 'amp2', target: 'amp3', label: 'audits & logs' } as Edge,
  ];
  return { id: 'probe-amp', nodes, edges };
}

/**
 * One narrow tile per script, each paired with the same name in English.
 *
 * Round 56 measured Latin-1 and Latin Extended-A and stopped at U+017F. That
 * left Cyrillic, Greek, Hebrew, Thai and Vietnamese on the 1 em fallback in a
 * pipeline whose draw/skip test reads that fallback twice with opposite
 * meanings: the widest-glyph clause charged an unknown character the LOWER
 * bound and the mean clause the UPPER one, and the mean clause binds. So a
 * Cyrillic name measured 96% over, first became drawable at 0.1945in against
 * an honest 0.0993in, and a 95 mil band of ordinary column widths refused it.
 *
 * The pairing is what makes this a fixture rather than a demonstration. Both
 * tiles in a pair are the same size and their names are the same length, so
 * any rule that reports the English one drawn and the other blank is
 * describing the width model and nothing else.
 *
 * The sheet is the format that matters here. A deck that cannot fit a name
 * cuts it and the whole name is still on the index slide; Visio has one page,
 * and a name it declines to draw survives only in the shape's Name attribute,
 * which is a handle for automation and never appears on paper.
 */
function probeScriptScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const pairs: Array<[string, string]> = [
    ['\u0412\u0438\u0440\u0442\u0443\u0430\u043b\u044c\u043d\u0430\u044f \u0441\u0435\u0442\u044c', 'Virtual network'],
    ['\u0395\u03b9\u03ba\u03bf\u03bd\u03b9\u03ba\u03cc \u03b4\u03af\u03ba\u03c4\u03c5\u03bf', 'Virtual network'],
    ['\u05e8\u05e9\u05ea \u05d5\u05d9\u05e8\u05d8\u05d5\u05d0\u05dc\u05d9\u05ea', 'Virtual network'],
    ['\u0e40\u0e04\u0e23\u0e37\u0e2d\u0e02\u0e48\u0e32\u0e22\u0e40\u0e2a\u0e21\u0e37\u0e2d\u0e19', 'Virtual network'],
    ['M\u1ea1ng \u1ea3o ri\u00eang', 'Virtual network'],
    ['Re\u021bea virtual\u0103', 'Virtual network'],
  ];
  const nodes: Node[] = [];
  pairs.forEach(([foreign, english], i) => {
    nodes.push({
      id: `scr${i}a`,
      type: 'azureNode',
      position: { x: 0, y: i * 90 },
      width: 26,
      height: 60,
      data: { label: foreign, serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
    } as unknown as Node);
    nodes.push({
      id: `scr${i}b`,
      type: 'azureNode',
      position: { x: 200, y: i * 90 },
      width: 26,
      height: 60,
      data: { label: english, serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
    } as unknown as Node);
  });
  // Chips too. The same over-charge sizes an edge label, where a name that
  // measures twice its ink produces a ribbon twice as wide as it needs to be
  // lying across the middle of the drawing: this one went from 1.784in of
  // reserved width to a measured 0.990in. The wording is kept short on
  // purpose - a fifteen-character sentence between two 26px tiles is
  // disproportionate in any script, and that is a different rule's business.
  const edges: Edge[] = pairs.map((_, i) => ({
    id: `scre${i}`,
    source: `scr${i}a`,
    target: `scr${i}b`,
    label: '\u0434\u0430\u043d\u043d\u044b\u0435',
  } as Edge));
  return { id: 'probe-script', nodes, edges };
}

/**
 * A long, shallow pipeline - the one drawing shape that forces the sheet scaler
 * down far enough to matter.
 *
 * A grid grows the page instead: 20x20 = 400 services still reports scale
 * 1.0001 and never triggers anything. 260 services in a line scales to 0.436,
 * which is where flat positioning constants and scaled dimensions come apart.
 */
function scaleDownPipelineScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const count = 260;
  const nodes: Node[] = Array.from({ length: count }, (_, i) => ({
    id: `ps${i}`,
    type: 'azureNode',
    position: { x: i * 170, y: 0 },
    width: 150,
    height: 75,
    data: {
      label: `Orders processing function ${i}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: count - 1 }, (_, i) => ({
    id: `ps-e${i}`,
    source: `ps${i}`,
    target: `ps${i + 1}`,
  } as Edge));
  return { id: 'probe-scaledown', nodes, edges };
}

/**
 * Labels that are not what they look like.
 *
 * Both exporters run every name through `singleLineName` before drawing it, so
 * what lands in the file is normalised: newlines and tabs become spaces, runs
 * of spaces collapse, the ends are trimmed. A gate that compares the RAW label
 * against the drawn text is therefore comparing two different strings and
 * matches nothing - the deck side of the cross-format rule went silently empty
 * and reported three clean names as missing from a deck that spells all three
 * out. Whitespace in a label is not exotic: it is what a pasted name arrives
 * with.
 */
function whitespaceLabelsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const labels = [
    'Zephyr order intake function ',
    'Quartz billing\nreconciliation hub',
    'Nimbus  telemetry  ingestion',
    'Cobalt fraud\tscoring service',
    'Verdant analytics warehouse',
    'Onyx configuration store',
  ];
  const widths = [14, 14, 14, 14, 160, 160];
  const nodes: Node[] = labels.map((label, i) => ({
    id: `ws${i}`,
    type: 'azureNode',
    position: { x: i * 260, y: (i % 2) * 200 },
    width: widths[i],
    height: widths[i] < 60 ? 30 : 110,
    data: { label, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges: Edge[] = labels.slice(1).map((_, i) => ({
    id: `ws-e${i}`, source: `ws${i}`, target: `ws${i + 1}`, label: 'invokes',
  } as Edge));
  return { id: 'probe-whitespace', nodes, edges };
}

/**
 * `probe-whitespace`, numbered.
 *
 * The same geometry the planner correctly refuses to split further -
 * `MIN_SERVICES_PER_TILED_SLIDE` rejects a six-window plan for six services -
 * but now carrying step callouts. The refusal is right; what was wrong is that
 * the deck it produces used to fall outside every badge rule, because the
 * conflict test was gated on the DRAWN tile clearing the markable bar and these
 * tiles draw at 0.1584in. Three discs shipped at 98% of the services they
 * number, on a corpus reporting no issues at all.
 */
function refusedRaiseScenario(): Scenario {
  const base = whitespaceLabelsScenario();
  return {
    id: 'probe-refused-raise',
    nodes: base.nodes,
    edges: base.edges.map((edge, i) => ({
      ...edge,
      id: `rr-e${i}`,
      data: { ...(edge.data ?? {}), stepNumber: i + 1, stepDescription: `Step ${i + 1} of the intake path` },
    } as unknown as Edge)),
  };
}

/**
 * A drawing that reaches its own bottom-left corner.
 *
 * The legend was once painted straight over the tiles for want of a reserved
 * strip, and the service-name panel repeated the mistake exactly: pinned to the
 * bottom-left, solid white, emitted after every tile, and reserved only in
 * `furnitureRects` - which keeps CONNECTOR LABELS off it, not shapes. On this
 * grid it covered 20 of 48 tiles completely and clipped 4 more, and the whole
 * corpus still reported clean, because every other fixture is a chain or a
 * handful of tiles that happen to sit above the corner.
 */
function panelBurialScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const words = ['Zephyr', 'Quartz', 'Nimbus', 'Cobalt', 'Verdant', 'Onyx', 'Amber', 'Slate'];
  const nodes: Node[] = [];
  for (let r = 0; r < 6; r += 1) {
    for (let c = 0; c < 8; c += 1) {
      const i = r * 8 + c;
      nodes.push({
        id: `pp${i}`,
        type: 'azureNode',
        position: { x: c * 120, y: r * 110 },
        width: 40,
        height: 60,
        data: {
          label: `${words[i % 8]} order intake reconciliation ${i}`,
          serviceName: 'Azure Functions',
          category: 'compute',
          iconPath: icon,
        },
      } as unknown as Node);
    }
  }
  const edges: Edge[] = Array.from({ length: nodes.length - 1 }, (_, k) => ({
    id: `pp-e${k + 1}`, source: `pp${k}`, target: `pp${k + 1}`,
  } as Edge));
  return { id: 'probe-panel-burial', nodes, edges };
}

/**
 * Twelve numbered steps, each of them two syllables long.
 *
 * BREVITY is what makes the column minimiser take its cap: with one-line
 * descriptions the stack height falls monotonically in the column count, so it
 * splits all the way to MAX_WORKFLOW_COLUMNS and sets each sentence in a
 * column a quarter of an inch wide. A sweep over step COUNT and chain length
 * never finds this - it was the fixture that could not produce the condition,
 * not the condition that did not exist - and it is the case where the removed
 * `Math.max(colW - 0.6, 0.4)` floor measured a row at 0.4in while drawing it
 * in 0.2583in, 1.55 times wider than the column it is set in.
 */
/**
 * Long, realistic service names on hairline tiles.
 *
 * The index rows read "<stub>  =  <full name>", which makes them the longest
 * strings the exporter emits, and the column they are set in was a CONSTANT
 * 3.4in - the one text decision in the file that never asked how wide its text
 * was. Four of five rows wrapped to two lines of 0.135in inside a 0.2in box,
 * and since rows are pitched at exactly the box height, the spill went through
 * the neighbouring row: two services' names drawn through each other by 35% of
 * the type size, in the panel whose entire purpose is that a shortened name
 * stays readable.
 */
/**
 * Eight services whose names share a long common prefix, on tiles too narrow to
 * print it. Every one of them shortens to the same stub, which makes the stub
 * useless as the index lookup key it is supposed to be: the index is on another
 * page, so the reader cannot see both at once and the drawing has to be
 * self-consistent on its own.
 */
function collidingStubsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const suffixes = ['alpha', 'bravo', 'charlie', 'delta', 'echo', 'foxtrot', 'golf', 'hotel'];
  const nodes: Node[] = suffixes.map((suffix, i) => ({
    id: `cs${i}`,
    type: 'azureNode',
    position: { x: i * 200, y: (i % 2) * 180 },
    width: 18,
    height: 30,
    data: {
      label: `Contoso platform shared services region ${suffix}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  return { id: 'probe-colliding-stubs', nodes, edges: [] };
}

function hairlineStubsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const suffixes = ['alpha', 'bravo', 'charlie', 'delta', 'echo', 'foxtrot', 'golf', 'hotel'];
  // The same eight colliding names as `probe-colliding-stubs`, on tiles a third
  // the width. The numeric key that fixture proved out is refused here: the
  // fallback asked for room for TWO of the key's widest glyph, and a one-digit
  // key never stacks, so the test bought nothing and cost every one of these
  // eight tiles its mark. All eight index rows then read "(not drawn)" - eight
  // services on the drawing, not one of them identifiable, which is the very
  // collision the key was introduced to end, one page over.
  const nodes: Node[] = suffixes.map((suffix, i) => ({
    id: `hs${i}`,
    type: 'azureNode',
    position: { x: i * 200, y: (i % 2) * 180 },
    width: 12,
    height: 24,
    data: {
      label: `Contoso platform shared services region ${suffix}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  return { id: 'probe-hairline-stubs', nodes, edges: [] };
}

/**
 * Emoji clusters an author actually pastes into a service name.
 *
 * A cluster is one glyph and is priced as one. Three separate ways of getting
 * that wrong all ended in a name drawn at the wrong size: VARIATION SELECTOR-16
 * promotes its base to the emoji advance rather than merely vanishing (a heart
 * is 1.000 em as a dingbat and 1.373 as an emoji, and a keycap digit is 0.539
 * against 1.373 - 61% under); a ZERO WIDTH JOINER absorbs what follows it
 * whatever plane that lives in (the staff in a health worker is U+2695, in the
 * BMP, so the cluster was charged twice - 73% over); and a promoted cluster is
 * MEASURED even though its base code point is not, or the coverage rule reports
 * a correctly priced glyph as a guess in every run forever.
 *
 * Under-charging paints a line out past its box; over-charging withholds the
 * name altogether. Both are here.
 */
function emojiClusterScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const names = [
    'Status \u2764\ufe0f monitor',
    'Cloud \u2601\ufe0f ingest gateway',
    'Step 1\ufe0f\u20e3 intake validation',
    'Clinic \u{1f468}\u200d\u2695\ufe0f records service',
    'Warning \u26a0\ufe0f alert router',
    'Region \u{1f1ef}\u{1f1f5}\u{1f1e9}\u{1f1ea} failover pair',
    // Text that is NOT emoji and was measured as if it were. The word joiner is
    // what documentation tooling emits for &NoBreak;, the ZWJ here forms a
    // Devanagari conjunct, and a variation selector over an ASCII base has no
    // emoji to select. All three were charged the emoji face.
    'Contoso\u2060Platform\u2060ingest\u2060relay\u2060service\u2060tier',
    'Archive\ufe0f tier\ufe0f rotation\ufe0f service',
    // A subdivision flag carries no joiner at all: it is a base flag followed
    // by six TAG code points that select "gbsct". Every clause that keeps a
    // sequence together looked for a joiner, so all seven were billed at the
    // astral face and one glyph cost seven - the name measured 600% wide and
    // the deck reserved a strip for it that nothing was ever drawn into.
    'Region \u{1f3f4}\u{e0067}\u{e0062}\u{e0073}\u{e0063}\u{e0074}\u{e007f} standby',
    // A byte order mark pasted in from a UTF-8 file. It is zero width, but it
    // is also inside JavaScript's own \s class, so the whitespace test claimed
    // it first and charged it a space.
    'Contoso\ufeff ingest\ufeff relay tier',
  ];
  const nodes: Node[] = names.map((label, i) => ({
    id: `ec${i}`,
    type: 'azureNode',
    position: { x: (i % 3) * 260, y: Math.floor(i / 3) * 200 },
    width: 150,
    height: 110,
    data: { label, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = [
    {
      id: 'ece1', source: 'ec0', target: 'ec1', label: 'heartbeat \u2764\ufe0f',
      data: { stepNumber: 1, stepDescription: 'The monitor sends a heartbeat \u2764\ufe0f every 30s.' },
    },
    {
      id: 'ece2', source: 'ec2', target: 'ec3', label: 'step 1\ufe0f\u20e3',
      data: { stepNumber: 2, stepDescription: 'Intake hands step 1\ufe0f\u20e3 to the clinic \u{1f468}\u200d\u2695\ufe0f service.' },
    },
  ] as unknown as Edge[];
  return { id: 'probe-emoji-clusters', nodes, edges };
}

/**
 * ASK-61-B, half one: twelve one-word steps.
 *
 * The band minimises stack height and nothing else, so brief descriptions push
 * it to its 12 column cap and each sentence gets 0.2583in of text column - two
 * characters. NO LINE-COUNT BOUND CAN SEE THIS: "ack" sets in one line in two
 * characters' width just as happily as in ten inches. Only the width can.
 */
function briefWorkflowScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 13 }, (_, i) => ({
    id: `bw${i}`,
    type: 'azureNode',
    position: { x: (i % 5) * 240, y: Math.floor(i / 5) * 200 },
    width: 150,
    height: 110,
    data: { label: `Step service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = Array.from({ length: 12 }, (_, i) => ({
    id: `bwe${i}`,
    source: `bw${i}`,
    target: `bw${i + 1}`,
    label: 'ack',
    data: { stepNumber: i + 1, stepDescription: 'ack' },
  })) as unknown as Edge[];
  return { id: 'probe-brief-workflow', nodes, edges };
}

/**
 * ASK-61-B, half two: twelve long steps.
 *
 * Long prose keeps the column wide enough to clear any width floor, and every
 * sentence still wraps to eight or ten lines - a stack of ribbons no reader
 * follows. NO WIDTH FLOOR CAN SEE THIS. The two fixtures exist as a pair so
 * that neither rule can be deleted on the grounds that the other covers it.
 */
function shreddedWorkflowScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const sentence = (i: number): string =>
    `The regional ingestion tier validates the payload envelope, resolves the partner tenant `
    + `against the directory, writes an audit record for step ${i + 1}, and forwards the request `
    + `to the downstream settlement processor over a private endpoint.`;
  const nodes: Node[] = Array.from({ length: 13 }, (_, i) => ({
    id: `sw${i}`,
    type: 'azureNode',
    position: { x: (i % 5) * 240, y: Math.floor(i / 5) * 200 },
    width: 150,
    height: 110,
    data: { label: `Ingest service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = Array.from({ length: 12 }, (_, i) => ({
    id: `swe${i}`,
    source: `sw${i}`,
    target: `sw${i + 1}`,
    label: `step ${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: sentence(i) },
  })) as unknown as Edge[];
  return { id: 'probe-shredded-workflow', nodes, edges };
}

function longIndexRowsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const names = [
    'Payments reconciliation and settlement processing function app',
    'Customer identity and access management gateway for partner tenants',
    'Regional telemetry ingestion and cold storage archival pipeline',
    'Order intake validation service (EMEA production, zone redundant)',
    'Azure Database for PostgreSQL flexible server - analytics replica',
    'Nimbus', 'Quartz', 'Zephyr',
  ];
  const widths = [8, 10, 14, 22, 40, 160, 160, 160];
  const nodes: Node[] = names.map((label, i) => ({
    id: `ln${i}`,
    type: 'azureNode',
    position: { x: i * 210, y: (i % 2) * 200 },
    width: widths[i],
    height: widths[i] < 60 ? 30 : 110,
    data: { label, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  return { id: 'probe-long-index', nodes, edges: [] };
}

/**
 * An index row too long for the sheet at the index's top size.
 *
 * `probe-long-index` proves rows are DEFINED; this proves they are DRAWN ON
 * THE PAGE. Names here are long enough that "<mark>  =  <name>" passes 170
 * characters, which at 10pt is about 8.9in of ink in a 12.6in column - fine on
 * its own, but the marks are cut stubs of the SAME leading words, so the pairs
 * are near-identical and the index cannot split into two columns without each
 * column falling under the row's natural width. That is the shape that drew
 * 3.127in past the right edge.
 */
function overlongIndexRowsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const stem = 'Contoso regional payments reconciliation, settlement and dispute resolution '
    + 'processing function application for the European Union production estate, '
    + 'zone redundant across three availability zones with cross region read replicas '
    + 'and a customer managed encryption key held in the shared platform vault';
  const nodes: Node[] = Array.from({ length: 6 }, (_, i) => ({
    id: `xi${i}`,
    type: 'azureNode',
    position: { x: i * 190, y: (i % 2) * 180 },
    // Narrow enough that every name is cut to a stub, so every one earns a row.
    width: 26,
    height: 30,
    data: {
      label: `${stem} - ring ${i}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  return { id: 'probe-overlong-index', nodes, edges: [] };
}

/**
 * An index row that wraps at 10pt and fits on one line at 7pt.
 *
 * `probe-overlong-index` is past saving at any size and proves the WRAP;
 * this one is inside the range and proves the SHRINK. Without the pair, one
 * mutation covers for the other and neither half is really tested.
 */
function shrinkableIndexRowsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const stem = 'Contoso regional payments reconciliation and settlement processing function '
    + 'application for the European Union production estate, zone redundant with a '
    + 'customer managed key held in the shared platform vault and paired for recovery';
  const nodes: Node[] = Array.from({ length: 6 }, (_, i) => ({
    id: `si${i}`,
    type: 'azureNode',
    position: { x: i * 190, y: (i % 2) * 180 },
    width: 26,
    height: 30,
    data: {
      label: `${stem} ${i}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  return { id: 'probe-shrinkable-index', nodes, edges: [] };
}

/**
 * One very long name among many ordinary ones.
 *
 * The index used to give every row the TALLEST row's pitch, so this shape -
 * 44 short rows and one that wraps - re-pitched the whole page: 0.2819in of
 * row for 0.1410in of ink, and a third index page for a list that fits on two.
 * The rule that catches it is the density one below; the fixture is here so
 * the packing cannot quietly regress to a fixed grid.
 */
function mixedLengthIndexScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const long = 'Contoso regional payments reconciliation, settlement and dispute resolution '
    + 'processing function application for the European Union production estate, '
    + 'zone redundant across three availability zones with cross region read replicas '
    + 'and a customer managed encryption key held in the shared platform vault';
  const nodes: Node[] = Array.from({ length: 45 }, (_, i) => ({
    id: `mi${i}`,
    type: 'azureNode',
    position: { x: (i % 9) * 190, y: Math.floor(i / 9) * 150 },
    width: 26,
    height: 30,
    data: {
      label: i === 0 ? long : `Contoso regional service unit number ${i} for the shared platform`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  return { id: 'probe-mixed-index', nodes, edges: [] };
}

/**
 * A workflow whose typical step is one word and whose important steps are
 * paragraphs.
 *
 * The median line bound cannot see this: 48 acknowledgements set in one line
 * hold the median at one however narrow the column gets, so the 3 sentences
 * the drawing exists to explain were folded into 9-line ribbons and no rule
 * said a word. The terse steps are not evidence about the column - they were
 * never going to wrap at any width - which is exactly why a statistic over all
 * of them is the wrong instrument.
 */
function bimodalWorkflowScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const TERSE = ['ack', 'ok', 'store', 'emit', 'log', 'retry', 'drop', 'queue', 'sign', 'flush'];
  const long = (i: number): string =>
    'The regional ingestion tier validates the payload envelope, resolves the partner tenant '
    + `against the corporate directory, writes an immutable audit record for step ${i + 1}, applies the `
    + 'tenant specific throttling policy, and forwards the request to the downstream settlement '
    + 'processor over a private endpoint before acknowledging the original caller, after which the '
    + 'reconciliation job compares the settled totals against the ledger of record, raises a variance '
    + 'alert when the two disagree by more than the agreed tolerance, and archives the matched batch '
    + 'to cold storage under the retention schedule the compliance team publishes each quarter.';
  const nodes: Node[] = Array.from({ length: 92 }, (_, i) => ({
    id: `pb${i}`,
    type: 'azureNode',
    position: { x: (i % 8) * 200, y: Math.floor(i / 8) * 160 },
    width: 150,
    height: 110,
    data: { label: `Ingest service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = Array.from({ length: 91 }, (_, i) => ({
    id: `pbe${i}`,
    source: `pb${i}`,
    target: `pb${i + 1}`,
    label: 'hands the payload downstream',
    data: { stepNumber: i + 1, stepDescription: i < 88 ? TERSE[i % 10] : long(i) },
  })) as unknown as Edge[];
  return { id: 'probe-bimodal-workflow', nodes, edges };
}

/**
 * A workflow whose connector labels are long enough to change the RESERVATION
 * without changing the band that is finally drawn.
 *
 * The reservation appends every labelled edge's wording to its row, because
 * whether a label is muted is not decided until the arrows are routed - which
 * is after the page has been sized. The panel appends it only for the edges
 * actually muted. So the two are measured over different sentences, and when
 * the extra wording crosses a line boundary the reservation gains a whole line
 * per row. Holding everything else fixed and lengthening ONLY the label took
 * the sheet from 17.09in to 26.39in for a drawing that spans 9.1in, and 4.2in
 * of the difference was printed as blank paper between the drawing and the
 * band. Both rules that should have caught it were already correct; nothing in
 * the corpus reached them.
 */
function bandGapScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const label = 'hands the settled payload to the downstream reconciliation processor';
  const description = 'Step validates the payload envelope, resolves the partner tenant against the '
    + 'corporate directory, writes an immutable audit record, applies the tenant specific throttling '
    + 'policy, and forwards the request to the downstream settlement processor.';
  const nodes: Node[] = Array.from({ length: 40 }, (_, i) => ({
    id: `bg${i}`,
    type: 'azureNode',
    position: { x: (i % 8) * 240, y: Math.floor(i / 8) * 200 },
    width: 150,
    height: 75,
    data: { label: `Settlement service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = Array.from({ length: 30 }, (_, i) => ({
    id: `bge${i}`,
    source: `bg${i}`,
    target: `bg${i + 1}`,
    label,
    data: { stepNumber: i + 1, stepDescription: `${i + 1}. ${description}` },
  })) as unknown as Edge[];
  return { id: 'probe-band-gap', nodes, edges };
}

/**
 * The same defect where SPLITTING CANNOT ABSORB IT.
 *
 * `probe-band-gap` is answered by laying the band out in more columns once the
 * real page width is known, which is the right first answer and shrinks the
 * sheet from 26.4in to 16.7in. It is not a guarantee: with 79 steps of
 * 400-character prose the median bound refuses to split any further, so the
 * whole of the reservation's over-estimate stays in one column and 2.58in of it
 * was printed as blank paper. The panel filling the height the page reserved
 * for it is what makes the guarantee structural rather than incidental, and
 * this is the fixture that reaches it.
 */
function bandFillScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const label = 'hands the settled payload to the downstream reconciliation and dispute '
    + 'resolution processor for the tenant of record for the settlement window';
  const description = 'validates the payload envelope, resolves the partner tenant against the '
    + 'corporate directory, writes an immutable audit record, applies the tenant specific throttling '
    + 'policy, forwards the request to the downstream settlement processor over a private endpoint, '
    + 'and acknowledges the original caller once the ledger has been written, after which the '
    + 'reconciliation job compares the settled totals against the ledger of record, raises a variance '
    + 'alert when the two disagree by more than the agreed tolerance, and archives the matched batch '
    + 'to cold storage under the retention schedule the compliance team publishes each quarter.';
  const nodes: Node[] = Array.from({ length: 80 }, (_, i) => ({
    id: `bf${i}`,
    type: 'azureNode',
    position: { x: (i % 10) * 240, y: Math.floor(i / 10) * 200 },
    width: 150,
    height: 75,
    data: { label: `Ledger service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = Array.from({ length: 79 }, (_, i) => ({
    id: `bfe${i}`,
    source: `bf${i}`,
    target: `bf${i + 1}`,
    label,
    data: { stepNumber: i + 1, stepDescription: `${i + 1}. Step ${description}` },
  })) as unknown as Edge[];
  return { id: 'probe-band-fill', nodes, edges };
}

/**
 * Names whose combining marks are still there after composition.
 *
 * The exporter now composes every drawn string, which retires the whole
 * Latin question - but it does NOT retire combining marks, and treating that
 * as the same problem would leave the cluster rules with no fixture that
 * reaches them. Arabic harakat, Devanagari matras and Thai vowel signs have no
 * precomposed forms at all: they survive NFC untouched, they are ordinary
 * spelling rather than decoration, and they are what the cluster pricing and
 * the cluster-boundary cut actually have to be right about.
 *
 * The tiles are small and their widths vary on purpose. Over-charging does not
 * merely mis-measure a name, it CUTS it, and a cut lands wherever the tile
 * happens to be narrow - at one width a code-point cut printed a mark stacked
 * on the ellipsis, and at the next width it did not.
 */
/**
 * A drawing authored too small to carry a mark, spread wide enough that the
 * SPLIT is what binds and not the cap.
 *
 * The planner and the renderer each hold a ceiling on how coarse a window may
 * be drawn, and they had drifted: the planner still capped at one authored
 * pixel per screen pixel while the renderer had started raising that for tiles
 * narrower than 19.2px, which cannot carry so much as a key at the natural
 * cap. A planner that stops splitting below the ceiling the renderer will
 * actually use leaves tiles under the markable bar with slides still in the
 * budget, and the deck then draws anonymous dots and an index that can only
 * say "(not drawn)".
 *
 * Sixty 14x10px tiles over a 4200x2600px sheet: small enough for the raised
 * ceiling to bind, large enough that the window fit decides the grid.
 */
/**
 * The header triple at four lengths.
 *
 * The gate had priced one 16-character name on all 227 runs, so the cover
 * title, the section headers and the diagram slide header had never been
 * measured against anything. The lengths are the ones where the behaviour
 * changes: 20 fits at full size, 70 crosses the two-line boundary on the
 * 22pt section header, 95 crosses it on the 40pt cover, and 130 reached
 * three lines and painted the content header off the top of the slide.
 *
 * The names are plausible rather than padded. A diagram called "Contoso
 * Global Retail Platform - Production Landing Zone with Paired Region
 * Disaster Recovery" is shorter than several service names already in the
 * corpus, and the field is free text with no cap anywhere on the path.
 */
function longTitleScenario(chars: number): Scenario {
  const full = 'Contoso Global Retail Platform - Production Landing Zone with '
    + 'Paired Region Disaster Recovery and Zone Redundant Ingestion Tier';
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [
    { id: 'lt1', type: 'azureNode', position: { x: 0, y: 0 }, width: 200, height: 120,
      data: { label: 'Ingestion function', serviceName: 'Azure Functions', category: 'compute', iconPath: icon } },
    { id: 'lt2', type: 'azureNode', position: { x: 420, y: 0 }, width: 200, height: 120,
      data: { label: 'Analytics store', serviceName: 'Azure Functions', category: 'compute', iconPath: icon } },
  ] as unknown as Node[];
  const edges = [
    { id: 'lte1', source: 'lt1', target: 'lt2', label: 'writes', data: { stepNumber: 1, stepDescription: 'The function writes each batch to the analytics store.' } },
  ] as unknown as Edge[];
  return {
    id: `probe-title-${chars}`,
    nodes,
    edges,
    title: full.slice(0, chars).trimEnd(),
    // Long enough to be a real person and a real team, which is what the
    // field holds in practice - "Audit" was five characters and never met
    // the right edge of the band it shares with the date.
    author: chars >= 95 ? 'Swarm Data SE, Jiayi Yang' : 'Audit',
  };
}

function tinyTileSpreadScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 60 }, (_, i) => ({
    id: `tt${i}`,
    type: 'azureNode',
    position: { x: (i % 10) * 420, y: Math.floor(i / 10) * 430 },
    width: 14,
    height: 16,
    data: {
      label: `Regional ingest worker ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  // No step numbers and no edge labels. A chip carrying even a five letter
  // word is wider than a 14px tile, and the smallest disc that still holds a
  // legible digit is wider still - so on this sheet a badge cannot be both
  // readable and smaller than the service it numbers. Those are real limits
  // and they are covered elsewhere at realistic sizes; this fixture is about
  // how small a TILE may become before the planner has to stop shrinking, and
  // it keeps that one variable moving on its own.
  const edges = Array.from({ length: 8 }, (_, i) => ({
    id: `tte${i}`,
    source: `tt${i}`,
    target: `tt${i + 10}`,
  })) as unknown as Edge[];
  return { id: 'probe-tiny-spread', nodes, edges };
}

function widthCliffScenario(): Scenario {
  // The one-pixel cliff, kept where it can be measured.
  //
  // Sixty tiles of 14px with a single 57px node among them. Under the
  // max-over-min ratio that once decided whether to raise the drawing ceiling
  // this reads 4.07 and the raise is refused, so 59 of 60 services come out
  // anonymous across 26 slides - and at 56px the same drawing reads 4.00, the
  // raise is allowed, and all 60 are named. One pixel on one node out of sixty
  // decided the whole deck, which is what an extremum-over-extremum statistic
  // does and is why the width now takes the median, as the height target
  // beside it already did.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 60 }, (_, i) => ({
    id: `wc${i}`,
    type: 'azureNode',
    position: { x: (i % 10) * 420, y: Math.floor(i / 10) * 430 },
    width: i === 59 ? 57 : 14,
    height: i === 59 ? 65 : 16,
    data: {
      label: `Regional ingest worker ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges = Array.from({ length: 8 }, (_, i) => ({
    id: `wce${i}`,
    source: `wc${i}`,
    target: `wc${i + 10}`,
  })) as unknown as Edge[];
  return { id: 'probe-width-cliff', nodes, edges };
}

function slaveredBadgeScenario(): Scenario {
  // One realistic workflow at a realistic size, with the slivers OUTNUMBERING
  // the tiles the workflow runs among.
  //
  // The first version of this fixture parked one 14px node beside an eight
  // stage chain, because reading the badge ceiling off the narrowest tile on
  // the sheet let that one node cut every badge in the chain from 0.240in to
  // 0.1119in - 53% off seven discs drawn between tiles nowhere near it. The
  // median fixed that and then had the same defect one node further out: four
  // large tiles beside five slivers has a sliver for a median, so the discs
  // collapse again, and the gate reports it as the OPPOSITE fault because the
  // badge is now 76.7% of a tile it is not drawn near. So the ceiling is taken
  // over the tiles at the ends of the numbered connectors, which is the only
  // set the measurement was ever about.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 9 }, (_, i) => ({
    id: `sb${i}`,
    type: 'azureNode',
    position: i >= 4
      ? { x: 2600 + (i - 4) * 90, y: 1900 }
      : { x: (i % 2) * 400, y: Math.floor(i / 2) * 320 },
    width: i >= 4 ? 14 : 200,
    height: i >= 4 ? 16 : 128,
    data: {
      label: i >= 4 ? `Retired probe ${i - 3}` : `Order pipeline stage ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges = Array.from({ length: 3 }, (_, i) => ({
    id: `sbe${i}`,
    source: `sb${i}`,
    target: `sb${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Stage ${i + 1} hands off to stage ${i + 2}` },
  })) as unknown as Edge[];
  return { id: 'probe-badge-sliver', nodes, edges };
}

function threeTierBadgeScenario(): Scenario {
  // An ORDINARY diagram. Three tiers of six services with four numbered hops,
  // and one small node - a private DNS zone, which every real Azure landing
  // zone has - dropped between the tiers.
  //
  // Nothing here is adversarial, and that is the point. Deciding which tile a
  // badge belongs to by asking which tile is nearest looks obviously right and
  // is false on any hop longer than the gap to a bystander: a hop on this grid
  // is over three inches, the DNS node sits under one inch from the badge, and
  // so a disc drawn between two 1.56in tiles was reported as 165% of a 0.146in
  // tile it has no connection to. A badge belongs to its edge, and the edge
  // names its own endpoints.
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const tiers = [
    ['tt-fd', 'Front Door', 0, 0],
    ['tt-app', 'App Service plan', 800, 0],
    ['tt-sql', 'Azure SQL Database', 1600, 0],
    ['tt-waf', 'Web Application Firewall', 0, 260],
    ['tt-fn', 'Function App', 800, 260],
    ['tt-blob', 'Blob Storage account', 1600, 260],
  ] as const;
  const nodes: Node[] = tiers.map(([id, label, x, y]) => ({
    id,
    type: 'azureNode',
    position: { x, y },
    width: 150,
    height: 96,
    data: { label, serviceName: label, category: 'networking', iconPath: icon },
  } as unknown as Node));
  nodes.push({
    id: 'tt-dns',
    type: 'azureNode',
    position: { x: 400, y: 130 },
    width: 14,
    height: 16,
    data: {
      label: 'Private DNS zone',
      serviceName: 'Private DNS zone',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node);
  const edges = [
    ['tt-fd', 'tt-app', 1],
    ['tt-waf', 'tt-fn', 2],
    ['tt-app', 'tt-sql', 3],
    ['tt-fn', 'tt-blob', 4],
  ].map(([source, target, step], i) => ({
    id: `tte${i}`,
    source,
    target,
    data: { stepNumber: step, stepDescription: `Step ${step} of the request path` },
  })) as unknown as Edge[];
  return { id: 'probe-three-tier', nodes, edges };
}

function twoChainBadgeScenario(): Scenario {
  // Two numbered chains on one sheet, at different sizes: a three service
  // cloud pipeline and a six sensor field bus drawn small.
  //
  // A ceiling taken over "the tiles at the ends of the numbered connectors" is
  // still a statistic over the whole sheet, and this is the deck that shows
  // it. Numbering the sensors dragged the pipeline's discs from 0.2400in to
  // 0.1375in - 43% off three badges because of six shapes on the other side of
  // the page - and BOTH badge rules stayed silent, because a ceiling bound
  // badge beside the tile that set it is 55% of it by construction, exactly
  // the bar the oversize rule uses. The ceiling has to be per connector.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  for (let i = 0; i < 3; i += 1) {
    nodes.push({
      id: `tc-c${i}`,
      type: 'azureNode',
      position: { x: i * 500, y: 0 },
      width: 150,
      height: 96,
      data: {
        label: `Cloud pipeline stage ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
  }
  for (let i = 0; i < 6; i += 1) {
    nodes.push({
      id: `tc-d${i}`,
      type: 'azureNode',
      position: { x: i * 140, y: 700 },
      width: 24,
      height: 96,
      data: {
        label: `Field sensor ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
  }
  const edges: Edge[] = [];
  for (let i = 0; i < 5; i += 1) {
    edges.push({
      id: `tcd${i}`,
      source: `tc-d${i}`,
      target: `tc-d${i + 1}`,
      data: { stepNumber: i + 1, stepDescription: `Sensor ${i + 1} forwards its reading` },
    } as unknown as Edge);
  }
  for (let i = 0; i < 2; i += 1) {
    edges.push({
      id: `tcc${i}`,
      source: `tc-c${i}`,
      target: `tc-c${i + 1}`,
      data: { stepNumber: 6 + i, stepDescription: `Pipeline stage ${i + 1} hands off` },
    } as unknown as Edge);
  }
  return { id: 'probe-two-chains', nodes, edges };
}

function numberedSpreadScenario(): Scenario {
  // `probe-tiny-spread` with the step numbers put back.
  //
  // Round 68 removed them and recorded why: at 14px a chip and a legible badge
  // are both wider than the tile, so the fixture could not isolate the one
  // variable it was built for. That reasoning was right about the fixture and
  // wrong about the corpus - it left the ONE deck that could expose a callout
  // wider than its service with the exposing element taken out, and PowerPoint
  // then sized its discs with no reference to the tile at all for four more
  // rounds. Measured at the time: a 0.3566in disc on a 0.2000in tile, 178%,
  // with the PowerPoint gate clean and the same drawing failing 51 times as a
  // Visio sheet.
  //
  // A tile this small has no diameter that is both legible and proportionate,
  // so both exporters draw the legibility floor and the rule stays quiet: the
  // only fix is a coarser split, which is the planner's trade, and the deck
  // must keep every cited step findable on the canvas. `probe-numbered-mid` is
  // the sibling that holds the ceiling itself to account.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 60 }, (_, i) => ({
    id: `ns${i}`,
    type: 'azureNode',
    position: { x: (i % 10) * 60, y: Math.floor(i / 10) * 60 },
    width: 14,
    height: 16,
    data: {
      label: `Regional ingest worker ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges = Array.from({ length: 59 }, (_, i) => ({
    id: `nse${i}`,
    source: `ns${i}`,
    target: `ns${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Worker ${i + 1} forwards to worker ${i + 2}` },
  })) as unknown as Edge[];
  return { id: 'probe-numbered-spread', nodes, edges };
}

function numberedMidSpreadScenario(): Scenario {
  // The same spread at 38px, which is the size that can actually hold the
  // ceiling to account.
  //
  // `probe-numbered-spread` cannot: 55% of a 0.2000in tile is 0.11in, under the
  // 0.18in that holds a digit, so both exporters draw the floor and the rule
  // is correctly silent - a fixture that is exempt cannot witness a regression.
  // At 38px the tile draws about 0.40in, 55% of it is 0.22in, and a legible
  // disc is comfortably inside that, so the ceiling is the only thing standing
  // between the reader and a callout at 65% of the service it points at, which
  // is what PowerPoint drew here for four rounds. The window is real but not
  // wide - under a 0.33in tile no disc is both legible and proportionate, and
  // over a 0.47in one the natural diameter is already inside the bar - so the
  // size is chosen to sit in the middle of it rather than on an edge.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 24 }, (_, i) => ({
    id: `nm${i}`,
    type: 'azureNode',
    position: { x: (i % 6) * 60, y: Math.floor(i / 6) * 60 },
    width: 38,
    height: 40,
    data: {
      label: `Regional ingest worker ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges = Array.from({ length: 23 }, (_, i) => ({
    id: `nme${i}`,
    source: `nm${i}`,
    target: `nm${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Worker ${i + 1} forwards to worker ${i + 2}` },
  })) as unknown as Edge[];
  return { id: 'probe-numbered-mid', nodes, edges };
}

function duplicateLabelFanScenario(): Scenario {
  // Four consumers off one bus, every hop numbered, and every consumer
  // carrying the SAME label.
  //
  // Resolving a badge's tiles by the visible name made this deck invisible.
  // The identical drawing with four distinct labels reported four discs at
  // 76.7% of the tiles they number; renaming the nodes turned the rule off,
  // and worse, the one hop with a resolvable large end was measured against
  // the 1.5625in bus - the end it is not dwarfing - read 7.2%, and was then
  // exempted as floored. Over the corpus that key left 732 of 2258 badges
  // unmeasured. The tile has carried its node id in shape data all along.
  //
  // The consumers are 38px rather than the 14px this was first written with,
  // for the same reason `probe-numbered-mid` is: at 14px no disc is both
  // legible and proportionate, the rule is correctly exempt, and a fixture
  // that is exempt cannot witness anything. At 38px they draw about 0.40in,
  // the rule is live, and the bus is still 150px - so a badge resolved to the
  // wrong end still reads as comfortably proportionate and still hides the
  // defect.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [{
    id: 'dfbus',
    type: 'azureNode',
    position: { x: 0, y: 0 },
    width: 150,
    height: 96,
    data: {
      label: 'Azure Service Bus',
      serviceName: 'Azure Service Bus',
      category: 'integration',
      iconPath: icon,
    },
  } as unknown as Node];
  for (let i = 0; i < 4; i += 1) {
    nodes.push({
      id: `dffn${i}`,
      type: 'azureNode',
      position: { x: 400, y: i * 120 },
      width: 38,
      height: 40,
      data: {
        label: 'Azure Function',
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
  }
  const edges = Array.from({ length: 4 }, (_, i) => ({
    id: `dfe${i}`,
    source: i === 0 ? 'dfbus' : `dffn${i - 1}`,
    target: `dffn${i}`,
    data: { stepNumber: i + 1, stepDescription: `Consumer ${i + 1} takes the message` },
  })) as unknown as Edge[];
  return { id: 'probe-dup-fan', nodes, edges };
}

function hubSpokeCalloutScenario(): Scenario {
  // One wide hub with eight narrow workers hanging off it, every hop numbered.
  //
  // This is the deck the badge rules could not see. Both exporters cap a disc
  // at 55% of the narrower end and raise it to a legibility floor when 55% is
  // under that floor, and both gates then decline to measure any badge sitting
  // on the floor. That pair is not a narrow window, it is an identity: a disc
  // can only exceed the bar by being on the floor, and being on the floor is
  // the exemption. Measured here before the fix, with the gate reporting PASS:
  // 93% of the tile on the overview and 62% on both reading windows, and the
  // Visio sheet at 55% on the same input - the two formats disagreeing by 31
  // points on the slides a reader actually reads.
  //
  // It reaches the other half of the question too. The badge is measured
  // against the NARROWER end, which is the spoke; against the hub it is 9.3%,
  // and with 14px spokes the drawing's disc is 2.7% of the hub and exempted
  // as floored on top. A callout on a hop into a 4.17in tile that is a tenth
  // of an inch across is not proportionate, it is lost.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [{
    id: 'hsc-hub',
    type: 'azureNode',
    position: { x: 900, y: 500 },
    width: 400,
    height: 96,
    data: {
      label: 'Azure Front Door',
      serviceName: 'Azure Front Door',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node];
  const edges: Edge[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push({
      id: `hsc-sp${i}`,
      type: 'azureNode',
      position: { x: (i % 4) * 500, y: i < 4 ? 0 : 1100 },
      width: 40,
      height: 96,
      data: {
        label: `Regional worker ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    edges.push({
      id: `hsce${i}`,
      source: 'hsc-hub',
      target: `hsc-sp${i}`,
      data: { stepNumber: i + 1, stepDescription: `Front Door routes to worker ${i + 1}` },
    } as unknown as Edge);
  }
  return { id: 'probe-hub-spoke', nodes, edges };
}

function numberedEstateScenario(): Scenario {
  // A hundred and twenty ordinary services, ordinary size, every hop numbered.
  //
  // Nothing about this deck is a probe: 150x75 tiles on a 400x260 grid is what
  // the app's own layout engines author. What it exposes is the overview. The
  // badge block reads `perSlide`, which is `allSlides.slice(overviewAt + 1)`,
  // so the first slide of every tiled deck has never been inside it - and the
  // overview is where the tiles get small. Measured before the fix, gate
  // reporting PASS: the overview drew its discs at 56% of a tile at 120
  // services, 67% at 160 and 78% at 240, growing without bound because the
  // disc is pinned to the legibility floor while the tile shrinks with N.
  //
  // The deck's own headline metric missed it by the same slice:
  // `minTileWidthIn` reported 1.178in while the smallest tile the reader is
  // handed is 0.319in, on slide one.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const cols = 11;
  for (let i = 0; i < 120; i += 1) {
    nodes.push({
      id: `es${i}`,
      type: 'azureNode',
      position: { x: (i % cols) * 400, y: Math.floor(i / cols) * 260 },
      width: 150,
      height: 75,
      data: {
        label: `Contoso platform service ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    if (i > 0) {
      edges.push({
        id: `ese${i}`,
        source: `es${i - 1}`,
        target: `es${i}`,
        data: { stepNumber: i, stepDescription: `Stage ${i} hands off to stage ${i + 1}` },
      } as unknown as Edge);
    }
  }
  return { id: 'probe-estate-120', nodes, edges };
}

/**
 * One wide service beside six ordinary narrow ones, every hop numbered.
 *
 * Round 73's refutation of my claim that the `max`-denominator undersize rule
 * was inert. It is not inert, it is wrong: every disc here is drawn at its
 * NATURAL size - 0.2063in, strictly between a 0.1556in floor and a 0.2728in
 * ceiling - and the rule reported three of them at 7.8% because a DIFFERENT
 * shape on the same slide is wide.
 *
 * The trigger is scale-free, which is the same shape of law as the vacuous bar
 * round 72 removed: `natural = 0.26 * px` and `px = 96 * scale`, so
 * `badge / widest = 24.96 / Wmax` in authored pixels whatever the transform
 * does. It crosses 10% at 250 authored px. A full-width Front Door or
 * Application Gateway banner beside 180px services is an ordinary Architecture
 * Center shape, so the false positive was one wide node away on real input.
 */
function wideHubCalloutScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10076-icon-service-Application-Gateways.svg';
  const nodes: Node[] = [{
    id: 'wh-hub',
    type: 'azureNode',
    position: { x: 600, y: 400 },
    width: 320,
    height: 96,
    data: {
      label: 'Contoso edge application gateway',
      serviceName: 'Azure Application Gateway',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node];
  const edges: Edge[] = [];
  for (let i = 0; i < 6; i += 1) {
    nodes.push({
      id: `wh-sp-${i}`,
      type: 'azureNode',
      position: { x: (i % 3) * 420, y: i < 3 ? 0 : 900 },
      width: 60,
      height: 96,
      data: {
        label: `Workload ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    edges.push({
      id: `whe-${i}`,
      source: 'wh-hub',
      target: `wh-sp-${i}`,
      data: { stepNumber: i + 1, stepDescription: `Route ${i + 1} leaves the gateway` },
    } as unknown as Edge);
  }
  return { id: 'probe-wide-hub', nodes, edges };
}

/**
 * A hundred and twenty slivers on an ordinary grid, numbered past a hundred.
 *
 * Round 73's second refutation. The callout bar `markableTileWIn` raises the
 * planner to is not reachable below about 17 authored pixels, because
 * `legibleScaleFor` takes `finestPerIn = min(frame) / (2 * WINDOW_BLEED_PX +
 * target)` and a flat 100px bleed caps a 14px tile at 0.3740in however many
 * windows are spent. A three-digit callout needs 0.4457in. So the planner
 * spends 55 windows, saturates below the bar, and the gate reports the
 * shortfall as though a coarser split were available.
 *
 * The step numbers are the only thing that changes: unnumbered this drawing is
 * 14 slides of 0.2000in tiles, one-digit 37 slides, two-digit 53, three-digit
 * 65 - and only the last one failed, on the edge that crosses from 99 to 100.
 */
function sliverEstateScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const cols = 11;
  for (let i = 0; i < 120; i += 1) {
    nodes.push({
      id: `sv${i}`,
      type: 'azureNode',
      position: { x: (i % cols) * 120, y: Math.floor(i / cols) * 120 },
      width: 14,
      height: 16,
      data: {
        label: `Edge sensor ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    if (i > 0) {
      edges.push({
        id: `sve-${i}`,
        source: `sv${i - 1}`,
        target: `sv${i}`,
        data: { stepNumber: i, stepDescription: `Sensor ${i} forwards to sensor ${i + 1}` },
      } as unknown as Edge);
    }
  }
  return { id: 'probe-sliver-120', nodes, edges };
}

/**
 * The same estate one authored pixel wider, and one wider again.
 *
 * Round 74 found two separate cliffs in that pixel. `probe-band-15` is where
 * the exporter clamps and a gate that modelled the frame without its
 * connection legend did not, failing a correct 21 slide deck 98 times.
 * `probe-band-16` is where the clamp itself used to miss: it compared against
 * the asymptote instead of `finestPerIn`, leaving a band about one pixel wide
 * in which the bar is unreachable and unclamped - 15px planned 21 slides, 16px
 * planned 64, and the 0.4274in tile that bought was still short of the
 * 0.4457in bar it was spent chasing.
 */
function bandEstateScenario(px: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const cols = 11;
  for (let i = 0; i < 120; i += 1) {
    nodes.push({
      id: `bd${i}`,
      type: 'azureNode',
      position: { x: (i % cols) * 120, y: Math.floor(i / cols) * 120 },
      width: px,
      height: 16,
      data: {
        label: `Edge sensor ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    if (i > 0) {
      edges.push({
        id: `bde-${i}`,
        source: `bd${i - 1}`,
        target: `bd${i}`,
        data: { stepNumber: i, stepDescription: `Sensor ${i} forwards to sensor ${i + 1}` },
      } as unknown as Edge);
    }
  }
  return { id: `probe-band-${px}`, nodes, edges };
}

/**
 * A bimodal estate: 20px tiles with a 14px sliver every thirtieth node.
 *
 * The gate used to mirror the planner's clamp with a MINIMUM where the planner
 * uses a MEDIAN, and on this drawing the two disagree: the median tile clears
 * its bar at 0.4471in, the plan is correct at 53 slides, and the gate accused
 * it of chasing a bar it had already reached while suppressing the four real
 * conflicts on the 0.3130in slivers. It is also the fixture on which the Visio
 * magnifier's own median declined half an available move.
 */
function blindSliverScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const cols = 11;
  for (let i = 0; i < 120; i += 1) {
    const sliver = i % 30 === 0;
    nodes.push({
      id: `bs${i}`,
      type: 'azureNode',
      position: { x: (i % cols) * 120, y: Math.floor(i / cols) * 120 },
      width: sliver ? 14 : 20,
      height: sliver ? 16 : 20,
      data: {
        label: `Edge sensor ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
    if (i > 0) {
      edges.push({
        id: `bse-${i}`,
        source: `bs${i - 1}`,
        target: `bs${i}`,
        data: { stepNumber: i, stepDescription: `Sensor ${i} forwards to sensor ${i + 1}` },
      } as unknown as Edge);
    }
  }
  return { id: 'probe-blind-sliver', nodes, edges };
}

/**
 * Six ordinary services and one small icon in the numbered flow.
 *
 * The shape every Architecture Center reference draws, with a private DNS zone
 * sized like the glyph it is. The Visio magnifier's median is 150px here, so
 * it declined to move at all and left two discs at 77% of the zone they point
 * at, when the move it needed was a 15.5in sheet instead of an 11.1in one.
 */
function mixedSliverScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const names = ['Azure Front Door', 'Azure App Service', 'Azure SQL Database',
    'Azure Key Vault', 'Azure Blob Storage', 'Azure Monitor'];
  const nodes: Node[] = names.map((label, i) => ({
    id: `mx${i}`,
    type: 'azureNode',
    position: { x: (i % 3) * 400, y: Math.floor(i / 3) * 300 },
    width: 150,
    height: 96,
    data: { label, serviceName: label, category: 'networking', iconPath: icon },
  } as unknown as Node));
  nodes.push({
    id: 'mx-dns',
    type: 'azureNode',
    position: { x: 200, y: 160 },
    width: 14,
    height: 16,
    data: {
      label: 'Private DNS zone',
      serviceName: 'DNS Zones',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node);
  const hops: Array<[string, string]> = [
    ['mx0', 'mx1'], ['mx1', 'mx-dns'], ['mx-dns', 'mx2'],
    ['mx2', 'mx3'], ['mx3', 'mx4'], ['mx4', 'mx5'],
  ];
  const edges: Edge[] = hops.map(([source, target], i) => ({
    id: `mxe-${i}`,
    source,
    target,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the request path` },
  } as unknown as Edge));
  return { id: 'probe-mixed-sliver', nodes, edges };
}

/**
 * A glyph-sized service in a numbered chain of ordinary ones, at three pitches
 * and two rows.
 *
 * The shape the Architecture Center draws constantly - six services in a row
 * and a private DNS zone sized like the glyph it is - and the deck draws the
 * zone's step disc WIDER THAN THE ZONE. At a 290px pitch the tile comes out
 * 0.1293in under a 0.1556in disc, 120% of the thing it numbers; move the same
 * zone off the row to x=1740 and it is 0.1178in under the same disc, 132%.
 *
 * Every gate rule missed it. The proportionality rule cannot arm, because a
 * tile that small can never satisfy `tile * BADGE_SHARE >= floor`. The conflict
 * rule armed and was then exempted, because the zone is narrower than the
 * planner's served tile.
 *
 * A plan exists and costs four slides: `probe-inline` is the same seven
 * services at the same pitch with the zone on the row, and it reaches 0.2829in
 * with its worst disc at 16%. Visio draws every one of these correctly at 55%.
 */
function spreadGlyphScenario(id: string, pitch: number, dnsX: number, dnsY: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const names = ['Azure Front Door', 'Azure App Service', 'Azure SQL Database',
    'Azure Key Vault', 'Azure Blob Storage', 'Azure Monitor'];
  const nodes: Node[] = names.map((label, i) => ({
    id: `sv${i}`,
    type: 'azureNode',
    position: { x: i * pitch, y: 0 },
    width: 150,
    height: 96,
    data: { label, serviceName: label, category: 'networking', iconPath: icon },
  } as unknown as Node));
  nodes.push({
    id: 'sv-dns',
    type: 'azureNode',
    position: { x: dnsX, y: dnsY },
    width: 16,
    height: 96,
    data: {
      label: 'Private DNS zone',
      serviceName: 'DNS Zones',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node);
  const chain = ['sv0', 'sv1', 'sv-dns', 'sv2', 'sv3', 'sv4', 'sv5'];
  const edges: Edge[] = chain.slice(1).map((target, i) => ({
    id: `sve-${i}`,
    source: chain[i],
    target,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the request path` },
  } as unknown as Edge));
  return { id, nodes, edges };
}

/**
 * Twenty ordinary services and one glyph, on a chain that runs THROUGH it.
 *
 * The reviewer's isolation of the exempt-band defect, kept because it is the
 * only shape in the corpus that puts a callout beside a service two orders of
 * magnitude narrower than its neighbours without tripping any other rule -
 * twenty full-size tiles keep the naming rule and the density rules quiet, so
 * whatever the callout rule says here is the callout rule's own verdict.
 *
 * `reachableTileW` was derived from `markIn` and then used to guard a clause
 * that fires on `floor`, and `markIn = floor / BADGE_SHARE`, so the bound came
 * out 1.82x too generous: every authored width in the top 45% of the exempt
 * band was waved through while the frame was demonstrably drawing that tile
 * wider than the disc. At 16 authored px the frame reaches 0.2978in against a
 * 0.1556in disc, 1.91x the headroom needed, and the plan that would have taken
 * it was refused by the density floor rather than by the page.
 */
/**
 * Two glyph-sized services a hair apart, so the callout lands ON them.
 *
 * The liveness proof for the contact term. Every other fixture that satisfies
 * `badge.w > tile` draws its disc in clear space between two icons, which is
 * why requiring contact silenced them; this one puts the hop so short that the
 * disc has nowhere to sit but on top of the services it numbers, which is the
 * shape round 71 was opened for - a disc drawn OVER the icon at 178%. Without
 * it the contact term would be indistinguishable from switching the rule off.
 */
function touchingBadgeScenario(id: string, pairs: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const nodes: Node[] = Array.from({ length: pairs * 2 }, (_, i) => ({
    id: `tb${i}`,
    type: 'azureNode',
    position: { x: (i % 2) * 34 + Math.floor(i / 2) * 900, y: Math.floor(i / 2) * 260 },
    width: 18,
    height: 150,
    data: {
      label: `Edge relay probe ${i + 1}`,
      serviceName: 'Traffic Manager',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: pairs }, (_, i) => ({
    id: `tbe-${i}`,
    source: `tb${i * 2}`,
    target: `tb${i * 2 + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the relay path` },
  } as unknown as Edge));
  return { id, nodes, edges };
}

function glyphChainScenario(id: string, glyphW: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const nodes: Node[] = Array.from({ length: 20 }, (_, i) => ({
    id: `rg${i}`,
    type: 'azureNode',
    position: { x: i * 700, y: 0 },
    width: 150,
    height: 110,
    data: {
      label: `Azure regional gateway ${i + 1}`,
      serviceName: 'Application Gateway',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node));
  nodes.push({
    id: 'rg-dns',
    type: 'azureNode',
    position: { x: 350, y: 0 },
    width: glyphW,
    height: 110,
    data: {
      label: 'Private DNS zone',
      serviceName: 'DNS Zones',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node);
  const chain = ['rg0', 'rg-dns', ...Array.from({ length: 19 }, (_, i) => `rg${i + 1}`)];
  const edges: Edge[] = chain.slice(1).map((target, i) => ({
    id: `rge-${i}`,
    source: chain[i],
    target,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the regional path` },
  } as unknown as Edge));
  return { id, nodes, edges };
}

/**
 * The same drawing at five, six and seven services.
 *
 * `planWindowsAtCeiling` prefers the COMFORTABLE grid whenever the deck can
 * afford it and it carries `MIN_SERVICES_PER_SLIDE` services, and the
 * comfortable grid is derived from the median tile's HEIGHT alone. So a drawing
 * whose legibility floor demands a finer grid than comfort does gets the coarse
 * one anyway - and then returns `{ windows: [], legible: true }`, "it fits
 * whole", for a scale it decided two lines earlier does not fit.
 *
 * The bounding box, the node sizes and the sliver are identical across all
 * three; only the count of ordinary services changes, and it crosses
 * `MIN_SERVICES_PER_SLIDE` between the first and the second. Adding one
 * ordinary service makes the sliver's tile narrower, its disc worse, and
 * switches the planner's own report of affordability from true to false.
 */
function mixCountScenario(bigCount: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const nodes: Node[] = [];
  const ids: string[] = [];
  for (let i = 0; i < bigCount; i += 1) {
    const x = bigCount > 1 ? Math.round(((950 - 150) * i) / (bigCount - 1)) : 0;
    nodes.push({
      id: `mc${i}`,
      type: 'azureNode',
      position: { x, y: 0 },
      width: 150,
      height: 96,
      data: {
        label: `Regional service ${i + 1}`,
        serviceName: 'Virtual Networks',
        category: 'networking',
        iconPath: icon,
      },
    } as unknown as Node);
    ids.push(`mc${i}`);
  }
  nodes.push({
    id: 'mc-dns',
    type: 'azureNode',
    position: { x: bigCount > 1 ? Math.round((950 - 150) / (bigCount - 1) / 2) : 60, y: 300 },
    width: 14,
    height: 96,
    data: {
      label: 'Private DNS zone',
      serviceName: 'DNS Zones',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node);
  const chain = [...ids.slice(0, 1), 'mc-dns', ...ids.slice(1)];
  const edges: Edge[] = chain.slice(1).map((target, i) => ({
    id: `mce-${i}`,
    source: chain[i],
    target,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the request path` },
  } as unknown as Edge));
  return { id: `probe-mix-${bigCount}`, nodes, edges };
}

/**
 * Two regions with a wide void between them, numbered.
 *
 * The magnifier's paper bound used to be measured on a span that still carried
 * this void, because gutter compaction ran after it. Above roughly 5754px of
 * separation `roomK` fell below 1 and the magnifier declined a move it plainly
 * had - and then compaction deleted the void anyway and shipped a 14.8in sheet
 * with 45in of the budget unused and seven discs at 77% of their tiles. The
 * larger the gap, the LESS magnification: 700px of void magnified 1.395, 4000px
 * magnified 1.282, and 7000px magnified not at all.
 */
function gutterRegionScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const nodes: Node[] = Array.from({ length: 8 }, (_, i) => ({
    id: `gk${i}`,
    type: 'azureNode',
    position: { x: i < 4 ? i * 160 : 7000 + (i - 4) * 160, y: 0 },
    width: 14,
    height: 30,
    data: {
      label: `Regional endpoint ${i + 1}`,
      serviceName: 'Virtual Networks',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: 7 }, (_, i) => ({
    id: `gke-${i}`,
    source: `gk${i}`,
    target: `gk${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the cross-region path` },
  } as unknown as Edge));
  return { id: 'probe-gutter-region', nodes, edges };
}

/**
 * A gutter that is under the void bar when it is measured and over it once the
 * drawing is magnified.
 *
 * `compactEmptyGutters` closes a band wider than `VOID_GUTTER_PX` (320 authored
 * px). Two clusters 250px apart survive that test, and then the callout
 * magnifier scales the whole drawing by 1.395 and hands the sheet a 349px void
 * it would have closed. This is the case the SECOND compaction exists for -
 * every other numbered fixture magnifies a drawing whose surviving gaps are far
 * enough under the bar to stay under it.
 */
function magnifiedGutterScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const xs = [0, 160, 320, 584, 744, 904, 1064, 1224];
  const nodes: Node[] = xs.map((x, i) => ({
    id: `mg${i}`,
    type: 'azureNode',
    position: { x, y: 0 },
    width: 14,
    height: 30,
    data: {
      label: `Segment endpoint ${i + 1}`,
      serviceName: 'Virtual Networks',
      category: 'networking',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: xs.length - 1 }, (_, i) => ({
    id: `mge-${i}`,
    source: `mg${i}`,
    target: `mg${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the segment path` },
  } as unknown as Edge));
  return { id: 'probe-magnified-gutter', nodes, edges };
}

/**
 * The same four services, drawn with and without zone rectangles around them.
 *
 * The gate used to take its "did the planner decline to serve this tile"
 * median over every node in the scenario, groups included, so adding two zones
 * moved the median from 24px to 150px and silenced a real 56% conflict without
 * moving one drawn pixel. Both members must report identically.
 */
function zoneMedianScenario(zones: number): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const sizes: Array<[number, number]> = [[150, 96], [24, 24], [24, 24], [24, 24]];
  const nodes: Node[] = sizes.map(([width, height], i) => ({
    id: `zm${i}`,
    type: 'azureNode',
    position: { x: i * 900, y: 0 },
    width,
    height,
    data: {
      label: `Workload component ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  for (let z = 0; z < zones; z += 1) {
    nodes.push({
      id: `zmz${z}`,
      type: 'groupNode',
      position: { x: z * 1800, y: -80 },
      width: 700,
      height: 300,
      data: { label: `Subscription ${z + 1}`, zoneKind: 'subscription' },
    } as unknown as Node);
  }
  const edges: Edge[] = Array.from({ length: 3 }, (_, i) => ({
    id: `zme-${i}`,
    source: `zm${i}`,
    target: `zm${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the workload path` },
  } as unknown as Edge));
  return { id: `probe-zone-median-${zones}`, nodes, edges };
}

/**
 * Half the services below the median, numbered.
 *
 * The "the planner declined to serve this tile" exemption has no cost term, and
 * `sorted[floor(n/2)]` puts up to half the drawing below the line by
 * construction. Here it covered two of four services and all three hops, while
 * the plan that reaches the bar costs exactly ONE extra window - and widens
 * every other tile on the deck by 39% in the bargain.
 */
function halfTailScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const sizes: Array<[number, number]> = [[150, 96], [24, 24], [150, 96], [24, 24]];
  const nodes: Node[] = sizes.map(([width, height], i) => ({
    id: `ht${i}`,
    type: 'azureNode',
    position: { x: i * 500, y: 0 },
    width,
    height,
    data: {
      label: `Pipeline stage ${i + 1}`,
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
    },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: 3 }, (_, i) => ({
    id: `hte-${i}`,
    source: `ht${i}`,
    target: `ht${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1} of the pipeline` },
  } as unknown as Edge));
  return { id: 'probe-half-tail', nodes, edges };
}

function bimodalSidecarScenario(): Scenario {
  // Six services at a normal size with a sidecar of twenty four tiny ones.
  //
  // The sidecar is the majority, so a width taken as an extreme over the sheet
  // reads 14px and a width taken as a median reads 14px too - what separates
  // them is that the median is not moved by a single node, and this deck is
  // here to hold the median honest. Before the planner learned to raise its
  // own ceiling this drew 10 slides with 6 of 30 services named; it now draws
  // 18 with all 30 named, and the density floor accepts the raise at 1.875
  // services a window.
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = [];
  for (let i = 0; i < 6; i += 1) {
    nodes.push({
      id: `bs-big${i}`,
      type: 'azureNode',
      position: { x: (i % 3) * 400, y: Math.floor(i / 3) * 300 },
      width: 150,
      height: 96,
      data: {
        label: `Core platform service ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
  }
  for (let i = 0; i < 24; i += 1) {
    nodes.push({
      id: `bs-tiny${i}`,
      type: 'azureNode',
      position: { x: 1600 + (i % 6) * 260, y: Math.floor(i / 6) * 260 },
      width: 14,
      height: 16,
      data: {
        label: `Sidecar collector ${i + 1}`,
        serviceName: 'Azure Functions',
        category: 'compute',
        iconPath: icon,
      },
    } as unknown as Node);
  }
  const edges = Array.from({ length: 5 }, (_, i) => ({
    id: `bse${i}`,
    source: `bs-big${i}`,
    target: `bs-big${i + 1}`,
  })) as unknown as Edge[];
  return { id: 'probe-bimodal-sidecar', nodes, edges };
}

function decomposedNameScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const dense = [
    // Arabic, fully vocalised: ten marks in thirty characters.
    'مَرْكَز البَيَانَات المُشْتَرَك',
    'خِدْمَة التَّخْزِين المُدَارَة',
    'شَبَكَة خَاصَّة مُؤَمَّنَة',
    // Thai, where the vowel signs stack above and below the consonant.
    'บริการจัดเก็บข้อมูลที่ใช้ร่วมกัน',
    'ระบบเครือข่ายส่วนตัวที่ปลอดภัย',
    'บริการรักษาความปลอดภัยของข้อมูล',
    // Yoruba, the case that SURVIVES composition. `ẹ̀` is U+1EB9 followed by
    // U+0300 and there is no single code point for it, so NFC leaves the mark
    // standing on a Latin base - which is the only remaining way a combining
    // mark can reach the exporter now that every entry point composes. The
    // table prices these marks at half an em when they are not recognised as
    // marks, and that is a Latin-width error on a Latin name.
    'Ẹ̀rọ ìsopọ̀ àwọn ìlú',
    'Ìṣàkóso ìpamọ́ dátà',
    'Ẹ̀rọ ìdánilójú ààbọ̀',
    'Ìtọ́jú ìwọ̀n ìlànà',
    // Cyrillic with the stress mark, the same class and a commoner one. U+0301
    // has no precomposed form on any Cyrillic base, so `Москва́` survives NFC
    // exactly as the Yoruba does. The class is broader than either: Lithuanian
    // and Navajo `ą́ ę́ į́ ǫ́`, Igbo `ị́ ọ́ ụ́`, and any triple-mark stack. All of
    // them are priced correctly by construction, because a mark costs zero and
    // the base is tabled - so the fixture's job is only to keep that path
    // reachable, and two scripts do it more convincingly than one.
    'Москва́ регион узел',
    'Хранилище да́нных',
  ];
  const nodes: Node[] = dense.map((label, i) => ({
    id: `nd${i}`,
    type: 'azureNode',
    position: { x: (i % 2) * 220, y: Math.floor(i / 2) * 120 },
    width: [44, 60, 76, 104, 132, 168, 52, 68, 92, 120, 84, 112][i],
    height: 30,
    data: { label, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges = [
    {
      id: 'nde1', source: 'nd0', target: 'nd1', label: 'يَنْسَخ',
      data: { stepNumber: 1, stepDescription: 'يَنْسَخ مَرْكَز البَيَانَات السِّجِلَّات إِلَى المِنْطَقَة التَّوْأَم.' },
    },
    {
      id: 'nde2', source: 'nd3', target: 'nd4', label: 'จัดเก็บ',
      data: { stepNumber: 2, stepDescription: 'บริการนี้จัดเก็บข้อมูลเก่าไว้ในที่จัดเก็บแบบเย็น' },
    },
  ] as unknown as Edge[];
  return { id: 'probe-nfd', nodes, edges };
}

/**
 * Round 66, issue 1: Turkish and Vietnamese names, authored decomposed.
 *
 * Pricing a combining mark at zero was correct and was not enough. The font
 * gives a precomposed glyph and its base different advances - "ş" is 0.5070
 * composed and 0.4240 decomposed, a 16.4% gap - so the same visible name still
 * measured differently depending on where it was typed, and BOTH models read
 * the same table, so the divergence rule agreed with the exporter. At 48x24
 * one spelling drew "Payla...si" and the other drew a bare "4".
 *
 * Authored decomposed only, with no composed twin. The two spellings are the
 * SAME NAME, so a diagram carrying both is a diagram with duplicate services
 * and every rule about distinct names fires on the fixture rather than on the
 * defect. What is checked here is single-file and observable: whatever was
 * authored, nothing decomposed may reach the slide. The counterfactual half -
 * that the two spellings produce identical decks - is a differential between
 * two files and lives in `tests/glyphAdvances.test.ts`.
 *
 * Turkish on purpose. The round-65 test used four names whose accents happen
 * to be in the class where the table agrees with itself, and passed - proving
 * the sample rather than the property. These are from the other class.
 */
function normFormScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10061-icon-service-Virtual-Networks.svg';
  const names = [
    'Şişli şube şebeke sunucusu',
    'Paylaşılan şifreleme şeması',
    'Şirket şubesi bağlantı noktası',
    'Bağlantı ağ geçidi dağıtım',
  ];
  // Vietnamese, for the same reason and a different cause: the two bare horn
  // vowels are the one Latin gap in the measured table, and no name can be
  // written in Vietnamese without them.
  const viet = [
    'Dịch vụ lưu trữ đối tượng',
    'Cổng kết nối riêng tư',
    'Tường lửa ứng dụng web',
  ];
  const labels = [...names, ...viet].map((name) => name.normalize('NFD'));
  const nodes: Node[] = labels.map((label, i) => ({
    id: `nf${i}`,
    type: 'azureNode',
    position: { x: (i % 2) * 420, y: Math.floor(i / 2) * 260 },
    // Wide enough that a realistic two-word edge label is not automatically
    // dominant. The tiles were 48x24 and the labels were cut to three letters
    // to stop the dominance rule firing - which is the one change that stops
    // this fixture exercising the thing it exists for, since the subject here
    // is how a name is SPELLED, not how small a tile can get. Widen the tile
    // and keep the real label instead.
    width: 200,
    height: 120,
    data: { label, serviceName: 'Virtual Network', category: 'networking', iconPath: icon },
  } as unknown as Node));
  const edges = [
    {
      id: 'nfe1', source: 'nf0', target: 'nf1', label: 'veri aktarımı'.normalize('NFD'),
      data: { stepNumber: 1, stepDescription: 'Şirket şubesi şifreleme şemasını paylaşır.'.normalize('NFD') },
    },
    {
      id: 'nfe2', source: 'nf4', target: 'nf5', label: 'sao lưu dữ liệu'.normalize('NFD'),
      data: { stepNumber: 2, stepDescription: 'Dịch vụ lưu trữ đối tượng ghi dữ liệu vào vùng lưu trữ lạnh.' },
    },
  ] as unknown as Edge[];
  return { id: 'probe-normform', nodes, edges };
}

function briefWorkflowStepsScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const nodes: Node[] = Array.from({ length: 13 }, (_, i) => ({
    id: `bw${i}`,
    type: 'azureNode',
    position: { x: (i % 3) * 130, y: Math.floor(i / 3) * 90 },
    width: 90,
    height: 50,
    data: { label: `Service ${i}`, serviceName: 'Azure Functions', category: 'compute', iconPath: icon },
  } as unknown as Node));
  const edges: Edge[] = Array.from({ length: 12 }, (_, k) => ({
    id: `bw-e${k + 1}`,
    source: `bw${k}`,
    target: `bw${k + 1}`,
    data: { stepNumber: k + 1, stepDescription: 'ack' },
  } as unknown as Edge));
  return { id: 'probe-brief-steps', nodes, edges };
}

function hairlineTilesScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/compute/10029-icon-service-Function-Apps.svg';
  const widths = [8, 10, 14, 22, 40, 160, 160, 160];
  // SIX DIFFERENT NAMES, not one repeated eight times.
  //
  // A shared label is a general anaesthetic for two of the best rules in this
  // file. The cross-format naming rule asks whether a name appears in both
  // drawings and the "name is nowhere on the page" rule asks whether it appears
  // at all - and when every tile carries the same string, ONE tile that draws it
  // satisfies both questions on behalf of the seven that did not. The fixture
  // was reporting the corpus clean by giving every rule the same answer eight
  // times.
  const names = [
    'Zephyr order intake function', 'Quartz billing reconciliation',
    'Nimbus telemetry ingestion hub', 'Cobalt fraud scoring service',
    'Verdant analytics warehouse', 'Onyx configuration store',
    'Amber checkout orchestrator', 'Slate inventory projection',
  ];
  const nodes: Node[] = widths.map((width, i) => ({
    id: `hair${i}`,
    type: 'azureNode',
    position: { x: i * 260, y: (i % 2) * 200 },
    width,
    height: width < 60 ? 30 : 110,
    data: {
      label: names[i],
      serviceName: 'Azure Functions',
      category: 'compute',
      iconPath: icon,
      sku: 'EP2',
      region: 'japaneast',
    },
  } as unknown as Node));
  const edges: Edge[] = widths.slice(1).map((_, i) => ({
    id: `hair-e${i}`,
    source: `hair${i}`,
    target: `hair${i + 1}`,
    label: 'invokes',
  } as Edge));
  return { id: 'hairline-tiles', nodes, edges };
}

/**
 * Tiles that are tall and narrow rather than the editor's default 2:1.
 *
 * `AzureNode` has no resize handle — the only `NodeResizer` is on groups — so
 * a non-2:1 tile arrives from loaded JSON, an MCP scene or a generated
 * blueprint, and never from dragging. That is why the corpus had none, and
 * why a width guard keyed on the HEIGHT-derived font went unmeasured: past
 * h = 1.0833in the four-character bar saturates at 0.7222in, so a 0.7813 x
 * 3.1250in tile — two and a half square inches — carried no text at all,
 * missing the bar by 0.0009in, while at the 7pt floor its column sets 7.9
 * capitals per line with 33 lines of room.
 *
 * The last two nodes carry no icon, which is the branch that used to fall
 * back to the OVERVIEW's 6pt floor on a full-size reading slide.
 */
function tallNarrowTilesScenario(): Scenario {
  const icon = '/Azure_Public_Service_Icons/Icons/networking/10063-icon-service-Firewalls.svg';
  const shapes: { w: number; h: number; icon: boolean }[] = [
    { w: 150, h: 75, icon: true },
    { w: 75, h: 300, icon: true },
    { w: 60, h: 200, icon: true },
    { w: 100, h: 400, icon: true },
    { w: 75, h: 300, icon: false },
    { w: 40, h: 120, icon: false },
    // The stub ladder. An icon-less tile is refused a name by `namedWidth` and
    // then drew one anyway with no column test: at 14px and 10px the widest
    // glyph in the drawn string is the ellipsis, which is wider than the whole
    // column, so nothing could set on one line at all.
    { w: 22, h: 120, icon: false },
    { w: 14, h: 120, icon: false },
    { w: 10, h: 200, icon: false },
    // Wide enough to keep its name, short enough that fitting it at 6pt and
    // painting it at 7pt does not agree: 0.84in of box is 8 lines at 6pt and
    // 7 at 7pt, so the eighth line is drawn below the box it was measured in.
    { w: 60, h: 191, icon: false },
  ];
  const nodes: Node[] = shapes.map((shape, i) => ({
    id: `tall${i}`,
    type: 'azureNode',
    position: { x: i * 320, y: 0 },
    width: shape.w,
    height: shape.h,
    data: {
      label: 'Azure Firewall Premium',
      serviceName: 'Azure Firewall',
      category: 'networking',
      ...(shape.icon ? { iconPath: icon } : {}),
      sku: 'Premium',
      region: 'japaneast',
    },
  } as unknown as Node));
  const edges: Edge[] = shapes.slice(1).map((_, i) => ({
    id: `tall-e${i}`,
    source: `tall${i}`,
    target: `tall${i + 1}`,
    label: 'inspects',
  } as Edge));
  return { id: 'tall-narrow-tiles', nodes, edges };
}

function unlabelledStepInflationScenario(): Scenario {
  const base = tightSubnetsScenario();
  // Deliberately along the same grain as the labelled hops, so the drawing has
  // no routing complaint of its own and the only thing under test is how many
  // numbers these arrows consume.
  const pairs: Array<[string, string]> = [
    ['web-1', 'app-1'], ['web-2', 'app-2'], ['app-2', 'data-2'], ['app-3', 'data-3'],
  ];
  const plain = pairs.map(([source, target], i) => ({ id: `plain-${i}`, source, target } as Edge));
  return { id: 'unlabelled-step-inflation', nodes: base.nodes, edges: [...base.edges, ...plain] };
}

/**
 * The same stack with the subnets sharing their edges.
 *
 * `stacked-subnets` and `tight-subnets` both leave a 23px gutter between tiers,
 * so a zone name that drifts off its own band lands in blank paper and no rule
 * about *whose* box it landed in can fire. Subnets drawn flush inside a virtual
 * network share an edge — it is how the Architecture Center draws them — and
 * then the paper above a band is not blank, it belongs to the tier above.
 */
function flushSubnetsScenario(): Scenario {
  const nodes: Node[] = [];
  const rows: Array<[string, string, number]> = [
    ['fweb', 'Web subnet', 3],
    ['fapp', 'Application subnet', 3],
    ['fdata', 'Data subnet', 3],
  ];
  const names = [
    'Azure Application Gateway', 'Azure App Service', 'Azure Functions',
    'Azure SQL Database', 'Azure Cosmos DB', 'Azure Key Vault',
  ];
  rows.forEach(([id, label, count], tier) => {
    nodes.push(grp(id, label, 0, tier * 95, 620, 95));
    for (let i = 0; i < count; i += 1) {
      nodes.push(svc(`${id}-${i}`, names[(tier * 2 + i) % names.length], 10 + i * 200, 10, id));
    }
  });
  const edges: Edge[] = [
    { id: 'fs1', source: 'fweb-0', target: 'fapp-0', label: 'Forwards' } as Edge,
    { id: 'fs2', source: 'fapp-0', target: 'fdata-0', label: 'Queries' } as Edge,
  ];
  return { id: 'flush-subnets', nodes, edges };
}

/**
 * A long diagonal cascade — every hop stepping down and across, the shape a
 * hand-dragged flow takes once it outgrows a screen.
 *
 * Nothing here is an outlier and no band is empty on either projection, so
 * neither trimming nor gutter compaction has anything to remove: the drawing
 * really is this large, and the only lever left is how many slides it is shown
 * on. The fixed-page deck used to compute the grid that would make it readable,
 * find it past the shared slide ceiling, and throw it away in favour of a grid
 * that reads at four points — which is what the customer deck then shipped.
 *
 * The 27-service variant is the same shape one size larger, and it is here
 * because the cell cap binds before the slide ceiling does. Stepping the grid
 * toward a square took the axis a diagonal is long in, so this drawing once
 * came out at 6.0pt on *fewer* slides than the 26-service one at 6.6pt: adding
 * a service made the deck both shorter and less readable, which is a plan
 * nobody would choose on purpose.
 */
function diagonalCascadeScenario(count = 16, id = 'diagonal-cascade'): Scenario {
  const names = [
    'Azure Front Door', 'Application Gateway', 'Azure App Service', 'Azure Functions',
    'Azure Service Bus', 'Azure SQL Database', 'Azure Cosmos DB', 'Azure Data Factory',
    'Azure Synapse Analytics', 'Azure Blob Storage', 'Azure Key Vault', 'Azure Monitor',
    'Azure Cache for Redis', 'Azure Event Hubs', 'Azure Logic Apps', 'Azure API Management',
  ];
  const nodes: Node[] = Array.from({ length: count }, (_, i) => svc(`d-${i}`, names[i % names.length], i * 900, i * 620));
  const edges: Edge[] = Array.from({ length: count - 1 }, (_, i) => ({
    id: `d-e-${i}`,
    source: `d-${i}`,
    target: `d-${i + 1}`,
    label: 'Hands off',
  } as Edge));
  return { id, nodes, edges };
}

/**
 * Two regions with the corridor between them labelled.
 *
 * "ExpressRoute circuit", "Internet", "On-premises", "Customer boundary" — the
 * Architecture Center labels the space between regions as often as it labels
 * the regions, and the editor makes one in a single click. It is by
 * construction a childless box standing in the widest empty band of the
 * drawing, which is exactly the band an exporter wants to remove: judging
 * emptiness by services alone crushed a 900px corridor to a 1px vertical line.
 */
function corridorZoneScenario(): Scenario {
  const nodes: Node[] = [
    ...Array.from({ length: 6 }, (_, i) => svc(
      `p-${i}`,
      ['Azure Front Door', 'Azure App Service', 'Azure SQL Database'][i % 3],
      (i % 3) * 200,
      Math.floor(i / 3) * 180,
    )),
    grp('link', 'ExpressRoute circuit', 2600, 60, 900, 240),
    ...Array.from({ length: 6 }, (_, i) => svc(
      `d-${i}`,
      ['Azure Traffic Manager', 'Azure Functions', 'Azure Cosmos DB'][i % 3],
      6000 + (i % 3) * 200,
      Math.floor(i / 3) * 180,
    )),
  ];
  const edges: Edge[] = [
    { id: 'k1', source: 'p-0', target: 'p-1', label: 'Routes' } as Edge,
    { id: 'k2', source: 'p-2', target: 'd-2', label: 'Replicates' } as Edge,
    { id: 'k3', source: 'd-0', target: 'd-1', label: 'Serves' } as Edge,
  ];
  return { id: 'corridor-zone', nodes, edges };
}

function strayZonePairScenario(): Scenario {
  const nodes: Node[] = [
    ...Array.from({ length: 12 }, (_, i) => svc(
      `p-${i}`,
      `Primary service ${i}`,
      (i % 4) * 220,
      Math.floor(i / 4) * 180,
    )),
    grp('sovereign', 'Sovereign data boundary', 5960, -60, 700, 560),
    grp('dr-vnet', 'DR virtual network', 6000, 0, 900, 420),
    svc('dr-a', 'DR gateway', 20, 20, 'dr-vnet'),
    svc('dr-b', 'DR cache', 20, 220, 'dr-vnet'),
    svc('dr-c', 'DR database', 720, 20, 'dr-vnet'),
    svc('dr-d', 'DR analytics', 720, 220, 'dr-vnet'),
  ];
  const edges: Edge[] = [
    { id: 'd1', source: 'p-0', target: 'p-1', label: 'Serves traffic' } as Edge,
    { id: 'd2', source: 'p-1', target: 'dr-a', label: 'Fails over' } as Edge,
    { id: 'd3', source: 'dr-a', target: 'dr-c', label: 'Replicates' } as Edge,
  ];
  return { id: 'stray-zone-pair', nodes, edges };
}

/**
 * An empty annotation band drawn across the top of the whole drawing.
 *
 * One click in the editor (`addGroupBoxAtPosition`) makes an empty group box,
 * and a sovereignty or tenant caption stretched over an architecture is among
 * the commonest things drawn on top of one. It is empty in both senses that
 * matter: nothing declares it as a parent and no tile is inside it.
 *
 * That combination made the band count as *occupied* — the rule that keeps a
 * labelled corridor from being crushed — so it bridged every void it spanned
 * and turned gutter compaction off for the whole drawing. Two clusters 5,450px
 * apart stayed 5,450px apart, and the deck that had been one legible slide
 * became three at 6.93pt.
 */
function bandAboveScenario(): Scenario {
  const names = ['Azure Front Door', 'Azure App Service', 'Azure SQL Database'];
  const cluster = (prefix: string, atX: number): Node[] => Array.from({ length: 6 }, (_, i) => (
    svc(`${prefix}-${i}`, names[i % names.length], atX + (i % 3) * 200, Math.floor(i / 3) * 180)
  ));
  return {
    id: 'band-above',
    nodes: [
      grp('scope', 'Sovereign boundary', -80, -600, 7060, 400),
      ...cluster('east', 0),
      ...cluster('west', 6000),
    ],
    edges: [
      { id: 'b-1', source: 'east-0', target: 'east-1', label: 'Routes' },
      { id: 'b-2', source: 'east-2', target: 'west-2', label: 'Replicates' },
      { id: 'b-3', source: 'west-0', target: 'west-1', label: 'Serves' },
    ] as Edge[],
  };
}

/**
 * The same diagonal cascade, inside the frame everyone draws around one.
 *
 * A subscription or "Azure" rectangle around the whole architecture is the most
 * ordinary annotation there is, and it made `fitBoxesWithin` a no-op: the frame
 * is one span covering the drawing, so the union of the shapes was the entire
 * axis, there was no whitespace left to spend, and the identity map came back.
 * The sheet then went to the uniform scaler, which takes the tiles down while
 * the label point size stays where it is.
 */
function framedCascadeScenario(count = 40, id = 'framed-cascade'): Scenario {
  const inner = diagonalCascadeScenario(count, id);
  return {
    ...inner,
    nodes: [grp('azure', 'Azure subscription', -80, -80, count * 900 + 160, count * 620 + 235), ...inner.nodes],
  };
}

/**
 * A grid packed so tightly that the gutters are narrower than a stub.
 *
 * 150x75 tiles on a 160x85 pitch leaves 10px between neighbours — below the
 * router's own 6px clearance margin on both sides, so `clearLanes` merges every
 * column into one span and offers no lane at all. `countBlocked` inflates each
 * tile by the same margin, so on this pitch every candidate reports the same
 * maximal count and the router cannot tell a route that grazes a corner from
 * one that runs the full height of three tiles it does not connect.
 *
 * A clean route exists — the 310..320 gutter between columns 1 and 2 — which is
 * what makes this a defect rather than an impossible drawing.
 */
function tightSeamScenario(): Scenario {
  const names = ['Azure Front Door', 'Azure App Service', 'Azure SQL Database', 'Azure Functions'];
  const nodes: Node[] = Array.from({ length: 20 }, (_, i) => (
    svc(`s-${i}`, names[i % names.length], (i % 5) * 160, Math.floor(i / 5) * 85)
  ));
  return {
    id: 'tight-seam',
    nodes,
    edges: [
      { id: 'x1', source: 's-0', target: 's-12', label: 'Calls' },
      { id: 'x2', source: 's-4', target: 's-15', label: 'Reads' },
    ] as Edge[],
  };
}

function grp(id: string, label: string, x: number, y: number, w: number, h: number): Node {
  return { id, type: 'groupNode', position: { x, y }, style: { width: w, height: h }, data: { label } } as Node;
}

/**
 * A drawing no page can hold even after every gap is closed.
 *
 * `fitBoxesWithin` gives up distance, which costs nothing but proximity. It has
 * nothing left to give when the shapes ALONE are over the limit: 150px tiles
 * are 1.5625in each, so past about 127 in a row the sheet is over Visio's 200in
 * ceiling with the tiles already touching. `scaleBoxesWithin` is the only
 * remaining answer and it shrinks every shape.
 *
 * This is the fixture that says what happens to the type when it does. The
 * point size used to be a fixed constant, so a 0.33in tile still carried 7.56pt
 * and printed its name almost three times wider than its own box, across
 * several neighbours. Nothing in the corpus reached the scaler, so both the
 * exporter and the two rules below carried the assumption untested.
 */
function overRowScenario(count = 150, id = 'over-row'): Scenario {
  const names = [
    'Azure Front Door', 'Azure App Service', 'Azure SQL Database', 'Azure Functions',
    'Azure Key Vault', 'Azure Service Bus',
  ];
  const nodes: Node[] = Array.from({ length: count }, (_, i) => (
    svc(`w-${i}`, names[i % names.length], i * 200, 0)
  ));
  return {
    id,
    nodes,
    edges: Array.from({ length: 8 }, (_, i) => (
      { id: `o${i}`, source: `w-${i * 4}`, target: `w-${i * 4 + 2}`, label: 'Calls', data: { stepNumber: i + 1 } }
    )) as Edge[],
  };
}

/**
 * Deep scale, where the two pieces of drawing furniture that are not tiles get
 * measured: a zone caption and a numbered step badge.
 *
 * `over-row` only reaches 85%, which is nowhere near the regime either of them
 * fails in. The zones are one service each so that the box a caption has to
 * fit inside comes down with the drawing: 360 of them in a row take the sheet
 * to 15%, and there the caption — held at its natural 9.4pt because nothing
 * scaled it — wraps to seven lines and stands 462% of the height of the zone
 * it names, printed straight over the service inside it and its neighbours.
 *
 * A row is also the shape that squeezes hardest, so this is the fixture that
 * catches two tiles welded flush and the zero-length connector between them.
 */
function scaledZoneRowScenario(): Scenario {
  return zoneRowScenario('scaled-zone-row', 480);
}

/**
 * The same row at 40 subnets, which is an ordinary landing-zone estate rather
 * than a stress test — and the band width the corpus never sampled.
 *
 * At 480 the zones are 0.027in wide and the exporter drops the caption
 * outright; at 20 they are half an inch and everything fits. Between them sits
 * the whole range where the band is narrower than 0.4in but wide enough to
 * draw, and that is where a caption column floored at 0.2in promised the
 * fitter a line that does not exist: the name was fitted onto one line and
 * PowerPoint stacked it one letter per line, out of the bottom of the band and
 * across the neighbours. Sampling one end of a range and calling it covered is
 * how that shipped.
 */
function midZoneRowScenario(): Scenario {
  return zoneRowScenario('mid-zone-row', 40);
}

function zoneRowScenario(id: string, count: number): Scenario {
  const names = ['Azure App Service', 'Azure SQL Database', 'Azure Key Vault', 'Azure Functions', 'Azure Cache for Redis'];
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let s = 0; s < count; s += 1) {
    const originX = s * 260;
    nodes.push(grp(`sub-${s}`, `Perimeter network subnet ${s}`, originX, 0, 200, 130));
    nodes.push(svc(`z-${s}`, names[s % names.length], originX + 25, 40));
    if (s > 0) {
      edges.push({ id: `ze-${s}`, source: `z-${s - 1}`, target: `z-${s}`, label: 'Peers', data: { stepNumber: s } } as Edge);
    }
  }
  return { id, nodes, edges };
}

/**
 * Real Architecture-Center step prose, long enough that every row wraps. The
 * whole corpus used `step N` and one-clause labels, so the workflow list was
 * only ever measured with sentences that fit on one line — and pagination
 * assumed exactly that.
 */
/**
 * A tight grid under a fan of twenty numbered arrows between one pair.
 *
 * The workflow band is opaque white and drawn last, and its reservation was
 * measured from the *authored* sentences while the panel is drawn from the
 * sentences plus the wording muted labels hand back to it. A fan mutes heavily
 * — twenty labels on one chord — so the panel grew 1.8in past its reservation
 * and painted out six of the nine services. Nothing in the corpus could see it:
 * the band was well-formed by every rule that judged the band.
 *
 * The grid is deliberately dense (150px tiles on a 220px pitch) so the drawing
 * is short and the band is the tall thing on the page. The fan is eight arrows
 * rather than the twenty that first exposed this: twenty badges on one chord is
 * a separate, inherent crowding problem, and it drowned the defect this guards
 * in noise. Eight still mutes the whole fan, which is all Issue 1 needs.
 */
function workflowFanScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) nodes.push(svc(`f${r}-${c}`, `Azure Service ${r}${c}`, c * 220, r * 120));
  }
  const edges: Edge[] = [];
  let step = 1;
  for (let i = 1; i < nodes.length; i += 1) {
    edges.push({
      id: `fh${i}`, source: nodes[i - 1].id, target: nodes[i].id,
      label: `Hop ${i}`,
      data: { stepNumber: step++, stepDescription: `Hop ${i}: traffic is inspected at the perimeter and forwarded to the next tier.` },
    } as Edge);
  }
  for (let i = 0; i < 8; i += 1) {
    edges.push({
      id: `ff${i}`, source: 'f1-1', target: 'f1-2',
      label: `The workload queries the reference store through a managed identity ${i + 1}`,
      data: {
        stepNumber: step++,
        stepDescription: `The workload queries the reference store through a managed identity, retrying on throttling ${i + 1}.`,
      },
    } as Edge);
  }
  return { id: 'workflow-fan', nodes, edges };
}

/**
 * Text carrying the code points XML 1.0 cannot represent.
 *
 * Every one of these arrives without the user typing anything unusual. U+000B
 * is Word and PowerPoint's own manual line break, so it comes in on a
 * copy-pasted service name; it is also a legal JSON escape, so it survives an
 * IaC or prototype import intact. A lone surrogate is what a string sliced at a
 * fixed character count leaves behind when it cuts an emoji in half.
 *
 * The failure they cause is invisible at export time and total at open time,
 * which is why this is a fixture and not a unit test: the point is that the
 * whole package — slides, drawing, and the document properties written from the
 * diagram name and the author — comes out openable.
 */
function controlCharScenario(): Scenario {
  const vt = '\u000b';
  const nodes = [
    svc('cc-web', `Payments${vt}gateway`, 0, 0),
    svc('cc-app', `Orders\u000cservice \u{1F680}`, 320, 0),
    // The id, not the label. Shape ids reach `<p:cNvPr name>` through
    // `objectName` and the Visio `NameU`, and they are the half of this that
    // looks like it came from us — it did not: an imported template or a model
    // response names its own nodes, and a name is as fatal to the parse as a
    // caption is.
    svc(`cc-db${vt}1`, `Ledger\u0001store\uD83D`, 640, 0),
    svc('cc-log', `Audit\u001ftrail`, 960, 0),
  ];
  const edges = [
    {
      id: 'cc1', source: 'cc-web', target: 'cc-app', label: `writes${vt}orders`,
      data: { stepNumber: 1, stepDescription: `The gateway writes${vt}orders to the service.` },
    },
    {
      id: `cc2${vt}b`, source: 'cc-app', target: `cc-db${vt}1`, label: 'commits\u0000rows',
      data: { stepNumber: 2, stepDescription: 'The service commits\u0000rows to the ledger.' },
    },
    {
      id: 'cc3', source: `cc-db${vt}1`, target: 'cc-log', label: 'emits\uDC00events',
      data: { stepNumber: 3, stepDescription: 'The ledger emits\uDC00events to the audit trail.' },
    },
  ] as Edge[];
  return { id: 'control-chars', nodes, edges };
}

/**
 * Eighty-one services, one of which is a 20px sliver.
 *
 * The cheapest shape that reaches the only branch where `drop()`'s
 * axis-awareness decides anything. The legible scale is set by the *shortest*
 * service on the sheet, so one short node explodes the grid the planner starts
 * from — 81 ordinary nodes plus one sliver is enough to walk past
 * `MAX_TILED_CELLS`, which is the loop that coarsens with `drop()` and nothing
 * else. Every other caller breaks on the first step, because the grid it starts
 * from is legible by construction.
 *
 * The rule this is a fixture for is the 7pt floor that already exists. Stepping
 * toward a square instead of dropping the slack axis costs between 17% and
 * 4196% of the type size on drawings shaped like this, and sometimes emits
 * *more* slides for the privilege. Without a fixture that reaches the branch,
 * the axis-aware version reads as unreachable cleverness and gets deleted.
 */
function shortServiceGridScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 81; i += 1) {
    nodes.push(svc(`ss-${i}`, `Service ${i}`, (i % 9) * 260, Math.floor(i / 9) * 190));
  }
  // The sliver. A 20px-tall service is what a collapsed annotation or a
  // hand-resized node looks like, and it is the whole reason the grid explodes.
  nodes.push({ ...svc('ss-thin', 'Tag', 9 * 260, 0), height: 20 } as Node);
  const edges: Edge[] = [];
  for (let i = 1; i < 9; i += 1) {
    edges.push({ id: `ss${i}`, source: `ss-${i - 1}`, target: `ss-${i}`, label: 'Calls' } as Edge);
  }
  return { id: 'short-service-grid', nodes, edges };
}

/**
 * A two-hundred-stage pipeline, a fifth of it collapsed to slivers.
 *
 * The regression guard for `drop()`'s axis-awareness, which nothing else in the
 * corpus reaches: mutate it to step toward a square and every other scenario,
 * including `short-service-grid`, emits byte-identical decks.
 *
 * Three things have to be true at once, and each was found by a fixture that
 * failed to fire without it. The grid must exceed `MAX_TILED_CELLS`, because
 * that is the only coarsening loop with no legibility break — the other two
 * call `drop()` once and discard the result, since a grid built to meet the
 * floor is by construction one step above it. Reaching 22500 cells needs a
 * genuinely short representative tile, so a fifth of the estate is collapsed
 * rather than one stray node, and the tenth-percentile target moves with it.
 * And the drawing must be shaped unlike the frame: at 100 stages across and 2
 * deep the width axis binds by a wide margin, so dropping the axis that already
 * binds spends scale for nothing.
 *
 * Square-stepping on this shape narrows the tiles from 1.29in to 0.51in and
 * truncates all two hundred labels.
 */
/**
 * Badges squeezed onto their own arrows in gaps narrower than a badge.
 *
 * Tiles are pitched 170x95 against a 150x75 node, so every gap is 20px — under
 * the natural badge diameter, which forces the placement search to hand
 * `stepBadgeXml` a reduced diameter and puts `badgeMinDiameterIn` on the
 * critical path. Nothing else in the corpus reaches that floor, so without this
 * the rule guarding it would be measuring geometry that is never built. Steps
 * run past nine because a two-digit number is where solving the disc for width
 * alone stops being nearly right.
 */
function squeezedBadgeScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 12; i += 1) {
    nodes.push(svc(`sb-${i}`, 'Azure Container Registry', (i % 6) * 158, Math.floor(i / 6) * 83));
    if (i > 0) {
      edges.push({
        id: `sbe-${i}`, source: `sb-${i - 1}`, target: `sb-${i}`, label: 'private endpoint', data: { stepNumber: i + 10 },
      } as Edge);
    }
  }
  return { id: 'squeezed-badges', nodes, edges };
}

function cascadeScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 200; i += 1) {
    const base = svc(`cs-${i}`, `Azure Kubernetes Service ${i}`, (i % 100) * 2000, Math.floor(i / 100) * 1600);
    nodes.push(i % 5 === 0 ? ({ ...base, height: 12 } as Node) : base);
  }
  const edges: Edge[] = [];
  for (let i = 1; i < 20; i += 1) {
    edges.push({ id: `cse${i}`, source: `cs-${i - 1}`, target: `cs-${i}`, label: 'Calls' } as Edge);
  }
  return { id: 'cascade', nodes, edges };
}

/**
 * Four hundred services with one name between them, a ninth of them collapsed.
 *
 * The collapsed fraction is the whole point. `target` is the tenth percentile
 * of tile heights, so at 40 of 400 it is still 75px and everything is ordinary;
 * at 45 of 400 it is 12px, `LEGIBLE_TILE_PT / 12 / 12` exceeds anything the
 * frame can deliver, and before the cap on `legibleScale` the legibility break
 * in both coarsening loops could never fire. The deck came out as 49 slides of
 * 0.47in tiles on which all 400 names read `"Azure…"` — one string for four
 * hundred services, and every width-based rule passed it at 4.24 characters per
 * line, which is why the rule guarding this is written on identity instead.
 *
 * The shared `"Azure "` prefix is not decoration: it is what makes a generous-
 * looking character budget collapse into a single string.
 */
function sharedPrefixEstateScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 400; i += 1) {
    const base = svc(`sp-${i}`, `Azure Kubernetes Service ${i}`, (i % 20) * 900, Math.floor(i / 20) * 900);
    nodes.push(i % 9 === 0 ? ({ ...base, height: 12 } as Node) : base);
  }
  const edges: Edge[] = [];
  for (let i = 1; i < 12; i += 1) {
    edges.push({ id: `spe${i}`, source: `sp-${i - 1}`, target: `sp-${i}`, label: 'Calls' } as Edge);
  }
  return { id: 'shared-prefix-estate', nodes, edges };
}

/**
 * Sixty services authored 20px tall, on a pitch that forces the planner to tile.
 *
 * A short tile makes `LEGIBLE_TILE_PT / 12 / target` demand a magnification the
 * renderer never grants — every window is drawn through a transform capped at
 * natural size — so the planner split, found the tiles no larger, and split
 * again, down to one tile per slide on a page 0.3% inked. Nothing about the
 * result improved across those extra thirty-six slides: same tile widths, same
 * 7pt floor, same zero truncations, same sixty distinct names.
 *
 * Twenty pixels is not a size the canvas can author — `NodeResizer` sits on
 * groups only, at `minHeight={150}` — but the model can, since
 * `blueprintArchitectureAI.ts:173` emits a height per node, and the same
 * runaway starts from an entirely ordinary 40px.
 */
function shortTileEstateScenario(): Scenario {
  const names = ['Front Door', 'API Management', 'App Service', 'Functions', 'Service Bus', 'Event Hubs',
    'Cosmos DB', 'SQL Database', 'Key Vault', 'Storage Account', 'Redis Cache', 'Container Apps'];
  const nodes: Node[] = Array.from({ length: 60 }, (_, i) => ({
    ...svc(`st-${i}`, `${names[i % 12]} ${i}`, (i % 10) * 400, Math.floor(i / 10) * 400),
    height: 20,
  } as Node));
  const edges: Edge[] = Array.from({ length: 11 }, (_, k) => ({
    id: `ste${k}`, source: `st-${k}`, target: `st-${k + 1}`, label: 'Calls',
  } as Edge));
  return { id: 'short-tile-estate', nodes, edges };
}

/**
 * Tiles authored between the icon threshold and the standard height.
 *
 * `serviceGroupXml`'s icon arithmetic is proportional, so an icon fits from
 * 0.43in — 41.28px — upward, not from the standard 75px. Everything authored in
 * between drew an icon that no rule watched: 45% of a standard tile, wide open.
 *
 * The second half of the fixture is the same rule's opposite failure. These
 * nodes carry no `height` at all and are sized through `style`, which is what
 * `readSize` reads and what the canvas writes when a layout engine sets a size.
 * Reading only `height` saw the default 75 and demanded icons the exporter is
 * right not to draw at 30px.
 */
function compactEstateScenario(): Scenario {
  const names = ['Front Door', 'API Management', 'App Service', 'Functions', 'Service Bus', 'Event Hubs',
    'Cosmos DB', 'SQL Database', 'Key Vault', 'Storage Account', 'Redis Cache', 'Container Apps'];
  const nodes: Node[] = [];
  for (let i = 0; i < 6; i += 1) {
    nodes.push({
      ...svc(`ce-${i}`, `${names[i]} Tier`, (i % 3) * 260, Math.floor(i / 3) * 190),
      height: 50,
    } as Node);
  }
  for (let i = 0; i < 12; i += 1) {
    const base = svc(`ces-${i}`, `${names[i]} Probe`, (i % 4) * 260, 450 + Math.floor(i / 4) * 190);
    delete (base as { height?: number }).height;
    nodes.push({ ...base, style: { width: 150, height: 30 } } as Node);
  }
  const edges: Edge[] = Array.from({ length: 5 }, (_, k) => ({
    id: `cee${k}`, source: `ce-${k}`, target: `ce-${(k + 1) % 6}`, label: 'Calls',
  } as Edge));
  return { id: 'compact-estate', nodes, edges };
}

/**
 * Service names long enough to wrap the customer deck's inventory table.
 *
 * That table is `colW: [5.2, ...]` at 12pt with 0.1in cell margins on each
 * side, so column zero holds 55 ASCII characters on a line, and it declares no
 * autofit — PowerPoint reads `<a:tr h>` as a minimum and grows the row. One
 * name in two lines costs 0.13in more than it was budgeted, and eighteen of
 * them put the last rows below the bottom of the slide, still in the file and
 * invisible on the page.
 *
 * The corpus was 5.5% away from this on its own: `long-name-fan` measures
 * 4.727in against the 5.00in budget, so three more ASCII characters would have
 * shipped it. These names are 77 characters, which is what an estate that names
 * its tier, region and redundancy in the label looks like.
 */
function wrappedInventoryScenario(): Scenario {
  const nodes: Node[] = Array.from({ length: 34 }, (_, i) => svc(
    `wi-${i}`,
    `Azure Database for PostgreSQL Flexible Server (Production, Zone Redundant) ${i}`,
    (i % 6) * 320,
    Math.floor(i / 6) * 200,
  ));
  const edges: Edge[] = Array.from({ length: 5 }, (_, k) => ({
    id: `wie${k}`, source: `wi-${k}`, target: `wi-${k + 1}`, label: 'Replicates',
  } as Edge));
  return { id: 'wrapped-inventory', nodes, edges };
}

/**
 * Thirty names whose words defeat ratio wrapping.
 *
 * `ceil(width / column)` is exact for CJK and for one over-long token, and a
 * lower bound for everything else: real text breaks between words and abandons
 * the rest of the line when the next word will not fit. Three runs each wider
 * than half the column take three lines where the ratio says two, and the error
 * compounds down the page. These are the names a platform team actually writes
 * — service, cluster, environment, region, pool — each run long enough to force
 * the break.
 */
function tokenWrapInventoryScenario(): Scenario {
  const nodes: Node[] = Array.from({ length: 30 }, (_, i) => svc(
    `tw-${i}`,
    `Azure Kubernetes Service aks${String(i).padStart(3, '0')}contosoplatformprodeastus2 `
    + `nodepool${String(i).padStart(3, '0')}systemsurgeeastus2`,
    (i % 6) * 320,
    Math.floor(i / 6) * 200,
  ));
  const edges: Edge[] = Array.from({ length: 4 }, (_, k) => ({
    id: `twe${k}`, source: `tw-${k}`, target: `tw-${k + 1}`, label: 'Scales',
  } as Edge));
  return { id: 'token-wrap-inventory', nodes, edges };
}

/**
 * Forty workflow steps of long prose, which is what the customer deck's
 * workflow slide never had to survive.
 *
 * The slide paginated on a row *count* and sized its type off the resulting
 * pitch, so nineteen rows of ordinary description overlapped by a fifth of
 * their pitch and the top row printed over the header bar. Long descriptions
 * are the normal case for a generated dataflow, not an extreme one.
 */
function longWorkflowScenario(): Scenario {
  // The chain topology `workflow-long-prose` already proves is clean, extended
  // to twenty-four hops. The diagram is deliberately unremarkable: what is
  // under test is the twenty-four rows of prose the workflow slide has to lay
  // out from it. Below about fifteen the rows still clear each other, which is
  // why the existing twelve-hop fixture never caught this.
  const nodes: Node[] = [];
  for (let i = 0; i < 25; i += 1) {
    nodes.push(svc(`lw${i}`, `Workflow service ${i}`, (i % 6) * 220, Math.floor(i / 6) * 150));
  }
  const edges: Edge[] = [];
  for (let i = 0; i < 24; i += 1) {
    edges.push({
      id: `lwe${i}`, source: `lw${i}`, target: `lw${i + 1}`, label: '手順を引き渡す',
    } as Edge);
  }
  return { id: 'workflow-long-rows', nodes, edges };
}

/**
 * Nine Visio workflow rows of unbreakable resource names.
 *
 * The Visio band and its guard both counted lines as `ceil(width / column)`,
 * which is a lower bound rather than a count: prose breaks between words and
 * abandons the rest of the line. Ordinary prose drew five lines in a budget of
 * four; a description made of 30-character resource names — the kind a naming
 * convention produces and a generated dataflow quotes verbatim — drew eight in
 * a budget of six. Visio does not clip, so every row painted through the rows
 * above and below it, and the guard reported nothing.
 */
function visioTokenWorkflowScenario(): Scenario {
  const nodes: Node[] = Array.from({ length: 10 }, (_, i) => svc(
    `vtw${i}`,
    `Contoso platform service ${i}`,
    (i % 5) * 300,
    Math.floor(i / 5) * 200,
  ));
  const edges: Edge[] = Array.from({ length: 9 }, (_, i) => ({
    id: `vtwe${i}`,
    source: `vtw${i}`,
    target: `vtw${i + 1}`,
    label: 'Hands off',
    data: {
      stepNumber: i + 1,
      // Long unbreakable runs: a greedy renderer cannot fit two on a line that
      // holds one and a half, so it wastes the tail of every line. The ratio
      // count never sees that waste.
      stepDescription: Array.from(
        { length: 8 },
        (_, k) => `contoso-platform-production-eastus2-instance-${String(i * 8 + k).padStart(2, '0')}`,
      ).join(' '),
    },
  } as Edge));
  return { id: 'visio-token-workflow', nodes, edges };
}

/**
 * Sixteen services whose names carry hard line breaks.
 *
 * `\n` survives the sanitiser — which scrubs U+000B while reasoning explicitly
 * about copy-paste, and lets the far commoner U+000A through — and pptxgenjs
 * turns each one into a real `<a:p>`. Every line counter in the codebase split
 * on whitespace, so a newline merely ended a run and never started a line, and
 * the shape scrape joined `<a:t>` with nothing at all so the guard could not
 * see that paragraphs had existed. The Services table measured 5.83in and drew
 * 10.33in: 3.9in of rows below the sheet, invisible and unrecoverable.
 */
function hardBreakInventoryScenario(): Scenario {
  const nodes: Node[] = Array.from({ length: 16 }, (_, i) => svc(
    `hb${i}`,
    `Azure SQL MI ${i + 1}\nProd ring\nEast US 2\nZone redundant`,
    (i % 6) * 280,
    Math.floor(i / 6) * 180,
  ));
  const edges: Edge[] = Array.from({ length: 6 }, (_, i) => ({
    id: `hbe${i}`,
    source: `hb${i}`,
    target: `hb${i + 1}`,
    label: 'Replicates',
    data: {
      stepNumber: i + 1,
      // Visio renders a raw newline in `<Text>` as a paragraph break too, and
      // the band budgeted two lines for rows that drew three.
      stepDescription: 'クライアントは Front Door を経由して要求を送信します。\n'
        + 'ゲートウェイが要求を検証し、認証トークンを確認します。\n'
        + 'API がバックエンドのサービスを呼び出します。',
    },
  } as Edge));
  return { id: 'hard-break-inventory', nodes, edges };
}

/**
 * A fan of four parallel hops whose connector labels carry hard breaks.
 *
 * `labelSize` was the third copy of `ceil(width / column)` in the file and the
 * one that survived the sweep. `estimateTextWidthIn` returns a *width*, and a
 * newline has none, so a four-paragraph label measured as one line. That height
 * is not decoration: it is the chip's collision rectangle when it is seated,
 * the rung pitch of a fan, the walk step when a chip is settled, and the
 * emitted `TxtHeight`. The pitch stayed frozen at 0.490in while the text grew
 * to 0.570in, and the chips were written through each other.
 *
 * The guard could not see it either, because it read `TxtHeight` back out of
 * the file — asking the exporter how tall its own text is. It measures the ink
 * now, the way the PowerPoint side has since the painted-ink rule.
 */
function visioBrokenLabelFanScenario(): Scenario {
  const nodes: Node[] = [svc('vbfa', 'Alpha', 0, 0), svc('vbfb', 'Bravo', 520, 0)];
  const edges: Edge[] = Array.from({ length: 4 }, (_, i) => ({
    id: `vbf${i}`,
    source: 'vbfa',
    target: 'vbfb',
    label: `L1 p${i + 1}\nL2 p${i + 1}\nL3 p${i + 1}\nL4 p${i + 1}`,
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1}.` },
  } as Edge));
  return { id: 'visio-broken-label-fan', nodes, edges };
}

/**
 * Realistic service names on the app's own default tile.
 *
 * `AzureNode` has no resizer and every layout engine in the repo places nodes
 * at 180x75, so this is not a stress case — it is what a user gets. The Visio
 * tile sized its label band with `min(needed, room - minIcon)`, a clamp with no
 * font reduction and no cut behind it, and Visio does not clip text to a text
 * block: a 120-character name overran the band by 0.271in and was drawn 0.076in
 * past the bottom of the tile, through the icon. The 77-character name the rest
 * of the corpus happens to exercise cleared the clamp by 0.002in, which is why
 * nothing had ever caught it.
 *
 * The names are graded — one that fits, one just over, one well over, and a
 * Japanese one — so a regression shows up as a step rather than a cliff.
 */
/**
 * Realistic service names on a tile that also carries a SKU and a region.
 *
 * The tile planner asked `Math.ceil(inkWidth / column)` how many lines a name
 * needs. That is the break-anywhere assumption — the count the ink would need
 * if a word could be split at any character — so it is a lower bound and never
 * the answer. Word wrap abandons the tail of a line whenever the next word will
 * not fit, and on these six names it costs one or two whole lines each: five of
 * the six were planned at three lines and drew four or five.
 *
 * The lines have to go somewhere. `labelBlockH` is what sizes the icon and, in
 * the top-aligned branch, what is left over becomes the text box's height —
 * so the surplus is drawn straight down through the "P1v3 · eastus" sub-line
 * and out of the bottom of the tile. "Azure Kubernetes Service Automatic
 * cluster" covered 100% of its sub-line's band.
 *
 * Nothing caught it because three separate things were wrong in the same
 * direction: the ink rule skipped every `service-` shape, `drawnTextRect` used
 * the identical ratio so the chip rules agreed with the defect, and all three
 * meta scenarios used names short enough to wrap the way the ratio predicts.
 */
function tileNameWithMetaScenario(): Scenario {
  const names = [
    'Azure Database for PostgreSQL Flexible Server',
    'Azure Kubernetes Service Automatic cluster',
    'Azure Monitor Application Insights workspace',
    'Azure Container Apps Environment workload profile',
    'Azure Synapse Analytics dedicated SQL pool',
    'Azure Cache for Redis Enterprise Flash tier',
  ];
  const nodes: Node[] = names.map((label, i) => ({
    ...svc(`tnm${i}`, label, (i % 2) * 230, Math.floor(i / 2) * 190, undefined, true, 'compute'),
    width: 160,
    height: 110,
    data: {
      label,
      serviceName: label,
      category: 'compute',
      iconPath: '/Azure_Public_Service_Icons/Icons/compute/10021-icon-service-Virtual-Machine.svg',
      sku: 'P1v3',
      region: 'eastus',
    },
  } as Node));
  const edges: Edge[] = names.slice(1).map((_, i) => ({
    id: `tnm-e${i}`,
    source: `tnm${i}`,
    target: `tnm${i + 1}`,
    label: 'sync',
  } as Edge));
  return { id: 'tile-name-with-meta', nodes, edges };
}

/**
 * A wide zone whose tiles sit flush against its top edge, with a long name.
 *
 * Built to exercise the band search where the score cannot separate two
 * candidates: a strip above the tiles and clear paper below them both cover no
 * tile at all, so both score zero and something other than the score decides.
 * The search used to stop at the first zero it found — `runs(top)` is pushed
 * into the candidate array before `runs(foot)` — and so took whichever clear
 * band was built first rather than the one that holds more of the name.
 *
 * Honest about what it does and does not show: on this geometry the caption
 * lands in a 4.28in third-band that holds the whole name at 10pt, with or
 * without the tie-break, so this fixture pins the *outcome* rather than
 * reproducing the tie. It is here because a flush-top zone with a 96-character
 * name is otherwise unrepresented in the corpus, and because the rule below —
 * which compares a cut caption's band against the widest clear strip of its
 * own zone — needs a zone with clear strips to be exercised against at all.
 */
/**
 * The reviewer's round-47 ASK-2 geometry, reproduced verbatim: a wide zone whose
 * top row is 11 tiles with a single 60px gap in it, and whose foot row is three
 * tiles leaving 1320px clear. Both `runs(top)` and `runs(foot)` produce a band
 * that covers no tile, so both score zero, and the inter-row gap (0.179in) is
 * shorter than the caption band (0.240in) so no gapRow is produced either — the
 * fixture isolates the tie-break and nothing else.
 *
 * With the tie-break the foot's 9.424in band wins and all 67 characters survive.
 * Revert it and `runs(top)` wins on first-seen, giving a 0.431in band and 13
 * characters. This is the fixture the earlier version of this scenario failed
 * to be: that one landed in a third-band wide enough for the whole name and so
 * measured nothing.
 */
function flushTopZoneScenario(): Scenario {
  const zone: Node = {
    id: 'ftz',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    width: 1800,
    height: 184,
    data: { label: 'Production VNet - Application Subnet (10.0.1.0/24) - Zone Redundant' },
  } as Node;
  const nodes: Node[] = [zone];
  [10, 160, 310, 460, 610, 760, 910, 1060, 1210, 1360, 1570].forEach((x, i) => {
    nodes.push(svc(`ftz${i}`, `Top service ${i}`, x, 4, 'ftz', true, 'compute'));
  });
  [10, 170, 330].forEach((x, i) => {
    nodes.push(svc(`ftzf${i}`, `Foot service ${i}`, x, 104, 'ftz', true, 'compute'));
  });
  return { id: 'flush-top-zone', nodes, edges: [] };
}

/**
 * The reviewer's round-47 Issue 1 geometry: two tiers of tiles inside one zone
 * with edges running straight down between them. The corridor between the rows
 * is the widest clear paper in the zone, so the caption search wants it; it is
 * also where `connectorLabelBox` seats its chips, and a chip is drawn last at
 * 92% opacity. Three chips landed on the caption, covering it across its full
 * height, and nothing fired: chips are exempt from the painted-ink rule, the
 * pairwise rule only examined shapes that overflowed their own box (the caption
 * fitted exactly), and the one rule that tests a drawn-last opaque object did
 * not list zone captions.
 */
function zoneCaptionCorridorScenario(): Scenario {
  const zone: Node = {
    id: 'zcc',
    type: 'groupNode',
    position: { x: 0, y: 0 },
    width: 900,
    height: 226,
    data: { label: 'Production VNet - Application Subnet (10.0.1.0/24)' },
  } as Node;
  const nodes: Node[] = [zone];
  for (let i = 0; i < 3; i += 1) {
    nodes.push(svc(`zcca${i}`, `Front service ${i}`, 20 + i * 290, 8, 'zcc', true, 'compute'));
    nodes.push(svc(`zccb${i}`, `Back service ${i}`, 20 + i * 290, 143, 'zcc', true, 'databases'));
  }
  const edges: Edge[] = Array.from({ length: 3 }, (_, i) => ({
    id: `zcc-e${i}`,
    source: `zcca${i}`,
    target: `zccb${i}`,
    label: 'writes order documents',
    data: { stepNumber: i + 1, stepDescription: `Step ${i + 1}` },
  } as Edge));
  return { id: 'zone-caption-corridor', nodes, edges };
}

function visioDefaultTileNamesScenario(): Scenario {
  const names = [
    'Azure Database for PostgreSQL flexible server - Business Critical - East US 2',
    'Azure Database for PostgreSQL Flexible Server (Business Critical, Zone Redundant HA, East US 2)',
    'Azure Database for PostgreSQL Flexible Server - Business Critical tier with zone redundant high availability in East US 2',
    'Azure Database for PostgreSQL フレキシブル サーバー ビジネス クリティカル ゾーン冗長 高可用性 東日本リージョン',
  ];
  const nodes: Node[] = names.map((label, i) => ({
    ...svc(`vdt${i}`, label, (i % 2) * 320, Math.floor(i / 2) * 190, undefined, true, 'databases'),
    width: 180,
    height: 75,
  }));
  const edges: Edge[] = [
    { id: 'vdt-e0', source: 'vdt0', target: 'vdt1', label: 'replicates' } as Edge,
    { id: 'vdt-e1', source: 'vdt2', target: 'vdt3', label: 'fails over to' } as Edge,
  ];
  return { id: 'visio-default-tile-names', nodes, edges };
}

/**
 * Four wide zones, laid out by the app itself, whose names do not fit.
 *
 * Nothing here is authored to break anything: four groups of six services at
 * the sizes the layout engine emits, with a subnet name of the length a real
 * network diagram carries. The estate is wider than a slide, so the deck tiles
 * it — and a window cut a zone down to a 0.924in band while leaving its
 * caption the full 60 characters the exporter's cell-count cap had
 * "shortened" it to. A cell count is not a width: those 60 characters wrapped
 * to 7 lines and painted 0.810in outside a 0.240in band, over the tiles below.
 *
 * The uncut zones overflowed too, by 0.060in each, so every zone caption on
 * the drawing was outside its band. None of it was visible to the gate, which
 * skipped any slide holding a service tile — one `continue` that exempted
 * every drawing in the corpus from the entire painted-ink apparatus.
 */
function zoneCaptionWideEstateScenario(): Scenario {
  const name = 'Production VNet - Application Subnet (10.0.1.0/24) - Zone Redundant';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let g = 0; g < 4; g += 1) {
    const gx = g * (40 + 6 * 200 + 40 + 90);
    nodes.push({
      id: `zc${g}`,
      type: 'groupNode',
      position: { x: gx, y: 0 },
      style: { width: 40 + 6 * 200 + 40, height: 90 + 75 + 40 },
      data: { label: `${name} ${g + 1}` },
    } as unknown as Node);
    for (let i = 0; i < 6; i += 1) {
      nodes.push({
        ...svc(`zc${g}_${i}`, `Azure Service ${g}-${i}`, 40 + i * 200, 90, `zc${g}`, true, 'compute'),
        width: 150,
        height: 75,
      });
      if (i) {
        edges.push({
          id: `zc-e${g}_${i}`,
          source: `zc${g}_${i - 1}`,
          target: `zc${g}_${i}`,
          label: 'calls',
        } as Edge);
      }
    }
  }
  return { id: 'zone-caption-wide-estate', nodes, edges };
}

/**
 * A 560-service estate under a 500-step CJK workflow.
 *
 * The band is sized twice — once at the narrowest page the exporter emits, to
 * hand the fit a budget, and once at the real width, which is what the page is
 * built from. The first was assumed to be an upper bound of the second because
 * narrow columns wrap longer. It is not, when the search stops at the first
 * split under its target rather than the shortest: wider columns wrap less, so
 * the wide pass reaches the target at *fewer* columns, and fewer columns is a
 * taller band. Both are under the target; the wide one was 6.55in taller.
 *
 * The fit is handed `198.3 − reserve` inches and the page is then built as
 * `content + drawn`, so every inch of divergence lands straight on the sheet:
 * this came out at 50 x 206in, which Visio refuses to open. The margin is
 * 0.5in and `ladder-in-grid` was already at 68% of it.
 */
function workflowWideBandScenario(): Scenario {
  const prose = '受注要求はエッジで認証されプライベートエンドポイント経由でワークロードへ転送されます詳細は運用手順書の該当節を参照してください追加の注記もあります';
  const nodes: Node[] = [];
  for (let r = 0; r < 28; r += 1) {
    for (let c = 0; c < 20; c += 1) nodes.push(svc(`w${r}-${c}`, `Service ${r}${c}`, c * 240, r * 760));
  }
  const edges: Edge[] = [];
  for (let i = 1; i <= 500; i += 1) {
    edges.push({
      id: `we${i}`, source: nodes[i - 1].id, target: nodes[i].id, label: `Hop ${i}`,
      data: { stepNumber: i, stepDescription: prose.slice(0, 44) },
    } as Edge);
  }
  return { id: 'workflow-wide-band', nodes, edges };
}

function workflowProseScenario(): Scenario {
  const sentences = [    'The client sends the request to Azure Front Door, which terminates TLS at the edge and applies the WAF ruleset before anything reaches the origin.',
    'Front Door forwards the validated request to the App Service origin over Private Link, so the origin is never reachable from the public internet.',
    'The web tier exchanges its managed identity for an access token and calls the API tier, which authorises the caller against the roles in the token.',
    'The API tier writes the order document to Azure Cosmos DB and the accompanying blob to Azure Storage in the same logical transaction boundary.',
    'A change feed trigger raises an event on Azure Service Bus so downstream processing is decoupled from the request path and can be retried safely.',
    'Azure Functions consumes the message, enriches it against the reference data cache and hands the result to the fulfilment system for dispatch.',
  ];
  const nodes: Node[] = [];
  for (let i = 0; i < 24; i += 1) {
    nodes.push(svc(`p${i}`, `Azure Service ${i}`, (i % 6) * 220, Math.floor(i / 6) * 150));
  }
  const edges: Edge[] = [];
  for (let i = 0; i < 23; i += 1) {
    edges.push({
      id: `q${i}`, source: `p${i}`, target: `p${i + 1}`,
      label: 'forwards the validated request',
      data: { stepNumber: i + 1, stepDescription: sentences[i % sentences.length] },
    } as Edge);
  }
  return { id: 'workflow-prose', nodes, edges };
}

/**
 * A step whose sentence needs more rows than the 0.62in row cap allowed. The
 * cap silently overrode the pagination reserve, so the text was printed
 * outside its own box and, past about 800 Latin characters, over the row
 * below it. `workflow-prose` cannot reach this: its sentences all fit at 12pt.
 */
function workflowLongProseScenario(): Scenario {
  const clause = 'The regional ingestion tier authenticates the caller with its managed identity, validates the payload against the published schema, '
    + 'writes the accepted document to Azure Cosmos DB, emits a change-feed event onto Azure Service Bus for the downstream fulfilment pipeline, '
    + 'and records the correlation identifier in Application Insights so the whole hop can be traced end to end afterwards. ';
  const nodes: Node[] = [];
  for (let i = 0; i < 13; i += 1) {
    nodes.push(svc(`g${i}`, `Azure Service ${i}`, (i % 6) * 220, Math.floor(i / 6) * 150));
  }
  const edges: Edge[] = [];
  for (let i = 0; i < 12; i += 1) {
    edges.push({
      id: `r${i}`, source: `g${i}`, target: `g${i + 1}`,
      label: 'hands off the payload',
      // Long enough to need five lines at the 9pt floor, which is more than
      // the capped row could ever hold.
      data: { stepNumber: i + 1, stepDescription: `${clause}${clause}`.slice(0, 800) },
    } as Edge);
  }
  return { id: 'workflow-long-prose', nodes, edges };
}

/**
 * One tile per category, so every accent in `CATEGORY_STYLES` is actually
 * rendered. Nothing else in the corpus sets `data.category`, so all 31
 * scenarios fell through to `other` and fifteen of the sixteen palettes had
 * never been drawn, let alone measured for contrast.
 */
function allCategoriesScenario(): Scenario {
  const names = Object.keys(CATEGORY_STYLES);
  const nodes = names.map((category, i) => svc(
    `k${i}`, `Azure Service ${i}`, (i % 4) * 240, Math.floor(i / 4) * 170, undefined, true, category,
  ));
  const edges: Edge[] = [];
  for (let i = 1; i < names.length; i += 1) {
    edges.push({
      id: `c${i}`, source: `k${i - 1}`, target: `k${i}`,
      label: `step ${i}`, data: { stepNumber: i },
    } as Edge);
  }
  return { id: 'all-categories', nodes, edges };
}

/**
 * Wide tiles carrying a full sub-line, stacked in rows barely further apart
 * than a chip is tall. The only clear paper for a label is the strip directly
 * under a tile — which is exactly where the SKU, the region and the price are
 * drawn. `meta-subline` does not discriminate here: its tiles are narrow
 * enough that the sub-line leaves slack at both ends, so a chip can settle
 * beside the words without touching them.
 */
function metaChipScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    const node = svc(`w${i}`, `Service ${i}`, (i % 4) * 300, Math.floor(i / 4) * 168);
    Object.assign(node.data as Record<string, unknown>, {
      sku: 'Standard_D2s',
      region: 'japaneast',
      pricing: { estimatedCost: 64.2, quantity: 1, region: 'japaneast' },
    });
    nodes.push(node);
  }
  const edges: Edge[] = [];
  for (let i = 0; i < 4; i += 1) {
    edges.push({
      id: `d${i}`, source: `w${i}`, target: `w${i + 4}`,
      label: 'replicates state', data: { stepNumber: i + 1, stepDescription: `step ${i + 1}` },
    } as Edge);
  }
  return { id: 'meta-chip', nodes, edges };
}

/**
 * Every tile carrying the SKU/region/cost sub-line. Nothing else in the corpus
 * sets `meta`, so the second character row — the smallest type either exporter
 * draws — was never emitted and no rule about it could ever fire.
 */
function metaSublineScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 9; i += 1) {
    const node = svc(`m${i}`, `Azure Service ${i}`, (i % 3) * 260, Math.floor(i / 3) * 200);
    Object.assign(node.data as Record<string, unknown>, {
      sku: i % 2 ? 'Standard_D4s_v5' : 'P1v3',
      region: 'japaneast',
      pricing: { estimatedCost: 128.4, quantity: 1, region: 'japaneast' },
    });
    nodes.push(node);
  }
  const edges: Edge[] = [];
  for (let i = 1; i < 9; i += 1) {
    edges.push({ id: `c${i}`, source: `m${i - 1}`, target: `m${i}`, label: 'マネージド ID で参照系を照会します', data: { stepNumber: i, stepDescription: `手順 ${i}` } } as Edge);
  }
  return { id: 'meta-subline', nodes, edges };
}

/**
 * Two dense clusters joined by one long bridge, so the middle of the grid
 * holds nothing. A part that owns only its own fitted cell leaves the bridge's
 * label and callout belonging to no slide at all: the arrow is drawn, the
 * number is missing, and the workflow list still cites it.
 */
/**
 * One front door fanning out to six services stacked on the far side. The
 * commonest Architecture Center shape, and the one that exposes port dealing:
 * if the six east-side ports are not handed out in the same top-to-bottom
 * order as the targets, the hops braid on their way across the paper.
 */
function hubFanScenario(): Scenario {
  const nodes: Node[] = [svc('hub', 'Azure Front Door', 0, 500)];
  const edges: Edge[] = [];
  for (let i = 0; i < 6; i += 1) {
    nodes.push(svc(`h${i}`, `Backend Service ${i}`, 600, i * 200));
    edges.push({ id: `hf${i}`, source: 'hub', target: `h${i}`, label: `route ${i}`, data: { stepNumber: i + 1 } } as Edge);
  }
  return { id: 'hub-fan', nodes, edges };
}

/**
 * One shared service consumed from all over a wide estate. Every hop is long
 * and several cross window seams, which is the arrangement that exposed the
 * dropped-hop class: the shared service is the single most important box on
 * the page, and an arrow into it that is not drawn takes its chip with it.
 */
function sharedServiceScenario(): Scenario {
  const nodes: Node[] = [svc('m', 'Azure Key Vault', 1500, 700)];
  const edges: Edge[] = [];
  const spots: Array<[number, number]> = [[0, 0], [3000, 0], [0, 1400], [3000, 1400], [1500, 0], [1500, 1400]];
  spots.forEach(([x, y], i) => {
    nodes.push(svc(`t${i}`, `Consumer Workload ${i}`, x, y));
    edges.push({
      id: `m-t${i}`,
      source: `t${i}`,
      target: 'm',
      label: `シークレットを取得 ${i}`,
      data: { stepNumber: i + 1 },
    } as Edge);
  });
  return { id: 'shared-service', nodes, edges };
}

/**
 * Twenty-four services on a 165 x 105 grid: a gutter of about fifteen pixels,
 * narrower than a callout disc. The reviewer's case for the buried-badge work
 * — with nowhere clear to sit, a ring search that gives up too early parks the
 * numbers on the tiles and the reader cannot match them to the step list.
 */
function tightGridScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 24; i += 1) {
    nodes.push(svc(`tg${i}`, `Service ${i}`, (i % 6) * 165, Math.floor(i / 6) * 105));
  }
  for (let i = 1; i < 24; i += 1) {
    edges.push({
      id: `tg${i - 1}-${i}`,
      source: `tg${i - 1}`,
      target: `tg${i}`,
      label: `step ${i}`,
      data: { stepNumber: i },
    } as Edge);
  }
  return { id: 'tight-grid', nodes, edges };
}

/**
 * The reviewer's two-stray case: a thirty-node banded estate plus a pair of
 * far-placed services with an edge of their own. Clamping pulls the strays back
 * onto the page but the router plans from where they used to be, so the hop
 * between them lands outside every window at once — annotations and all.
 */
function bandedTwoStraysScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 30; i += 1) {
    nodes.push(svc(`b-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', i * 300, (i % 3) * 200));
    if (i > 0) {
      edges.push({
        id: `y-${i}`,
        source: `b-${i - 1}`,
        target: `b-${i}`,
        label: 'HTTPS token check',
        ...(i <= 8 ? { data: { stepNumber: i, stepDescription: `step ${i}` } } : {}),
      } as Edge);
    }
  }
  nodes.push(svc('b-stray', 'Copilot Studio', -14000, -6000));
  nodes.push(svc('b-stray2', 'Microsoft Fabric', -14000, -5600));
  edges.push({
    id: 'ss',
    source: 'b-stray',
    target: 'b-stray2',
    label: 'mirrors the analytics estate',
    data: { stepNumber: 9, stepDescription: 'mirrors the analytics estate into Fabric' },
  } as Edge);
  return { id: 'banded-two-strays', nodes, edges };
}

/**
 * Strays in OPPOSITE directions — the case a rigid translation cannot survive.
 * Moving the cloud as one body preserved the 18000px of empty space between the
 * two strays, so trimming the outliers *grew* the drawing from 189in to 198in:
 * a 199in Visio sheet (Visio refuses anything past 200in), a 56in slide, and
 * 4pt type on the fixed-size customer deck. Packing them into a strip in the
 * margin costs the width of the strip and nothing else.
 */
function oppositeStraysScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push(svc(`o-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', (i % 4) * 220, Math.floor(i / 4) * 180));
  }
  nodes.push(svc('o-west', 'Copilot Studio', -9000, 400));
  nodes.push(svc('o-east', 'Microsoft Fabric', 9000, 400));
  const edges: Edge[] = [
    { id: 'ow', source: 'o-0', target: 'o-west', label: 'agent actions', data: { stepNumber: 1, stepDescription: 'Copilot Studio calls the agent action' } } as Edge,
    { id: 'oe', source: 'o-3', target: 'o-east', label: 'mirrored to Fabric', data: { stepNumber: 2, stepDescription: 'Operational data is mirrored into Fabric' } } as Edge,
  ];
  return { id: 'opposite-strays', nodes, edges };
}

/**
 * Three strays off three different corners plus a zone that has drifted with
 * one of them, so the packing has to keep a group with the service it contains
 * while still collapsing the empty space in both directions at once.
 */
function cornerStraysScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push(svc(`x-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', (i % 4) * 220, Math.floor(i / 4) * 180));
  }
  nodes.push(svc('x-nw', 'Copilot Studio', -9000, -9000));
  nodes.push(grp('x-far', 'Remote Region', 9000, -9000, 520, 300));
  nodes.push(svc('x-ne', 'Microsoft Fabric', 120, 120, 'x-far'));
  nodes.push(svc('x-se', 'Power BI', 9000, 9000));
  const edges: Edge[] = [
    { id: 'cw', source: 'x-0', target: 'x-nw', label: 'agent actions', data: { stepNumber: 1, stepDescription: 'Copilot Studio calls the agent action' } } as Edge,
    { id: 'cn', source: 'x-1', target: 'x-ne', label: 'mirrored to Fabric', data: { stepNumber: 2, stepDescription: 'Operational data is mirrored into Fabric' } } as Edge,
    { id: 'cs', source: 'x-3', target: 'x-se', label: 'served to Power BI', data: { stepNumber: 3, stepDescription: 'Power BI reads the semantic model' } } as Edge,
  ];
  return { id: 'corner-strays', nodes, edges };
}

/**
 * A single row of services with one node far left and one far right. Every
 * service shares a row, so the vertical quartile range is zero, the fence has
 * zero width, and the one node on a second row used to count as an outlier —
 * which pushed the kept set under the majority bar and abandoned the trim on
 * BOTH axes. The drawing then sized an 85in page and shipped 0.24in tiles that
 * no rule could see, because trimming never ran and so nothing was ever parked.
 */
function symmetricStraysScenario(): Scenario {
  const nodes: Node[] = [
    svc('y-0', 'Azure Front Door', 0, 0),
    svc('y-1', 'Azure Functions', 260, 0),
    svc('y-2', 'Azure SQL Database', 520, 0),
    svc('y-3', 'Azure Key Vault', 260, 170),
    svc('y-west', 'Copilot Studio', -4000, 0),
    svc('y-east', 'Microsoft Fabric', 4000, 0),
  ];
  const edges: Edge[] = [
    { id: 'sw', source: 'y-0', target: 'y-west', label: 'agent actions', data: { stepNumber: 1, stepDescription: 'Copilot Studio calls the agent action' } } as Edge,
    { id: 'se', source: 'y-2', target: 'y-east', label: 'mirrored to Fabric', data: { stepNumber: 2, stepDescription: 'Operational data is mirrored into Fabric' } } as Edge,
    { id: 'sk', source: 'y-1', target: 'y-3', label: 'reads secrets', data: { stepNumber: 3, stepDescription: 'The function reads its secrets from Key Vault' } } as Edge,
  ];
  return { id: 'symmetric-strays', nodes, edges };
}

/**
 * The canonical Architecture Center hub-and-spoke: a hub, four spokes on a
 * 1400px radius, and four shared services. Nine services in 30x30in of mostly
 * whitespace, which is what makes it dangerous — the tiling planner's
 * services-per-slide floors are written to stop a twelve-service diagram
 * becoming a flip-book, and a deliberately sparse drawing trips every one of
 * them. The planner then reports "this frame cannot show the drawing legibly at
 * any grid", which is a *request to grow the page* — and the deck the export
 * button produces cannot grow its page. It read the empty window list as "it
 * already fits" and squeezed all nine services onto one 13.333x7.5in slide at
 * 0.315in with 4pt type, for every drawing past the point where the audited
 * diagram-only deck starts growing its page.
 */
function hubSpokeScenario(): Scenario {
  const R = 1400;
  const nodes: Node[] = [
    svc('hub', 'Azure Firewall', 0, 0),
    svc('spoke-n', 'Azure Kubernetes Service', 0, -R),
    svc('spoke-s', 'Azure App Service', 0, R),
    svc('spoke-e', 'Azure SQL Database', R, 0),
    svc('spoke-w', 'Azure Functions', -R, 0),
    ...[0, 1, 2, 3].map((i) => svc(`shared-${i}`, ['Azure Key Vault', 'Azure Monitor', 'Azure Bastion', 'Azure DNS'][i], 200 + i * 190, 400)),
  ];
  const edges: Edge[] = [
    { id: 'h1', source: 'hub', target: 'spoke-n', label: 'Peered', data: { stepNumber: 1, stepDescription: 'The hub peers with the container spoke' } } as Edge,
    { id: 'h2', source: 'hub', target: 'spoke-s', label: 'Peered', data: { stepNumber: 2, stepDescription: 'The hub peers with the web spoke' } } as Edge,
    { id: 'h3', source: 'hub', target: 'spoke-e', label: 'Peered', data: { stepNumber: 3, stepDescription: 'The hub peers with the data spoke' } } as Edge,
    { id: 'h4', source: 'hub', target: 'spoke-w', label: 'Peered', data: { stepNumber: 4, stepDescription: 'The hub peers with the integration spoke' } } as Edge,
    { id: 'h5', source: 'hub', target: 'shared-0', label: 'Inspects', data: { stepNumber: 5, stepDescription: 'Shared services sit behind the firewall' } } as Edge,
  ];
  return { id: 'hub-spoke', nodes, edges };
}

/**
 * A compliance boundary drawn across the drawing to a remote service, which is
 * how Architecture Center security diagrams show scope: two zones overlap, and
 * the wide one contains services it does not own.
 *
 * Parking grouped a stray zone with every box that sat inside its rectangle,
 * with no parent check at all, so half of a 4x2 grid of core services was torn
 * out and packed into the margin — 55% of the drawing moved, past the 40% the
 * majority floor is supposed to allow, because the claim happens after that
 * test. And because a cluster was packed as a unit but never compacted inside,
 * the 8800px-wide zone was translated whole and the parked drawing came out
 * 101.67in against 95.21in for never trimming at all.
 */
function scopeZoneScenario(): Scenario {
  const nodes: Node[] = [
    ...Array.from({ length: 8 }, (_, i) => svc(
      `v-${i}`,
      i % 2 ? 'Azure Functions' : 'Azure SQL Database',
      (i % 4) * 220,
      Math.floor(i / 4) * 180,
    )),
    grp('vnet', 'Hub virtual network', -40, -40, 900, 420),
    grp('pci-scope', 'Cardholder data scope', 300, -140, 8800, 540),
    svc('remote', 'Azure Payment HSM', 8800, 60),
  ];
  const edges: Edge[] = [
    { id: 'p1', source: 'v-0', target: 'v-1', label: 'Accepts card data', data: { stepNumber: 1, stepDescription: 'The gateway accepts card data' } } as Edge,
    { id: 'p2', source: 'v-1', target: 'v-2', label: 'Tokenises the PAN', data: { stepNumber: 2, stepDescription: 'The function tokenises the PAN' } } as Edge,
    { id: 'p3', source: 'v-2', target: 'remote', label: 'Signs with the HSM', data: { stepNumber: 3, stepDescription: 'The payment HSM signs the transaction' } } as Edge,
    { id: 'p4', source: 'v-5', target: 'v-6', label: 'Writes the ledger', data: { stepNumber: 4, stepDescription: 'The ledger is written back' } } as Edge,
  ];
  return { id: 'scope-zone', nodes, edges };
}

/**
 * Forty-eight services in two rows of twenty-four. The hop that turns the row
 * is the longest in the drawing and the only thing that explains how row one
 * reaches row two, and a seam filter expressed purely as a fraction of the
 * whole hop drops it from every window.
 */
function wideChainScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 48; i += 1) {
    nodes.push(svc(`w${i}`, `Service ${i}`, (i % 24) * 300, Math.floor(i / 24) * 220));
    if (i > 0) {
      edges.push({
        id: `w${i - 1}-${i}`,
        source: `w${i - 1}`,
        target: `w${i}`,
        label: `step ${i}`,
        data: { stepNumber: i },
      } as Edge);
    }
  }
  return { id: 'wide-chain', nodes, edges };
}

/**
 * A grid one node wider than `tight-grid`, at the pitch where the callouts sat
 * 87% inside a tile and the old 0.9 burial bar let them through.
 */
function grid5x5TightScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 25; i += 1) {
    nodes.push(svc(`a${i}`, `Service ${i}`, (i % 5) * 168, Math.floor(i / 5) * 108));
    if (i > 0) {
      edges.push({
        id: `a${i - 1}-${i}`,
        source: `a${i - 1}`,
        target: `a${i}`,
        label: `step ${i}`,
        data: { stepNumber: i },
      } as Edge);
    }
  }
  return { id: 'grid5x5-tight', nodes, edges };
}

function barbellScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 6; i += 1) nodes.push(svc(`l${i}`, `Left Service ${i}`, (i % 2) * 220, Math.floor(i / 2) * 200));
  for (let i = 0; i < 6; i += 1) nodes.push(svc(`r${i}`, `Right Service ${i}`, 3200 + (i % 2) * 220, Math.floor(i / 2) * 200));
  for (let i = 1; i < 6; i += 1) {
    edges.push({ id: `le${i}`, source: `l${i - 1}`, target: `l${i}`, label: `left hop ${i}`, data: { stepNumber: i } } as Edge);
  }
  edges.push({ id: 'bridge', source: 'l5', target: 'r0', label: 'private peering', data: { stepNumber: 6 } } as Edge);
  for (let i = 1; i < 6; i += 1) {
    edges.push({ id: `re${i}`, source: `r${i - 1}`, target: `r${i}`, label: `right hop ${i}`, data: { stepNumber: i + 6 } } as Edge);
  }
  return { id: 'barbell', nodes, edges };
}

/**
 * Six parallel edges between one close-together pair, each with a long CJK
 * label. The routes are already fanned apart by a fraction of a rung, so a
 * stagger measured from each route's own anchor lands the chips off the
 * lattice and half inside each other from the fourth rung on.
 */
function parallelScenario(): Scenario {
  const nodes = [svc('pa', 'Azure Front Door', 0, 0), svc('pb', 'Azure Kubernetes Service', 190, 0)];
  const label = 'ゲートウェイ経由の HTTPS';
  const edges = Array.from({ length: 6 }, (_, i) => ({
    id: `par${i + 1}`,
    source: 'pa',
    target: 'pb',
    label: `${label} ${i + 1}`,
    data: { stepNumber: i + 1 },
  })) as Edge[];
  return { id: 'parallel', nodes, edges };
}

/**
 * A deep fan dropped into a crowded grid. The ladder is far larger than any one
 * chip, so unless it is the thing that dodges — and unless the chips it still
 * lands on are then moved out from under it — it shunts unrelated labels into
 * each other well away from the fan itself.
 */
function ladderInGridScenario(): Scenario {
  const nodes: Node[] = [];
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col < 5; col += 1) nodes.push(svc(`g${row}-${col}`, `Service ${row}${col}`, col * 290, row * 180));
  }
  const edges: Edge[] = [];
  for (let i = 1; i < nodes.length; i += 1) {
    edges.push({
      id: `hop${i}`, source: nodes[i - 1].id, target: nodes[i].id, label: `ホップ ${i}`, data: { stepNumber: i },
    } as Edge);
  }
  for (let i = 0; i < 7; i += 1) {
    edges.push({
      id: `fan${i}`,
      source: 'g0-0',
      target: 'g0-1',
      label: `マネージド ID で参照系を照会します ${i + 1}`,
      data: { stepNumber: nodes.length + i },
    } as Edge);
  }
  return { id: 'ladder-in-grid', nodes, edges };
}

/**
 * Two deep fans on neighbouring rows of the same grid. Each ladder is larger
 * than the corridor it belongs to, so both have to step off it - and the clear
 * band one of them finds is the band the other one wanted. A bundle scored
 * only against the drawing parks itself straight on top of its neighbour.
 */
function twinLaddersScenario(): Scenario {
  const nodes: Node[] = [];
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col < 5; col += 1) nodes.push(svc(`t${row}-${col}`, `Service ${row}${col}`, col * 290, row * 180));
  }
  const edges: Edge[] = [];
  for (let i = 1; i < nodes.length; i += 1) {
    edges.push({
      id: `w${i}`, source: nodes[i - 1].id, target: nodes[i].id, label: `ホップ ${i}`, data: { stepNumber: i },
    } as Edge);
  }
  for (let i = 0; i < 4; i += 1) {
    edges.push({
      id: `u${i}`, source: 't1-0', target: 't1-1', label: `マネージド ID で参照系を照会します ${i + 1}`, data: { stepNumber: nodes.length + i },
    } as Edge);
  }
  for (let i = 0; i < 10; i += 1) {
    edges.push({
      id: `d${i}`, source: 't2-0', target: 't2-1', label: `イベントを Service Bus に発行します ${i + 1}`, data: { stepNumber: nodes.length + 20 + i },
    } as Edge);
  }
  return { id: 'twin-ladders', nodes, edges };
}
/**
 * A fan on a roomy grid. There is clear air a long way off in every direction,
 * so a ladder scored only on what it covers will happily walk to the far side
 * of the drawing and settle beside somebody else's arrow. Nothing about that
 * placement looks wrong to a collision check — it is perfectly clean — but the
 * reader matches the wording to the arrow nearest it and gets the wrong hop.
 */
function strayLadderScenario(): Scenario {
  const nodes: Node[] = [];
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col < 4; col += 1) nodes.push(svc(`s${row}-${col}`, `Service ${row}${col}`, col * 260, row * 170));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let row = 0; row < 3; row += 1) {
    for (let col = 0; col + 1 < 4; col += 1) {
      step += 1;
      edges.push({
        id: `h${row}_${col}`, source: `s${row}-${col}`, target: `s${row}-${col + 1}`,
        label: 'マネージド ID で参照系を照会します', data: { stepNumber: step },
      } as Edge);
    }
  }
  for (let i = 0; i < 8; i += 1) {
    step += 1;
    edges.push({
      id: `fan${i}`, source: 's1-0', target: 's1-1',
      label: `注文ドキュメントを Cosmos DB に書き込みます ${i + 1}`, data: { stepNumber: step },
    } as Edge);
  }
  return { id: 'stray-ladder', nodes, edges };
}
/**
 * A dense grid with all four connection types in play, so the colour key is at
 * its tallest, and enough labelled hops that the bottom-left corner is busy.
 * The key is drawn last and is all but opaque: anything under it is gone from
 * the finished deck, and a buried callout leaves the workflow band citing a
 * step the reader cannot find anywhere on the drawing.
 */
/**
 * A model asked for one flow twice hands several arrows the SAME step number,
 * each with its own sentence. The workflow list is keyed by number, so every
 * sentence after the first was dropped while all of those badges still read the
 * same digit.
 */
function duplicateStepsScenario(): Scenario {
  const nodes: Node[] = [
    svc('web', 'App Service', 0, 0),
    svc('api', 'API Management', 300, 0),
    svc('db', 'Azure SQL Database', 600, 0),
    svc('cache', 'Azure Cache for Redis', 300, 190),
    svc('log', 'Log Analytics', 600, 190),
  ];
  const hops: [string, string, string][] = [
    ['web', 'api', 'ユーザー要求をゲートウェイに転送します'],
    ['api', 'db', '注文レコードを読み書きします'],
    ['api', 'cache', 'セッション状態をキャッシュします'],
    ['api', 'log', '要求メトリックを送信します'],
    ['db', 'log', '監査ログを送信します'],
  ];
  const edges: Edge[] = hops.map(([source, target, label], i) => ({
    id: `dup${i}`, source, target, label,
    // Every one of them numbered 3, which is exactly what a re-prompted model emits.
    data: { stepNumber: 3, stepDescription: `${label}。` },
  } as Edge));
  return { id: 'duplicate-steps', nodes, edges };
}

function legendCornerScenario(): Scenario {
  const nodes: Node[] = [];
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col < 6; col += 1) nodes.push(svc(`g${row}${col}`, `Azure Service ${row}${col}`, col * 260, row * 190));
  }
  const kinds = ['sync', 'async', 'telemetry', 'data'];
  const edges: Edge[] = [];
  let step = 0;
  for (let row = 0; row < 4; row += 1) {
    for (let col = 0; col + 1 < 6; col += 1) {
      step += 1;
      edges.push({
        id: `h${row}${col}`, source: `g${row}${col}`, target: `g${row}${col + 1}`,
        label: 'マネージド ID で注文ドキュメントを書き込みます',
        data: { connectionType: kinds[row % 4], stepNumber: step, stepDescription: `手順 ${step}` },
      } as Edge);
    }
  }
  return { id: 'legend-corner', nodes, edges };
}
/**
 * One product group containing a dense field of services. A zone is a single
 * box, so the tiler used to see one shape it could not split and grew the page
 * into a plotter sheet the whole deck then inherited.
 */
function gridFanScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) nodes.push(svc(`g${r}${c}`, `Azure Service ${r}${c}`, c * 300, r * 200));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c + 1 < 3; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`, label: '注文ドキュメントを Cosmos DB に書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `g${r}${c}`, target: `g${r + 1}${c}`, label: '注文ドキュメントを Cosmos DB に書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 5; i += 1) {
    step += 1;
    edges.push({ id: `f${i}`, source: 'g11', target: 'g12', label: `注文ドキュメントを Cosmos DB に書き込みます ${i}`, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
  }
  return { id: 'grid-fan', nodes, edges };
}

/**
 * The same grid, but the fan is three deep instead of five. Three is the awkward
 * depth: too many to sit on the arrows as a single chip, too few to trip the
 * mute that turns a fan into loose numbers. So it stays a rigid three-rung
 * ladder on the most crowded row of the drawing, which is exactly the shape
 * that has nowhere to stand.
 */
function gridFan3Scenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) nodes.push(svc(`g${r}${c}`, `Azure Service ${r}${c}`, c * 300, r * 200));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 3; r += 1) {
    for (let c = 0; c + 1 < 3; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`, label: '注文ドキュメントを Cosmos DB に書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 3; r += 1) {
    for (let c = 0; c < 3; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `g${r}${c}`, target: `g${r + 1}${c}`, label: '注文ドキュメントを Cosmos DB に書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 3; i += 1) {
    step += 1;
    edges.push({ id: `f${i}`, source: 'g11', target: 'g12', label: `注文ドキュメントを Cosmos DB に書き込みます ${i}`, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
  }
  return { id: 'grid3x3-fan3-JA', nodes, edges };
}

/**
 * A 5x5 grid on tight spacing with a fan of eight in the middle of it. The
 * densest shape in the corpus: every hop has neighbours on all four sides, so
 * there is no clear air anywhere for anything to escape into, and the fan is
 * deep enough that its ladder is taller than the row it stands in.
 */
/**
 * The reviewer's caption fixture: a plain 5x5 grid on a 210x140 pitch, every
 * edge carrying the same sentence, no fan anywhere. The tight vertical pitch
 * leaves 65px between rows, which is less than a chip is tall, so chips are
 * pushed onto the tile below — and onto the one thing that says which service
 * that tile is.
 */
/**
 * Long names on a tight pitch. A short name is one centred line in the middle
 * of the tile, so a chip lapping the tile's edge misses the words entirely; a
 * name that wraps to three lines fills the tile, and then the same lap lands
 * squarely on the letters. This is the case where "the chip is only 8% over
 * the tile" and "the chip is sitting on the name" are the same event.
 */
function longNameGridScenario(): Scenario {
  const nodes: Node[] = [];
  const names = [
    'Azure Kubernetes Service 本番クラスター',
    'Azure Database for PostgreSQL フレキシブル サーバー',
    'Azure Container Registry プレミアム',
    'Microsoft Entra ID ワークロード ID',
  ];
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c < 4; c += 1) nodes.push(svc(`n${r}${c}`, names[(r + c) % names.length], c * 205, r * 135));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c + 1 < 4; c += 1) {
      step += 1;
      edges.push({ id: `lh${r}${c}`, source: `n${r}${c}`, target: `n${r}${c + 1}`, label: '注文ドキュメントを書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 4; r += 1) {
    for (let c = 0; c < 4; c += 1) {
      step += 1;
      edges.push({ id: `lv${r}${c}`, source: `n${r}${c}`, target: `n${r + 1}${c}`, label: '参照系を照会します', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  return { id: 'long-names-tight', nodes, edges };
}

function grid5x5CaptionScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) nodes.push(svc(`g${r}${c}`, `Azure Service ${r}${c}`, c * 210, r * 140));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `g${r}${c}`, target: `g${r}${c + 1}`, label: 'writes order documents to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `g${r}${c}`, target: `g${r + 1}${c}`, label: 'writes order documents to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  return { id: 'grid5x5-captions', nodes, edges };
}

/**
 * A 5×5 grid whose vertical hops carry an ordinary 45-character sentence.
 *
 * No fan, no metadata, no CJK, stock names — the least exotic diagram that can
 * be drawn, and the corridor between rows is still too narrow for the chip.
 */
function longLabelGridScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) nodes.push(svc(`p${r}${c}`, `Azure Service ${r}${c}`, c * 210, r * 140));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `p${r}${c}`, target: `p${r}${c + 1}`, label: 'writes order documents to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `p${r}${c}`, target: `p${r + 1}${c}`, label: 'queries the read model with a managed identity', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  return { id: 'long-label-grid', nodes, edges };
}

/**
 * The same grid with SKU / region / price on every tile, so the bottom-anchored
 * sub-line is present — the strip a chip lapping its endpoint tile from below
 * lands on, which nothing modelled and no rule measured.
 */
function metaTightScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      const node = svc(`q${r}${c}`, `Azure Database for PostgreSQL ${r}${c}`, c * 210, r * 140);
      Object.assign(node.data as Record<string, unknown>, {
        sku: 'Standard_D4s_v5',
        region: 'japaneast',
        pricing: { estimatedCost: 128.4, quantity: 1, region: 'japaneast' },
      });
      nodes.push(node);
    }
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `q${r}${c}`, target: `q${r}${c + 1}`, label: 'writes order documents to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `q${r}${c}`, target: `q${r + 1}${c}`, label: 'queries the read model with a managed identity', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 8; i += 1) {
    step += 1;
    edges.push({ id: `mf${i}`, source: 'q22', target: 'q23', label: `replicates the order stream ${i}`, data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
  }
  return { id: 'meta-tight', nodes, edges };
}

/** The same pressure with CJK names long enough to fill all three tile lines. */
function longNameFanScenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      nodes.push(svc(`w${r}${c}`, `Azure Database for PostgreSQL フレキシブル サーバー ${r}${c}`, c * 210, r * 140));
    }
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `w${r}${c}`, target: `w${r}${c + 1}`, label: '注文ドキュメントを Cosmos DB に書き込みます', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `w${r}${c}`, target: `w${r + 1}${c}`, label: 'マネージド ID で参照系を照会します', data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 8; i += 1) {
    step += 1;
    edges.push({ id: `wf${i}`, source: 'w22', target: 'w23', label: `注文ストリームを複製します ${i}`, data: { stepNumber: step, stepDescription: `手順 ${step}` } } as Edge);
  }
  return { id: 'long-name-fan', nodes, edges };
}

function fan8Tight5x5Scenario(): Scenario {
  const nodes: Node[] = [];
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) nodes.push(svc(`t${r}${c}`, `Azure Service ${r}${c}`, c * 215, r * 150));
  }
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 5; r += 1) {
    for (let c = 0; c + 1 < 5; c += 1) {
      step += 1;
      edges.push({ id: `h${r}${c}`, source: `t${r}${c}`, target: `t${r}${c + 1}`, label: 'writes the order document to Cosmos DB', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let r = 0; r + 1 < 5; r += 1) {
    for (let c = 0; c < 5; c += 1) {
      step += 1;
      edges.push({ id: `v${r}${c}`, source: `t${r}${c}`, target: `t${r + 1}${c}`, label: 'queries the read model with a managed identity', data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
    }
  }
  for (let i = 0; i < 8; i += 1) {
    step += 1;
    edges.push({ id: `f${i}`, source: 't22', target: 't23', label: `writes the order document to Cosmos DB ${i}`, data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
  }
  return { id: 'fan8-5x5-tight', nodes, edges };
}

/** A plain chain of 40 services, no fans at all — the least exotic estate there is. */
function estateChainScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 40; i += 1) nodes.push(svc(`n${i}`, `Azure Service ${i}`, (i % 8) * 240, Math.floor(i / 8) * 190));
  const edges: Edge[] = [];
  for (let i = 1; i < 40; i += 1) {
    edges.push({ id: `c${i}`, source: `n${i - 1}`, target: `n${i}`, label: 'マネージド ID で参照系を照会します', data: { stepNumber: i, stepDescription: `手順 ${i}` } } as Edge);
  }
  return { id: 'estate-chain', nodes, edges };
}

/**
 * The same shape as the estate chain, but six per row and with an English
 * sentence for a label. A scenario proves a fix for the string it carries, so
 * the chain is run at both a CJK width and a Latin one.
 */
function chain24Scenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 24; i += 1) nodes.push(svc(`n${i}`, `Azure Service ${i}`, (i % 6) * 240, Math.floor(i / 6) * 190));
  const edges: Edge[] = [];
  for (let i = 1; i < 24; i += 1) {
    edges.push({ id: `c${i}`, source: `n${i - 1}`, target: `n${i}`, label: 'writes the order document to Cosmos DB', data: { stepNumber: i, stepDescription: `Step ${i}` } } as Edge);
  }
  return { id: 'chain24-en', nodes, edges };
}

/**
 * Three fans of five stacked on adjacent rows. All three mute, and once a
 * muted fan is placed as loose numbers rather than a lattice, the lowest fan's
 * callouts are free to drift onto the hops of the fan above it. This is the
 * only shape in the corpus with more than one muted fan.
 */
function tripleMutedScenario(): Scenario {
  const EN = 'writes the order document to Cosmos DB';
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  let step = 0;
  for (let r = 0; r < 4; r += 1) {
    for (let c = 0; c < 4; c += 1) nodes.push(svc(`g${r}${c}`, `Azure Service ${r}${c}`, c * 250, r * 165));
  }
  const hop = (id: string, a: string, b: string, label: string): void => {
    step += 1;
    edges.push({ id, source: a, target: b, label, data: { stepNumber: step, stepDescription: `Step ${step}` } } as Edge);
  };
  for (let r = 0; r < 4; r += 1) for (let c = 0; c < 3; c += 1) hop(`h${r}${c}`, `g${r}${c}`, `g${r}${c + 1}`, EN);
  for (let r = 0; r < 3; r += 1) for (let c = 0; c < 4; c += 1) hop(`v${r}${c}`, `g${r}${c}`, `g${r + 1}${c}`, EN);
  ([['g11', 'g12'], ['g21', 'g22'], ['g31', 'g32']] as const).forEach(([a, b], fan) => {
    for (let i = 0; i < 5; i += 1) hop(`F${fan}_${i}`, a, b, `${EN} ${fan}${i}`);
  });
  return { id: 'triple-muted', nodes, edges };
}

/**
 * Past 72 services the overview clamps tile type to 4pt. This is the fixture
 * that decides what a tile does when it can no longer be named.
 */
function estate72Scenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 72; i += 1) {
    nodes.push(svc(`e${i}`, `Azure Container Apps ${i}`, (i % 9) * 230, Math.floor(i / 9) * 165));
    if (i > 0) edges.push({ id: `k${i}`, source: `e${i - 1}`, target: `e${i}`, label: 'HTTPS' } as Edge);
  }
  return { id: 'estate72', nodes, edges };
}


function denseZoneScenario(): Scenario {
  const nodes: Node[] = [grp('zone', 'Production landing zone', 0, 0, 2400, 1200)];
  const edges: Edge[] = [];
  for (let i = 0; i < 28; i += 1) {
    nodes.push(svc(
      `d-${i}`,
      i % 2 ? 'Azure Kubernetes Service' : 'Azure Container Registry',
      60 + (i % 7) * 320,
      90 + Math.floor(i / 7) * 260,
      'zone',
    ));
    if (i > 0) {
      edges.push({
        id: `dz-${i}`, source: `d-${i - 1}`, target: `d-${i}`, label: 'private endpoint', data: { stepNumber: i },
      } as Edge);
    }
  }
  return { id: 'dense-zone', nodes, edges };
}

/** Mirrors a real AI-generated enterprise diagram: wide, grouped, long labels. */
function wideScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  const names = [
    'Copilot Studio', 'Key Vault', 'Azure OpenAI Service', 'Azure AI Search',
    'Azure Kubernetes Service', 'Azure SQL Database', 'Application Gateway', 'Azure Front Door',
    'Azure Functions', 'Azure Service Bus', 'Azure Data Factory', 'Azure Synapse Analytics',
  ];
  const zones = ['Ingress zone', 'Application zone', 'Data zone', 'Integration zone'];
  zones.forEach((zone, z) => {
    nodes.push(grp(`zone-${z}`, zone, z * 900, 0, 820, 560));
    for (let i = 0; i < 3; i += 1) {
      const idx = z * 3 + i;
      nodes.push(svc(`svc-${idx}`, names[idx], 60 + (i % 2) * 380, 90 + Math.floor(i / 2) * 200, `zone-${z}`));
    }
  });
  for (let i = 0; i < 11; i += 1) {
    edges.push({
      id: `e-${i}`,
      source: `svc-${i}`,
      target: `svc-${i + 1}`,
      // Half the flow is numbered, so the audit sees both the badge path and
      // the unnumbered path in the same drawing.
      ...(i < 6 ? { data: { stepNumber: i + 1, stepDescription: `ステップ ${i + 1}: サービス間の呼び出しを実行します` } } : {}),
      label: i % 3 === 0 ? 'HTTPS 経由でトークン検証を実施' : i % 3 === 1 ? 'Private Link' : 'Managed identity authentication',
    } as Edge);
  }
  return { id: 'wide', nodes, edges };
}

function compactScenario(): Scenario {
  const nodes = [
    grp('z', 'Application zone', 0, 0, 520, 320),
    svc('a', 'API Management', 60, 80, 'z'),
    svc('b', 'Azure Functions', 320, 80, 'z'),
  ];
  return { id: 'compact', nodes, edges: [{ id: 'e', source: 'a', target: 'b', label: 'Invoke worker' } as Edge] };
}

/** Beyond the 56" page limit: proves the fallback downscale stays legible. */
function oversizeScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 40; i += 1) {
    nodes.push(svc(`n-${i}`, i % 2 ? 'Azure Kubernetes Service' : 'Copilot Studio', i * 260, (i % 4) * 220));
    if (i > 0) edges.push({ id: `x-${i}`, source: `n-${i - 1}`, target: `n-${i}`, label: 'Managed identity authentication' } as Edge);
  }
  return { id: 'oversize', nodes, edges };
}

/**
 * Banding, numbering and an outlier at once. Each rule existed but none had a
 * scenario where they interact: a stray belongs to no band under a plain range
 * test, and a shape straddling a seam is admitted by two bands at once.
 */
function bandedScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 30; i += 1) {
    nodes.push(svc(`b-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', i * 300, (i % 3) * 200));
    if (i > 0) {
      edges.push({
        id: `y-${i}`,
        source: `b-${i - 1}`,
        target: `b-${i}`,
        label: 'HTTPS 経由でトークン検証を実施',
        ...(i <= 8 ? { data: { stepNumber: i, stepDescription: `ステップ ${i}: 帯をまたぐ呼び出しを実行します` } } : {}),
      } as Edge);
    }
  }
  nodes.push(svc('b-stray', 'Copilot Studio', -14000, -6000));
  return { id: 'banded', nodes, edges };
}

/**
 * Twenty narrated steps: rows stop shrinking at the legible minimum, so the
 * list has to continue onto another slide rather than drop its tail.
 */
function narrativeScenario(): Scenario {
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let i = 0; i < 21; i += 1) {
    nodes.push(svc(`w-${i}`, i % 2 ? 'Azure Service Bus' : 'Azure Functions', (i % 7) * 300, Math.floor(i / 7) * 240));
    if (i > 0) {
      edges.push({
        id: `w-e-${i}`,
        source: `w-${i - 1}`,
        target: `w-${i}`,
        label: 'Private Link',
        data: { stepNumber: i, stepDescription: `ステップ ${i}: マネージド ID による認証を経てメッセージを転送します` },
      } as Edge);
    }
  }
  return { id: 'narrative', nodes, edges };
}

/** A dense cluster plus one far-placed node: nothing may fall off the page. */
function outlierScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push(svc(`c-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', (i % 4) * 220, Math.floor(i / 4) * 180));
  }
  nodes.push(svc('outlier', 'Copilot Studio', 9000, 4000));
  // Numbered, labelled edges on the clamped path: this is the only
  // configuration where the badge can be clamped back onto its own label chip,
  // so without these edges that rule was never actually evaluated.
  const edges: Edge[] = [
    { id: 'e-out', source: 'c-0', target: 'outlier', label: 'HTTPS 経由でトークン検証を実施', data: { stepNumber: 1, stepDescription: '外れ値のサービスへ接続します' } } as Edge,
    { id: 'e-in', source: 'c-1', target: 'c-2', label: 'Managed identity authentication', data: { stepNumber: 2, stepDescription: 'マネージド ID で認証します' } } as Edge,
  ];
  return { id: 'outlier', nodes, edges };
}

interface Report {
  scenario: string;
  format: string;
  issues: string[];
  metrics: Record<string, number>;
  /**
   * Every service name each format actually DRAWS, paired with the label it
   * was authored from.
   *
   * A name is compared by presence and by how much of it survives, never by
   * count: the two formats export at deliberately different page scales, so a
   * sheet legitimately names more tiles than the deck. What is never
   * legitimate is a service named in one file and absent from the other, or
   * one kept whole in one file and cut to a stub in the other.
   *
   * It lives on the report rather than inside either auditor because no rule
   * inside either auditor could see it: each one only ever reads the shapes
   * its own exporter emitted, so a name that is missing from one file is
   * invisible by construction.
   */
  drawnNames?: { authored: string; drawn: string }[];
}

/**
 * The shape the AI actually returns, run through the real layout engine.
 *
 * Every other scenario hand-places its nodes, so until this one existed the
 * audit never saw what a generated diagram looks like — and a linear flow is by
 * far the most common thing a model produces.
 */
async function generatedScenario(): Promise<Scenario> {
  const { applyLayoutPreset } = await import('../src/utils/layoutPresets');
  const names = [
    'Azure Front Door', 'Application Gateway', 'Azure Kubernetes Service', 'Azure Service Bus',
    'Azure Functions', 'Azure Cosmos DB', 'Azure Data Factory', 'Azure Synapse Analytics',
    'Azure OpenAI Service', 'Azure AI Search', 'Key Vault', 'Azure Monitor',
  ];
  const nodes: Node[] = names.map((name, i) => svc(`g-${i}`, name, 0, 0));
  const edges: Edge[] = names.slice(1).map((_, i) => ({
    id: `g-e-${i + 1}`,
    source: `g-${i}`,
    target: `g-${i + 1}`,
    label: 'HTTPS 経由でトークン検証を実施',
    data: { stepNumber: i + 1, stepDescription: `ステップ ${i + 1}: 次のサービスへ要求を引き渡します` },
  } as Edge));

  const laidOut = await applyLayoutPreset(nodes, edges, {
    preset: 'flow-lr', spacing: 'comfortable', edgeStyle: 'smooth', emphasizePrimaryPath: false,
  });
  return { id: 'generated', nodes: laidOut.nodes, edges: laidOut.edges, fromLayoutEngine: true };
}

/**
 * What the generator is actually told to produce: 3 zones, 10 services,
 * hub-and-spoke telemetry, numbered flow — then run through the real layout
 * preset. `wide` is grouped but hand-placed and `generated` is engine-laid-out
 * but flat, so until now no scenario exercised grouping and the layout engine
 * at the same time, which is every diagram a user actually gets.
 */
async function groupedGeneratedScenario(): Promise<Scenario> {
  const { applyLayoutPreset } = await import('../src/utils/layoutPresets');
  const zones: { id: string; label: string; members: string[] }[] = [
    { id: 'z-edge', label: 'Ingress zone', members: ['Azure Front Door', 'Application Gateway'] },
    { id: 'z-app', label: 'Application zone', members: ['Azure Kubernetes Service', 'Azure Functions', 'Azure Service Bus'] },
    { id: 'z-data', label: 'Data zone', members: ['Azure Cosmos DB', 'Azure SQL Database', 'Azure Data Lake Storage'] },
    { id: 'z-ops', label: 'Security & operations', members: ['Microsoft Entra ID', 'Key Vault', 'Azure Monitor'] },
  ];
  const nodes: Node[] = [];
  const flat: string[] = [];
  zones.forEach((zone, z) => {
    nodes.push(grp(zone.id, zone.label, z * 900, 0, 820, 560));
    zone.members.forEach((name, i) => {
      const id = `gg-${z}-${i}`;
      nodes.push(svc(id, name, 60 + (i % 2) * 380, 90 + Math.floor(i / 2) * 200, zone.id));
      flat.push(id);
    });
  });
  const link = (n: number, from: string, to: string, label: string): Edge => ({
    id: `gg-e-${n}`,
    source: from,
    target: to,
    label,
    data: { stepNumber: n, stepDescription: `ステップ ${n}: ${label}` },
  } as Edge);
  const edges: Edge[] = [
    link(1, 'gg-0-0', 'gg-0-1', 'WAF で検査した要求を転送'),
    link(2, 'gg-0-1', 'gg-1-0', 'コンテナー化された API へ負荷分散'),
    link(3, 'gg-1-0', 'gg-1-2', '注文イベントを非同期で発行'),
    link(4, 'gg-1-2', 'gg-1-1', 'キューの受信でハンドラーを起動'),
    link(5, 'gg-1-1', 'gg-2-0', 'マネージド ID で注文ドキュメントを書き込み'),
    link(6, 'gg-1-0', 'gg-2-1', 'Private Endpoint 経由で参照系を照会'),
    link(7, 'gg-1-1', 'gg-2-2', '分析用に生データを保管'),
    link(8, 'gg-3-0', 'gg-1-0', 'ワークロード ID にトークンを発行'),
    link(9, 'gg-1-0', 'gg-3-1', '接続シークレットをマネージド ID で取得'),
    link(10, 'gg-1-0', 'gg-3-2', 'ログとメトリックを送信'),
  ];
  const laidOut = await applyLayoutPreset(nodes, edges, {
    preset: 'flow-lr', spacing: 'comfortable', edgeStyle: 'smooth', emphasizePrimaryPath: false,
  });
  return { id: 'grouped-generated', nodes: laidOut.nodes, edges: laidOut.edges, fromLayoutEngine: true };
}

function countByName(shapes: { name: string }[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const shape of shapes) counts.set(shape.name, (counts.get(shape.name) ?? 0) + 1);
  return counts;
}

/**
 * Put an icon on every tile before the conversion is measured.
 *
 * Nothing rasterizes under Node — `canRasterize()` wants `document` and
 * `Image` — so a generated deck reaches the audit with zero `icon-*` shapes.
 * That is not the deck a user exports, and it silently switched off the whole
 * grouping half of the conversion: with no icon there is nothing to group, so
 * the group frame, the child z-order, and the question of gluing a connector
 * to a shape nested inside a `<p:grpSp>` were all scored at 0% coverage.
 *
 * The geometry mirrors `pptxExporter`'s own `addImage` call: square, centred
 * across the tile, sitting just below its top edge.
 */
function withSynthesizedIcons(slideXml: string): string {
  const tiles = parseShapes(slideXml).filter(
    (s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'),
  );
  if (tiles.length === 0) return slideXml;
  const usedIds = [...slideXml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => +m[1]);
  let nextId = Math.max(0, ...usedIds) + 1;
  const pics: string[] = [];
  for (const tile of tiles) {
    const size = Math.min(0.6, tile.w * 0.3, tile.h * 0.42);
    if (size <= 0) continue;
    const emu = (v: number): number => Math.round(v * EMU_PER_INCH);
    pics.push(
      `<p:pic><p:nvPicPr><p:cNvPr id="${nextId}" name="icon-${tile.name.slice('service-'.length)}"/>`
      + `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
      + `<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`
      + `<p:spPr><a:xfrm><a:off x="${emu(tile.x + (tile.w - size) / 2)}" y="${emu(tile.y + tile.h * 0.06)}"/>`
      + `<a:ext cx="${emu(size)}" cy="${emu(size)}"/></a:xfrm>`
      + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
    );
    nextId += 1;
  }
  if (pics.length === 0) return slideXml;
  return slideXml.replace('</p:spTree>', `${pics.join('')}</p:spTree>`);
}

/**
 * The deck is repaired into real PowerPoint objects after pptxgenjs has
 * written it — connectors glued to the services they join, service names
 * inside their tiles, tiles grouped with their icons. The rules the drawing is
 * measured by all address shapes that conversion moves or removes, so the
 * conversion gets its own rules, run on the converted XML.
 *
 * The one that matters most is the last: a conversion that quietly eats a
 * service name would satisfy every structural rule perfectly.
 */
function auditNativeConversion(
  rawSlides: readonly string[],
  edges: readonly Edge[] = [],
): { issues: string[]; glued: number; ungluable: number; groups: number } {
  const issues: string[] = [];
  const edgeById = new Map(edges.map((edge) => [String(edge.id), edge]));
  let glued = 0;
  let ungluable = 0;
  let groups = 0;
  const allSlides = rawSlides.map(withSynthesizedIcons);
  let seenAnchorable = false;

  allSlides.forEach((slideXml, index) => {
    const before = parseShapes(slideXml);
    const after = nativizeSlideXml(slideXml);
    const where = `slide ${index + 1}`;

    for (const tag of ['p:sp', 'p:cxnSp', 'p:grpSp', 'p:pic', 'p:txBody'] as const) {
      const open = (after.match(new RegExp(`<${tag}>`, 'g')) ?? []).length;
      const close = (after.match(new RegExp(`</${tag}>`, 'g')) ?? []).length;
      if (open !== close) issues.push(`${where}: converted XML has ${open} <${tag}> but ${close} </${tag}>`);
    }

    // A connector glued to an id that is not on the slide is dropped by
    // PowerPoint on open, which loses the arrow entirely.
    const ids = new Set([...after.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]));
    for (const glue of after.matchAll(/<a:(?:st|end)Cxn id="(\d+)" idx="(\d+)"\/>/g)) {
      if (!ids.has(glue[1])) issues.push(`${where}: connector glued to shape id ${glue[1]}, which is not on the slide`);
      if (+glue[2] > 3) issues.push(`${where}: connector glued to site ${glue[2]}, which a rectangle does not have`);
    }

    for (const cxn of after.matchAll(/<p:cxnSp>[\s\S]*?<\/p:cxnSp>/g)) {
      if (/<a:stCxn /.test(cxn[0]) && /<a:endCxn /.test(cxn[0])) glued += 1;
      else ungluable += 1;
    }
    // Glue is the difference between a deck the reader can rearrange and one
    // that falls apart on the first drag: an unglued line stays behind when its
    // tile moves. A hop cut by a window seam cannot be glued and must not be
    // counted against the exporter, so only the hops whose BOTH endpoints
    // already coincide with a tile drawn on this very slide are judged. Two
    // thirds of the arrows on the shared-service deck were reported unglueable
    // and nothing in the audit had ever looked at why.
    const siteSlack = 0.021;
    const tilesHere = before.filter(
      (s) => s.name.startsWith('service-') && !s.name.startsWith('service-label-') && !s.name.startsWith('service-meta-') && s.w > 0,
    );
    const onSite = (p: { x: number; y: number }): boolean => tilesHere.some((t) => [
      { x: t.x + t.w / 2, y: t.y },
      { x: t.x, y: t.y + t.h / 2 },
      { x: t.x + t.w / 2, y: t.y + t.h },
      { x: t.x + t.w, y: t.y + t.h / 2 },
    ].some((site) => Math.hypot(site.x - p.x, site.y - p.y) <= siteSlack));
    const nativeById = new Map<string, string>();
    for (const cxn of after.matchAll(/<p:(?:cxnSp|sp)>[\s\S]*?<\/p:(?:cxnSp|sp)>/g)) {
      const name = /<p:cNvPr id="\d+" name="(connector-[^"]*)"/.exec(cxn[0])?.[1];
      if (name) nativeById.set(name, cxn[0]);
    }
    for (const arrow of before.filter((s) => s.name.startsWith('connector-'))) {
      const path = arrow.path ?? [];
      if (path.length < 2) continue;
      if (!onSite(path[0]) || !onSite(path[path.length - 1])) continue;
      const xml = nativeById.get(arrow.name);
      if (xml && !(/<a:stCxn /.test(xml) && /<a:endCxn /.test(xml))) {
        issues.push(`${where}: arrow "${arrow.name}" starts and ends on connection sites but is not glued, so it detaches when the reader moves a tile`);
      }
    }
    // The rule above can only judge arrows that already land on a site, so it
    // would fall silent the day the router stopped putting them there — glue
    // would vanish deck-wide and the audit would report nothing. The overview
    // slide draws every tile, so on it every hop is anchorable; if most are
    // not, the endpoints themselves have drifted.
    const arrowsHere = before.filter((s) => s.name.startsWith('connector-') && (s.path ?? []).length >= 2);
    if (!seenAnchorable && arrowsHere.length >= 4) {
      seenAnchorable = true;
      const anchored = arrowsHere.filter((s) => onSite(s.path![0]) && onSite(s.path![s.path!.length - 1])).length;
      if (anchored < 0.6 * arrowsHere.length) {
        issues.push(`${where}: only ${anchored} of ${arrowsHere.length} arrows begin and end on a connection site, so most of the deck cannot be glued at all`);
      }
    }
    // Both rules above judge an arrow against the SITES, and an arrow that
    // reaches neither is exempt from the first and just a statistic in the
    // second. That is the shape a mis-planned route takes: on a clamped
    // drawing the router aimed a hop at a stray's declared position while the
    // tile was drawn somewhere else entirely, and the arrow finished 7in away
    // from the service it names, on a deck too small for the 60% floor to
    // apply. So measure the ends against the TILES they claim, whenever the
    // slide draws them — orientation-agnostic, because the exporter is free to
    // draw a hop from either end.
    const tileByName = new Map(tilesHere.map((t) => [t.name.slice('service-'.length), t]));
    for (const arrow of before.filter((s) => s.name.startsWith('connector-')
      && !s.name.startsWith('connector-label-') && !s.name.startsWith('connector-step-'))) {
      const path = arrow.path ?? [];
      if (path.length < 2) continue;
      const edge = edgeById.get(arrow.name.slice('connector-'.length));
      if (!edge) continue;
      const head = path[0];
      const tail = path[path.length - 1];
      for (const id of [String(edge.source), String(edge.target)]) {
        const tile = tileByName.get(id);
        if (!tile) continue;
        const gap = (p: { x: number; y: number }): number => Math.hypot(
          Math.max(tile.x - p.x, 0, p.x - (tile.x + tile.w)),
          Math.max(tile.y - p.y, 0, p.y - (tile.y + tile.h)),
        );
        const near = Math.min(gap(head), gap(tail));
        if (near > 0.2) {
          const at = `(${head.x.toFixed(2)},${head.y.toFixed(2)})->(${tail.x.toFixed(2)},${tail.y.toFixed(2)})`;
          const box = `[${tile.x.toFixed(2)},${tile.y.toFixed(2)} ${tile.w.toFixed(2)}x${tile.h.toFixed(2)}]`;
          issues.push(`${where}: arrow "${arrow.name}" ${at} ends ${near.toFixed(2)}in from "${id}" ${box}, the service it connects`);
        }
      }
    }
    groups += (after.match(/<p:grpSp>/g) ?? []).length;

    // A group frame that does not enclose its children clips them, and a
    // child offset that does not match the frame shifts every child by the
    // difference. Both are silent: the XML stays well formed and the deck
    // still opens, it just draws in the wrong place.
    for (const group of after.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)) {
      const name = /<p:cNvPr id="\d+" name="([^"]*)"/.exec(group[0])?.[1] ?? 'group';
      const frame = /<p:grpSpPr><a:xfrm><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/><a:chOff x="(-?\d+)" y="(-?\d+)"\/><a:chExt cx="(\d+)" cy="(\d+)"\/>/.exec(group[0]);
      if (!frame) {
        issues.push(`${where}: group "${name}" has no readable frame, so PowerPoint cannot place its children`);
        continue;
      }
      const [ox, oy, cx, cy, hx, hy, hcx, hcy] = frame.slice(1).map(Number);
      // Child coordinates are absolute only while the child origin and extent
      // equal the frame's. Any drift here scales and translates the contents.
      if (ox !== hx || oy !== hy || cx !== hcx || cy !== hcy) {
        issues.push(`${where}: group "${name}" child frame ${hx},${hy} ${hcx}x${hcy} differs from its own ${ox},${oy} ${cx}x${cy}, which shifts every child`);
      }
      for (const child of parseShapes(group[0])) {
        const cxEmu = child.x * EMU_PER_INCH;
        const cyEmu = child.y * EMU_PER_INCH;
        const slack = 0.01 * EMU_PER_INCH;
        if (
          cxEmu < ox - slack || cyEmu < oy - slack
          || cxEmu + child.w * EMU_PER_INCH > ox + cx + slack
          || cyEmu + child.h * EMU_PER_INCH > oy + cy + slack
        ) {
          issues.push(`${where}: group "${name}" does not enclose its child "${child.name}", which will be clipped`);
        }
      }
    }

    // The conversion groups a tile with its icon and also glues arrows to that
    // tile, so most glue ends up pointing at a shape nested inside a group.
    // If that ever stops resolving, every arrow silently detaches.
    for (const glue of after.matchAll(/<a:(?:st|end)Cxn id="(\d+)" idx="\d+"\/>/g)) {
      const target = new RegExp(`<p:cNvPr id="${glue[1]}" name="([^"]*)"`).exec(after)?.[1];
      if (target && !target.startsWith('service-')) {
        issues.push(`${where}: connector glued to "${target}", which is not a service tile`);
      }
    }

    // A duplicated shape id makes PowerPoint declare the file damaged and
    // repair it. The splice works by byte offset against the original shape
    // list, so a collision is exactly the failure this transform is most
    // likely to produce — and a glue check only proves ids *resolve*, not that
    // they resolve to one shape.
    const seenIds = new Set<string>();
    for (const decl of after.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g)) {
      if (seenIds.has(decl[1])) {
        issues.push(`${where}: conversion emitted shape id ${decl[1]} twice ("${decl[2]}")`);
      }
      seenIds.add(decl[1]);
    }

    // Nothing the reader could see may be lost by the conversion. The SKU /
    // region / price sub-line counts: a conversion that ate the cost figure
    // would otherwise pass every rule here.
    for (const label of before.filter(
      (s) => s.name.startsWith('service-label-') || s.name.startsWith('service-meta-'),
    )) {
      if (label.text.trim() === '') continue;
      // Per paragraph: the shape scrape now joins `<a:p>` with a newline, and
      // the emitted XML keeps each paragraph in its own `<a:t>`, so a
      // whole-string search for a multi-paragraph label can never match.
      const paragraphs = label.text.split('\n').filter((p) => p.trim() !== '');
      // Against the UNESCAPED runs, not against a re-escaped needle searched in
      // raw XML. `"` is legal unescaped in element content and `&` is not, so
      // re-escaping a name to look for it means guessing which of several legal
      // spellings the writer chose - and guessing wrong reports that a name the
      // conversion preserved perfectly was lost.
      const afterText = [...after.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
        .map((m) => unescapeXml(m[1]))
        .join('\n');
      if (!paragraphs.every((p) => afterText.includes(p))) {
        const kind = label.name.startsWith('service-meta-') ? 'service sub-line' : 'service name';
        issues.push(`${where}: conversion lost the ${kind} "${label.text}"`);
      }
    }
    for (const tile of before.filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'))) {
      const id = tile.name.slice('service-'.length);
      const grouped = new RegExp(`<p:grpSp>(?:(?!</p:grpSp>)[\\s\\S])*name="service-${escapeRe(id)}"`).test(after);
      if (before.some((s) => s.name === `icon-${id}`) && !grouped) {
        issues.push(`${where}: tile "${tile.name}" was not grouped with its icon, so dragging it leaves the icon behind`);
      }
    }
  });
  return { issues, glued, ungluable, groups };
}

function escapeRe(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Stand-in for the DOM rasteriser, which returns nothing under Node. Without
 * this every tile is drawn with no icon at all — a materially different tile
 * interior, with the caption band 2.1x too tall and 0.35in out of position —
 * so the chip walk, the spoiled-chip budget and every contrast composite were
 * being measured against a layout the user never receives.
 */
function synthesisedIcons(scenario: Scenario): Map<string, { bytes: Uint8Array; dataUrl: string; sizePx: number }> {
  const icons = new Map<string, { bytes: Uint8Array; dataUrl: string; sizePx: number }>();
  for (const node of scenario.nodes) {
    const path = (node.data as { iconPath?: string } | undefined)?.iconPath;
    if (path) icons.set(path, { bytes: PIXEL_PNG_BYTES, dataUrl: PIXEL_PNG, sizePx: 128 });
  }
  return icons;
}

/**
 * How wide and tall the author's own drawing is, in inches, before any export
 * decides anything. Trimming far-placed nodes out of the fit exists to make the
 * sheet *smaller*, so neither exporter can legitimately produce a page much
 * larger than this: when one does, outlier handling has grown the drawing it
 * was supposed to shrink, and the user gets a plotter sheet with a stamp of
 * architecture in the middle of it.
 */
function drawingSpanIn(scenario: Scenario): { w: number; h: number } {
  const byId = new Map(scenario.nodes.map((n) => [n.id, n]));
  const absolute = (node: Node): { x: number; y: number } => {
    let x = node.position.x;
    let y = node.position.y;
    const seen = new Set<string>([node.id]);
    let parent = node.parentNode ? byId.get(node.parentNode) : undefined;
    while (parent && !seen.has(parent.id)) {
      seen.add(parent.id);
      x += parent.position.x;
      y += parent.position.y;
      parent = parent.parentNode ? byId.get(parent.parentNode) : undefined;
    }
    return { x, y };
  };
  const rects = scenario.nodes.map((node) => {
    const at = absolute(node);
    return {
      ...at,
      w: node.width ?? (node.style?.width as number | undefined) ?? 150,
      h: node.height ?? (node.style?.height as number | undefined) ?? 75,
    };
  });
  if (rects.length === 0) return { w: 0, h: 0 };
  const minX = Math.min(...rects.map((r) => r.x));
  const maxX = Math.max(...rects.map((r) => r.x + r.w));
  const minY = Math.min(...rects.map((r) => r.y));
  const maxY = Math.max(...rects.map((r) => r.y + r.h));
  return { w: (maxX - minX) / PX_PER_IN, h: (maxY - minY) / PX_PER_IN };
}


/**
 * The deck the export button actually produces.
 *
 * `buildDiagramSlidePptx` is a diagram-only deck that may grow its page for a
 * large architecture; `buildArchitectureDeckPptx` carries title, workflow,
 * services, review and cost slides that are all designed for a standard 16:9
 * page, so it cannot. Every rule in this file was measured against the first
 * one, and the second — the one `App.tsx` calls — was drawing the whole
 * architecture squeezed onto one fixed slide: 0.05in tiles and 4pt type for a
 * drawing the audited deck showed at 0.44in. Same failure class as the icons:
 * the gate did not exercise the configuration that ships.
 *
 * Only the properties that can differ between the two are checked here, so
 * this stays cheap: the page must never grow, the type must clear the same
 * floor, and every service must reach exactly one window.
 */
async function auditCustomerDeck(scenario: Scenario): Promise<string[]> {
  // Built the way `App.tsx:3483` builds it. Passing an empty list audited a
  // configuration the product never ships, which left the Services slide — the
  // customer deck's only path for spelling out a name the drawing shortened —
  // permanently empty and so permanently unexercised by every rule below.
  const groupLabels = new Map<string, string>();
  for (const node of scenario.nodes) {
    if (node.type === 'groupNode') {
      groupLabels.set(node.id, String((node.data as { label?: string } | undefined)?.label ?? ''));
    }
  }
  const deckServices = scenario.nodes
    .filter((n) => n.type !== 'groupNode')
    .map((n) => {
      const parentId = (n as { parentNode?: string; parentId?: string }).parentNode
        ?? (n as { parentNode?: string; parentId?: string }).parentId;
      const data = n.data as { label?: string; iconPath?: string } | undefined;
      const category = data?.iconPath?.match(/\/Icons\/([^/]+)\//i)?.[1]
        .replace(/[-_]+/g, ' ')
        .replace(/\b\w/g, (c) => c.toUpperCase());
      return {
        name: data?.label || 'Unnamed service',
        category,
        group: (parentId ? groupLabels.get(parentId) : undefined) || undefined,
      };
    });
  // The deck's workflow, validation, findings and cost slides are all gated on
  // an option this audit never passed, so four of its eight slide types were
  // never built here at all — and a row-overlap defect on the workflow slide
  // shipped because of it. These are synthesised the way `App.tsx` does, from
  // the same scenario, so the gate exercises the whole deck.
  const nodeLabels = new Map<string, string>(scenario.nodes.map((n) => [
    n.id,
    String((n.data as { label?: string } | undefined)?.label ?? n.id),
  ]));
  const labelOf = (id: string): string => nodeLabels.get(id) ?? id;
  const stepText = (edge: { source: string; target: string; label?: unknown; data?: unknown }): string => {
    // An authored description wins, because that is what the app passes: the
    // 800-character sentences `workflow-long-prose` carries are the real shape
    // of Architecture-Center prose, and synthesising over them would test the
    // synthesiser instead of the slide.
    const authored = String((edge.data as { stepDescription?: unknown } | undefined)?.stepDescription ?? '').trim();
    if (authored) return authored;
    const from = labelOf(edge.source);
    const to = labelOf(edge.target);
    const detail = String(edge.label ?? '').trim();
    // Otherwise full prose, because that is what a generated dataflow produces —
    // the model writes a sentence per hop, not a caption. A synthesised step
    // short enough to fit whatever pitch the slide happened to pick would test
    // nothing about how the slide picks it.
    return `${from} ${detail ? `sends ${detail} to` : 'calls'} ${to} over the platform's `
      + 'private network. The call carries an idempotency key, so a retry after a partial '
      + 'failure is safe, and the response is returned to the caller only once the '
      + 'downstream write has been acknowledged by every replica in the region.';
  };
  const deckWorkflow = scenario.edges.slice(0, 60).map((edge, index) => ({
    step: index + 1,
    description: stepText(edge as unknown as { source: string; target: string; label?: unknown; data?: unknown }),
    services: [labelOf(edge.source), labelOf(edge.target)].filter(Boolean),
  }));
  // The corpus was blind to its own worst case: the synthesised assessment was
  // English boilerplate, so no scenario ever exercised a Japanese finding —
  // and `localization.ts` instructs the model, verbatim, to write findings and
  // recommendations in Japanese. CJK is roughly twice the width per character,
  // so a box that holds an English finding holds half a Japanese one.
  const hasCjk = (s: unknown) => /[\u3040-\u30ff\u4e00-\u9fff]/.test(String(s ?? ''));
  const isJa = scenario.nodes.some((n) => hasCjk((n.data as { label?: string } | undefined)?.label))
    || scenario.edges.some((e) => hasCjk(e.label)
      || hasCjk((e.data as { stepDescription?: string } | undefined)?.stepDescription));
  const issueTextFor = (name: string, ja: boolean): string => (ja
    ? `${name}は単一の可用性ゾーンにのみ配置されているため、ゾーン障害が発生した`
      + '場合には当該階層の全インスタンスが同時に停止し、復旧までの間サービス全体が'
      + '利用できなくなる構成となっています。またフェールオーバー手順が文書化されて'
      + 'おらず、障害発生時の復旧作業が特定の担当者の記憶に依存する状態です。'
    : `${name} is deployed to a single availability zone, so a zone outage `
      + 'takes the whole tier offline for the duration of the incident.');
  const deckValidation = {
    overallScore: 72,
    overallLabel: isJa ? '一部に改善余地あり' : 'Adequate, with gaps',
    summary: isJa
      ? 'この構成は信頼性とセキュリティの基本要件をおおむね満たしていますが、コスト最適化と'
        + '運用性の観点では改善の余地が残されています。特に単一ゾーン配置の資源が複数あり、'
        + 'ゾーン障害時には該当階層が全面的に停止する構成となっているため、次回のリリースまでに'
        + 'ゾーン冗長構成への移行を検討することを推奨します。また、監視とアラートの設定が'
        + '一部の資源に限定されており、障害の検知が運用チームへの通報に依存している箇所が'
        + '見受けられますので、併せて是正されることを推奨いたします。加えて、バックアップの'
        + '保持期間が既定値のままとなっている資源が存在し、監査要件で求められる期間を'
        + '下回る可能性がありますので、要件を確認のうえ保持ポリシーを明示的に設定して'
        + 'ください。ネットワーク境界については、パブリックエンドポイントが有効なままの'
        + 'データストアが確認されましたので、プライベートエンドポイントへの移行と'
        + 'パブリックアクセスの無効化を、移行計画に含めることを強く推奨いたします。'
        + '最後に、容量計画の根拠となる負荷試験の記録が確認できませんでしたので、'
        + '本番相当の負荷での試験結果を残されることをご検討ください。あわせて、'
        + '構成変更の履歴が追跡できる形で保存されていない資源が複数確認されましたので、'
        + 'コードとしてのインフラストラクチャによる管理へ段階的に移行し、'
        + '変更内容の審査と承認の記録が残る運用体制を整備されることを推奨いたします。'
        + 'また、identity の観点では、共有シークレットによる認証が残存しているため、'
        + 'マネージド ID への移行を優先度の高い課題として計画に含めてください。'
      : 'The estate meets the reliability and security baselines, with cost and '
        + 'operational-excellence gaps that are worth closing before the next release.',
    pillars: (isJa
      ? ['信頼性', 'セキュリティ', 'コスト最適化', 'オペレーショナルエクセレンス', 'パフォーマンス効率']
      : ['Reliability', 'Security', 'Cost Optimization', 'Operational Excellence', 'Performance Efficiency']
    ).map((pillar, index) => ({
      pillar,
      score: [78, 81, 58, 66, 74][index],
      maturity: isJa ? '一部に改善余地あり' : 'Adequate, with gaps',
    })),
    findings: deckServices.slice(0, 14).map((service, index) => ({
      severity: (['critical', 'high', 'medium', 'low'] as const)[index % 4],
      category: isJa ? '信頼性' : (service.category || 'Platform'),
      // One finding longer than a whole page, because a model handed a complex
      // service writes one. Shrinking cannot answer it and the packer places at
      // least one block per page, so without a physical cut it would be laid
      // out anyway and painted off the bottom of the slide.
      issue: index === 3 && isJa
        ? Array.from({ length: 80 }, () => `${service.name}の構成には可用性とセキュリティの`
          + '両面で是正が必要な点が確認されており、放置した場合には障害時の復旧が'
          + '長期化するおそれがあります。').join('')
        : issueTextFor(service.name, isJa),
      recommendation: isJa
        ? '主要リージョン内の少なくとも二つの可用性ゾーンにインスタンスを分散したうえで、'
          + 'ゾーン冗長構成に対応した SKU へ移行し、あわせてゾーン単位の正常性監視を'
          + '有効化してください。移行に際しては、事前に検証環境で切り替え手順を確認し、'
          + '復旧目標時間と復旧目標地点を関係者間で合意したうえで、手順書として'
          + '文書化されることを推奨いたします。'
        // Numbered remediation steps, one per line, which is how a model asked
        // for steps writes them and what `recommendationSteps()` already
        // expects. The hard breaks are real paragraphs in the emitted file and
        // used to be measured as though they were not there at all.
        : '1. Enable zone redundancy on the tier.\n'
          + '2. Add a second node pool in another zone.\n'
          + '3. Update the runbook and re-run the failover drill.\n'
          + '4. Record the achieved RTO against the agreed target.',
    })),
    modelUsed: 'audit',
  };
  const deckCost = {
    totalMonthly: 18432.55,
    annual: 221190.6,
    currency: 'USD',
    term: 'monthly',
    region: 'Japan East',
    pricesAsOf: '2026-08-01',
    fixedCost: 12100.2,
    usageCost: 6332.35,
    byCategory: [
      { category: 'Compute', cost: 9210.1, percentage: 50 },
      { category: 'Databases', cost: 4608.14, percentage: 25 },
      { category: 'Networking', cost: 2764.88, percentage: 15 },
      { category: 'Storage', cost: 1849.43, percentage: 10 },
    ],
    topServices: deckServices.slice(0, 10).map((service, index) => ({
      serviceName: service.name,
      cost: 1800 - index * 120,
      tier: 'Standard',
      percentage: 10 - index * 0.5,
    })),
    regions: [
      { name: 'Japan East', flag: '🇯🇵', monthly: 18432.55, annual: 221190.6, isCurrent: true },
      { name: 'Southeast Asia', flag: '🇸🇬', monthly: 17120.4, annual: 205444.8, isCheapest: true },
      { name: 'West Europe', flag: '🇳🇱', monthly: 19004.11, annual: 228049.32 },
      { name: 'East US 2', flag: '🇺🇸', monthly: 17988.02, annual: 215856.24 },
    ],
    // The two partial-comparison banners were the last character caps in the
    // exporter and no scenario had ever rendered them, in either script. A
    // Japanese deck gets the incomplete path (the caps were ~2x undersized in
    // CJK, where a name is a full em wide); a Latin deck keeps the
    // cheapest-region path, so both branches stay covered.
    ...(isJa
      ? {
        regionComparisonIncomplete: true,
        // Only the part before the colon is printed, so the names carry the
        // length. Long enough that a banner which does not shrink runs off its
        // own box and onto the table beneath it — which is what the 150-char
        // cap did, and what nothing in the corpus had ever rendered.
        unavailableRegions: [
          '東日本 (Japan East) セカンダリ: SKU 未提供',
          '西日本 (Japan West) プライマリ: SKU 未提供',
          '東南アジア (Southeast Asia): ゾーン冗長なし',
          '西ヨーロッパ (West Europe): 価格取得不可',
          '北ヨーロッパ (North Europe): 階層未提供',
          '米国東部 2 (East US 2): マネージド未提供',
          '米国西部 3 (West US 3): 予約価格なし',
          'オーストラリア東部 (Australia East): 未提供',
          '韓国中部 (Korea Central): ゾーン冗長なし',
          'インド中部 (Central India): SKU 未提供',
          'ブラジル南部 (Brazil South): 価格取得不可',
          'カナダ中部 (Canada Central): 階層未提供',
        ],
      }
      : {}),
  };
  // The cover's "Brief:" box had never been drawn — `prompt` appeared nowhere
  // in this file — and it was the last unmeasured character cap in the
  // exporter. It is the user's typed brief, straight out of a `<textarea>`, so
  // hard breaks are the normal case: twenty bulleted lines is 213 characters,
  // half the old 420-character cap, and drew 4.875in in a 1.700in box.
  const brief = [
    'Front Door + WAF', 'Container Apps', 'Azure SQL Managed Instance', 'Blob storage',
    'Key Vault', 'Microsoft Entra ID', 'Log Analytics', 'Application Insights',
    'Private Link', 'NAT gateway', 'Azure Bastion', 'Azure Firewall', 'Azure Cache for Redis',
    'Service Bus', 'Azure Functions', 'Event Grid', 'API Management', 'Front-end CDN',
    'Zone redundancy across three availability zones', 'Paired-region disaster recovery',
  ].join('\n');
  const pptx = await buildArchitectureDeckPptx(PIXEL_PNG, {
    diagramName: scenario.title ?? 'Contoso Platform',
    author: scenario.author ?? 'Audit',
    date: '2026-08-10',
    prompt: isJa
      ? ['フロント ドアと WAF を前段に配置します。', 'コンテナー アプリでアプリケーションを実行します。',
        'データベースは SQL Managed Instance を使用します。', 'ストレージは BLOB を使用します。',
        'シークレットは Key Vault に格納します。', '認証は Microsoft Entra ID で行います。',
        'ログは Log Analytics に集約します。', '監視は Application Insights を使用します。',
        'ネットワークは Private Link で閉域化します。', '送信は NAT ゲートウェイを経由します。',
        '運用アクセスは Azure Bastion に限定します。', '境界は Azure Firewall で保護します。',
        'キャッシュは Azure Cache for Redis を使用します。', '非同期処理は Service Bus を使用します。',
        'イベント処理は Azure Functions で実装します。', 'イベント配信は Event Grid を使用します。',
        'API は API Management で公開します。', '静的配信は CDN を使用します。',
        '可用性ゾーンは 3 つに分散します。', 'ディザスター リカバリーはペア リージョンに構成します。'].join('\n')
      : brief,
    model: 'gpt-5 (2026-06)',
    isDarkMode: scenario.dark === true,
    diagram: { nodes: scenario.nodes, edges: scenario.edges },
    presetIcons: synthesisedIcons(scenario),
    services: deckServices,
    workflow: deckWorkflow,
    validation: deckValidation,
    cost: deckCost,
  });
  const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const sldSz = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentation);
  const pageW = sldSz ? +sldSz[1] / EMU_PER_INCH : BASE_SLIDE_W_IN;
  const pageH = sldSz ? +sldSz[2] / EMU_PER_INCH : BASE_SLIDE_H_IN;

  const issues: string[] = [];
  issues.push(...xmlWellFormednessIssues(await zipXmlParts(zip), 'customer deck: '));
  if (pageW > BASE_SLIDE_W_IN + 0.01 || pageH > BASE_SLIDE_H_IN + 0.01) {
    issues.push(`customer deck: page is ${pageW.toFixed(2)}x${pageH.toFixed(2)}in — every other slide in this deck is laid out for ${BASE_SLIDE_W_IN}x${BASE_SLIDE_H_IN}in`);
  }

  const slides = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')))
      .map((name) => zip.file(name)!.async('string')),
  );
  // A tiled deck opens with the whole drawing shown small on purpose, so the
  // floors below are about the windows that follow it.
  const drawn = slides.filter((xml) => xml.includes('name="service-'));
  const windows = drawn.filter((xml) => !xml.includes('(Overview)'));
  if (drawn.length === 0) {
    issues.push('customer deck: the diagram slide carries no native shapes at all');
    return issues;
  }

  const shapes = windows.flatMap((xml) => parseShapes(xml));
  const labels = shapes.filter((s) => s.name.startsWith('service-label-'));
  const minFont = labels.length > 0 ? Math.min(...labels.map((l) => l.fontSize ?? 99)) : 99;
  if (minFont < 7) {
    issues.push(`customer deck: smallest label font is ${minFont}pt (below the 7pt legibility floor)`);
  }
  issues.push(...connectorLabelFontIssues(shapes, 'customer deck: '));

  // Same contract as the diagram-only deck: every hop the caller asked for has
  // to be drawn somewhere. This deck tiles on its own plan against a page that
  // cannot grow, so a route can be dropped here while the audited deck carries
  // it — and a workflow slide that describes a hop the reader cannot find is
  // exactly the defect the numbered-callout convention exists to prevent.
  const drawnArrows = new Set<string>();
  for (const xml of drawn) {
    for (const shape of parseShapes(xml)) {
      if (!shape.name.startsWith('connector-')) continue;
      if (shape.name.startsWith('connector-label-') || shape.name.startsWith('connector-step-')) continue;
      drawnArrows.add(shape.name.slice('connector-'.length));
    }
  }
  for (const edge of scenario.edges) {
    // Against the name the exporter writes, not the one the diagram authored.
    // Ids reach the package through the same sanitiser as prose, so an id
    // carrying a forbidden code point is drawn under its stripped name and a
    // raw-id lookup reports it missing from a deck that in fact contains it.
    const id = auditStrip(String(edge.id));
    if (!drawnArrows.has(id)) issues.push(`customer deck: edge "${id}" is in the diagram but drawn on no slide`);
  }

  for (const node of scenario.nodes) {
    if ((node.type ?? '') === 'groupNode') continue;
    const marker = `name="service-${auditStrip(String(node.id))}"`;
    const on = windows.filter((xml) => xml.includes(marker)).length;
    if (on === 0) issues.push(`customer deck: service "${node.id}" is drawn on no slide`);
    else if (on > 1) issues.push(`customer deck: service "${node.id}" is drawn on ${on} slides`);
  }

  // A `slice(0, N)` with nothing said about the remainder is the same silent
  // loss the cost and services tables were both fixed for: the reader is shown
  // five findings, the app selected six, and the sixth appears on no slide and
  // in no footnote. Either the deck carries a finding whole or it says how many
  // it left out — anything else reads as a complete assessment and is not one.
  {
    const said = slides
      .flatMap((xml) => [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unescapeXml(m[1])))
      .join('\n');
    const missing = deckValidation.findings.filter((f) => !said.includes(auditStrip(f.issue)));
    const declares = new RegExp(`\\b${deckValidation.findings.length}\\s+finding`).test(said);
    if (missing.length > 0 && !declares) {
      issues.push(
        `customer deck: ${missing.length} of ${deckValidation.findings.length} WAF finding(s) appear `
        + 'nowhere in the deck, and no heading says how many were shown',
      );
    }
  }

  // Every name the drawing shortened must be spelled out somewhere in the deck,
  // and in this deck the Services table is the only place that can do it.
  //
  // Read out of `<a:t>` rather than out of the shape scrape, because table text
  // lives in an `<a:tbl>` inside a `<p:graphicFrame>` and a `<p:sp>`/`<p:pic>`
  // scan cannot see it at all — which is how a table that was working got
  // reported as empty. The check is on the full authored string: a table that
  // stops at row twenty announces sixty components and discharges twenty.
  const deckText = new Set<string>();
  for (const xml of slides) {
    for (const match of xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) deckText.add(collapseWs(unescapeXml(match[1])));
  }
  const stranded: string[] = [];
  for (const node of scenario.nodes) {
    if ((node.type ?? '') === 'groupNode') continue;
    // Trimmed on both sides of the comparison. A label carrying a forbidden
    // code point comes back from `auditStrip` with the gap it left, and the
    // deck writes the same name without it; that is the sanitiser working, not
    // a name gone missing.
    //
    // Whitespace is normalised on both sides for the same reason: the deck
    // collapses a name onto one line by design, so a name authored with a hard
    // break is discharged in full when it reads the same. Normalising does not
    // weaken the rule — the deck side is still one entry per `<a:t>`, so a name
    // genuinely split across paragraphs still has no entry that holds it whole.
    const name = collapseWs(auditStrip(String((node.data as { label?: string } | undefined)?.label ?? '')));
    if (name && !deckText.has(name)) stranded.push(name);
  }
  if (stranded.length > 0) {
    issues.push(
      `customer deck: ${stranded.length} service name(s) appear nowhere in full — e.g. `
      + `${stranded.slice(0, 3).map((n) => `"${n}"`).join(', ')}`,
    );
  }

  // A row that is present in the file but below the bottom of the slide is a
  // name lost, and the rule above cannot tell: it proves the string is in the
  // package, not that it is on the page. The table declares no autofit, so
  // PowerPoint treats `<a:tr h>` as a minimum and grows any row whose text
  // wraps — one two-line name in eighteen rows is enough to push the last of
  // them off the sheet. Measure the wrap against each column's usable width
  // and add it up.
  const tableRects = new Map<number, { x: number; y: number; w: number; h: number }[]>();
  for (const [index, xml] of slides.entries()) {
    for (const frame of xml.matchAll(/<p:graphicFrame>([\s\S]*?)<\/p:graphicFrame>/g)) {
      const body = frame[1];
      if (!body.includes('<a:tbl>')) continue;
      const offY = /<a:off[^>]*\by="(-?\d+)"/.exec(body);
      const offX = /<a:off[^>]*\bx="(-?\d+)"/.exec(body);
      const top = offY ? +offY[1] / EMU_PER_INCH : BASE_SLIDE_H_IN;
      const cols = [...body.matchAll(/<a:gridCol[^>]*\bw="(\d+)"/g)].map((m) => +m[1] / EMU_PER_INCH);
      if (cols.length === 0) continue;
      const marginIn = 0.2;
      let total = 0;
      for (const row of body.matchAll(/<a:tr[^>]*>([\s\S]*?)<\/a:tr>/g)) {
        const declared = /<a:tr[^>]*\bh="(\d+)"/.exec(row[0]);
        const minH = declared ? +declared[1] / EMU_PER_INCH : 0;
        let lines = 1;
        let cell = 0;
        // The vertical cell insets PowerPoint charges to the row on top of the
        // text, read from the row's own `tcPr` rather than assumed. Budgeting
        // only the text is how a table can be measured onto the page and still
        // print below it.
        let insetV = 0;
        for (const tc of row[1].matchAll(/<a:tc[\s\S]*?<\/a:tc>/g)) {
          // Per paragraph, like the shape scrape: a cell whose name carries a
          // hard break is several lines tall and used to measure as one.
          const text = [...tc[0].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
            .map((p) => [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => unescapeXml(m[1])).join(''))
            .join('\n');
          const pt = +(/\bsz="(\d+)"/.exec(tc[0])?.[1] ?? 1200) / 100;
          const usable = Math.max(0.5, (cols[cell] ?? cols[0]) - marginIn);
          if (text) lines = Math.max(lines, auditWrappedLines(text, usable, pt));
          const marT = +(/<a:tcPr[^>]*\bmarT="(\d+)"/.exec(tc[0])?.[1] ?? 45720);
          const marB = +(/<a:tcPr[^>]*\bmarB="(\d+)"/.exec(tc[0])?.[1] ?? 45720);
          insetV = Math.max(insetV, (marT + marB) / EMU_PER_INCH);
          cell += 1;
        }
        const pt = +(/\bsz="(\d+)"/.exec(row[1])?.[1] ?? 1200) / 100;
        total += Math.max(minH, lines * pt * 1.35 / 72 + insetV);
      }
      if (top + total > BASE_SLIDE_H_IN + 0.01) {
        issues.push(
          `customer deck: slide ${index + 1}'s table wraps to ${total.toFixed(2)}in from y=${top.toFixed(2)}in `
          + `— ${(top + total - BASE_SLIDE_H_IN).toFixed(2)}in of it is below the bottom of the slide`,
        );
      }
      // Hand the table's footprint to the overlap rule below. A table lives in
      // a `<p:graphicFrame>` and the shape scrape reads `<p:sp>`/`<p:pic>`, so
      // the two paths could not see each other at all: a paragraph painting an
      // inch into a table was unreportable by construction.
      const rects = tableRects.get(index) ?? [];
      rects.push({
        x: offX ? +offX[1] / EMU_PER_INCH : 0,
        y: top,
        w: cols.reduce((sum, c) => sum + c, 0),
        h: total,
      });
      tableRects.set(index, rects);
    }
  }

  // The same defect one slide over, in a different shape. A prose slide lays
  // its rows out on a pitch derived from a *count* — `floor(body / minimum)` —
  // with nothing measuring the words that go in them. The boxes are
  // `wrap="square"` with no autofit, so a description longer than its row
  // renders at full size and spills out of both ends of the box, printing over
  // the row above and the row below. Nothing in the file is out of place; only
  // the ink is. So measure the ink: wrap each text box's contents at its own
  // point size and check that the band it paints stays inside the slide and
  // clear of the next box's.
  for (const [index, xml] of slides.entries()) {
    const painted = parseShapes(xml)
      // The drawing's tiles are laid out by their own planner and measured by
      // the tile rules below, which know about the icon they share the box
      // with. Everything *else* on a drawing slide — the zone captions, the
      // legend, the workflow band, the step callouts — had no ink measurement
      // at all, because this loop used to skip any slide that contained a tile.
      // One whole-slide `continue` exempted the entire apparatus from every
      // drawing in the corpus, which is how a 0.4x0.24in zone caption came to
      // paint 22 wrapped lines and 3.3in of ink down the page unnoticed.
      .filter((s) => s.text.trim().length > 0
        // Connector chips are exempt, and that exemption is a known hole, not a
        // judgement: measured, 22 of them paint outside their own rounded
        // rectangles, by up to 0.29in of a 0.385in box. The cause is the same
        // break-anywhere ratio as the tile name below — `connectorLabelBox`
        // sizes the chip with `Math.ceil(ink / perLine)` and a crowded chip
        // wraps to one line more than that. It is left here because correcting
        // it is not a one-line change: an honest count makes the chip taller,
        // 36 chips then cannot be seated between the tiles, and walking the
        // retry ladder down to the legibility floor to buy the room detaches 5
        // numbered callouts from their own arrows. It needs the placement
        // search, and it needs its own round.
        && !s.name.startsWith('connector-')
        // The tile *name* is measured here like everything else. It was the
        // last text on a drawing slide with no ink rule over it, and it was
        // wrong: `Math.ceil(ink / column)` planned three lines for a name that
        // wraps to five, and the two surplus lines were drawn straight through
        // the "P1v3 · eastus" sub-line under it. The sub-line's own box is
        // `wrap="none"` — one line however wide — so it is measured by the
        // chip rules instead, and the tile rectangle carries no text at all.
        && !(s.name.startsWith('service-') && !s.name.startsWith('service-label-')))
      .map((s) => {
        const pt = s.fontSize ?? (s.runs[0]?.sizePt ?? 12);
        const lines = auditWrappedLines(s.text.trim(), textColumnIn(s), pt);
        const need = lines * linePitchIn(s, pt);
        const over = Math.max(0, need - s.h);
        // `anchor="ctr"` grows the block from the middle, so an overflow spills
        // equally out of the top and the bottom; a top-anchored box spills down.
        const top = s.anchor === 'ctr' ? s.y - over / 2 : s.y;
        // The horizontal extent of the ink, not of the box. A box is as wide
        // as the row it was sized for; the ink inside it may be a third of
        // that and sitting at one end. Mixing the two — ink vertically, box
        // horizontally — is what let a "72 / 100" score readout report 85
        // collisions between two boxes whose glyphs are 0.2in apart.
        const drawn = drawnTextRect(s);
        const left = drawn ? drawn.x : s.x;
        const right = drawn ? drawn.x + drawn.w : s.x + s.w;
        // Where the glyphs are vertically, which is not where the box is. A
        // 58pt "72" centred in a 1.25in box paints about 0.8in in the middle
        // of it; `max(h, need)` claims the whole 1.25in, which is the right
        // answer for "did this text spill" and the wrong one for "did these
        // two collide". Reported 85 collisions between a score and the "/ 100"
        // beneath it whose glyphs are a fifth of an inch apart.
        const inkTop = drawn ? drawn.y : top;
        const inkBottom = drawn ? drawn.y + drawn.h : top + Math.max(s.h, need);
        return { shape: s, top, bottom: top + Math.max(s.h, need), over, left, right, inkTop, inkBottom };
      });
    // Two invariants that do not re-wrap anything.
    //
    // The whole line-counting apparatus — the two exporters' counters and this
    // file's copy — is one algorithm written three times, so a blind spot in
    // the algorithm is a blind spot in all three at once. That is not a theory:
    // it is exactly how a hard line break stayed invisible to every measurement
    // in the codebase simultaneously. These two are different in kind. The
    // first is *evidence* read straight out of the emitted file — the renderer
    // has already committed to one paragraph per `<a:p>`, and no wrap can put
    // two of them on one line, so the paragraph count is a hard lower bound on
    // the lines drawn. The second is physics: however the words are broken up,
    // the total ink has to fit inside the total line length. Both catch only
    // the dangerous direction, which is under-counting.
    for (const s of parseShapes(xml)) {
      if (!s.text.trim() || s.name.startsWith('connector-')) continue;
      const pt = s.fontSize ?? (s.runs[0]?.sizePt ?? 12);
      const box = textColumnIn(s);
      const counted = auditWrappedLines(s.text.trim(), box, pt);
      const drawn = s.paragraphs.filter((p) => p.trim().length > 0).length;
      if (drawn > counted) {
        issues.push(
          `customer deck: slide ${index + 1} counts ${counted} line(s) for "${s.text.trim().slice(0, 24)}" `
          + `but the file already contains ${drawn} non-empty paragraphs`,
        );
      }
      const ink = s.runs.reduce((sum, r) => sum + auditTextWidthIn(r.text, r.sizePt || pt), 0);
      if (ink > counted * box + 0.01) {
        issues.push(
          `customer deck: slide ${index + 1} counts ${counted} line(s) for "${s.text.trim().slice(0, 24)}" `
          + `— ${ink.toFixed(2)}in of type cannot fit in ${(counted * box).toFixed(2)}in of line`,
        );
      }
    }
    // Off the sheet is off the sheet whether the box overflowed or the row
    // pitch simply walked past the bottom — both lose the words either way.
    for (const p of painted) {
      if (p.bottom > BASE_SLIDE_H_IN + 0.01 || p.top < -0.01) {
        issues.push(
          `customer deck: slide ${index + 1} paints "${p.shape.text.trim().slice(0, 32)}" from `
          + `y=${p.top.toFixed(2)}in to ${p.bottom.toFixed(2)}in — outside the ${BASE_SLIDE_H_IN}in slide`,
        );
      }
    }
    // On a drawing slide the two rules above can both be blind at once. A zone
    // caption that overruns its band paints down the middle of the page, so it
    // never leaves the sheet, and everything it lands on is a service tile,
    // which is measured by the tile rules and not by this loop — so a 0.40in
    // band holding 3.30in of ink over 22 wrapped lines was invisible to every
    // check in the file. The band is a box the exporter chose for itself, and
    // ink outside a box the exporter chose is ink nothing accounted for, so
    // measure that directly rather than waiting for it to hit something.
    if (xml.includes('name="service-')) {
      for (const p of painted) {
        if (p.over <= 0.02) continue;
        const pt = p.shape.fontSize ?? (p.shape.runs[0]?.sizePt ?? 12);
        const drawnLines = auditWrappedLines(p.shape.text.trim(), textColumnIn(p.shape), pt);
        issues.push(
          `customer deck: slide ${index + 1} draws "${p.shape.text.trim().slice(0, 32)}" as `
          + `${drawnLines} line(s) at ${pt}pt in a ${p.shape.w.toFixed(3)}x${p.shape.h.toFixed(3)}in band `
          + `— ${p.over.toFixed(3)}in of it is outside`,
        );
      }
    }
    // Overlap needs only *one* side to have spilled. Two boxes may legitimately
    // be stacked, so a collision between two well-behaved boxes is a layout —
    // but ink that has left its own box and landed on a neighbour is never
    // legitimate, and the neighbour has no reason to be overflowing too. The
    // stricter pairing missed a Japanese assessment paragraph painting 0.30in
    // over a section heading that fitted its own box perfectly.
    //
    // That reasoning was half right and shipped a hole. "A collision between
    // two well-behaved boxes is a layout" is true of boxes placed side by side,
    // which the shared-column test already excludes — it is not true of two
    // boxes that were simply put on top of each other, and that is the more
    // common failure because it needs no arithmetic to go wrong, only a
    // position. Gating on `over > 0.01` made it unreportable by construction:
    // a zone caption fitted its band exactly, three chips were drawn over it,
    // and the pair was never examined.
    const reported = new Set<string>();
    // Tables are opaque blocks of ink drawn from their own frame. Any box whose
    // ink lands inside one is a collision whether or not either overflowed —
    // the table is not going to move out of the way.
    for (const rect of tableRects.get(index) ?? []) {
      for (const p of painted) {
        if (p.shape.y >= rect.y - 0.01) continue; // the table's own heading sits above it
        const sharesX = p.shape.x < rect.x + rect.w - 0.02 && rect.x < p.shape.x + p.shape.w - 0.02;
        if (!sharesX) continue;
        const into = p.bottom - rect.y;
        if (into > 0.02) {
          issues.push(
            `customer deck: slide ${index + 1} paints "${p.shape.text.trim().slice(0, 32)}" `
            + `${into.toFixed(2)}in into a table`,
          );
        }
      }
    }
    // Pictures are ink too, and `painted` can never hold one — it is filtered
    // to shapes carrying text, so an icon overlapping a table was blind by
    // construction on the one slide that has both. A picture has no wrapping to
    // measure, so its own rectangle is the whole story.
    for (const rect of tableRects.get(index) ?? []) {
      for (const pic of parseShapes(xml)) {
        // Named as a picture, not merely un-texted. The earlier reading let any
        // untexted shape through — a background band, a divider, a rule — and
        // compared it against a table as if it were an icon, which is a false
        // positive waiting for the first decoration drawn behind a table.
        if (!pic.name.startsWith('icon-') && !pic.name.startsWith('picture-')) continue;
        const ox = Math.min(pic.x + pic.w, rect.x + rect.w) - Math.max(pic.x, rect.x);
        const oy = Math.min(pic.y + pic.h, rect.y + rect.h) - Math.max(pic.y, rect.y);
        if (ox > 0.02 && oy > 0.02) {
          issues.push(
            `customer deck: slide ${index + 1} draws picture "${pic.name}" ${oy.toFixed(2)}in into a table`,
          );
        }
      }
    }
    for (const a of painted) {
      for (const b of painted) {
        if (a === b) continue;
        // Only rows that share a column can collide; two captions side by side
        // in the same band are a layout, not an overlap.
        const sharesX = a.left < b.right - 0.02 && b.left < a.right - 0.02;
        if (!sharesX) continue;
        // Ink against ink. A box that overflows still paints past its own
        // bottom, so the spill is folded into the ink extent rather than
        // dropped: `over` grows the block, `max(h, need)` does not.
        const aBottom = Math.max(a.inkBottom, a.over > 0.01 ? a.bottom : a.inkBottom);
        const bBottom = Math.max(b.inkBottom, b.over > 0.01 ? b.bottom : b.inkBottom);
        const overlap = Math.min(aBottom, bBottom) - Math.max(a.inkTop, b.inkTop);
        if (overlap <= 0.02) continue;
        // Both orderings reach this when two boxes each spill onto the other;
        // report the pair once so a count of issues stays a count of defects.
        const key = [painted.indexOf(a), painted.indexOf(b)].sort((x, y) => x - y).join(':');
        if (reported.has(key)) continue;
        reported.add(key);
        issues.push(
          `customer deck: slide ${index + 1} overlaps "${a.shape.text.trim().slice(0, 24)}" (${a.shape.name}) and `
          + `"${b.shape.text.trim().slice(0, 24)}" (${b.shape.name}) by ${overlap.toFixed(2)}in of type`,
        );
      }
    }
    // The footer band is chrome: it is drawn on every slide, always last, and
    // nothing in the body may reach it. A body box that is sized correctly for
    // its own contents but positioned over the footer spills nothing, so the
    // pair rule above cannot see it — and yet the footer prints on top and the
    // last line of the body is gone. This is the arm that catches a block too
    // tall for a page rather than a box too small for its text.
    //
    // It lives here, and not where it was found: an earlier edit landed the
    // whole block inside `xmlWellFormednessIssues`, whose scope has no
    // `painted` and no `index`. It could therefore never run — and worse, the
    // one branch that would have reached it was the branch that reports a
    // character XML cannot encode, so the first genuinely unopenable package
    // would have crashed the gate with a ReferenceError instead of naming the
    // defect.
    const footer = painted.find((p) => p.shape.text.includes('Generated by Microsoft Product'));
    if (footer) {
      for (const p of painted) {
        if (p === footer) continue;
        if (p.shape.y > footer.shape.y - 0.01) continue; // chrome and notes that live there
        const into = p.bottom - footer.top;
        if (into > 0.02) {
          issues.push(
            `customer deck: slide ${index + 1} paints "${p.shape.text.trim().slice(0, 32)}" `
            + `${into.toFixed(2)}in into the footer band`,
          );
        }
      }
    }
  }
  // Every measurement in this file assumes CJK draws at one em. That holds for
  // Yu Gothic UI and not for the Calibri the theme's minor font falls back to,
  // whose `ea` is empty — so a run that reached the file without an explicit
  // East Asian typeface would be measured 60% too wide and every fit above it
  // would be quietly wrong in the unsafe direction. pptxgenjs writes `<a:ea>`
  // alongside every `<a:latin>` today, but that is its behaviour, not a
  // guarantee, so pin it rather than depend on it.
  {
    const latin = slides.reduce((n, xml) => n + (xml.match(/<a:latin\b/g) ?? []).length, 0);
    const ea = slides.reduce((n, xml) => n + (xml.match(/<a:ea\b/g) ?? []).length, 0);
    if (ea < latin) {
      issues.push(
        `customer deck: ${latin - ea} run(s) declare a Latin typeface with no East Asian `
        + 'typeface, so CJK falls through to the theme font and is measured wrong',
      );
    }
  }

  return issues;
}

/**
 * What the audit expects a sanitised string to look like.
 *
 * Independent of the exporter's own strip for the same reason as the regex
 * above: the rules that compare authored text against emitted text have to
 * agree with the *specification*, not with whatever the shipped code currently
 * does, or a broken strip would move the goalposts to meet itself.
 */
function auditStrip(value: string): string {
  return value.replace(
    /[\uD800-\uDBFF][\uDC00-\uDFFF]|[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDFFF]/g,
    (m) => (m.length === 2 ? m : ' '),
  // COMPOSED, because the exporters compose. Every use of this function asks
  // "what is the exporter right to emit for this authored string?", and the
  // answer now includes a canonical spelling. Left decomposed, a name or a
  // sentence authored on a Mac was compared against a composed drawn string
  // and reported as never having reached a slide. This is ICU answering a
  // question about the bytes, not a copy of the exporter's normaliser.
  ).normalize('NFC');
}
async function zipXmlParts(zip: JSZip): Promise<Array<{ path: string; text: string }>> {
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir && /\.(xml|rels)$/i.test(n));
  return Promise.all(names.map(async (path) => ({ path, text: await zip.files[path].async('string') })));
}

async function auditPptx(scenario: Scenario): Promise<Report> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: scenario.title ?? 'Contoso Platform',
    author: scenario.author ?? 'Audit',
    date: '2026-08-10',
    isDarkMode: scenario.dark === true,
    diagram: { nodes: scenario.nodes, edges: scenario.edges },
    presetIcons: synthesisedIcons(scenario),
  });
  const buffer = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer;
  writeFileSync(path.join(OUT, `${scenario.id}.pptx`), buffer);
  const zip = await JSZip.loadAsync(buffer);
  const presentation = await zip.file('ppt/presentation.xml')!.async('string');
  const sldSz = /<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(presentation);
  const pageW = sldSz ? +sldSz[1] / EMU_PER_INCH : 13.333;
  const pageH = sldSz ? +sldSz[2] / EMU_PER_INCH : 7.5;
  const allSlides = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')))
      .map((name) => zip.file(name)!.async('string')),
  );
  const slideCount = allSlides.length;
  // The icons have to actually reach the deck. `rasterizeIcons` needs a DOM and
  // returns an empty map under Node, so for most of this audit's life every
  // tile was measured with no icon in it — a different tile interior from the
  // one the user receives. Assert the pictures are there, so the harness can
  // never silently go blind that way again.
  const drawnPics = allSlides.reduce((sum, slideXml) => sum + (slideXml.match(/<p:pic>/g) ?? []).length, 0);
  // A tiled deck opens with the whole drawing shown small on purpose, so the
  // legibility and one-slide-per-service rules below are about the slides that
  // follow it. Measuring the overview against them would report every tiled
  // deck as broken, and dropping the rules to accommodate it would stop them
  // measuring anything.
  const overviewAt = allSlides.findIndex((slideXml) => slideXml.includes('(Overview)'));
  const xml = overviewAt < 0 ? allSlides : allSlides.slice(overviewAt + 1);
  const perSlide = xml.map((slideXml) => parseShapes(slideXml));
  const shapes = perSlide.flat();
  // The overview is excluded from the rules above for a real reason, but the
  // exclusion is total: it removes slide 1 from `perSlide` AND from `shapes`,
  // so not one of the ~700 lines of per-slide rules below has ever looked at
  // it. That is where a whole class hid — 1,523 connector chips emitted
  // narrower than a single glyph of their own type, on the first slide of
  // eleven decks, invisible to every rule in this file.
  //
  // So the overview gets its own reduced rule set instead of none. It may not
  // be measured for legibility of *names* — it is deliberately small — but a
  // box that cannot hold one character of the type inside it is not small, it
  // is broken: PowerPoint does not clip, so it stacks one letter per line and
  // spills them sideways out of the shape.
  const overviewShapes: Shape[] = overviewAt < 0 ? [] : parseShapes(allSlides[overviewAt]);

  const issues: string[] = [];
  issues.push(...xmlWellFormednessIssues(await zipXmlParts(zip), ''));
  // The header band, on the path most users actually take.
  //
  // Every containment rule in this file lived inside the customer-deck audit,
  // so the main export path's own title had never been measured - and it is
  // the one string on the slide that is free text the user typed, with no cap
  // anywhere between the name box and the XML. At three wrapped lines a 24pt
  // header centred in a 0.730in band paints above the top edge of the slide.
  // Restricted to the band because everything below it is a tile, a chip or a
  // caption with rules of its own; this is about the furniture at the top.
  for (const [index, shapesOnSlide] of perSlide.entries()) {
    for (const shape of shapesOnSlide) {
      const text = shape.text.trim();
      if (text.length === 0 || shape.y > 1.0 || shape.h <= 0) continue;
      if (shape.name.startsWith('service-') || shape.name.startsWith('connector-')) continue;
      const pt = shape.fontSize ?? (shape.runs[0]?.sizePt ?? 12);
      const lines = auditWrappedLines(text, textColumnIn(shape), pt);
      const blockIn = (lines * pt * 1.35) / 72;
      if (blockIn > shape.h + 0.02) {
        issues.push(
          `slide ${index + 1} header "${text.slice(0, 32)}" draws ${lines} line(s) at ${pt}pt `
          + `— ${blockIn.toFixed(3)}in of type in a ${shape.h.toFixed(3)}in band, so it grows out of it`,
        );
      }
    }
  }
  // Every character the deck DRAWS must have a measured advance.
  //
  // The oracle this file provides is a per-glyph table measured from the font,
  // and its weakness is not that the exporter might use a different table -
  // it is that both tables can be INCOMPLETE in the same places. Both fell
  // back to a flat average for anything outside printable ASCII, so the
  // ellipsis the exporter appends at every truncation point was charged 0.54
  // em against a real 0.733, and an arrow in a connector label 0.54 against a
  // real 1.0. Neither side could see it, because a shared blind spot is not a
  // disagreement.
  //
  // This rule is sound whether or not both sides are correct: it asserts
  // COVERAGE rather than agreement, and it fails on the commit that first
  // draws a character nobody has measured.
  //
  // ASKED ABOUT THE CLUSTER, not the code point. A promoted heart is U+2764
  // plus VS16 and is charged the emoji advance exactly, but no table entry for
  // U+2764 can ever read 1.373, so per code point this reported a correctly
  // priced glyph as a guess in every future run - an issue nobody could clear.
  const unmeasured = new Map<string, number>();
  for (const slideXml of allSlides) {
    for (const match of slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
      const run = unescapeXml(match[1]);
      for (const cluster of auditClusters(run)) {
        if (cluster.measured) continue;
        unmeasured.set(cluster.text, (unmeasured.get(cluster.text) ?? 0) + 1);
      }
    }
  }
  if (unmeasured.size > 0) {
    const worst = [...unmeasured.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6);
    issues.push(
      `the deck draws ${unmeasured.size} character(s) with no measured advance, so every width `
      + `and wrap that touches them is a guess: `
      + worst.map(([character, n]) => `U+${(character.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')} x${n}`).join(', '),
    );
  }
  // AND THE TWO MEASUREMENTS MUST AGREE, not merely both exist.
  //
  // Coverage is one property and correctness is another. Every string the deck
  // draws is now priced by both models - the exporter's own walk, and this
  // file's ICU-segmented one - and a divergence is reported wherever they
  // disagree by more than rounding. This is the rule that could have caught the
  // joiner and variation-selector defects on the commit that introduced them:
  // a word joiner made a 57-character name measure as one glyph, and a
  // variation selector charged the letter "z" 204% of its width.
  //
  // Only meaningful because the two are no longer the same algorithm. While
  // this file transcribed the exporter's walk, a divergence rule could not have
  // failed on any input at all.
  const diverged = new Map<string, string>();
  for (const slideXml of allSlides) {
    for (const match of slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
      const run = unescapeXml(match[1]);
      if (!run.trim()) continue;
      const mine = auditClusters(run).reduce((sum, c) => sum + c.em, 0);
      const theirs = advanceWidthIn(run, 72);
      if (Math.abs(mine - theirs) > Math.max(0.01, mine * 0.02)) {
        diverged.set(
          run.slice(0, 32),
          `${theirs.toFixed(4)} against ${mine.toFixed(4)} em`,
        );
      }
    }
  }
  if (diverged.size > 0) {
    issues.push(
      `${diverged.size} drawn string(s) are priced differently by the exporter and the gate, `
      + `so one of the two models is wrong about text the reader will see: `
      + [...diverged.entries()].slice(0, 3)
        .map(([run, how]) => `${JSON.stringify(run)} ${how}`).join('; '),
    );
  }
  // AND NO MARK MAY BE LEFT STANDING ON ITS OWN.
  //
  // Pricing is one property; cutting is another, and a name is cut by code
  // point wherever it is too long for its tile. A combining mark carries no
  // glyph of its own - it is drawn on whatever precedes it - so a cut between a
  // letter and its accent does not lose an accent, it MOVES it: the decomposed
  // spelling of "Passerelle securisee ... partagees" came out of a 60x30 tile
  // as an acute stacked on the ellipsis. Nothing in the deck could see that,
  // because the string still held every code point the model had priced.
  //
  // Asked of what is DRAWN rather than of the truncation routine, so it holds
  // for every path that shortens a name and for the ones not written yet.
  const strandedMarks = new Set<string>();
  for (const slideXml of allSlides) {
    for (const match of slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
      const run = unescapeXml(match[1]);
      const points = [...run];
      for (const [i, point] of points.entries()) {
        if (!AUDIT_COMBINING_RE.test(point)) continue;
        const before = i === 0 ? '' : points[i - 1];
        if (before === '' || before === '\u2026' || /\s/.test(before) || AUDIT_COMBINING_RE.test(before)) {
          if (before !== '' && AUDIT_COMBINING_RE.test(before)) continue;
          strandedMarks.add(run.slice(0, 32));
        }
      }
    }
  }
  if (strandedMarks.size > 0) {
    issues.push(
      `${strandedMarks.size} drawn string(s) put a combining mark on nothing, so it lands on the ellipsis `
      + `or on the space before it: ${[...strandedMarks].slice(0, 3).map((run) => JSON.stringify(run)).join('; ')}`,
    );
  }
  // AND THE SPELLING MUST NOT BE OBSERVABLE AT ALL.
  //
  // The two rules above ask whether the width model prices a mark correctly.
  // That question was answered twice and was wrong twice, because the
  // disagreement kept retreating a level: first the walk, then the pricing
  // rule, then the TABLE both models read - and the font gives a precomposed
  // glyph and its base different advances for 532 characters, so "a mark costs
  // nothing" still prices "Şişli şube şebeke sunucusu" 2.8% apart depending on
  // where it was typed. 480 of those 532 are cheaper decomposed, which is the
  // direction that paints outside the tile.
  //
  // This rule stops asking. Every drawn string must already be in composed
  // form, so no measurement, cut or comparison downstream can tell which
  // spelling was authored. It re-derives nothing and shares no table with the
  // exporter: it asks ICU one question about the bytes on the slide.
  const decomposed = new Set<string>();
  for (const slideXml of allSlides) {
    for (const match of slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) {
      const run = unescapeXml(match[1]);
      if (!run.trim() || run.normalize('NFC') === run) continue;
      decomposed.add(run.slice(0, 32));
    }
  }
  if (decomposed.size > 0) {
    issues.push(
      `${decomposed.size} drawn string(s) are not in composed form, so the same visible name is `
      + `measured and cut differently depending on where it was typed: `
      + `${[...decomposed].slice(0, 3).map((run) => JSON.stringify(run)).join('; ')}`,
    );
  }
  // The audit ran icon-blind for its whole life: `canRasterize()` is false
  // under Node, so `rasterizeIcons` returned an empty map and every rule that
  // measures a tile — the caption band's position and height, and therefore
  // every chip, callout and contrast composite derived from it — was tuned
  // against a deck nobody is ever sent. This is the tripwire for that.
  //
  // It cannot simply demand pictures: the exporter deliberately drops an icon
  // the tile is too small to render legibly and keeps the words instead, which
  // is right and is what `meta-tight` (a 5x5 grid whose tiles also carry an
  // SKU, a region and a price) does on every tile. So the rule fires only for a
  // deck with a tile roomy enough that no such trade was needed.
  const roomyTile = perSlide.flat().some(
    (s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta')
      && s.h >= 0.55 && s.w >= 0.9,
  );
  if (synthesisedIcons(scenario).size > 0 && drawnPics === 0 && roomyTile) {
    issues.push(`deck embeds no icon pictures for ${scenario.nodes.length} nodes`);
  }
  const native = auditNativeConversion(allSlides, scenario.edges);
  issues.push(...native.issues);
  // A box whose text column cannot hold one glyph of its own type. PowerPoint
  // clips nothing: it stacks the word one letter per line and paints the tail
  // outside the shape on every side, so a row of these becomes a smear of type
  // rather than a set of labels. Scale is not the excuse — the box shrinks with
  // the drawing but the font does not, so past a certain scale the trade the
  // exporter thinks it is making is not available.
  //
  // Applied to the overview, which no other rule in this file reaches.
  for (const shape of overviewShapes) {
    const text = shape.text.trim();
    const pt = shape.fontSize;
    if (text === '' || !pt) continue;
    const widest = measuredWidestGlyphIn(text, pt);
    const column = Math.max(0, shape.w - shape.insetX);
    // Same 0.01in tolerance the exporter's own guard uses, so the rule and the
    // thing it measures cannot disagree about the borderline case.
    //
    // TWO of the widest glyph, matching the exporter's bar. Asking whether one
    // letter fits cannot see a chip that spells its sentence one character per
    // line, and 97 of them did exactly that while clearing the one-glyph test
    // by 0.0022in.
    const needs = text.length > 1 ? widest * 2 : widest;
    if (needs > column + 0.01) {
      const many = text.length > 1;
      issues.push(
        `overview draws "${text}" at ${pt}pt in a ${shape.w.toFixed(3)}in box — `
        + `its widest letter needs ${widest.toFixed(3)}in`
        + (many ? `, so a line of them needs ${needs.toFixed(3)}in,` : '')
        + ` and the column is ${column.toFixed(3)}in`,
      );
    }
  }
  // A chip whose text takes more lines than its box is tall. Every other rule
  // that could see this exempts `connector-` (`:3595`, `:3644`), and the
  // overview slice hid the rest, so 317 chips across 18 decks overflowed in
  // silence — one of them painting 47% of its box below itself, over the tiles
  // and arrows it runs between.
  //
  // The column is read from the emitted `<a:bodyPr lIns/rIns>`, not from the
  // exporter's model of it: the two disagreed by 0.12in for the whole life of
  // this file, and a rule that reads the model cannot see that.
  for (const [at, slideShapes] of [...perSlide.entries(), [-1, overviewShapes] as [number, Shape[]]]) {
    for (const chip of slideShapes) {
      if (!chip.name.startsWith('connector-label-')) continue;
      const text = chip.text.trim();
      const pt = chip.fontSize;
      if (text === '' || !pt) continue;
      const column = Math.max(MIN_TEXT_COLUMN_IN, chip.w - chip.insetX);
      const rows = measuredWrappedLines(text, column, pt);
      const need = rows * ((pt * 1.3) / 72);
      if (need > chip.h + 0.02) {
        const where = at < 0 ? 'the overview' : `slide ${at + 2}`;
        issues.push(
          `${where} draws "${text}" at ${pt}pt in a ${chip.w.toFixed(3)}x${chip.h.toFixed(3)}in chip — `
          + `${rows} line(s) in a ${column.toFixed(3)}in column need ${need.toFixed(3)}in, `
          + `so ${(need - chip.h).toFixed(3)}in is painted below it`,
        );
      }
    }
  }
  // The PPTX equivalent of the Visio sheet-size invariant below is unreachable
  // and deliberately absent: the page is clamped to PowerPoint's 56in limit,
  // and the fit only trims outliers once the drawing is already past ~52in, so
  // "page larger than the drawing plus chrome" cannot happen. Rigid-translation
  // parking is caught here by the font floor and the page-count rule instead.
  issues.push(...await auditCustomerDeck(scenario));
  // A chip or a numbered callout with no arrow anywhere in the deck is worse
  // than a missing label: the reader sees a sentence and a ① floating on blank
  // paper and goes looking for a hop that was never drawn. This caught a route
  // dropped from EVERY window at once by a per-window "don't draw a flattened
  // hop" rule that assumed some other window would carry it.
  const drawnArrows = new Set<string>();
  const annotatedArrows = new Set<string>();
  for (const slideXml of allSlides) {
    for (const shape of parseShapes(slideXml)) {
      if (shape.name.startsWith('connector-label-')) annotatedArrows.add(shape.name.slice('connector-label-'.length));
      else if (shape.name.startsWith('connector-step-')) annotatedArrows.add(shape.name.slice('connector-step-'.length));
      else if (shape.name.startsWith('connector-')) drawnArrows.add(shape.name.slice('connector-'.length));
    }
  }
  for (const id of annotatedArrows) {
    if (!drawnArrows.has(id)) issues.push(`arrow "connector-${id}" is annotated but drawn on no slide`);
  }
  // Annotation is not enough on its own. A hop dropped before its chip is even
  // placed loses the annotation too, so nothing is left to be orphaned and the
  // rule above stays silent while the step list still describes the hop. The
  // real contract is the scenario's own edge list: every edge the caller asked
  // for has to appear somewhere in the deck.
  for (const edge of scenario.edges) {
    const id = auditStrip(String(edge.id));
    if (!drawnArrows.has(id)) issues.push(`edge "${id}" is in the diagram but drawn on no slide`);
  }
  // A tile asked to show a SKU, a region and a price and showing none of them
  // is silent content loss: unlike a muted chip, whose wording is handed to the
  // step list, a dropped sub-line has no carrier anywhere in the deck. The
  // numbers the reader came for are simply absent. Exempt when the tiles are
  // too small to carry a second character row at all.
  const wantsMeta = scenario.nodes.filter((node) => {
    const data = node.data as Record<string, unknown> | undefined;
    return !!data && (data.sku !== undefined || data.region !== undefined);
  }).length;
  if (wantsMeta > 0 && roomyTile) {
    const drawnMeta = allSlides.reduce(
      (sum, slideXml) => sum + parseShapes(slideXml).filter((s) => s.name.startsWith('service-meta-')).length,
      0,
    );
    if (drawnMeta === 0) issues.push(`deck drops the SKU/region sub-line on all ${wantsMeta} tiles that declare one`);
  }
  for (const slideXml of allSlides) {
    const bg = /<p:bg>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(slideXml)?.[1]?.toLowerCase() ?? 'ffffff';
    issues.push(...contrastIssues(parseShapes(slideXml), bg));
  }
  // A zone caption that was cut while the zone had wider clear paper to put it
  // on. The band search scores candidates by how much of a tile each would
  // cover, and two completely clear bands both score zero — so the score cannot
  // separate them and something else decides. That something used to be the
  // order they happen to sit in the array, with a short-circuit on the first
  // zero, which handed the caption to whichever clear band was built first
  // rather than to the one that holds the name.
  //
  // The loss is invisible to every other rule here: the cut name reaches the
  // index slide, the band is inside its zone, the type is above the floor, and
  // nothing overflows. Only the comparison the search declined to make shows
  // it, so make it: for a caption that is cut, find the widest strip of the
  // zone that is free of tiles and tall enough for the band, and require the
  // band to be within a fifth of it.
  for (const [index, slideXml] of allSlides.entries()) {
    const shapes = parseShapes(slideXml);
    const tiles = shapes.filter((s) => /^service-[^-]/.test(s.name));
    if (tiles.length === 0) continue;
    for (const caption of shapes.filter((s) => s.name.startsWith('zone-label-'))) {
      if (!caption.text.includes('…')) continue;
      const zone = shapes.find((s) => s.name === `zone-${caption.name.slice('zone-label-'.length)}`);
      if (!zone) continue;
      // Every horizontal strip of the zone that could seat this band: above the
      // tiles, below them, and in any gap between two rows of them. A strip is
      // usable only if it is at least as tall as the band already drawn.
      //
      // And within a strip, the widest *run* clear of tiles, not merely the
      // strips that are clear from edge to edge. A zone whose foot row carries
      // three tiles at one end has no full-width clear row at all, yet nine
      // clear inches beside those tiles — which is exactly the band the search
      // is supposed to find and exactly the one the earlier version of this
      // rule could not see. It scored that case as "no clear paper" and passed
      // a caption cut to 44 of its 76 characters.
      const inZone = tiles.filter((t) => t.x < zone.x + zone.w && zone.x < t.x + t.w
        && t.y < zone.y + zone.h && zone.y < t.y + t.h);
      const edges = [zone.y, zone.y + zone.h, ...inZone.flatMap((t) => [t.y, t.y + t.h])]
        .filter((y) => y >= zone.y - 1e-6 && y <= zone.y + zone.h + 1e-6)
        .sort((a, b) => a - b);
      // The widest band the exporter is able to build, which is the zone less
      // the inset it always keeps. Measuring against the raw zone width accuses
      // it of a choice it was never offered — and on a zone under ~0.6in wide
      // that inset alone is more than the fifth of slack this rule allows.
      const reachable = Math.max(0.4, zone.w - 0.12);
      let widest = 0;
      for (let i = 0; i + 1 < edges.length; i += 1) {
        const top = edges[i];
        const bot = edges[i + 1];
        if (bot - top < caption.h - 1e-6) continue;
        const across = inZone
          .filter((t) => t.y < bot - 1e-6 && top < t.y + t.h - 1e-6)
          .sort((a, b) => a.x - b.x);
        let cursor = zone.x;
        for (const tile of across) {
          widest = Math.max(widest, Math.min(reachable, tile.x - cursor));
          cursor = Math.max(cursor, tile.x + tile.w);
        }
        widest = Math.max(widest, Math.min(reachable, zone.x + zone.w - cursor));
      }
      if (widest > caption.w * 1.25 + 0.01) {
        issues.push(
          `customer deck: slide ${index + 1} cuts zone caption "${caption.text.trim().slice(0, 28)}" `
          + `into a ${caption.w.toFixed(2)}in band when ${widest.toFixed(2)}in of the zone is clear`,
        );
      }
    }
  }
  // The overview is exempt from the legibility floor because it is a map, not
  // a reading surface — but "smaller than the floor" is not the same as "ink
  // the reader cannot resolve at all". Type this small is grey mush that makes
  // the thumbnail harder to read, not easier, so it must not be drawn: show
  // the shapes and let the slides that follow carry the names.
  const OVERVIEW_FLOOR_PT = 6;
  let overviewMinFont = 0;
  let overviewEmptyTiles = 0;
  if (overviewAt >= 0) {
    const overviewShapes = parseShapes(allSlides[overviewAt]);
    const sized = overviewShapes.filter((s) => s.text.trim() !== '' && s.fontSize !== null);
    overviewMinFont = sized.length ? Math.min(...sized.map((s) => s.fontSize ?? 99)) : 0;
    const illegible = sized.filter((s) => (s.fontSize ?? 99) < OVERVIEW_FLOOR_PT);
    if (illegible.length) {
      issues.push(
        `overview draws ${illegible.length} text run(s) at ${Math.min(...illegible.map((s) => s.fontSize ?? 99))}pt, under the ${OVERVIEW_FLOOR_PT}pt the reader can resolve: e.g. "${illegible[0].text}"`,
      );
    }
    // The rule above counts type, so it is satisfied by drawing none — an empty
    // grey box scores perfectly. This is the rule that cannot be satisfied by
    // deleting content: whatever else it does, a tile must say something.
    const named = new Set(
      overviewShapes.filter((s) => s.name.startsWith('service-label-')).map((s) => s.name.slice('service-label-'.length)),
    );
    const iconed = new Set(
      [...allSlides[overviewAt].matchAll(/name="icon-([^"]+)"/g)].map((m) => m[1]),
    );
    const blank = overviewShapes
      .filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'))
      .filter((s) => {
        const id = s.name.slice('service-'.length);
        return !named.has(id) && !iconed.has(id) && s.text.trim() === '';
      });
    overviewEmptyTiles = blank.length;
    if (blank.length) {
      issues.push(
        `overview draws ${blank.length} tile(s) with neither a name nor an icon, e.g. "${blank[0].name}" — an empty box says less than small type`,
      );
    }
  }
  const tiles = shapes.filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'));
  // A shape that draws nothing is not a name. The exporter emits a
  // `service-label-` box whenever the tile is `named || stub`; when the stub
  // column cannot hold two glyphs that box now carries an empty string, and
  // counting it as a name let four blank slivers back into `namedWidths` and
  // drag the median tile width from 0.503in down to 0.200in — which reported
  // eight perfectly ordinary chips as "2.8x wider than the smallest node tile".
  const labels = shapes.filter((s) => s.name.startsWith('service-label-') && s.text.trim() !== '');
  const chips = shapes.filter((s) => s.name.startsWith('connector-label-'));

  const minTileW = Math.min(...tiles.map((t) => t.w));
  const namedTiles = new Set(labels.map((l) => l.name.replace('service-label-', '')));
  const namedBoxes = tiles.filter((t) => namedTiles.has(t.name.replace('service-', '')));
  const namedWidths = namedBoxes.map((t) => t.w).sort((a, b) => a - b);
  // The MEDIAN named tile, not the narrowest. "Wider than the smallest tile"
  // is a statement about the drawing's unit of scale, and one authored sliver
  // is not that unit: a deck with a legitimate 0.42in tile beside ordinary
  // 1.7in ones reported every perfectly proportioned chip as a violation while
  // saying nothing about whether any chip dominates the drawing. The median is
  // the typical tile, which is what the rule has always meant.
  const namedMinTileW = namedWidths.length > 0
    ? namedWidths[Math.floor((namedWidths.length - 1) / 2)]
    : minTileW;
  // Weighed by AREA, because that is what "dominates the drawing" means and
  // width alone is only a proxy for it. The proxy holds while tiles are
  // roughly as wide as they are tall and breaks the moment they are not: a
  // deck of deliberately tall, narrow tiles reported a 0.560 x 0.131in chip
  // as outweighing a 0.377 x 1.510in service — 0.073 in^2 against 0.570 in^2,
  // one eighth of it — purely because the chip is 0.18in wider. Same failure
  // as the Visio type-to-tile ratio: measure the thing, not the stand-in.
  const namedAreas = namedBoxes.map((t) => t.w * t.h).sort((a, b) => a - b);
  const namedMedianArea = namedAreas.length > 0
    ? namedAreas[Math.floor((namedAreas.length - 1) / 2)]
    : 0;
  const minFont = Math.min(...labels.map((l) => l.fontSize ?? 99));

  // A tile with room for a name and no name on it. The width guard in the
  // exporter used to be keyed on the font its HEIGHT implied, so a taller tile
  // demanded a wider column and a 0.78 x 3.13in shape — two and a half square
  // inches — drew no text at all. Nothing could see it: every label rule reads
  // the labels that WERE emitted, so a missing one is invisible by
  // construction. This reads the tiles instead and asks whether each one had
  // the column the exporter's own floor asks for.
  // Per SLIDE, and MEASURED rather than restated. A tile keeps the same
  // `service-<id>` shape name on every slide it appears on, so a set built
  // from the whole deck says "this tile is named" when only one slide names
  // it, and the deck-wide overview set judged reading-slide copies by the
  // overview's 6pt floor and the overview's height test — two errors that
  // happened to cancel, which is not a reason to keep either.
  //
  // The old bar, `column >= 4 * floorPt/72`, was the exporter's own predicate
  // copied verbatim: it could only ever agree, and it caught the round-52 bug
  // by luck because that bug changed which font the operand used. This asks
  // instead whether the WHOLE authored name, wrapped into the column the tile
  // has at the floor size, fits the height the tile has — a statement about
  // the text and the box, which the exporter cannot make true by changing its
  // mind about a constant.
  //
  // "Two of the widest glyph fit" alone is necessary but not sufficient, and
  // saying otherwise reported two hairline tiles whose names would have needed
  // thirteen lines of a two-line shape. A tile that cannot hold its name is
  // entitled to hand it to the index slide; a tile with room for all of it and
  // nothing drawn is the round-53 defect, where 2.44 square inches went blank.
  // "The whole name fits" has to mean fits BESIDE THE ICON. An overview tile
  // draws its icon at up to 78% of its height and hands the name to the
  // reading slide, which is the design; measuring the name against the tile's
  // full height called all 354 of them a defect. The icon's drawn height is on
  // the slide, so this reads it rather than modelling it.
  const iconHeights = new Map<string, number>();
  for (const slideXml of allSlides) {
    for (const pic of slideXml.matchAll(
      /<p:pic>[\s\S]*?name="icon-([^"]+)"[\s\S]*?<a:ext cx="\d+" cy="(\d+)"\/>[\s\S]*?<\/p:pic>/g,
    )) {
      iconHeights.set(`${allSlides.indexOf(slideXml)}:${pic[1]}`, +pic[2] / EMU_PER_INCH);
    }
  }
  const authoredById = new Map(
    scenario.nodes.map((n) => [auditStrip(String(n.id)), String(n.data?.label ?? '')]),
  );
  const overviewIndex = overviewAt;
  const slidesForTiles: { shapes: Shape[]; overview: boolean }[] = allSlides.map((xml, at) => ({
    shapes: parseShapes(xml),
    overview: at === overviewIndex,
  }));
  for (const [at, slide] of slidesForTiles.entries()) {
    const slideNamed = new Set(
      slide.shapes.filter((s) => s.name.startsWith('service-label-') && s.text.trim() !== '')
        .map((s) => s.name.slice('service-label-'.length)),
    );
    const slideTiles = slide.shapes.filter(
      (s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'),
    );
    for (const tile of slideTiles) {
      const id = tile.name.slice('service-'.length);
      if (slideNamed.has(id)) continue;
      const authored = authoredById.get(id);
      if (!authored) continue;
      const floorPt = slide.overview ? 6 : 7;
      const column = Math.max(0.05, tile.w - 0.06);
      const lines = measuredWrappedLines(authored, column, floorPt);
      const needed = (lines * floorPt * 1.35) / 72;
      const room = tile.h - 0.06 - (iconHeights.get(`${at}:${id}`) ?? 0);
      if (measuredDrawableInColumn(authored, floorPt, column) && needed <= room) {
        issues.push(
          `tile "${tile.name}" is ${tile.w.toFixed(3)}x${tile.h.toFixed(3)}in and draws no name — `
          + `"${authored}" wraps to ${lines} line(s) needing ${needed.toFixed(3)}in of the `
          + `${room.toFixed(3)}in its icon leaves at the ${floorPt}pt floor, so the whole name fits`,
        );
      }
    }
  }

  for (const label of labels) {
    const font = label.fontSize ?? 11;
    // Independent of the exporter's inequality, not a restatement of it.
    //
    // `charsPerLine = w / (pt/72) >= 4` was a TAUTOLOGY: the exporter
    // guarantees `nameColumn >= 4 * fontSize/72`, only ever lowers the drawn
    // font from there, and emits the box at exactly `innerW === nameColumn`.
    // Every named tile in the corpus reported exactly 4.0 or comfortably
    // above, and it could not land below whatever the exporter did — so it
    // was checking its own premise. Worse, an exemption then had to be carved
    // out for stub tiles, which made the rule blind to five genuinely
    // unreadable labels at once.
    //
    // Two of the widest glyph the string actually contains, from the measured
    // table, is a statement about the drawn text rather than about the
    // exporter's arithmetic — the same move `measuredWidestGlyphIn` made for
    // the chip guard. It passes on every named tile in the corpus and reports
    // the stub ladder.
    // Two questions, not one, and an outlier may not answer either.
    //
    // A box narrower than the widest glyph cannot set that glyph at all; a box
    // that holds fewer than two of the string's TYPICAL characters spells it
    // one letter per line. The old form asked only `w >= 2 * widest`, which
    // lets a single `m` at 0.861 em speak for a string whose mean is 0.55 -
    // and that made this rule report a name the exporter was right to draw,
    // in a column setting 2.8 characters a line.
    const widest = measuredWidestGlyphIn(label.text, font);
    const glyphs = auditClusters(label.text).filter((cluster) => !/^\s*$/.test(cluster.text));
    // The LOWER bound in both clauses. Reading the widest glyph one way and
    // the mean the other made every character of an untabled script argue
    // both sides of the same test, and the mean is the binding one. By
    // CLUSTER in both, too: counting a keycap as three glyphs divided one
    // glyph's advance across three of them.
    let meanEm = 0;
    for (const glyph of glyphs) meanEm += glyph.measured ? glyph.em : 0;
    const mean = glyphs.length > 0 ? (meanEm * font) / 72 / glyphs.length : 0;
    if (label.text && glyphs.length > 0
      && (label.w + 0.005 < widest || label.w + 0.005 < 2 * mean)) {
      const bar = Math.max(widest, 2 * mean);
      issues.push(
        `label "${label.text}" is drawn at ${font}pt in a ${label.w.toFixed(3)}in box — `
        + `its widest glyph is ${widest.toFixed(3)}in and its mean is ${mean.toFixed(3)}in, `
        + `so a readable line needs ${bar.toFixed(3)}in and only `
        + `${(label.w / Math.max(mean, 1e-9)).toFixed(2)} typical characters fit`,
      );
    }
    // 1.35, matching every other line-height in this file and in the exporter:
    // `Yu Gothic UI`'s own hhea and OS/2 win metrics give 1.3301. 1.25 was 6%
    // optimistic and the tile path was reserving at 1.22, 8% optimistic, and
    // both errors were invisible while the two sides shared a width model that
    // over-counted lines by more than the shortfall.
    const lineHeight = (font * 1.35) / 72;
    // Counted by WRAPPING the string, not by dividing its total ink by the
    // column. `ceil(width / w)` is the break-anywhere assumption: it says how
    // many lines the ink would need if words could be split at any character,
    // which is a lower bound and never the answer. The chip rule was moved off
    // it in round 52 and the tile-label rule was left behind, so a name fitted
    // at one size and painted at another - eight lines believed, seven real -
    // drew its last line 0.104in below the box it was measured in and 0.044in
    // past the bottom of its own tile, and nothing reported it.
    const realLines = measuredWrappedLines(label.text, label.w, font);
    if (realLines * lineHeight > label.h + 0.02) {
      issues.push(
        `label "${label.text}" wraps to ${realLines} line(s) at ${font}pt in its `
        + `${label.w.toFixed(3)}in column, needing ${(realLines * lineHeight).toFixed(3)}in `
        + `of a ${label.h.toFixed(3)}in box - ${((realLines * lineHeight) - label.h).toFixed(3)}in is painted below it`,
      );
    }
  }
  for (const chip of chips) {
    // Measured against the smallest tile that still carries a NAME. An unnamed
    // tile is not the drawing's unit of scale — it is a box the fitter has
    // already decided cannot hold type — and letting an authored 8px sliver
    // define "the size of a service" made every chip in the deck a violation
    // while saying nothing about whether the chip dominates the drawing.
    if (namedMinTileW > 0 && chip.w > namedMinTileW && chip.w * chip.h > namedMedianArea) {
      issues.push(
        `edge chip "${chip.text}" is ${(chip.w / namedMinTileW).toFixed(1)}x wider than the typical node tile `
        + `and covers ${(chip.w * chip.h).toFixed(3)}in^2 against its ${namedMedianArea.toFixed(3)}in^2`,
      );
    }
  }
  // Which two services each numbered arrow runs between, by id. The deck names
  // its shapes `connector-step-<routeId>` and `service-<nodeId>`, so a callout
  // can be measured against the tiles it actually calls out rather than
  // against a statistic over the slide.
  const endsOfRoute = new Map<string, [string, string]>();
  const stepOfRoute = new Map<string, number>();
  for (const edge of scenario.edges) {
    endsOfRoute.set(auditStrip(String(edge.id)), [
      auditStrip(String(edge.source)),
      auditStrip(String(edge.target)),
    ]);
    const step = (edge as unknown as { data?: { stepNumber?: number | string } }).data?.stepNumber;
    stepOfRoute.set(auditStrip(String(edge.id)), Number(step) || 1);
  }
  // The exporter's own derived floor, at the smallest type a callout may be set
  // in. Replicated rather than imported because this file measures the FILE and
  // must keep working if the exporter's arithmetic changes underneath it.
  const BADGE_SHARE = 0.55;
  // The planner's own bar for a tile worth chasing, replicated for the same
  // reason. Below it the exporter has already decided the tile is an authored
  // sliver and left it to the renderer's type floor.
  const MARKABLE_TILE_W_IN = 0.2;
  const badgeFloorIn = (stepNumber: number, fontPt: number): number => {
    const digits = String(Math.max(1, Math.abs(Math.trunc(stepNumber)))).length;
    return (fontPt / 72) * (Math.hypot(digits * 0.62, 1.3) / 0.9);
  };
  const conflicts: Array<{ name: string; tile: number; floor: number; ratio: number }> = [];
  // Authored widths, for the two rules that have to reason about the transform
  // the gate cannot see: what diameter the sizing model would have chosen, and
  // what tile width the window planner could ever have delivered.
  const authoredW = new Map<string, number>();
  for (const node of scenario.nodes) {
    const w = Number((node as unknown as { width?: number }).width);
    if (Number.isFinite(w) && w > 0) authoredW.set(auditStrip(String(node.id)), w);
  }
  // Whether the planner gave the callout bar up for this whole drawing.
  //
  // `legibleScaleFor` chases `markableTileWIn(widest step)` only while that bar
  // is below `finestPerIn`, and falls back to the plain markable bar when it is
  // not, because a target no grid can hit does not raise the floor, it switches
  // the floor off. So the clamp is one decision taken once for the drawing, and
  // when it engages every tile on every sheet is planned to the plain bar.
  //
  // ASKED, not replicated. The replicated form disagreed with the exporter on
  // two independent statistics and was wrong in both directions: it modelled a
  // 6.04in frame where a numbered drawing gets 5.77in - the connection legend
  // is 0.24 + 0.03in of it - and so failed a correctly clamped 21 slide deck 98
  // times at 15 authored px; and it took the NARROWEST node where the planner
  // takes the MEDIAN, so on a drawing of 20px tiles with a 14px sliver every
  // thirtieth node it declared a clamp the exporter never made, accused a
  // correct 53 slide plan of chasing a bar it had already reached, and
  // suppressed the four real conflicts on the same deck. Both bugs are
  // unreachable now: there is one copy of the arithmetic and it is the copy
  // that ran.
  const calloutPlan = calloutPlanFor(scenario);
  const calloutBarClamped = calloutPlan.clamped;
  // The tile the planner was serving, in authored pixels.
  //
  // `planWindowsAtCeiling` targets the MEDIAN service, deliberately and at
  // length: an extremum has a neighbour, so one sliver among eighty would
  // otherwise decide the deck, and `probe-whitespace` measures that trade as
  // four slivers dragging their 160px neighbours to 2.3in each and putting one
  // tile on a slide. A tile below the median is therefore one the planner
  // declined to serve, and the deck has no move for it that does not cost more
  // than it buys - chasing the 14px slivers on `probe-blind-sliver` is
  // reachable, so the clamp would not even stop it, and it would take a correct
  // 53 slide deck past 80 to serve three discs out of 119.
  //
  // Note the asymmetry with Visio, which is not an inconsistency but the two
  // formats' different currencies. A sheet pays for the same fix in INCHES, so
  // `magnifiedForCallouts` now chases the narrowest badged tile and serves
  // exactly these hops. A deck pays in slides, and cannot.
  // ASKED, not replicated, for the second time and the same reason. The gate's
  // own median spanned every entry in `scenario.nodes` with a positive width,
  // groups included, while the planner medians over `partitionBoxes(...)
  // .services`, which excludes `kind === 'group'`. A zone rectangle is wider
  // than a service, so the gate's median was always the higher of the two and
  // the divergence was blindness only: drawing two 700px zones around four
  // services moved the gate's median from 24px to 150px, put every service
  // below it, and turned a reported 56% conflict into a silent pass without
  // changing one drawn pixel. On a landing zone diagram, where subscription,
  // VNet and subnet frames outnumber the services inside them, the gate's
  // median was a zone width and the rule was off for the whole deck.
  //
  // And it is the SERVED width, not the median. The median was a proxy for
  // "the planner declined to serve this tile" with no cost behind it, and
  // `sorted[floor(n/2)]` puts up to half a drawing below its own median: on
  // four services of 150, 24, 150, 24 authored px it excused two of the four
  // and all three hops when the plan that serves them costs exactly one extra
  // window. The planner now tries the narrowest badged tile first and keeps
  // that plan whenever it clears the same density floor that already judges
  // the median raise, so this number is the outcome of a measurement rather
  // than an assumption - and where it still sits at the median, the deck
  // measured the finer plan and could not afford it.
  //
  // The served width is no longer consulted directly. A width is the wrong
  // question: read deck-wide it exempted every hop at once when it came back
  // as a sentinel, and read per hop it exempts exactly the glyph-sized service
  // that most needs the callout to be proportionate. What the rule asks for
  // instead is `chaseAffordable`, the RESULT of the chase the planner actually
  // ran, which cannot be true for a plan that was never attempted.
  // A label may lean on the two services its own arrow connects. The reader
  // still attributes it correctly — it is touching the very icons it is about
  // — and on a hop shorter than the label there is nowhere else for it to go.
  // Leaning on a THIRD service is a different thing entirely: it hides an
  // unrelated icon and reads as that service's caption. So the bar for a
  // stranger's tile stays at a couple of percent, and an endpoint of the
  // arrow itself is allowed a tenth of its area before it counts as hidden.
  //
  // Hoisted out of the per-slide loop, which recomputed all three from the
  // scenario on every sheet, so that the badge loop below can share them.
  const membersOfZone = new Map<string, number>();
  for (const node of scenario.nodes) {
    if (!node.parentNode) continue;
    membersOfZone.set(node.parentNode, (membersOfZone.get(node.parentNode) ?? 0) + 1);
  }
  const annotationEnds = new Map<string, Set<string>>();
  for (const edge of scenario.edges) {
    annotationEnds.set(edge.id, new Set([`service-${edge.source}`, `service-${edge.target}`]));
  }
  const tileBudget = (annotation: string, tile: Shape): number => {
    const routeId = annotation.replace(/^connector-(label|step)-/, '');
    return annotationEnds.get(routeId)?.has(tile.name) ? 0.1 : 0.02;
  };
  // Collisions are only real between shapes printed on the same sheet. Reading
  // every slide as one pile reported a chip on part 1 as covering a tile on
  // part 3, which turned every tiled deck into a wall of phantom issues and
  // hid whatever was genuinely wrong.
  for (const slideShapes of perSlide) {
    const slideTiles = slideShapes.filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'));
    const slideChips = slideShapes.filter((s) => s.name.startsWith('connector-label-'));
    for (const chip of slideChips) {
      for (const tile of slideTiles) {
        const area = overlapArea(chip, tile);
        if (area > tileBudget(chip.name, tile) * tile.w * tile.h) {
          issues.push(`edge chip "${chip.text}" overlaps node "${tile.name}" by ${((area / (tile.w * tile.h)) * 100).toFixed(0)}%`);
        }
      }
    }
    // Leaning on a tile is tolerable — the reader can still see which service
    // it is. Leaning on the tile's *name* is not, because the name is the only
    // thing that says which service it is, and a chip is drawn on top of it in
    // a near-solid fill. Measured against the words themselves, not the box
    // they are laid out in: with no icon that box is nearly the whole tile, so
    // scoring against it would just restate the tile rule at a tighter budget.
    // This also closes the gap in "a tile with neither a name nor an icon is an
    // issue" — a name present in the XML but painted over satisfies that rule
    // while telling the reader nothing.
    for (const chip of slideChips) {
      for (const caption of slideShapes.filter((s) => s.name.startsWith('service-label-'))) {
        const words = drawnTextRect(caption);
        if (!words) continue;
        const area = overlapArea(chip, words);
        const share = area / Math.max(words.w * words.h, 1e-6);
        if (share > 0.05) {
          issues.push(`edge chip "${chip.text}" covers ${(share * 100).toFixed(0)}% of the name "${caption.text}"`);
        }
      }
      // The SKU / region / price sub-line was drawn, modelled by no obstacle
      // and measured by no rule, so a chip could sit squarely on it and every
      // check passed. It is the second reason the tile is on the slide — an
      // architecture the reader cannot cost or place in a region is a different
      // document — but it is recoverable from the service itself in a way the
      // name is not, so it is budgeted a little more loosely than the name.
      for (const meta of slideShapes.filter((s) => s.name.startsWith('service-meta-'))) {
        const words = drawnTextRect(meta, true);
        if (!words) continue;
        const dx = Math.min(chip.x + chip.w, words.x + words.w) - Math.max(chip.x, words.x);
        const dy = Math.min(chip.y + chip.h, words.y + words.h) - Math.max(chip.y, words.y);
        if (dx <= 0 || dy <= 0) continue;        // Which way the chip bites matters, and an area share cannot tell the
        // two apart. A deep bite over a run of columns costs the reader those
        // characters outright — the region, or the price. A shallow one across
        // the whole line clips every character instead: still a defect, but a
        // different one, and it is invisible to a rule that only sums area
        // because a thin sliver of a thin line is a very small number.
        const columns = dx / words.w;
        const rows = dy / words.h;
        const tile = meta.name.slice('service-meta-'.length);
        if (rows > 0.5 && columns > 0.12) {
          issues.push(`edge chip "${chip.text}" covers ${(columns * 100).toFixed(0)}% of ${tile}'s sub-line "${meta.text}"`);
        } else if (rows > 0.25 && columns > 0.6) {
          issues.push(`edge chip "${chip.text}" clips ${(rows * 100).toFixed(0)}% off every character of ${tile}'s sub-line "${meta.text}"`);
        }
      }
    }
    // A `wrap="none"` line that outgrows its box does not wrap and does not
    // clip: PowerPoint draws it centred at full width, spilling out of both
    // sides of the tile over whatever the neighbours put there. Nothing else
    // measures it, because every other rule scores against the shape's box and
    // the box is the one thing this text ignores.
    for (const meta of slideShapes.filter((s) => s.name.startsWith('service-meta-'))) {
      const drawn = textWidthIn(meta.text.trim(), meta.fontSize ?? 0);
      if (meta.text.trim() === '' || !meta.fontSize) continue;
      if (drawn > meta.w + 0.01) {
        issues.push(`sub-line "${meta.text}" overflows its tile by ${((drawn - meta.w) * 100 / meta.w).toFixed(0)}% (${drawn.toFixed(2)}in of text in a ${meta.w.toFixed(2)}in box)`);
      }
    }

    // tiled deck the box is redrawn on every slide a member landed on, so a
    // zone of six services can appear five times, each time as a closed box
    // around one tile — the reader has no way to tell the fragment from the
    // whole. A fragment has to hold most of what it claims.
    for (const zone of slideShapes.filter((s) => s.name.startsWith('zone-') && !s.name.startsWith('zone-label-'))) {
      const zoneId = zone.name.replace(/^zone-/, '');
      const total = membersOfZone.get(zoneId);
      if (!total || total <= 2) continue;
      const held = slideTiles.filter((tile) => overlapArea(tile, zone) > 0.5 * tile.w * tile.h).length;
      // Saying so is enough: a title that reads "Data zone (3 / 28)" tells the
      // reader this is a slice, which is all the closed box failed to do.
      const title = slideShapes.find((s) => s.name === `zone-label-${zoneId}`);
      if (/\(\s*\d+\s*\/\s*\d+\s*\)/.test(title?.text ?? '')) continue;
      if (held > 0 && held < Math.ceil(total / 2)) {
        issues.push(`zone "${zone.name}" is drawn closed around ${held} of its ${total} services`);
      }
    }
    // A zone title that a member tile is standing on is a zone with no name.
    // The band belongs to the container, so nothing the container holds may be
    // drawn across it — but only what it holds. Architecture Center security
    // diagrams routinely draw a compliance boundary straight across a drawing,
    // overlapping tiles that belong to a different container, and blaming the
    // exporter for an overlap the author drew turns this rule into noise.
    for (const title of slideShapes.filter((s) => s.name.startsWith('zone-label-'))) {
      const zoneId = title.name.replace(/^zone-label-/, '');
      const members = new Set(
        scenario.nodes.filter((node) => node.parentNode === zoneId).map((node) => `service-${node.id}`),
      );
      let covered = 0;
      for (const tile of slideTiles) if (members.has(tile.name)) covered += overlapArea(title, tile);
      if (covered > 0.25 * title.w * title.h) {
        const pct = ((covered / (title.w * title.h)) * 100).toFixed(0);
        issues.push(`zone title "${title.text}" is ${pct}% covered by the tiles inside it`);
      }
    }
    // A zone's name written inside a *different* zone is the same false
    // containment claim as a service drawn outside its own boundary, in the
    // other direction — and it is exactly what scoring title placement against
    // service tiles alone produced: in a stacked drawing the clear band just
    // above a zone belongs to the zone above it, so "Data subnet" was printed
    // inside the Application subnet's box.
    for (const title of slideShapes.filter((s) => s.name.startsWith('zone-label-'))) {
      const zoneId = title.name.replace(/^zone-label-/, '');
      const own = slideShapes.find((s) => s.name === `zone-${zoneId}`);
      if (!own) continue;
      const area = Math.max(1e-6, title.w * title.h);
      const inside = overlapArea(title, own);
      // A fragment's drawn rectangle is not the zone, it is what survived the
      // window cut, so its name may legitimately sit just outside the cut.
      const fragment = /\(\s*\d+\s*\/\s*\d+\s*\)/.test(title.text ?? '');
      if (!fragment && inside < 0.9 * area) {
        issues.push(`zone title "${title.text}" is only ${((inside / area) * 100).toFixed(0)}% inside the "${zoneId}" boundary it names`);
      }
      for (const other of slideShapes) {
        if (!other.name.startsWith('zone-') || other.name.startsWith('zone-label-')) continue;
        if (other.name === `zone-${zoneId}`) continue;
        const trespass = overlapArea(title, other);
        // Only when the name is not in its own box as well — and only when the
        // two boxes actually overlap. An author who draws a compliance band
        // across half a virtual network has drawn two rectangles that overlap,
        // and every point inside one of them is inside the other: there is no
        // placement the exporter could choose that would satisfy a flat "never
        // inside another zone", so demanding it turns this rule into noise.
        //
        // Gating that exemption on the title's own containment instead of on
        // the zones' was too generous in the other direction. Two boxes that do
        // not overlap at all have a placement that satisfies the rule by
        // construction — inside one is outside the other — so a name 90% in its
        // own box and 10% in a disjoint neighbour is an avoidable false claim
        // that the old gate waved through.
        const authored = overlapArea(own, other) > 0;
        if (trespass > 0.01 * area && (!authored || inside < 0.9 * area)) {
          issues.push(`zone title "${title.text}" is written ${((trespass / area) * 100).toFixed(0)}% inside "${other.name}" and only ${((inside / area) * 100).toFixed(0)}% inside the "${zoneId}" it names`);
        }
      }
    }
  }
  // The badge rules run over one more sheet than the collision rules above:
  // the overview.
  //
  // Its exclusion from `perSlide` is structural - that array is
  // `allSlides.slice(overviewAt + 1)` - and it is right for every rule in the
  // loop above, which measure whether shapes collide. The overview is a dense
  // index by design; measuring chip-against-tile there reported four phantom
  // overlaps on `pipeline-region` the moment this loop was widened, which is
  // how the separation earned its own loop rather than a shared one.
  //
  // It is exactly wrong for badge proportion, because the overview is where
  // the tiles are SMALL: on a 120-service estate the discs drew at 56% of a
  // tile, 67% at 160 and 78% at 240, growing without bound as the tile shrinks
  // with N, while the reading windows never moved off 18-20%. The rule was
  // armed on the slides where nothing can go wrong and blind on the one slide
  // where it does.
  const badgeSlides = overviewShapes.length > 0 ? [...perSlide, overviewShapes] : perSlide;
  for (const slideShapes of badgeSlides) {
    const slideTiles = slideShapes.filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'));
    const slideChips = slideShapes.filter((s) => s.name.startsWith('connector-label-'));
    for (const badge of slideShapes.filter((s) => s.name.startsWith('connector-step-'))) {
      let buried = 0;
      for (const tile of slideTiles) {
        buried += overlapArea(badge, tile);
        // Two bars, both of which have to be crossed. The tile bar catches a
        // number printed over a service; the badge bar keeps a number that is
        // merely lapping a rim from being reported as one.
        //
        // A callout stands on the arrow it numbers. On a dense grid that arrow
        // runs through a row gutter narrower than the disc, so the disc has to
        // lap the tile above or below it — `chain24-en`'s wrap-around hop laps
        // its neighbour by half a disc, which is 3% of the tile. There is no
        // clear slot within reach in any direction, so failing that case only
        // rewards moving the number away from its own hop. A number genuinely
        // printed over an icon is 90-100% of the disc, so the badge bar leaves
        // that firmly caught while dropping the rim laps.
        const onTile = overlapArea(badge, tile);
        if (onTile > tileBudget(badge.name, tile) * tile.w * tile.h
          && onTile > 0.6 * badge.w * badge.h) {
          issues.push(`step badge "${badge.name}" covers node "${tile.name}" by ${((onTile/(tile.w*tile.h))*100).toFixed(0)}% (badge area ${((onTile/(badge.w*badge.h))*100).toFixed(0)}%)`);
        }
      }
      // The rule above is measured against the TILE, so a disc swallowed whole
      // by a large tile is only 4% of it and never fires — which is how a
      // callout came to sit 100% inside an unrelated service with the gate
      // still green. Readability is a property of the disc, so measure it that
      // way too.
      //
      // The bar is "swallowed", not "touching". A disc straddling a tile edge
      // still reads as a callout on an arrow; one wholly inside a tile reads as
      // that tile's own number and hides the icon underneath it. Measured
      // residue at the time of writing: `twin-ladders` rests two rungs of a
      // ten-deep ladder at ~50% over a tile edge, which no weight in the walk
      // moves and which is the accepted cost of routing its wrap-around hops
      // through the row gutter instead of straight through three services.
      //
      // The bar was 0.9 and that was fixture-tuned: a grid one node wider
      // buried five callouts at 87% and passed, although 0.87 and 0.93 are the
      // same picture and a disc can be moved 3% of its diameter to satisfy the
      // rule. Against the shipping corpus the worst single-tile burial is 0.50,
      // so 0.7 leaves 0.20 of headroom and still catches those grids.
      if (buried > 0.7 * badge.w * badge.h) {
        const worst = slideTiles.reduce((a, b) => (overlapArea(badge, b) > overlapArea(badge, a) ? b : a), slideTiles[0]);
        issues.push(`step badge "${badge.name}" is ${((buried / (badge.w * badge.h)) * 100).toFixed(0)}% buried inside "${worst?.name}"`);
      }
      // A callout must not be wider than the service it calls out.
      //
      // This rule lived only in the drawing exporter for four rounds while the
      // deck sized its disc as `clamp(0.26 * px, 0.18, 0.42)` - an absolute
      // floor and an absolute ceiling with no reference to the tile at all.
      // Because `px` is the drawn width of 96 authored pixels, the ratio is
      // `0.26 * 96 / W` and does not move with the scale, so the defect was
      // invisible to every rule that measures scale: a 14px node drew a
      // 0.3566in disc on a 0.2000in tile, 178%, and the deck passed while the
      // identical drawing failed 51 times as a Visio sheet.
      //
      // Resolved by identity and not by position. `connector-step-<routeId>`
      // and `service-<nodeId>` both carry the id directly, so unlike the sheet
      // this needs no name matching at all.
      const routeId = badge.name.slice('connector-step-'.length);
      const ends = endsOfRoute.get(routeId);
      if (ends) {
        const drawn = ends
          .map((nodeId) => ({ nodeId, tile: slideTiles.find((t) => t.name.slice('service-'.length) === nodeId) }))
          .filter((d): d is { nodeId: string; tile: typeof slideTiles[number] } => !!d.tile && d.tile.w > 0)
          .map((d) => ({ nodeId: d.nodeId, w: d.tile.w, shape: d.tile }));
        const widths = drawn.map((d) => d.w);
        // Measure against the ends that ARE drawn; skip only when neither is.
        //
        // This used to require both ends, on the reasoning that "the end that
        // was cut is not evidence about the end drawn". True for a rule about
        // the hop, false for this one: the disc and the tile are both drawn on
        // this slide, and that pair is the entire measurement. The consequence
        // was perverse, because splitting the deck is the exporter's remedy for
        // a tile too small, and splitting is exactly what puts a hop's two ends
        // on different windows - so the better the chase worked, the blinder
        // the gate got. Measured on one drawing, all four badges were outside
        // the field of view and one of them satisfied the conflict predicate
        // exactly, at 77% of the service it numbers; across the shipping corpus
        // 34 badges were invisible for this reason alone.
        if (widths.length > 0) {
          const tile = Math.min(...widths);
          // Only where a proportionate disc could also have been a legible one.
          //
          // Round 72 established that this precondition is not a narrow window,
          // it is the exact complement of the rule: the exporter emits
          // `max(floor, min(natural, 0.55 * tile))`, so the bar can only be
          // exceeded by a disc sitting on the floor, and a disc on the floor is
          // what the precondition exempts. Armed implies cannot fire. So this
          // clause stays - measuring the disc there points at the only object
          // in the picture that is behaving correctly - and the conflict itself
          // is reported below, against the tile, where it can be acted on.
          //
          // The clause is not dead weight. It is the mutation guard on the
          // ceiling: restore the tile-blind clamp and it fires on every
          // ordinary numbered deck.
          const floor = badgeFloorIn(stepOfRoute.get(routeId) ?? 1, 7);
          // Does the disc actually MEET a service it numbers?
          //
          // Six rounds of this rule family compared a diameter to a width and
          // called the ratio a defect. Measured on the emitted slides, the four
          // discs it was still reporting at 98% lay entirely off their icons -
          // `overlapArea` zero for every one - at the author's own faithful
          // proportion, 14/160 authored and 0.1584/1.81 drawn, both 8.75%. A
          // number that never comes near the thing it numbers is not competing
          // with it, and the failure this family exists for, round 71's 178%,
          // was a disc drawn ON TOP of an icon. So contact is now part of the
          // claim. It also retires an exemption: with contact required, the
          // rule goes quiet on that deck on its MERITS, and the exporter's
          // `chaseAffordable` - which was the median order statistic re-entering
          // one level down, false on every drawing whose narrowest badged tile
          // ties the median, with no plan measured at all - is no longer needed
          // to silence it. Occlusion past 70% stays with the buried rule; this
          // one owns the band where a disc touches an icon it dwarfs.
          const touches = drawn.some((d) => overlapArea(badge, d.shape) > 0);
          if (tile * BADGE_SHARE >= floor - 1e-6 && badge.w > tile * BADGE_SHARE + 2e-4) {
            issues.push(`step badge "${badge.name}" is ${badge.w.toFixed(4)}in across on a `
              + `${tile.toFixed(4)}in tile — ${((badge.w / tile) * 100).toFixed(0)}% of the `
              + 'service it is calling out');
          } else if (badge.w > tile + 1e-6 && touches && drawn.every((d) => {
            const w = authoredW.get(d.nodeId);
            return w === undefined || w >= calloutPlan.reachableTileW * BADGE_SHARE - 1e-9;
          })) {
            // A disc WIDER THAN THE SERVICE IT NUMBERS, reported ahead of every
            // exemption that is about the plan.
            //
            // There is no reading of a deck in which this is acceptable, so it
            // is tested before the served-tile and affordability clauses, which
            // all ask whether the exporter had a better PLAN available. This one
            // does not care. Measured on six ordinary services and one
            // glyph-sized DNS zone - the commonest shape the Architecture
            // Center draws - the zone came out 0.1293in under a 0.1556in disc
            // at a 290px pitch and 0.1178in off the row: 120% and 132% of the
            // thing being pointed at, with the gate silent, because every rule
            // was keyed on the tile the planner served and the zone is narrower
            // than that by construction.
            //
            // The one condition is a bound on the FRAME, not on the plan. Below
            // `reachableTileW` the bleed alone fills the window before the bar
            // is met, so no split reaches it and the page size is the only
            // remedy - which is why Visio draws these same drawings correctly.
            // It is per hop, against this hop's own endpoints. Applying it
            // caught a shipped defect the round it was written: 120 sensors of
            // 14 authored px number 119 hops, and three-digit discs came out
            // 0.2451in on 0.2006in tiles, 122%, on a corpus reporting clean.
            // That deck needs a 16.7px tile for a proportionate disc and has
            // 14, so it is exempt here and reported nowhere - the honest next
            // move is to let a callout's digits take the same 7pt floor its
            // tiles already take, which would size that disc 0.1907in.
            conflicts.push({ name: badge.name, tile, floor, ratio: badge.w / tile });
          } else if (tile * BADGE_SHARE < floor - 1e-6
            && !calloutBarClamped
            // Only where this FRAME can deliver the bar for this hop's own
            // tiles. Below `reachableTileW` no split reaches it - the bleed
            // alone fills the window first - so the deck has no move and the
            // page size is the only remedy, which is why Visio draws these
            // same drawings at 55%. Per hop, against its own endpoints: read
            // deck-wide it would repeat the `servedW: Infinity` defect and
            // exempt hops between tiles ten times wider.
            && drawn.every((d) => {
              const w = authoredW.get(d.nodeId);
              return w === undefined || w >= calloutPlan.reachableTileW * BADGE_SHARE - 1e-9;
            })
            && touches) {
            // The empty intersection, stated as what it is. A tile this small
            // admits no disc that is both readable and proportionate, so there
            // is nothing the callout can do about it and the fix is a coarser
            // split or no callout on this slide at all.
            //
            // Only above the markable bar, which is the line between a tile the
            // planner chases and a tile it has already declined to. Below it
            // sits an authored sliver: `probe-two-chains` numbers six 24px
            // sensors, and no split widens them, because a 24x96 node is capped
            // by the type ceiling at a scale set by its HEIGHT - the aspect is
            // the author's and the plan cannot touch it. The renderer floors
            // such a tile's type at 7pt regardless, so the sliver reads; what
            // it cannot do is carry a proportionate disc, and reporting that
            // would be another bar with no move behind it.
            //
            // At and above the bar the planner met its old contract and this is
            // a stricter one, which is exactly what `markableTileWIn` raises.
            //
            // And only where the stricter one is REACHABLE. Round 73 measured
            // a deck where it is not: 120 nodes of 14 authored px saturate at
            // 0.42in against a 0.4457in three-digit bar, and the same drawing
            // was clean at 99 steps and failed at 100 - a labelling act with no
            // geometric content. The premise written above, that at and above
            // the bar the stricter contract can be met, is simply false in that
            // region, so the rule now says so instead of asserting it.
            //
            // Read from `calloutBarClamped`, which is the exporter's own
            // decision and not a second opinion about it: exactly where the
            // planner stops chasing the callout bar, the gate stops requiring
            // it. Every other drawing stays armed, including one authored at
            // 16px, where the bar is reachable and a miss is a real miss.
            conflicts.push({ name: badge.name, tile, floor, ratio: badge.w / tile });
          }
          // The reader hunts a callout against the LARGEST thing beside it, not
          // the smallest, so the undersize test takes the other denominator
          // from the ceiling.
          //
          // Exempt at every value the sizing model can produce, which is all
          // three of `max(floor, min(natural, ceiling))`. At any of them the
          // disc is exactly what the model asked for and the fault, if there is
          // one, is not the disc's.
          //
          // Round 73 shipped this rule with only two of the three and it
          // false-positived at once. `probe-wide-hub` draws every callout at
          // its natural 0.2063in, between a 0.1556in floor and a 0.2728in
          // ceiling, and three of them were reported at 7.8% because ONE OTHER
          // SHAPE on the slide is 320 authored px wide. The trigger is
          // scale-free - `natural / maxTile = 24.96 / Wmax` whatever the
          // transform does - so it crosses 10% at 250 authored px and a
          // full-width gateway banner beside ordinary services was enough.
          //
          // What is left is a disc at none of the three: one cut by something
          // that is not its own hop. That is the two-chain fault, where a
          // sheet-wide ceiling capped discs at 0.1375in with a 0.24in ceiling,
          // a 0.1119in floor and a 0.26in natural size - clear of all of them.
          //
          // `natural` is recovered from the drawing rather than assumed,
          // because the gate cannot see the transform: every tile on a slide
          // shares one scale, so `scale = drawn / authored` on any endpoint and
          // `natural = 0.26 * 96 * scale`.
          const widest = Math.max(...widths);
          const ceiling = tile * BADGE_SHARE;
          const authored = authoredW.get(drawn[0].nodeId);
          const natural = authored && authored > 0
            ? Math.min(0.42, Math.max(floor, 0.26 * 96 * (drawn[0].w / authored)))
            : NaN;
          if (badge.w < widest * 0.1
            && badge.w < ceiling - 1e-4
            && badge.w > floor + 1e-4
            && !(Number.isFinite(natural) && Math.abs(badge.w - natural) <= 1e-3)) {
            issues.push(`step badge "${badge.name}" is ${badge.w.toFixed(4)}in across beside a `
              + `${widest.toFixed(4)}in tile — ${((badge.w / widest) * 100).toFixed(1)}% of it, `
              + 'too small to find on the hop it numbers');
          }
        }
      }
      for (const chip of slideChips) {
        if (overlapArea(badge, chip) > 0.25 * badge.w * badge.h) {
          issues.push(`step badge "${badge.name}" collides with edge chip "${chip.text}"`);
        }
      }
      // A callout whose number runs outside its own bubble reads as a smear
      // over whatever the arrow passes through.
      const digits = badge.text.length;
      const wide = digits * 0.62 * ((badge.fontSize ?? 9) / 72);
      const tall = ((badge.fontSize ?? 9) * 1.3) / 72;
      if (wide > badge.w + 0.005 || tall > badge.h + 0.005) {
        issues.push(
          `step badge "${badge.name}" draws "${badge.text}" at ${(badge.fontSize ?? 9).toFixed(1)}pt needing ${wide.toFixed(3)}x${tall.toFixed(3)}in inside a ${badge.w.toFixed(3)}in circle`,
        );
      }
    }
  }
  // The empty intersection, reported once per deck rather than once per disc.
  //
  // A tile under `floor / 0.55` admits no diameter that is both readable and
  // proportionate, so the exporter draws the floor and the ratio rule above
  // correctly declines to blame the disc. Something still has to be said, or
  // the deck ships callouts at 62% and 78% of the services they number with a
  // clean gate - which is exactly what rounds 68 to 71 did. This is the rule
  // that can actually fire on exporter output, and it points at the tile.
  //
  // Worst first, and capped, because one bad plan produces one line per hop
  // and the number of hops is not information.
  if (conflicts.length > 0) {
    const worst = conflicts.sort((a, b) => b.ratio - a.ratio)[0];
    issues.push(`${conflicts.length} step callout(s) sit on a tile too small to carry one: `
      + `"${worst.name}" needs at least ${worst.floor.toFixed(4)}in to stay readable and at most `
      + `${(worst.tile * BADGE_SHARE).toFixed(4)}in to stay proportionate on its `
      + `${worst.tile.toFixed(4)}in tile, so it draws at ${(worst.ratio * 100).toFixed(0)}% of `
      + 'the service it is calling out');
  }
  // A target the frame cannot reach must not be paid for anyway.
  //
  // Where `calloutBarClamped` holds, `legibleScaleFor` has stopped chasing the
  // callout bar and fallen back to the markable one, so the deck should cost
  // what the same architecture costs unnumbered. Under the mutation that
  // removes the clamp it does not: 120 nodes of 14 authored px came out as 64
  // slides at a 0.374in tile, against 21 slides at 0.201in with the clamp in
  // place - three times the deck, bought by chasing a 0.4457in bar that
  // saturates at 0.42in and is missed either way.
  //
  // Stated as the tile, because that is what the extra windows were spent on
  // and it is the one number the file carries. A clamped plan lands ON the
  // markable bar by construction; half as much again is not a rounding.
  if (calloutBarClamped && slideCount > 1 && minTileW > MARKABLE_TILE_W_IN * 1.5) {
    issues.push(`the callout bar is out of reach for this drawing, so the plan should stop at the `
      + `${MARKABLE_TILE_W_IN.toFixed(2)}in markable bar — instead ${slideCount} slides were spent `
      + `reaching a ${minTileW.toFixed(3)}in tile that still cannot carry its callout`);
  }
  // A shape reduced to a hairline is worse than one drawn too big: the reader
  // cannot see that anything is missing — the band is simply gone, and so is
  // whatever it said. Measured across the whole deck rather than per slide,
  // because a zone wider than a window is legitimately a sliver on the slides
  // at its edges; what is never legitimate is a zone that is a sliver on every
  // slide it appears on, which is what a compaction bug produces.
  const bestZone = new Map<string, { w: number; h: number }>();
  for (const slideShapes of perSlide) {
    for (const zone of slideShapes.filter((s) => s.name.startsWith('zone-') && !s.name.startsWith('zone-label-'))) {
      const id = zone.name.replace(/^zone-/, '');
      const best = bestZone.get(id);
      bestZone.set(id, { w: Math.max(best?.w ?? 0, zone.w), h: Math.max(best?.h ?? 0, zone.h) });
    }
  }
  for (const [id, seen] of bestZone) {
    const node = scenario.nodes.find((n) => n.id === id);
    // Absolute only. A deck legitimately cuts a zone at a window edge, so the
    // proportions it is drawn in on any one slide are not the author's, and
    // comparing them reports every wide scope band as crushed.
    if (seen.w < 0.05 || seen.h < 0.05) {
      issues.push(`zone "${String(node?.data?.label ?? id)}" is never drawn larger than ${seen.w.toFixed(3)}x${seen.h.toFixed(3)}in — a shape flattened to a line is a shape deleted`);
    }
  }
  // Two annotations on one arrow's worth of space is a pile, not a ladder. The
  // tile rules never look at annotation-on-annotation, so a fan restacking on
  // itself - or one ladder parking on another - used to pass at zero issues.
  for (const slideShapes of perSlide) {
    const annotations = slideShapes.filter((s) => /^connector-(label|step)-/.test(s.name));
    for (let i = 0; i < annotations.length; i += 1) {
      for (let j = i + 1; j < annotations.length; j += 1) {
        const a = annotations[i];
        const b = annotations[j];
        const hit = overlapArea(a, b);
        if (hit > 0.01) issues.push(`annotation "${a.name}" and "${b.name}" overlap by ${hit.toFixed(3)}sq in`);
      }
    }
  }
  // On the Azure Architecture Center an arrow never runs through a service it
  // does not connect: a line that disappears under a tile and comes out the
  // other side reads as touching it, and on a generated layout that is the most
  // visible difference between a reference drawing and a sketch. Nothing
  // measured this — the connector path was parsed only to attribute chips.
  const endpointsOf = new Map<string, { source: string; target: string }>();
  for (const edge of scenario.edges) {
    endpointsOf.set(String(edge.id), { source: String(edge.source), target: String(edge.target) });
  }
  for (const slideShapes of perSlide) {
    const tiles = slideShapes.filter(
      (s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'),
    );
    for (const arrow of slideShapes.filter((s) => s.name.startsWith('connector-') && !/^connector-(label|step)-/.test(s.name))) {
      const path = arrow.path;
      if (!path || path.length < 2) continue;
      const ends = endpointsOf.get(arrow.name.slice('connector-'.length));
      if (!ends) continue;
      for (const tile of tiles) {
        const id = tile.name.slice('service-'.length);
        if (id === ends.source || id === ends.target) continue;
        // Shrink the tile so an arrow that merely grazes a corner or runs along
        // an edge is not reported; only a line that genuinely goes under the
        // service counts.
        const inset = 0.04;
        const box = { x: tile.x + inset, y: tile.y + inset, w: tile.w - 2 * inset, h: tile.h - 2 * inset };
        if (box.w <= 0 || box.h <= 0) continue;
        let inside = 0;
        for (let i = 1; i < path.length; i += 1) {
          const a = path[i - 1];
          const b = path[i];
          const segLen = Math.hypot(b.x - a.x, b.y - a.y);
          const steps = Math.max(2, Math.ceil(segLen / 0.02));
          for (let s = 0; s <= steps; s += 1) {
            const t = s / steps;
            const px = a.x + (b.x - a.x) * t;
            const py = a.y + (b.y - a.y) * t;
            if (px >= box.x && px <= box.x + box.w && py >= box.y && py <= box.y + box.h) {
              inside += segLen / steps;
            }
          }
        }
        if (inside > 0.15) {
          issues.push(`arrow "${arrow.name}" runs ${inside.toFixed(2)}in through node "${tile.name}", which it does not connect`);
        }
      }
    }
  }
  // Two arrows drawn one on top of the other are one arrow to the reader. The
  // router is a pure function of a single edge, so two hops that meet the same
  // port of the same service — a chain hop leaving head-on and an elbow hop
  // arriving around the corner — are handed the identical centre line and the
  // shorter of the two simply disappears. Reference architectures never do
  // this: every arrow into a box lands on its own point.
  //
  // Fan siblings are exempt. A bundle of parallel edges between one pair of
  // services is deliberately drawn as a set of parallel lines and is already
  // spread by `fanOffset`; measuring them against each other would report the
  // feature as the defect.
  for (const slideShapes of perSlide) {
    const arrows = slideShapes
      .filter((s) => s.name.startsWith('connector-') && !/^connector-(label|step)-/.test(s.name))
      .filter((s) => (s.path?.length ?? 0) >= 2);
    const pairKey = (name: string): string => {
      const ends = endpointsOf.get(name.slice('connector-'.length));
      if (!ends) return name;
      return ends.source < ends.target ? `${ends.source}|${ends.target}` : `${ends.target}|${ends.source}`;
    };
    const lengthOf = (s: Shape): number => {
      const path = s.path ?? [];
      let sum = 0;
      for (let i = 1; i < path.length; i += 1) sum += Math.hypot(path[i].x - path[i - 1].x, path[i].y - path[i - 1].y);
      return sum;
    };
    const bboxOf = (s: Shape): { x0: number; y0: number; x1: number; y1: number } => {
      const path = s.path ?? [];
      return {
        x0: Math.min(...path.map((p) => p.x)),
        y0: Math.min(...path.map((p) => p.y)),
        x1: Math.max(...path.map((p) => p.x)),
        y1: Math.max(...path.map((p) => p.y)),
      };
    };
    for (const short of arrows) {
      const shortLen = lengthOf(short);
      if (shortLen < 0.05) continue;
      const bs = bboxOf(short);
      const others = arrows.filter((other) => {
        if (other === short) return false;
        if (pairKey(other.name) === pairKey(short.name)) return false;
        if (lengthOf(other) < shortLen - 0.001) return false;
        // Same length: break the tie by name so a genuinely coincident pair is
        // reported once, against one of the two, rather than twice or not at all.
        if (Math.abs(lengthOf(other) - shortLen) <= 0.001 && other.name <= short.name) return false;
        const bo = bboxOf(other);
        return !(bs.x1 + 0.05 < bo.x0 || bo.x1 + 0.05 < bs.x0 || bs.y1 + 0.05 < bo.y0 || bo.y1 + 0.05 < bs.y0);
      });
      if (others.length === 0) continue;
      const path = short.path ?? [];
      let coincident = 0;
      for (let k = 1; k < path.length; k += 1) {
        const p = path[k - 1];
        const q = path[k];
        const segLen = Math.hypot(q.x - p.x, q.y - p.y);
        const steps = Math.max(2, Math.ceil(segLen / 0.03));
        for (let s = 0; s < steps; s += 1) {
          const t = (s + 0.5) / steps;
          const at = { x: p.x + (q.x - p.x) * t, y: p.y + (q.y - p.y) * t };
          if (others.some((other) => pathGap(other, at) < 0.02)) coincident += segLen / steps;
        }
      }
      // Both bars matter. A long absolute run is unreadable however long the
      // arrow is, and a short hop swallowed whole is invisible however short
      // the run. The share bar is above a half so two hops that merely share a
      // corner, or cross, are never reported.
      if (coincident > 0.3 && coincident > 0.55 * shortLen) {
        const under = others
          .filter((other) => {
            const p2 = short.path ?? [];
            return p2.some((pt) => pathGap(other, pt) < 0.02);
          })
          .map((other) => other.name)
          .slice(0, 4)
          .join(', ');
        issues.push(`arrow "${short.name}" is drawn under other arrows (${under || 'unnamed'}) for ${coincident.toFixed(2)}in of its ${shortLen.toFixed(2)}in length`);
      }
    }
    // Two hops leaving the SAME side of the SAME service must leave it in the
    // same order they arrive. A fan out of a front door is one of the commonest
    // shapes on the Architecture Center, and when the ports are dealt in the
    // wrong order the arrows braid: every line crosses its neighbours on the
    // way out and the reader has to trace each one through the knot. The braid
    // forms on the shared jog column, where the segments are collinear rather
    // than properly crossing, so it is caught by ranking rather than by
    // intersection.
    const sourceOf = new Map<string, string>();
    for (const edge of scenario.edges) sourceOf.set(String(edge.id), String(edge.source));
    const fans = new Map<string, { name: string; depart: number; arrive: number }[]>();
    for (const arrow of arrows) {
      const path = arrow.path ?? [];
      if (path.length < 4) continue;
      const from = sourceOf.get(arrow.name.slice('connector-'.length));
      if (from === undefined) continue;
      const stubIsHorizontal = Math.abs(path[1].x - path[0].x) > Math.abs(path[1].y - path[0].y);
      const key = `${from}#${stubIsHorizontal ? (path[1].x > path[0].x ? 'E' : 'W') : (path[1].y > path[0].y ? 'S' : 'N')}`;
      const list = fans.get(key) ?? [];
      const end = path[path.length - 1];
      list.push({
        name: arrow.name,
        depart: stubIsHorizontal ? path[2].y - path[0].y : path[2].x - path[0].x,
        arrive: stubIsHorizontal ? end.y : end.x,
      });
      fans.set(key, list);
    }
    for (const [key, fan] of fans) {
      if (fan.length < 2) continue;
      const ranked = [...fan].sort((a, b) => a.arrive - b.arrive);
      for (let i = 1; i < ranked.length; i += 1) {
        const prev = ranked[i - 1];
        const here = ranked[i];
        // Only judge pairs whose destinations are genuinely apart, so a tie
        // broken either way is never called a braid.
        if (here.arrive - prev.arrive < 0.05) continue;
        if (here.depart < prev.depart - 0.001) {
          issues.push(`arrows "${prev.name}" and "${here.name}" leave "service-${key.split('#')[0]}" in the opposite order to their destinations and cross each other`);
        }
      }
    }
  }
  // A chip has to be readable AS the label of the arrow it belongs to. One
  // parked beside a different hop is worse than one overlapping a tile: the
  // reader matches it to the wrong arrow and never knows they did.
  for (const slideShapes of perSlide) {
    const arrows = slideShapes.filter((s) => s.name.startsWith('connector-') && !/^connector-(label|step)-/.test(s.name));
    if (arrows.length === 0) continue;
    for (const chip of slideShapes.filter((s) => s.name.startsWith('connector-label-'))) {
      const own = arrows.find((arrow) => arrow.name === `connector-${chip.name.replace('connector-label-', '')}`);
      if (!own) continue;
      const at = { x: chip.x + chip.w / 2, y: chip.y + chip.h / 2 };
      const mine = pathGap(own, at);
      const nearest = arrows.reduce((best, arrow) => (pathGap(arrow, at) < pathGap(best, at) ? arrow : best), arrows[0]);
      if (nearest.name !== own.name && pathGap(nearest, at) < mine - 0.25) {
        issues.push(`edge chip [${chip.name}] "${chip.text}" is ${pathGap(nearest, at).toFixed(2)}in from ${nearest.name} but ${mine.toFixed(2)}in from its own arrow`);
      }
    }
  }
  // The numbered callout has to point at the same hop its wording does. It is
  // measured against arrows from *other* bundles only: a fan of parallel edges
  // between one pair of services is a single object to the reader, so a rung
  // sitting nearer fan-sibling 5 than fan-sibling 6 misleads nobody.
  const bundleOf = new Map<string, string>();
  for (const edge of scenario.edges) {
    bundleOf.set(edge.id, [edge.source, edge.target].sort().join('|'));
  }
  const bundleKey = (arrowName: string): string => bundleOf.get(arrowName.replace('connector-', '')) ?? arrowName;
  for (const slideShapes of perSlide) {
    const arrows = slideShapes.filter((s) => s.name.startsWith('connector-') && !/^connector-(label|step)-/.test(s.name));
    if (arrows.length === 0) continue;
    for (const badge of slideShapes.filter((s) => s.name.startsWith('connector-step-'))) {
      const own = arrows.find((arrow) => arrow.name === `connector-${badge.name.replace('connector-step-', '')}`);
      if (!own) continue;
      const ownBundle = bundleKey(own.name);
      const at = { x: badge.x + badge.w / 2, y: badge.y + badge.h / 2 };
      const mine = pathGap(own, at);
      const strangers = arrows.filter((arrow) => bundleKey(arrow.name) !== ownBundle);
      if (strangers.length === 0) continue;
      const nearest = strangers.reduce((best, arrow) => (pathGap(arrow, at) < pathGap(best, at) ? arrow : best), strangers[0]);
      if (pathGap(nearest, at) < mine - 0.25) {
        issues.push(`callout "${badge.name}" is ${pathGap(nearest, at).toFixed(2)}in from ${nearest.name} but ${mine.toFixed(2)}in from its own arrow`);
      }
    }
  }
  // The colour key is drawn last and is 92% opaque, so whatever it lands on is
  // invisible in the finished deck. A callout it buries leaves the workflow
  // band citing a step number that appears nowhere on the drawing.
  for (const slideShapes of perSlide) {
    const legend = slideShapes.find((s) => s.name === 'connection-legend');
    if (!legend) continue;
    for (const other of slideShapes) {
      if (!/^(connector-label-|connector-step-|service-|zone-label-)/.test(other.name)) continue;
      const hit = overlapArea(legend, other);
      if (hit <= 0.001) continue;
      issues.push(`connection legend covers ${((hit / Math.max(other.w * other.h, 1e-6)) * 100).toFixed(0)}% of "${other.name}"`);
    }
  }
  // A chip is drawn after everything except the colour key and is 92% opaque,
  // so a caption underneath one is not dimmed, it is gone. The exporter keeps
  // chips off tile captions by handing every tile's band to the placement
  // search as an obstacle; zone captions were handed over to nobody, which
  // did not matter while every candidate band sat in a zone's margin — chips
  // are seated in the corridors between tile rows, and captions were never
  // there. Offering those corridors as caption rows put the two in the same
  // place, so the obstacle has to be real.
  for (const slideShapes of perSlide) {
    const chips = slideShapes.filter((s) => s.name.startsWith('connector-label-'));
    if (chips.length === 0) continue;
    for (const caption of slideShapes) {
      if (!/^(zone-label-|service-label-)/.test(caption.name)) continue;
      // The ink, not the band. A caption is left-aligned in a band sized for
      // the widest name in its row, so a chip parked over the empty right end
      // of that band covers no glyph and is not a defect — reporting it would
      // make the rule fire on layouts that read perfectly and teach the
      // exporter to move chips away from nothing.
      const ink = drawnTextRect(caption);
      if (!ink) continue;
      for (const chip of chips) {
        const hit = overlapArea(chip, ink);
        if (hit <= 0.001) continue;
        issues.push(
          `chip "${chip.name}" covers ${((hit / Math.max(ink.w * ink.h, 1e-6)) * 100).toFixed(0)}% `
          + `of the drawn name "${caption.text.trim().slice(0, 28)}"`,
        );
      }
    }
  }
  // Wording may never simply vanish. A label the exporter decided not to draw
  // has to survive as a numbered callout that the workflow slide explains -
  // that is the only trade the Architecture Center makes - and if it does
  // neither, the export has quietly lost content the author wrote.
  const drawnChips = new Set(shapes.filter((s) => s.name.startsWith('connector-label-')).map((s) => s.name.replace('connector-label-', '')));
  const drawnBadges = new Map(shapes.filter((s) => s.name.startsWith('connector-step-')).map((s) => [s.name.replace('connector-step-', ''), s.text]));
  const explained = new Set(
    shapes.map((s) => /^workflow-step-(\d+)$/.exec(s.name)?.[1]).filter((n): n is string => !!n),
  );
  // A row that exists is not a row that says anything. The trade the exporter
  // makes when it mutes a chip is "the workflow slide carries this wording
  // instead", so the wording has to actually be on that slide - otherwise the
  // deck reads "13. Step 13" and the author's sentence is simply gone.
  const foldWording = (s: string): string => s
    .toLowerCase()
    .replace(/[\s\u3000]+/g, '')
    .replace(/[.,;:!?、。（）()[\]「」"'`´’‘“”\-…]/g, '');
  const deckWording = foldWording(shapes.map((s) => s.text).join(' '));
  for (const edge of scenario.edges) {
    // Both sides sanitised: the deck's chips and badges are keyed by the name
    // the exporter wrote, and its wording is the sanitised wording.
    const eid = auditStrip(String(edge.id));
    const label = auditStrip(readEdgeLabel(edge)).trim();
    if (!label || drawnChips.has(eid)) continue;
    const badge = drawnBadges.get(eid);
    if (badge !== undefined && explained.has(badge)) {
      if (!deckWording.includes(foldWording(label))) {
        issues.push(`edge "${eid}" was muted to callout ${badge}, but its wording "${label}" appears nowhere in the deck`);
      }
      continue;
    }
    issues.push(
      badge === undefined
        ? `edge "${eid}" is labelled "${label}" but the deck has neither a chip nor a callout for it`
        : `edge "${eid}" lost its label "${label}" to callout ${badge}, which no workflow row explains`,
    );
  }
  const authoredNames = new Map<string, string>();
  for (const node of scenario.nodes) {
    if (node.type !== 'azureNode') continue;
    const label = (node.data as { label?: string } | undefined)?.label;
    // Stored in the form the exporters will actually draw, not as authored.
    // Every comparison below asks "did this name reach a shape?", and an
    // authored spelling that no file contains answers no for every one of
    // them. `singleLineName` is the exporter's own function, called rather
    // than copied, and it is idempotent, so the later `drawnForm` calls are
    // unaffected.
    if (typeof label === 'string' && label) authoredNames.set(auditStrip(String(node.id)), singleLineName(label));
  }
  const authoredZones = new Map<string, string>();
  for (const node of scenario.nodes) {
    if (node.type === 'azureNode') continue;
    const label = (node.data as { label?: string } | undefined)?.label;
    if (typeof label === 'string' && label) authoredZones.set(auditStrip(String(node.id)), label);
  }
  // Truncation is only acceptable when the full wording survives somewhere the
  // reader can reach. A chip clipped to 42 cells with no workflow row carrying
  // the rest has silently thrown away what the author wrote.
  // An index row is now "<mark>  =  <full name>", the same pair the Visio index
  // prints, so membership is tested by SUFFIX rather than by equality. Matched
  // on the suffix and not by splitting on the separator, because a name may
  // contain the separator itself and an undrawable name leaves an empty mark.
  const indexRows = shapes
    .filter((s) => s.name.startsWith('index-name-'))
    .map((s) => s.text.trim());
  const indexed = {
    has: (name: string): boolean => indexRows.some(
      (row) => row === name || row.endsWith(`  =  ${name}`) || row.endsWith(`=  ${name}`),
    ),
  };
  // An index row now QUOTES the mark, and the mark is a cut string, so the row
  // legitimately contains an ellipsis - "Az…  =  Azure Firewall Premium". It is
  // the recovery route, not a label that lost its tail, and counting it as one
  // made the index slide report itself as damage in 17 corpus scenarios. What
  // must never be cut is the NAME half, and that is checked below against the
  // authored strings rather than by looking for an ellipsis anywhere in the row.
  const truncated = shapes.filter(
    (s) => s.text.includes('…') && !s.name.startsWith('index-name-'),
  );
  const stranded = truncated.filter((s) => {
    const svcId = /^service-label-(.*)$/.exec(s.name)?.[1];
    if (svcId !== undefined) {
      // The deck's index slide spells this one out, so the drawing is free to
      // abbreviate it.
      const name = authoredNames.get(svcId);
      return !name || !indexed.has(name);
    }
    // A zone whose band cannot hold its name is cut like any other label, and
    // is spelled out in the same place. The band is not negotiable — a window
    // can cut a zone to a 0.39in sliver — so the index is the only way the
    // name survives the cut.
    const zoneId = /^zone-label-(.*)$/.exec(s.name)?.[1];
    if (zoneId !== undefined) {
      const name = authoredZones.get(zoneId);
      return !name || !indexed.has(name);
    }
    const id = /^connector-label-(.*)$/.exec(s.name)?.[1];
    if (!id) return true;
    const badge = drawnBadges.get(id);
    return badge === undefined || !explained.has(badge);
  });
  if (stranded.length) {
    issues.push(`${stranded.length} truncated label(s) have no workflow row carrying the rest: ${stranded.slice(0, 3).map((s) => s.name).join(', ')}`);
  }

  // Workflow numbering: an arrow that the AI numbered must carry its callout,
  // and the callout must not sit on top of a node or its own label chip —
  // either way the reader cannot match the arrow to the workflow prose.
  //
  // Expectations come from the repaired edges, not the raw scenario: the
  // exporter renumbers duplicate step numbers before drawing, so raw data
  // would assert that five arrows all still read "3".
  const numberedEdges = narrateEdgeCallouts(scenario.edges).filter(
    (e) => readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber) !== undefined,
  );
  const badges = shapes.filter((s) => s.name.startsWith('connector-step-'));
  // A labelled edge the author never numbered may be PROMOTED: when its chip
  // has nowhere legible to stand, the exporter gives it the next free number
  // and a workflow row rather than leaving the words on top of a service name.
  // So the model no longer predicts the badge count exactly — it predicts a
  // floor, plus the set of arrows allowed to appear above it.
  const promotableIds = new Set(
    narrateEdgeCallouts(scenario.edges)
      .filter((e) => readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber) === undefined
        && readEdgeLabel(e) !== '')
      .map((e) => auditStrip(String(e.id))),
  );
  // Membership alone is permutation-blind: swapping every badge onto the wrong
  // arrow would pass. The object name carries the route id, so check the exact
  // arrow-to-number correspondence instead.
  const expectedByRoute = new Map(
    numberedEdges.map((e) => [auditStrip(String(e.id)), String(readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber))]),
  );
  const badgedRoutes = new Set(badges.map((b) => b.name.replace(/^connector-step-/, '')));
  for (const [routeId, want] of expectedByRoute) {
    if (!badgedRoutes.has(routeId)) {
      issues.push(`connector ${routeId} is workflow step ${want} but the drawing has no badge for it`);
    }
  }
  const seenNumbers = new Map<string, string>();
  for (const badge of badges) {
    const routeId = badge.name.replace(/^connector-step-/, '');
    const want = expectedByRoute.get(routeId);
    if (want === undefined) {
      if (!promotableIds.has(routeId)) {
        issues.push(`step badge "${badge.name}" does not belong to any numbered connector`);
      }
    } else if (badge.text !== want) {
      issues.push(`connector ${routeId} is numbered "${badge.text}" but its workflow step is ${want}`);
    }
    // Two arrows may not read the same digit whatever their provenance: the
    // workflow list is keyed by number, so a promoted edge that collided with
    // an authored one would send the reader to somebody else's sentence.
    const already = seenNumbers.get(badge.text);
    if (already !== undefined && already !== routeId) {
      issues.push(`connectors ${already} and ${routeId} both carry the badge "${badge.text}"`);
    }
    seenNumbers.set(badge.text, routeId);
  }

  // A promoted number is only useful if the reader can find it in the list, and
  // a list that reads 1, 2, 3, 16 tells the reader twelve steps went missing.
  // The numbers available to promotion are bounded: the highest number the
  // author wrote, plus one per labelled arrow that could need one. Spending
  // them on the plain connectors — which carry no words and so can never
  // become a row — pushed the one real promotion off the end of the sequence.
  {
    const authoredMax = numberedEdges.reduce(
      (most, e) => Math.max(most, readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber) ?? 0),
      0,
    );
    const ceiling = authoredMax + promotableIds.size;
    for (const badge of badges) {
      const routeId = badge.name.replace(/^connector-step-/, '');
      if (!promotableIds.has(routeId) || expectedByRoute.has(routeId)) continue;
      const read = Number(badge.text);
      if (Number.isFinite(read) && read > ceiling) {
        issues.push(
          `promoted connector ${routeId} is numbered "${badge.text}", past the ${ceiling} `
          + `numbers this drawing can explain (${authoredMax} authored + ${promotableIds.size} labelled)`,
        );
      }
    }
  }

  // The Workflow list is the prose the reader matches the drawing against, so
  // every number it cites has to exist as a callout on the canvas and vice
  // versa. A hop whose callout was dropped leaves the prose pointing at nothing;
  // a callout with no prose leaves the reader with an unexplained number.
  const workflowNumbers = new Set(
    shapes
      .map((s) => /^workflow-step-(\d+)$/.exec(s.name)?.[1])
      .filter((n): n is string => !!n),
  );
  if (workflowNumbers.size > 0) {
    const callouts = new Set(badges.map((b) => b.text));
    const missing = [...workflowNumbers].filter((n) => !callouts.has(n));
    const unexplained = [...callouts].filter((n) => !workflowNumbers.has(n));
    if (missing.length) {
      issues.push(`workflow cites step${missing.length === 1 ? '' : 's'} ${missing.sort((a, b) => +a - +b).join(', ')} with no callout on the canvas`);
    }
    if (unexplained.length) {
      issues.push(`callout${unexplained.length === 1 ? '' : 's'} ${unexplained.sort((a, b) => +a - +b).join(', ')} appear on the canvas but not in the workflow`);
    }
  }

  // Nothing may be drawn outside the page: an off-slide shape is invisible in
  // PowerPoint, which reads as missing content.
  for (const shape of shapes) {
    if (shape.x < -0.01 || shape.y < -0.01 || shape.x + shape.w > pageW + 0.01 || shape.y + shape.h > pageH + 0.01) {
      issues.push(
        `shape "${shape.name}" is off-page at (${shape.x.toFixed(2)}, ${shape.y.toFixed(2)}) ${shape.w.toFixed(2)}x${shape.h.toFixed(2)}in on a ${pageW.toFixed(2)}x${pageH.toFixed(2)}in page`,
      );
    }
  }
  // Absolute legibility, not just relative fit: sub-7pt body text is unreadable
  // when projected, and a warning note is not a substitute for a readable
  // drawing — an oversized architecture must be split across slides instead.
  if (Number.isFinite(minFont) && minFont < 7) {
    issues.push(`smallest label font is ${minFont}pt (below the 7pt legibility floor)`);
  }
  {
    // Same floor, applied to the arrow chips. It has to be measured separately
    // because the tile rule filters on `service-label-` and so never saw them:
    // the chip carried its own 4pt floor and a scaled-down drawing wrote arrow
    // labels at 6.39pt while every tile beside them was held at 7.
    const chipShapes = allSlides
      .filter((xml) => !xml.includes('(Overview)'))
      .flatMap((xml) => parseShapes(xml));
    issues.push(...connectorLabelFontIssues(chipShapes, ''));
  }

  // A deck nobody can open in PowerPoint is not an export. PowerPoint gives a
  // deck exactly one page size, so an oversized drawing drags the title and
  // workflow slides onto the plotter sheet with it. Splitting the drawing
  // across ordinary slides is the way out, so the rule has an escape hatch --
  // but a single grown page has none, and this rule is deliberately not scoped
  // to the layout-engine scenarios: a hand-placed canvas exports through the
  // same code path and deserves the same deck.
  const standardPage = Math.abs(pageW - 13.333) < 0.05 && Math.abs(pageH - 7.5) < 0.05;
  const diagramSlides = perSlide.filter((slideShapes) =>
    slideShapes.some((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta')),
  ).length;
  if (!standardPage) {
    // Tiling is the escape hatch, but tiling a plotter sheet into more plotter
    // sheets is not: every part still inherits the page size, so the deck is
    // still one nobody can open. A grown page is only defensible when the
    // drawing genuinely cannot be tiled onto ordinary slides at all.
    issues.push(
      diagramSlides <= 1
        ? `the deck is a single ${pageW.toFixed(2)}x${pageH.toFixed(2)}in page instead of standard 13.33x7.5in slides`
        : `the deck is ${diagramSlides} parts of a ${pageW.toFixed(2)}x${pageH.toFixed(2)}in page instead of standard 13.33x7.5in slides`,
    );
  }

  // Learn pairs every numbered callout with the sentence it points at. A badge
  // without its row is an unexplained digit.
  const narrated = new Set(
    scenario.edges
      .map((e) => e.data as { stepNumber?: number; stepDescription?: string } | undefined)
      .filter((d) => Number.isInteger(d?.stepNumber) && !!d?.stepDescription)
      .map((d) => d!.stepNumber!),
  );
  for (const step of narrated) {
    if (!shapes.some((s) => s.name === `workflow-text-${step}`)) {
      issues.push(`workflow step ${step} is numbered on the drawing but missing from the workflow list`);
    }
  }
  for (const row of shapes.filter((s) => s.name.startsWith('workflow-text-'))) {
    if (!row.text.trim()) issues.push(`workflow row "${row.name}" is blank`);
    // PowerPoint does not clip a `valign: middle` box — it spills the overflow
    // symmetrically past both edges. A step sentence that wraps to more lines
    // than its row is tall therefore runs into the rows above and below it, and
    // the list stops being readable exactly when the prose gets real.
    // Measured unclamped on purpose: `drawnTextRect` caps height at the box,
    // which is what makes it useless for asking whether the text fits the box.
    if (row.fontSize && row.text.trim()) {
      // Greedy wrap and 1.35, not `ceil(ink / w)` and 1.22. The ratio is the
      // break-anywhere lower bound the rest of this file was moved off, and
      // 1.22 is below `Yu Gothic UI`'s own 1.3301 win metric, so this rule
      // simultaneously under-counted the lines and under-charged for each one.
      const lines = auditWrappedLines(row.text.trim(), row.w, row.fontSize);
      const needed = (lines * row.fontSize * 1.35) / 72;
      if (needed > row.h + 0.01) {
        issues.push(`workflow row "${row.name}" needs ${needed.toFixed(2)}in of text (${lines} lines at ${row.fontSize}pt) in a ${row.h.toFixed(2)}in row, so it spills onto its neighbours`);
      }
    }
  }
  // Numbers are the only handle a reader has on the prose, so two arrows may
  // never wear the same one, and no sentence the author wrote may go missing.
  // A duplicate used to be silent twice over: both badges read the same digit
  // and the workflow list, keyed by number, kept only the first sentence.
  const badgeCounts = new Map<string, number>();
  for (const badge of badges) badgeCounts.set(badge.text, (badgeCounts.get(badge.text) ?? 0) + 1);
  for (const [text, count] of badgeCounts) {
    if (count > 1) issues.push(`${count} callouts all read "${text}", so the reader cannot tell which row is which`);
  }
  const authored = new Set(
    scenario.edges
      .map((e) => (e.data as { stepDescription?: string } | undefined)?.stepDescription?.trim())
      .filter((d): d is string => !!d)
      // Compare against what the exporter is *right* to emit. A description
      // carrying an XML-forbidden code point has to be sanitised on the way
      // out, so the sentence on the slide legitimately differs from the one the
      // author typed, and demanding they match byte-for-byte would make the
      // only correct behaviour look like a dropped sentence.
      .map((d) => auditStrip(d)),
  );
  if (authored.size > 0) {
    const rowText = new Set(shapes.filter((s) => s.name.startsWith('workflow-text-')).map((s) => s.text.replace(/…$/, '').trim()));
    // Either direction is a match: a row truncated with an ellipsis is a prefix
    // of what the author wrote, and a row that has had the arrow's own label
    // appended to it starts with what the author wrote.
    const lost = [...authored].filter((d) => ![...rowText].some((r) => r.length > 0 && (d.startsWith(r) || r.startsWith(d))));
    if (lost.length) {
      issues.push(`${lost.length} authored step description(s) reach no slide: ${lost.slice(0, 3).join(' | ')}`);
    }
  }

  // Banding must not lose or duplicate anything. A service that falls between
  // two bands is silently absent from the deck; one that straddles a seam is
  // drawn twice, once shoved against a page edge on top of whatever is there.
  const serviceIds = scenario.nodes.filter((n) => n.type === 'azureNode').map((n) => auditStrip(String(n.id)));
  const drawnTiles = new Map<string, number>();
  for (const tile of tiles) {
    const id = tile.name.replace(/^service-/, '');
    drawnTiles.set(id, (drawnTiles.get(id) ?? 0) + 1);
  }
  for (const id of serviceIds) {
    const drawn = drawnTiles.get(id) ?? 0;
    if (drawn === 0) issues.push(`service "${id}" is drawn on no slide`);
    else if (drawn > 1) issues.push(`service "${id}" is drawn on ${drawn} slides`);
  }
  // Truncation is fine; truncation that stops telling two services apart is
  // not. Bar this on identity rather than on a character budget, because Azure
  // names share the prefix "Azure " and a tile with room for four characters
  // still looks generous while every name on the sheet has collapsed to
  // "Azure…". Measured per slide, since that is the unit a reader looks at.
  for (const [index, slideShapes] of perSlide.entries()) {
    const labels = slideShapes.filter((s) => s.name.startsWith('service-label-'));
    if (labels.length < 2) continue;
    const authored = new Set<string>();
    for (const shape of labels) {
      const name = authoredNames.get(shape.name.slice('service-label-'.length));
      if (name) authored.add(name);
    }
    const drawnDistinct = new Set(labels.map((s) => s.text.trim()).filter(Boolean));
    if (authored.size > 1 && drawnDistinct.size < authored.size) {
      issues.push(
        `slide ${index + 1} draws ${authored.size} differently-named services as only `
        + `${drawnDistinct.size} distinct string(s) — e.g. "${[...drawnDistinct][0] ?? ''}"`,
      );
    }
  }
  // DISTINCTNESS WAS ONLY A PROXY FOR RESOLVABILITY. The rule above asks
  // whether two services look different; it cannot ask whether either of them
  // means anything. When the exporter started drawing a numeric key on a tile
  // it could not name, the strings became distinct and this rule fell silent -
  // and the deck now carried marks like "3" that the index slide never defined,
  // so the reader held a box marked 3 and a list of sentences containing no 3.
  // Fixing a proxy is how a caught defect becomes an uncaught one.
  //
  // The bar: every string a tile draws must resolve to exactly one service
  // somewhere in the same FILE - either by being the authored name itself, or
  // by being defined on the index slide as "<mark>  =  <name>".
  const indexRowText = shapes
    .filter((s) => s.name.startsWith('index-name-'))
    .map((s) => s.text.trim());
  const definedMarks = new Map<string, Set<string>>();
  for (const row of indexRowText) {
    const cut = row.indexOf('  =  ');
    if (cut < 0) continue;
    // The mark half is a LIST. One service drawn on two slides at two widths
    // shortens to two stubs, and the row defines both against the one name.
    // Split on "  |  ", not on a comma: cut names contain commas, and splitting
    // on one shredded "…n, Zone Redundant) 0" into fragments matching no tile.
    const marks = row.slice(0, cut).split('  |  ').map((m) => m.trim()).filter(Boolean);
    const authoredName = row.slice(cut + 5).trim();
    for (const mark of marks) {
      if (!definedMarks.has(mark)) definedMarks.set(mark, new Set());
      definedMarks.get(mark)!.add(authoredName);
    }
  }
  // Compared through the two functions the EXPORTER runs every label through,
  // in the order it runs them. A label carrying a tab, a hard break, a control
  // character or a lone surrogate is authored one way and drawn another, and
  // comparing the raw strings reported a tile drawing its own name in full as a
  // mark standing for nothing. Calling the real functions rather than copying
  // their regexes, so this cannot drift from what the file actually contains.
  const drawnForm = (text: string): string => singleLineName(stripXmlForbidden(text));
  const authoredSet = new Set([...authoredNames.values()].map(drawnForm));
  const undefinedMarks = new Map<string, string>();
  const ambiguousMarks = new Map<string, number>();
  // THE OVERVIEW DRAWS MARKS TOO, AND IS WHERE MOST OF THEM ARE. It is excluded
  // from the legibility rules for a real reason - it is deliberately small -
  // but "this mark is defined somewhere" is not a legibility question. Reading
  // only the window slides made the blind spot line up exactly with the defect:
  // the overview holds the smallest tiles, so it needs the most keys, and it
  // was the one slide none of whose keys were checked. `wide-chain` drew 38
  // bare integers there in a deck with no index slide at all and reported PASS.
  const markShapes = [...shapes, ...overviewShapes];
  for (const shape of markShapes) {
    if (!shape.name.startsWith('service-label-')) continue;
    const mark = shape.text.trim();
    if (!mark) continue;
    const owner = authoredNames.get(shape.name.slice('service-label-'.length)) ?? '';
    if (mark === drawnForm(owner) || authoredSet.has(mark)) continue;
    const defined = definedMarks.get(mark);
    if (!defined) undefinedMarks.set(mark, owner);
  }
  // AMBIGUITY IS READ OFF THE INDEX, not off the tiles. Asking only about marks
  // a tile draws cannot see a mark standing for two services while only one of
  // them is on the slide in front of the reader; the index is the one place
  // both are written down.
  //
  // "(not drawn)" is excluded because it is not a mark - it is the row admitting
  // this name reached no shape at all, and several such rows are several
  // admissions rather than one ambiguous key. The condition it stands for is
  // fixed where it is caused: `MARKABLE_TILE_W_IN` lifts the transform cap so a
  // tile arrives wide enough to carry a key, and `probe-hairline-stubs` holds
  // the geometry that used to produce eight of these rows.
  for (const [mark, names] of definedMarks) {
    if (mark !== UNLABELLED_ROW && names.size > 1) ambiguousMarks.set(mark, names.size);
  }
  if (undefinedMarks.size > 0) {
    issues.push(
      `${undefinedMarks.size} mark(s) drawn on tiles are defined nowhere in the deck, so the reader has no `
      + `way to tell what they stand for: `
      + `${[...undefinedMarks].slice(0, 3).map(([m, o]) => `${JSON.stringify(m)} (${o})`).join('; ')}`,
    );
  }
  if (ambiguousMarks.size > 0) {
    issues.push(
      `${ambiguousMarks.size} mark(s) stand for more than one service in the deck: `
      + `${[...ambiguousMarks].slice(0, 3).map(([m, n]) => `${JSON.stringify(m)} for ${n} names`).join('; ')}`,
    );
  }
  // AND THE CONVERSE. A row promising a mark the reader cannot find anywhere on
  // the drawing sends them hunting for a shape that does not exist, which is
  // worse than a row admitting the name was never drawn - "(not drawn)" at
  // least tells the truth. The exporter reported the label its sizing arrived
  // at rather than the one it emitted, and an iconed tile below the naming
  // floor draws no text at all, so eight rows named eight marks and eight tiles
  // carried none of them.
  const drawnMarks = new Set(
    markShapes
      .filter((s) => s.name.startsWith('service-label-') || s.name.startsWith('zone-label-'))
      .map((s) => s.text.trim())
      .filter(Boolean),
  );
  const phantomMarks = [...definedMarks.keys()]
    .filter((mark) => mark !== UNLABELLED_ROW && !drawnMarks.has(mark));
  if (phantomMarks.length > 0) {
    issues.push(
      `${phantomMarks.length} index row(s) define a mark that is drawn on no shape in the deck, so the `
      + `reader is sent looking for something that is not there: `
      + `${phantomMarks.slice(0, 3).map((m) => JSON.stringify(m)).join('; ')}`,
    );
  }
  // The NAME half of an index row is the last copy of the wording in the file.
  // A mark is a cut string by design, so the generic truncation rule cannot look
  // for an ellipsis in these rows; this asks the stronger question instead - the
  // half after the separator must BE an authored name, character for character.
  // An index that abbreviates is an index that recovers nothing.
  // Every authored label in the diagram, zones included: a zone whose band
  // cannot hold its caption is spelled out on this same index, so a set built
  // from services alone reported every zone row as a cut name.
  const allAuthored = new Set(
    scenario.nodes
      .map((n) => drawnForm(String((n.data as { label?: string } | undefined)?.label ?? '')))
      .filter(Boolean),
  );
  const cutIndexNames = indexRowText
    .filter((row) => row.includes('  =  '))
    .map((row) => row.slice(row.indexOf('  =  ') + 5).trim())
    .filter((name) => name !== '' && !allAuthored.has(name));
  if (cutIndexNames.length > 0) {
    issues.push(
      `${cutIndexNames.length} index row(s) do not spell a service name out in full, so the one place `
      + `the reader can recover the wording has lost it too: `
      + `${cutIndexNames.slice(0, 3).map((n) => JSON.stringify(n)).join('; ')}`,
    );
  }
  // THE INDEX PAGE MUST BE MEASURED, NOT ONLY READ. Every rule above this one
  // reads the index's TEXT - is the mark defined, is the name spelled in full -
  // and not one of them looks at where the row is DRAWN. So the index sized its
  // type at a hard-coded 10pt, drew each row with `wrap="none"`, and PowerPoint
  // did what `wrap="none"` means: it painted the row at its natural width,
  // straight off the right edge of the sheet. Nothing clips it and nothing
  // warns; the characters past the edge are simply not on the page. The corpus
  // came within 0.040in of the margin, which is one long service name away.
  //
  // This is the same blind spot the Visio index had in round 60 and the same
  // one page 2 had before it: a page that some rule reads is not a page any
  // rule MEASURES.
  const rowOverruns: string[] = [];
  const rowOverflows: string[] = [];
  for (const row of shapes.filter((s) => s.name.startsWith('index-name-'))) {
    if (!row.text.trim()) continue;
    const pt = row.fontSize ?? 10;
    const column = Math.max(0.05, row.w - row.insetX);
    if (row.wrapNone) {
      // Unwrapped ink takes its natural width from the box's left edge,
      // whatever the box says, so the box width is not the bound - the sheet
      // is.
      const ink = measuredTextWidthIn(row.text, pt);
      const past = (row.x + row.insetX / 2 + ink) - (pageW - 0.05);
      if (past > 0.005) {
        rowOverruns.push(
          `${row.name} draws ${past.toFixed(3)}in past the right edge at ${pt.toFixed(1)}pt `
          + `(${row.text.length} chars): ${JSON.stringify(row.text.slice(0, 40))}`,
        );
      }
    } else {
      // A wrapping row must be given the height its lines take, or the row
      // below is printed over it.
      const lines = measuredWrappedLines(row.text, column, pt);
      const need = lines * pt * 1.45 / 72;
      if (need > row.h + 0.01) {
        rowOverflows.push(
          `${row.name} wraps to ${lines} line(s) needing ${need.toFixed(3)}in in a `
          + `${row.h.toFixed(3)}in row`,
        );
      }
    }
  }
  if (rowOverruns.length > 0) {
    issues.push(
      `${rowOverruns.length} index row(s) are drawn off the right edge of the slide, so the one `
      + `page that defines the drawing's marks loses its own text: ${rowOverruns.slice(0, 3).join('; ')}`,
    );
  }
  if (rowOverflows.length > 0) {
    issues.push(
      `${rowOverflows.length} index row(s) are taller than the row they are given, so each is `
      + `printed over the row beneath it: ${rowOverflows.slice(0, 3).join('; ')}`,
    );
  }
  // AND THE ROW MUST NOT WRAP WHEN IT DID NOT HAVE TO. Wrapping stops a row
  // being lost, but a lookup row broken across lines is still the wrong answer
  // when a smaller size in the range the index already allows would have kept
  // it whole: the reader scans this page by running down the left column, and
  // a wrapped row puts the next mark two lines below where the eye expects it
  // while halving how many pairs the page holds. So the type shrinks FIRST and
  // wrapping is the last resort, not the first.
  const avoidablyWrapped: string[] = [];
  for (const row of shapes.filter((s) => s.name.startsWith('index-name-'))) {
    if (!row.text.trim() || row.wrapNone) continue;
    const pt = row.fontSize ?? 10;
    const column = Math.max(0.05, row.w - row.insetX);
    if (measuredWrappedLines(row.text, column, pt) <= 1) continue;
    // The smallest size this page is allowed to set. Anything below it is a
    // different defect - the legibility floor - and is not proposed here.
    const floorPt = 7;
    if (pt <= floorPt + 0.01) continue;
    if (measuredWrappedLines(row.text, column, floorPt) <= 1) {
      avoidablyWrapped.push(
        `${row.name} wraps at ${pt.toFixed(1)}pt but fits one line at ${floorPt}pt `
        + `(${row.text.length} chars in ${column.toFixed(2)}in)`,
      );
    }
  }
  if (avoidablyWrapped.length > 0) {
    issues.push(
      `${avoidablyWrapped.length} index row(s) are broken across lines at a size the page did not `
      + `have to use, so the lookup column the reader scans is interrupted for nothing: `
      + `${avoidablyWrapped.slice(0, 3).join('; ')}`,
    );
  }
  // AND THE SHRINK MUST HAVE BOUGHT SOMETHING. The rule above says the type
  // shrinks before a row is allowed to wrap; this one says the opposite thing
  // and is the reason that one is safe. The fit loop stepped down half a point
  // at a time and also stopped at the 7pt floor, so a name too long for a
  // full-width column at 10pt walked all the way to the bottom of the range and
  // wrapped there anyway - the reader was handed the smallest type the deck can
  // set AND the wrapping the shrink existed to prevent. Nothing was bought and
  // three points of legibility were spent, on the one page a reader opens
  // precisely because a mark on the drawing meant nothing to them.
  //
  // Stated as a property of the printed page rather than of the loop: below the
  // maximum size, no row may wrap. Shrinking is permitted only where it ends
  // the wrapping outright.
  const pointlessShrink: string[] = [];
  for (const row of shapes.filter((s) => s.name.startsWith('index-name-'))) {
    if (!row.text.trim() || row.wrapNone) continue;
    const pt = row.fontSize ?? 10;
    if (pt >= 10 - 0.01) continue;
    const column = Math.max(0.05, row.w - row.insetX);
    const lines = measuredWrappedLines(row.text, column, pt);
    if (lines <= 1) continue;
    pointlessShrink.push(
      `${row.name} is set at ${pt.toFixed(1)}pt and still wraps to ${lines} lines `
      + `(${row.text.length} chars in ${column.toFixed(2)}in)`,
    );
  }
  if (pointlessShrink.length > 0) {
    issues.push(
      `${pointlessShrink.length} index row(s) pay for a shrink that bought nothing: the page dropped `
      + `below 10pt and the rows wrapped regardless, so the reader gets the smallest type the deck `
      + `sets and the broken column too: ${pointlessShrink.slice(0, 3).join('; ')}`,
    );
  }
  // AND AGAINST THE PAGE'S OTHER FURNITURE, not only against itself. The three
  // rules above measure a row against its own box and against the sheet edge,
  // and between those two the row walked into the footer: the grid was fitted
  // to the space below `listTop` and then placed 0.24in lower to clear the
  // note that heads the list, so the last row of every column overhung the
  // footer band. Seven shipped scenarios were 0.0403in into it, and once the
  // pitch became a variable it reached 0.1309in - two thirds of a row.
  const footerBound: string[] = [];
  for (const slideShapes of perSlide) {
    const foot = slideShapes.find((s) => s.text.includes('Generated by Microsoft Product'));
    if (!foot) continue;
    for (const row of slideShapes.filter((s) => s.name.startsWith('index-name-'))) {
      if (!row.text.trim()) continue;
      const into = (row.y + row.h) - foot.y;
      if (into > 0.02) {
        footerBound.push(
          `${row.name} overhangs the footer by ${into.toFixed(4)}in `
          + `(${((into / Math.max(row.h, 0.001)) * 100).toFixed(0)}% of its ${row.h.toFixed(4)}in row)`,
        );
      }
    }
  }
  if (footerBound.length > 0) {
    issues.push(
      `${footerBound.length} index row(s) are drawn into the footer band, which is painted last and `
      + `over them: ${footerBound.slice(0, 3).join('; ')}`,
    );
  }
  // AND THE ROW MUST BE THE SIZE OF ITS OWN TEXT. A row given a pitch it does
  // not use is not a drawing defect the reader can see on that row - it is a
  // page they are handed for nothing. One 300-character name gave all 44 other
  // rows a 0.2819in pitch for 0.1410in of ink and bought a third index page
  // for a list that fits on two, and at five columns the same arithmetic
  // strands four fifths of the sheet. Measured as the fraction of each row's
  // box its own lines actually fill, so a genuinely tall row is not accused.
  const slackRows: string[] = [];
  for (const row of shapes.filter((s) => s.name.startsWith('index-name-'))) {
    if (!row.text.trim() || row.h <= 0) continue;
    const pt = row.fontSize ?? 10;
    const column = Math.max(0.05, row.w - row.insetX);
    const ink = measuredWrappedLines(row.text, column, pt) * pt * 1.45 / 72;
    if (ink < row.h * 0.75 - 0.005) {
      slackRows.push(
        `${row.name} fills ${ink.toFixed(4)}in of a ${row.h.toFixed(4)}in row `
        + `(${((ink / row.h) * 100).toFixed(0)}%)`,
      );
    }
  }
  if (slackRows.length > 0) {
    issues.push(
      `${slackRows.length} index row(s) are given a pitch their own text does not use, so the page `
      + `holds a fraction of the pairs it has room for: ${slackRows.slice(0, 3).join('; ')}`,
    );
  }
  for (const [name, count] of countByName(badges)) {
    if (count > 1) issues.push(`step badge "${name}" is drawn ${count} times`);
  }
  // Splitting that buys nothing. The planner is allowed to spend slides to make
  // tiles bigger, but a window is drawn through a transform capped at natural
  // size, so once a tile is already as large as it was authored, splitting
  // again cannot enlarge it — it only moves the same ink onto more pages. Sixty
  // services authored 20px tall came out as sixty-one slides carrying one tile
  // each, on a page 0.3% inked, with tiles no wider and type no larger than the
  // twenty-five slides they needed. Bar the shape of that: a window slide alone
  // with its tile while that tile is already at natural width.
  const authoredWidths = new Map(
    scenario.nodes
      .filter((n) => n.type === 'azureNode')
      .map((n) => [
        auditStrip(String(n.id)),
        n.width ?? (n.style?.width as number | undefined) ?? 150,
      ]),
  );
  let lonelyWindows = 0;
  // "As large as this can get" is the renderer's ceiling, not natural width.
  // The rule used to count a lone tile drawn at 99% of its authored size, and
  // on a drawing of sub-19.2px tiles that is simply wrong: the renderer lifts
  // such a tile above natural to reach the markable bar, so a lone tile at
  // 1.03x natural still has 1.37x available and splitting further genuinely
  // does enlarge it. Measuring against natural width made the rule fire on
  // exactly the decks that had further to go, and stay silent on the deck that
  // had spent 24 slides to draw nothing.
  const authoredNarrowest = Math.min(
    ...[...authoredWidths.values()].filter((w) => w > 0),
    Infinity,
  );
  const ceilingInPerPx = Math.max(1 / 96, Number.isFinite(authoredNarrowest) && authoredNarrowest > 0
    ? 0.2 / authoredNarrowest
    : 0);
  for (const slideShapes of perSlide) {
    const slideTiles = slideShapes.filter(
      (s) => s.name.startsWith('service-')
        && !s.name.startsWith('service-label-')
        && !s.name.startsWith('service-meta-')
        && s.w > 0,
    );
    if (slideTiles.length !== 1) continue;
    const authoredPx = authoredWidths.get(slideTiles[0].name.replace(/^service-/, '')) ?? 150;
    if (slideTiles[0].w >= authoredPx * ceilingInPerPx * 0.99) lonelyWindows += 1;
  }
  // One lone tile is ordinary — a drawing whose last window holds the remainder
  // will have exactly that. A deck built out of them is the defect.
  if (lonelyWindows > 2 && lonelyWindows > slideCount * 0.5) {
    issues.push(
      `${lonelyWindows} of ${slideCount} slides carry a single service tile already at natural `
      + `width — the deck split past the point where splitting can enlarge anything`,
    );
  }
  // A deck of anonymous dots that promises otherwise.
  //
  // Every continuation slide is captioned "this architecture needs more than
  // one readable slide, so it continues across N of them". The gate measured
  // whether the tiles were large enough, whether the names that WERE drawn
  // fitted, and whether the index carried the rest - but never whether the
  // slides the caption was written for drew a single character. A 26 slide
  // deck with 120 tiles, none of them carrying any type and all 60 index rows
  // reading "(not drawn)", passed clean, because each individual rule was
  // satisfied by a drawing with nothing on it.
  //
  // The bar is 30% of DISTINCT services, both parts measured rather than
  // chosen. Instrumenting all 125 pptx runs put 120 of them at 100% named and
  // the remaining five at 88.9, 75, 70, 62.5 and 33.3, while two decks built
  // to defeat a zero-threshold rule sat at 20.0 and 1.7 - an empty gap, and 30
  // leaves the lowest legitimate deck passing on the strength of its index.
  // Distinct services rather than label shapes because one service drawn on
  // twelve windows is one name to a reader, and counting instances inflates
  // with the very split the rule exists to judge.
  const namedServices = new Set(perSlide.flat()
    .filter((s) => s.name.startsWith('service-label-') && s.text.trim().length > 0)
    .map((s) => s.name.slice('service-label-'.length)));
  const drawnServices = new Set(perSlide.flat()
    .filter((s) => s.name.startsWith('service-')
      && !s.name.startsWith('service-label-')
      && !s.name.startsWith('service-meta-')
      && s.w > 0)
    .map((s) => s.name.replace(/^service-/, '')));
  const namedShare = drawnServices.size > 0 ? namedServices.size / drawnServices.size : 1;
  if (slideCount > 2 && drawnServices.size > 0 && namedShare < 0.3) {
    issues.push(
      `${slideCount} slides draw ${drawnServices.size} service(s) and name only ${namedServices.size} `
      + `of them — ${(namedShare * 100).toFixed(1)}% named on a deck captioned as readable`,
    );
  }
  for (const [name, count] of countByName(chips)) {
    if (count > 1) issues.push(`edge chip "${name}" is drawn ${count} times`);
  }
  // Every narrated step must reach the deck, however long the workflow is:
  // rows that stop shrinking have to continue onto another slide, not vanish.
  const narratedRows = shapes.filter((s) => s.name.startsWith('workflow-text-')).length;
  if (narrated.size > 0 && narratedRows < narrated.size) {
    issues.push(`${narratedRows} workflow rows drawn for ${narrated.size} narrated steps`);
  }

  // The strip symptom is a SHAPE, not an area. Area fill barely moves when a
  // drawing is stretched, because the page is sized from the drawing and both
  // the numerator and the denominator shrink together — a one-rank-per-service
  // strip still measures 3-4% full, so an area rule is silent exactly when it
  // matters. Aspect ratio is what actually changes, and WRAP_TRIGGER_RATIO is
  // the number the product itself uses to decide a layout needs folding.
  //
  // A deck that had to be banded is exempt from *how it was drawn*, but not
  // from the layout itself: measuring per-slide tile bounds meant a wrap
  // regression escaped the moment the strip grew long enough to be split, so
  // the shape is measured on the layout the engine produced, in its own
  // coordinates, whatever the exporter then did with it.
  const tileArea = tiles.reduce((sum, tile) => sum + tile.w * tile.h, 0);
  const density = tileArea / Math.max(pageW * pageH * slideCount, 1);
  const laidOut = scenario.nodes.filter((n) => n.type === 'azureNode');
  if (scenario.fromLayoutEngine && laidOut.length >= 4) {
    // React Flow keeps a child's position relative to its parent, and every
    // service in a grouped scenario has a `parentNode`. Reading raw positions
    // measured the union of intra-zone offsets, so moving zones — which is
    // exactly what wrapping does — changed nothing the rule could see, and it
    // was simultaneously blind to a 72:1 grouped strip and noisy on a healthy
    // wrapped layout. Resolve the parent chain first.
    const byId = new Map(scenario.nodes.map((n) => [n.id, n]));
    const absolute = (node: (typeof scenario.nodes)[number]): { x: number; y: number } => {
      let x = node.position.x;
      let y = node.position.y;
      const seen = new Set<string>([node.id]);
      let parent = node.parentNode ? byId.get(node.parentNode) : undefined;
      while (parent && !seen.has(parent.id)) {
        seen.add(parent.id);
        x += parent.position.x;
        y += parent.position.y;
        parent = parent.parentNode ? byId.get(parent.parentNode) : undefined;
      }
      return { x, y };
    };
    const points = laidOut.map((n) => ({
      ...absolute(n),
      w: n.width ?? (n.style?.width as number | undefined) ?? 150,
      h: n.height ?? (n.style?.height as number | undefined) ?? 75,
    }));
    const minX = Math.min(...points.map((p) => p.x));
    const maxX = Math.max(...points.map((p) => p.x + p.w));
    const minY = Math.min(...points.map((p) => p.y));
    const maxY = Math.max(...points.map((p) => p.y + p.h));
    const aspect = (maxX - minX) / Math.max(maxY - minY, 1);
    if (aspect > WRAP_TRIGGER_RATIO) {
      issues.push(`layout is ${aspect.toFixed(1)}:1 — it was stretched into a strip`);
    }
  }

  return {
    scenario: scenario.id,
    format: 'pptx',
    issues,
    drawnNames: (() => {
      const byId = new Map(scenario.nodes.map((n) => [auditStrip(String(n.id)), String(n.data?.label ?? '')]));
      // The LONGEST rendition of each name wins. A tiled deck draws the same
      // tile twice - small on the overview, in full on its reading slide - and
      // scoring the overview's stub against Visio's full name would report
      // every tiled deck as a divergence.
      const best = new Map<string, string>();
      for (const shape of allSlides.flatMap((slideXml) => parseShapes(slideXml))) {
        if (!shape.name.startsWith('service-label-') || shape.text.trim() === '') continue;
        const id = shape.name.slice('service-label-'.length);
        const authored = byId.get(id) ?? id;
        const drawn = shape.text.trim();
        if ((best.get(authored)?.length ?? -1) < drawn.length) best.set(authored, drawn);
      }
      // The index slide names the service too, and that is the whole reason it
      // exists: "Names shortened on the drawing, in full."
      //
      // Harvesting only `service-label-*` shapes made the cross-format rule
      // claim a service was "on no shape at all in the PowerPoint deck" while
      // its name sat in full on the index. The deck's recovery route for a tile
      // too small to caption is the index; Visio has no index, which is exactly
      // why the sheet needed a rule of its own. Ignoring the index here does not
      // make the rule stricter, it makes it wrong - and it fired on one of the
      // five shortened names in a fixture where all five were indexed.
      // Scoped to the INDEX SLIDE and to SERVICES. Harvesting all text, or all
      // nodes, let zone and boundary captions - which are drawn by a different
      // shape on the sheet and were never in this comparison - register as
      // service names, and the rule reported 565 phantom divergences.
      const INDEX_MARKER = 'Names shortened on the drawing, in full.';
      // An index row is now "<mark>  =  <full name>", the same pair the Visio
      // index prints, so this matches by SUFFIX rather than by set membership -
      // and by suffix rather than by splitting on the separator, because a name
      // may contain the separator and an undrawable name leaves an empty mark.
      const indexRows = allSlides
        .filter((slideXml) => slideXml.includes(INDEX_MARKER))
        .flatMap((slideXml) => parseShapes(slideXml))
        .map((s) => s.text.trim());
      const indexed = {
        has: (name: string): boolean => indexRows.some(
          (row) => row === name || row.endsWith(`  =  ${name}`) || row.endsWith(`=  ${name}`),
        ),
      };
      // Normalised the way the EXPORTERS normalise, by calling the function they
      // call. The raw label is not what either file contains: collectExportBoxes
      // runs every name through singleLineName, so a label with a newline, a
      // double space or a trailing space is stored one way and compared another,
      // and the deck side of this comparison silently matched nothing at all.
      // Copying the regex here would work until the day someone changed it in
      // one place.
      const serviceNames = new Set(
        scenario.nodes
          .filter((n) => n.type !== 'groupNode')
          .map((n) => singleLineName(String(n.data?.label ?? ''))),
      );
      for (const authored of serviceNames) {
        if (!authored || !indexed.has(authored)) continue;
        if ((best.get(authored)?.length ?? -1) < authored.length) best.set(authored, authored);
      }
      return [...best]
        .map(([authored, drawn]) => ({ authored, drawn }))
        .sort((a, b) => a.authored.localeCompare(b.authored));
    })(),
    metrics: {
      slides: slideCount,
      shapes: shapes.length,
      tiles: tiles.length,
      pageWidthIn: +pageW.toFixed(3),
      pageHeightIn: +pageH.toFixed(3),
      minTileWidthIn: +minTileW.toFixed(3),
      minFontPt: minFont,
      overviewMinFontPt: overviewMinFont,
      overviewEmptyTiles,
      gluedConnectors: native.glued,
      unglueableConnectors: native.ungluable,
      shapeGroups: native.groups,
      chips: chips.length,
      maxChipWidthIn: chips.length ? +Math.max(...chips.map((c) => c.w)).toFixed(3) : 0,
      stepBadges: badges.length,
      fillPct: +(density * 100).toFixed(2),
    },
  };
}

async function auditVsdx(scenario: Scenario): Promise<Report> {
  // The drawing a user receives, not the one Node happens to be able to build.
  // Rasterisation needs a DOM, so every icon silently resolved to nothing and
  // the package shipped with no media and no page relationships — meaning the
  // icon wiring, which is exactly what "the icons are missing" was about, had
  // never once been measured.
  const iconPaths = new Set<string>();
  for (const node of scenario.nodes) {
    const path = (node.data as { iconPath?: string } | undefined)?.iconPath;
    if (path) iconPaths.add(path);
  }
  const icons = synthesisedIcons(scenario);
  const pkg = await buildVsdxPackage(scenario.nodes, scenario.edges, 'Contoso Platform', icons);
  const issues: string[] = [];
  issues.push(...xmlWellFormednessIssues(
    pkg.parts
      .filter((p) => typeof p.data === 'string' && /\.(xml|rels)$/i.test(p.path))
      .map((p) => ({ path: p.path, text: p.data as string })),
    '',
  ));
  const pagePart = pkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  // The index lives on its own PAGE now, so it is a different part. Every rule
  // below measures the DRAWING and must keep seeing page 1 alone - the index is
  // read only where the question is "can a reader find this name anywhere".
  const indexPart = pkg.parts.find((p) => /page2\.xml$/i.test(p.path));
  const indexXml = typeof indexPart?.data === 'string' ? indexPart.data : '';
  const media = pkg.parts.filter((p) => /\/media\//i.test(p.path));
  const serviceCount = scenario.nodes.filter((n) => n.type !== 'groupNode').length;
  if (iconPaths.size > 0 && media.length === 0) {
    issues.push(`no embedded icon media parts (expected ~${serviceCount})`);
  }
  // Counting the payload is not counting the picture. Media parts and their
  // relationships are pushed unconditionally, but the `<Rel>` that puts one on
  // the sheet is emitted only inside `iconChild`, so a drawing can ship 700
  // rasters and 700 relationships that no shape references and satisfy the rule
  // above while showing not one icon. That is exactly what happened below sheet
  // scale 0.5504, on two scenarios that were passing. Count the shapes.
  const pageXmlForIcons = typeof pagePart?.data === 'string' ? pagePart.data : '';
  const drawnIcons = (pageXmlForIcons.match(/NameU="Icon\.\d+"/g) ?? []).length;
  // Only tiles that had room, measured the way the exporter measures them.
  //
  // `serviceGroupXml`'s icon arithmetic is fully proportional, so the height at
  // which an icon stops fitting is scale-invariant and depends only on the
  // authored height: solving `0.78125h - 0.19h - 0.16h >= 0.08h` puts the
  // threshold at 0.43in, or 41.28px. Writing the standard tile height there
  // instead left a 34px band — 45% of a standard tile — in which the exporter
  // draws icons and no rule watched them, so the icon scaling could break for
  // every node in it and the gate would sleep through it. The fallback matters
  // as much as the number: the exporter reads `height ?? style.height ??
  // DEFAULT_SERVICE_H` (`diagramExportGeometry.ts:170`), so reading only
  // `height` made the rule fire on a correct sheet whose nodes carry their size
  // on `style` and are rightly too short for an icon.
  const ICON_MIN_PX = 0.43 * 96;
  const wantIcons = scenario.nodes.filter((n) => {
    const styled = (n.style as { height?: number } | undefined)?.height;
    const styledW = (n.style as { width?: number } | undefined)?.width;
    return n.type !== 'groupNode'
      && Boolean((n.data as { iconPath?: string } | undefined)?.iconPath)
      && (n.height ?? styled ?? 75) >= ICON_MIN_PX
      // Width as well as height, for the same reason height is here at all.
      //
      // The test asked only whether a node was TALL enough to host an icon,
      // which counts a 12px-wide glyph as owing one. It cannot carry one at any
      // scale: the exporter sizes the square by `min(h * 0.42, w * 0.34, ...)`
      // and drops it below `0.08 * px` as unreadable, so on a 0.0732in tile the
      // icon comes out 0.0249in against a 0.0468in floor and the words are kept
      // instead - the correct call, reported as a missing icon. The authored
      // aspect is the author's, not the exporter's, and this rule exists to
      // catch icons LOST in export, not icons that never had room.
      && (n.width ?? styledW ?? 75) >= ICON_MIN_PX;
  }).length;
  if (wantIcons > 0 && drawnIcons < wantIcons) {
    issues.push(`${wantIcons - drawnIcons} of ${wantIcons} service icon(s) are embedded but never drawn on the sheet`);
  }
  // Ink has to stay in proportion to the shape it outlines. Every LineWeight on
  // the sheet used to be a literal while the geometry around it scaled, so on a
  // deeply reduced drawing the border stopped being an outline and became the
  // tile: measured at 900 stages, a 0.0125in pen was a quarter of the tile's
  // height and a connector stroke was over half the gap it crossed. The reader
  // sees a grey mat, and zooming in does not recover it because the weight is
  // in the file rather than in the view.
  //
  // A sixteenth of the typical tile is the bar: a border at that width already
  // takes an eighth of the tile's vertical extent once both edges are counted.
  // Measured on `over-row-700` at 16.2% sheet scale, the flat literal is 9.9% of
  // the tile and a scaled pen is 2.8% — the bar sits between them with room on
  // both sides, and at natural size the literal is 1.6% and never near it.
  //
  // Typical, not shortest, and for the same reason the window planner stopped
  // sizing itself from the shortest tile: one authored 12px node makes every
  // pen on an otherwise ordinary sheet look ten times too heavy, and a rule
  // that fires there says nothing about the sheet the reader is holding.
  //
  // Pens are read per shape rather than in bulk. The legend and workflow band
  // are page furniture drawn at natural size whatever the drawing is reduced
  // to, so their pen is in the right proportion to them and measuring it
  // against a reduced tile is a category error — and they cannot be told apart
  // by weight, since the furniture's 0.01in is lighter than the drawing's
  // 0.0125in.
  const DRAWING_SHAPES = /^(Service|Tile|Zone|Connector|StepBadge)$/;
  const tileHeights: number[] = [];
  const drawingPens: number[] = [];
  for (const chunk of pageXmlForIcons.split('<Shape ').slice(1)) {
    const kind = /^[^>]*NameU="([A-Za-z]+)\./.exec(chunk)?.[1];
    if (!kind || !DRAWING_SHAPES.test(kind)) continue;
    if (kind === 'Tile') {
      const h = Number(/<Cell N="Height" V="([\d.]+)"/.exec(chunk)?.[1]);
      if (h > 0) tileHeights.push(h);
    }
    const w = Number(/<Cell N="LineWeight" V="([\d.]+)"/.exec(chunk)?.[1]);
    if (w > 0) drawingPens.push(w);
  }
  if (tileHeights.length > 0 && drawingPens.length > 0) {
    tileHeights.sort((a, b) => a - b);
    const typicalTile = tileHeights[Math.floor(tileHeights.length * 0.5)];
    const heaviest = Math.max(...drawingPens);
    if (heaviest > typicalTile / 16) {
      issues.push(
        `line weight ${heaviest.toFixed(4)}in is ${(heaviest / typicalTile * 100).toFixed(1)}% of the `
        + `typical tile (${typicalTile.toFixed(3)}in) — the outline has become the shape`,
      );
    }
  }
  // A drawing that names a relationship it does not ship is a drawing Visio
  // refuses to open. Neither half was ever checked, because under Node there
  // were no relationships to check.
  const relsPart = pkg.parts.find((p) => /page1\.xml\.rels$/i.test(p.path));
  const relsXml = typeof relsPart?.data === 'string' ? relsPart.data : '';
  const pageXmlForRels = typeof pagePart?.data === 'string' ? pagePart.data : '';
  const referenced = new Set([...pageXmlForRels.matchAll(/r:id="([^"]+)"/g)].map((m) => m[1]));
  if (referenced.size > 0 && relsXml === '') {
    issues.push(`${referenced.size} icon relationship(s) referenced but no page1.xml.rels part was written`);
  }
  const declared = new Set([...relsXml.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  for (const id of referenced) {
    if (!declared.has(id)) issues.push(`icon relationship "${id}" is used on the page but never declared`);
  }
  const targets = [...relsXml.matchAll(/Target="\.\.\/media\/([^"]+)"/g)].map((m) => m[1]);
  const shipped = new Set(media.map((p) => p.path.replace(/^.*\/media\//, '')));
  for (const target of targets) {
    if (!shipped.has(target)) issues.push(`icon relationship points at media/${target}, which is not in the package`);
  }
  const xml = typeof pagePart?.data === 'string' ? pagePart.data : '';
  // Visio text contrast. The Visio path carried its own hard-coded colours and
  // was never measured — the PowerPoint deck had a contrast rule, this one did
  // not, so a fix applied to one exporter could silently miss the other.
  // Each `<Shape>` fragment carries its own fill or inherits the enclosing
  // group's, and every character `Row` in it names a text colour.
  const hex6 = /^#[0-9a-fA-F]{6}$/;
  const seenVsdxContrast = new Set<string>();
  // Visio nests: a service group carries the label text, but the fill that text
  // is read against lives on the group's child tile, and a step badge is a flat
  // sibling with a fill of its own. So a shape's backdrop is its own fill, else
  // the first one among its descendants, else its ancestors', else the white
  // page. Attributing fills by document order instead reads a badge's dark disc
  // as the backdrop of whatever was drawn next.
  type Frame = { name: string; fill?: string; runs: { color: string; size: number }[] };
  const stack: Frame[] = [];
  const drawn: { name: string; fill: string; color: string; size: number }[] = [];
  const tokenRe = /<Shape\s[^>]*?(\/?)>|<\/Shape>|<Cell N="FillForegnd" V="(#[0-9a-fA-F]{6})"\/>|<Cell N="Color" V="(#[0-9a-fA-F]{6})"\/><Cell N="Size" V="([\d.]+)"/g;
  const closeFrame = (): void => {
    const frame = stack.pop();
    if (!frame) return;
    let fill = frame.fill;
    for (let i = stack.length - 1; i >= 0 && !fill; i -= 1) fill = stack[i].fill;
    for (const run of frame.runs) drawn.push({ name: frame.name, fill: fill ?? '#FFFFFF', ...run });
    const parent = stack[stack.length - 1];
    if (parent && !parent.fill && frame.fill) parent.fill = frame.fill;
  };
  for (const token of xml.matchAll(tokenRe)) {
    const [text, selfClosing, fillHex, colorHex, sizeIn] = token;
    if (text.startsWith('</Shape')) { closeFrame(); continue; }
    if (text.startsWith('<Shape')) {
      stack.push({ name: /NameU="([^"]*)"/.exec(text)?.[1] ?? 'shape', runs: [] });
      if (selfClosing === '/') closeFrame();
      continue;
    }
    const top = stack[stack.length - 1];
    if (!top) continue;
    if (fillHex) top.fill = fillHex;
    else if (colorHex) top.runs.push({ color: colorHex, size: +sizeIn });
  }
  while (stack.length > 0) closeFrame();
  for (const run of drawn) {
    if (!hex6.test(run.color) || !hex6.test(run.fill)) continue;
    const ratio = contrastRatio(run.color.slice(1), run.fill.slice(1));
    // Visio font sizes are inches; 18pt = 0.25in is the WCAG large-text bar.
    const bar = run.size >= 0.25 ? 3 : 4.5;
    if (ratio >= bar) continue;
    const key = `${run.color}|${run.fill}|${bar}`;
    if (seenVsdxContrast.has(key)) continue;
    seenVsdxContrast.add(key);
    issues.push(`${run.name} draws ${run.color} text on ${run.fill} — contrast ${ratio.toFixed(2)}:1, below the ${bar}:1 WCAG AA bar`);
  }
  const textCount = (xml.match(/<Text>/g) ?? []).length;
  if (textCount < serviceCount) issues.push(`only ${textCount} text blocks for ${serviceCount} services`);
  // Visio refuses pages larger than 200" on a side.
  if (pkg.pageWidthIn > 200 || pkg.pageHeightIn > 200) {
    issues.push(`page ${pkg.pageWidthIn.toFixed(0)}x${pkg.pageHeightIn.toFixed(0)}in exceeds Visio's 200in limit`);
  }
  // Visio draws 1 : 1, so the sheet is the drawing plus its margins and the
  // workflow band — it can never legitimately be much larger than the drawing
  // it carries. Parking two opposite strays by translating them as one body
  // took an 8.4in architecture onto a 199in sheet: 4% of the page width, i.e.
  // invisible at "fit to window", and 0.68in short of the file being rejected
  // outright. The 200in rule above only catches the very end of that range;
  // this catches the whole class, and unlike that rule it has no constant to
  // tune — the bar is the drawing itself.
  const span = drawingSpanIn(scenario);
  // Scaled by whatever `magnifiedForCallouts` did, because that growth is not
  // outlier growth either.
  //
  // A sheet cannot split, so the only way it can reach a tile wide enough to
  // carry its callout is to be printed bigger, and the exporter now does that
  // for a numbered drawing authored small. Replicated here rather than
  // measured off the page, because the whole point of this rule is that it
  // reads the AUTHORED drawing: a parked stray is drawn where the sheet says
  // it is, so a drawn span would grow with the fault and go blind on it.
  const calloutK = calloutMagnificationFor(scenario.nodes, scenario.edges).k;
  // The numbered workflow gets its own band across the top of the sheet, the
  // colour key gets a strip at the bottom, and the sheet has a minimum size;
  // none of that is outlier growth. Read from the panels the exporter drew
  // rather than modelled from a row pitch — rows are as tall as their sentences
  // — plus the slack the reservation is allowed to miss by, which the "band
  // sits on the drawing it describes" rule below is what actually bounds.
  const drawnBand = /NameU="Workflow\.\d+"[\s\S]*?<Cell N="Height" V="([\d.-]+)"\/>/.exec(xml);
  const drawnLegend = /NameU="Legend\.\d+"[\s\S]*?<Cell N="Height" V="([\d.-]+)"\/>/.exec(xml);
  const BAND_RESERVE_SLACK_IN = 1.2;
  const bandIn = (drawnBand ? +drawnBand[1] + 0.24 + BAND_RESERVE_SLACK_IN : 0)
    + (drawnLegend ? +drawnLegend[1] + 0.45 : 0);
  const allowedW = Math.max(11, span.w * calloutK + PAGE_CHROME_SLACK_IN);
  const allowedH = Math.max(8.5, span.h * calloutK + PAGE_CHROME_SLACK_IN + bandIn);
  if (pkg.pageWidthIn > allowedW) {
    issues.push(`Visio sheet is ${pkg.pageWidthIn.toFixed(1)}in wide for a drawing that spans ${(span.w * calloutK).toFixed(1)}in — trimming outliers must shrink the sheet, never grow it`);
  }
  if (pkg.pageHeightIn > allowedH) {
    issues.push(`Visio sheet is ${pkg.pageHeightIn.toFixed(1)}in tall for a drawing that spans ${(span.h * calloutK).toFixed(1)}in — trimming outliers must shrink the sheet, never grow it`);
  }

  // Every service group must sit on the page, or Visio simply shows nothing
  // where the user expects a service.
  const groupRe = /<Shape ID="(\d+)" NameU="Service\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g;
  let offPage = 0;
  let match: RegExpExecArray | null;
  let minFontIn = 1;
  while ((match = groupRe.exec(xml)) !== null) {
    const [, , pinX, pinY, w, h] = match;
    const left = +pinX - +w / 2;
    const bottom = +pinY - +h / 2;
    if (left < -0.01 || bottom < -0.01 || left + +w > pkg.pageWidthIn + 0.01 || bottom + +h > pkg.pageHeightIn + 0.01) {
      offPage += 1;
    }
  }
  if (offPage > 0) issues.push(`${offPage} service shape(s) sit outside the ${pkg.pageWidthIn.toFixed(1)}x${pkg.pageHeightIn.toFixed(1)}in page`);

  // A zone is a claim about who is inside it, and the exporter is free to move
  // shapes — trimming parks strays, and an overlapping band gets clipped. Both
  // can silently leave a service outside the boundary that owns it, and on the
  // sheet that is not a cosmetic slip: the reader is told the service is out of
  // scope. Membership here is what the author declared, never what happens to
  // overlap, so a compliance band drawn across the drawing is not read as
  // owning everything it crosses.
  const rectOf = (block: string): { x: number; y: number; w: number; h: number } | null => {
    const geo = /<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/.exec(block);
    if (!geo) return null;
    const [, px, py, w, h] = geo;
    return { x: +px - +w / 2, y: +py - +h / 2, w: +w, h: +h };
  };
  const namedRects = (prefix: string): Map<string, { x: number; y: number; w: number; h: number }> => {
    const out = new Map<string, { x: number; y: number; w: number; h: number }>();
    for (const m of xml.matchAll(new RegExp(`<Shape [^>]*NameU="${prefix}\\.\\d+"[^>]*Name="([^"]*)"[\\s\\S]*?<\\/Shape>`, 'g'))) {
      const rect = rectOf(m[0]);
      if (rect && !out.has(m[1])) out.set(m[1], rect);
    }
    return out;
  };
  const escAttr = (value: string): string => (value || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  const zoneRects = namedRects('Zone');
  const serviceRects = namedRects('Service');
  // Shapes are identified on the sheet by label, so only labels that name
  // exactly one shape can be checked without guessing which one moved.
  const labelUses = new Map<string, number>();
  for (const node of scenario.nodes) {
    const label = escAttr(String(node.data?.label ?? ''));
    labelUses.set(label, (labelUses.get(label) ?? 0) + 1);
  }
  for (const node of scenario.nodes) {
    if (!node.parentNode) continue;
    const zone = scenario.nodes.find((n) => n.id === node.parentNode);
    const zoneLabel = escAttr(String(zone?.data?.label ?? ''));
    const ownLabel = escAttr(String(node.data?.label ?? ''));
    if (labelUses.get(zoneLabel) !== 1 || labelUses.get(ownLabel) !== 1) continue;
    const zoneRect = zoneRects.get(zoneLabel);
    const own = serviceRects.get(ownLabel);
    if (!zoneRect || !own) continue;
    const cx = own.x + own.w / 2;
    const cy = own.y + own.h / 2;
    if (cx < zoneRect.x - 0.02 || cx > zoneRect.x + zoneRect.w + 0.02
      || cy < zoneRect.y - 0.02 || cy > zoneRect.y + zoneRect.h + 0.02) {
      issues.push(`service "${String(node.data?.label ?? node.id)}" is drawn outside the "${String(zone?.data?.label ?? node.parentNode)}" zone it belongs to`);
    }
  }

  // A shape reduced to a hairline is worse than one drawn too big, because the
  // reader cannot see that anything is missing — the band is simply gone, and
  // so is whatever it said. Checked against the drawing the author made as well
  // as against an absolute floor, since a zone scaled down with everything else
  // is fine and one scaled down on its own is a bug the page size hides.
  for (const node of scenario.nodes) {
    if (node.type !== 'groupNode') continue;
    const label = escAttr(String(node.data?.label ?? ''));
    if (labelUses.get(label) !== 1) continue;
    const rect = zoneRects.get(label);
    if (!rect) continue;
    const drawnW = Number(node.style?.width ?? node.width ?? 0);
    const drawnH = Number(node.style?.height ?? node.height ?? 0);
    // Absolute, plus a scale check for the one case where scale is knowable.
    //
    // Compaction legitimately shrinks a zone: a compliance band drawn across a
    // whole architecture loses whatever empty space was closed underneath it,
    // and comparing its proportions to the author's reports that as damage. But
    // a zone with no service standing inside it has nothing underneath it to
    // close, so its size on the sheet is fully determined — every service tile
    // is drawn at the same 150px, so one of them gives the sheet's scale, and
    // the band has to be exactly that many inches wide. This is the corridor
    // label between two regions, and it is the shape a void-closing bug
    // destroys, because it is by construction standing in the emptiest part of
    // the drawing.
    //
    // "Nothing underneath it" is per-axis, because the usual annotation is both
    // at once. A sovereign caption stretched over an architecture holds no
    // service, yet every gap between the clusters it covers lies within its
    // width and is closed under it — so its exported width is not determined
    // and demanding the author's is demanding the void back. Its height is
    // determined, because on that axis it genuinely stands clear.
    const spansOn = (
      pos: (n: Node) => number, size: (n: Node) => number, zoneAt: number, zoneSize: number,
    ): boolean => scenario.nodes.some((other) => {
      if (other === node || other.type === 'groupNode') return false;
      const over = Math.min(pos(other) + size(other), zoneAt + zoneSize) - Math.max(pos(other), zoneAt);
      return over > size(other) / 2;
    });
    const nx = Number(node.position?.x ?? 0);
    const ny = Number(node.position?.y ?? 0);
    const spansX = spansOn((n) => Number(n.position?.x ?? 0), (n) => Number(n.width ?? 150), nx, drawnW);
    const spansY = spansOn((n) => Number(n.position?.y ?? 0), (n) => Number(n.height ?? 75), ny, drawnH);
    const holdsAny = scenario.nodes.some((other) => {
      if (other === node || other.type === 'groupNode') return false;
      const ox = Number(other.position?.x ?? 0) + Number(other.width ?? 150) / 2;
      const oy = Number(other.position?.y ?? 0) + Number(other.height ?? 75) / 2;
      if (other.parentNode) return other.parentNode === node.id;
      return ox >= nx && ox <= nx + drawnW && oy >= ny && oy <= ny + drawnH;
    });
    const tileW = serviceRects.size > 0 ? Math.max(...[...serviceRects.values()].map((r) => r.w)) : 0;
    const scale = tileW > 0 ? tileW / 150 : 0;
    const measurable = !holdsAny && scale > 0 && drawnW > 0 && drawnH > 0;
    const starved = measurable
      && ((!spansX && rect.w < 0.6 * drawnW * scale) || (!spansY && rect.h < 0.6 * drawnH * scale));
    if (rect.w < 0.05 || rect.h < 0.05 || starved) {
      issues.push(`zone "${String(node.data?.label ?? node.id)}" is exported ${rect.w.toFixed(3)}x${rect.h.toFixed(3)}in for a ${drawnW}x${drawnH} box the sheet draws at ${(drawnW * scale).toFixed(3)}x${(drawnH * scale).toFixed(3)}in — a shape flattened to a line is a shape deleted`);
    }
  }

  for (const size of xml.matchAll(/<Cell N="Size" V="([\d.]+)"\/>/g)) {
    minFontIn = Math.min(minFontIn, +size[1]);
  }
  const minFontPt = +(minFontIn * 72).toFixed(2);
  // A drawing wider than 127 tiles is over Visio's 200in ceiling with its
  // shapes already touching, so it is scaled down bodily and no absolute point
  // size is attainable: 7pt type on a tile shrunk to a third of an inch is not
  // legible, it is three times wider than its own box and printed over the
  // neighbours. What the sheet owes the reader there is proportion — type
  // shrunk no harder than the drawing was, so zooming in recovers it.
  //
  // Measure the scale from the sheet rather than recomputing the exporter's
  // arithmetic, so the two cannot drift: a service tile is 150px = 1.5625in
  // when nothing has been given up.
  const tileWidths = [...xml.matchAll(/NameU="Service\.\d+"[\s\S]*?<Cell N="Width" V="([\d.]+)"/g)]
    .map((m) => +m[1]);
  const sheetScale = tileWidths.length > 0
    ? Math.min(1, Math.max(...tileWidths) / (150 / PX_PER_IN))
    : 1;
  const floorPt = sheetScale >= 0.999 ? 7 : 7 * sheetScale;
  // The floor is PowerPoint's, deliberately. Both exporters draw the same
  // drawing at the same scale, so type that is unreadable in the deck is
  // unreadable on the sheet, and the two must not disagree about where the
  // limit is.
  if (minFontPt < floorPt - 0.01) {
    issues.push(sheetScale >= 0.999
      ? `smallest Visio font is ${minFontPt}pt (below the 7pt floor the deck enforces)`
      : `smallest Visio font is ${minFontPt}pt on a sheet scaled to ${(sheetScale * 100).toFixed(0)}% `
        + `— type shrunk harder than the drawing it labels (floor ${floorPt.toFixed(2)}pt)`);
  }

  // The name a Visio tile draws has to fit the tile it names.
  //
  // Visio does not clip text to its text block - that is the premise the
  // connector chip was fixed on - so a text block sized one line short does
  // not truncate, it paints the surplus lines out through the bottom of the
  // shape and across whatever is under it. Every rule the sheet had for this
  // read `TxtHeight` back out of the file, which is the exporter's own answer
  // to the question and cannot disagree with the exporter about anything.
  //
  // This measures the sentence instead, the way the connector rule has since
  // the painted-ink round, and it is the rule the sheet's line counter needed:
  // `wrapOneLineIn` broke an over-wide word with `ceil(w / column)`, the third
  // copy of that defect in the repo, and NOTHING in this file could see it -
  // the mutation survived all 96 files of the corpus.
  // Scanned by shape CHUNK, not by one regex across the page. A Service shape
  // is a Visio group and its icon is a child shape with a `<Text/>` of its
  // own, so a lazy `[\s\S]*?<Text>` walks straight past the group's own text
  // into the child's empty one and the rule silently measures nothing. It read
  // as working because the only tile it had ever fired on was icon-less.
  // Splitting on the shape tag ends each chunk exactly where its first child
  // begins.
  // A rule that has silently measured NOTHING twice does not get to report
  // zero on trust. Every skip above is a `continue` with no record, so the
  // difference between "all clean" and "the scan matched nothing" is invisible
  // - which is exactly how the child-`<Text>` capture and the missing `<cp>`
  // marker both read as passing. Count what was reached and say so.
  let serviceChunks = 0;
  let measuredChunks = 0;
  for (const chunk of xml.split('<Shape ID=')) {
    if (!/NameU="Service\.\d+"/.test(chunk.slice(0, 200))) continue;
    serviceChunks += 1;
    const label = /Name="([^"]*)"/.exec(chunk)?.[1] ?? '';
    const cell = (name: string): number => {
      const hit = new RegExp(`<Cell N="${name}" V="([\\d.-]+)"`).exec(chunk);
      return hit ? +hit[1] : 0;
    };
    const tileH = cell('Height');
    const txtW = cell('TxtWidth');
    const txtH = cell('TxtHeight');
    const txtPinY = cell('TxtPinY');
    // Each run is drawn at its OWN size: the name and the sub-line are
    // separate character rows, and measuring both at the name's size reports a
    // block that is neither.
    //
    // Split on the row tag rather than spanning rows with a lazy quantifier.
    // That is the same defect as the child-`<Text>` capture above, one Section
    // down: a Row that happens to omit `Size` would let the scan walk into the
    // next Row and report the WRONG run's size, silently. Every Row emits one
    // today, which is exactly the kind of premise that fails without a sound.
    const characterSection = /<Section N="Character">([\s\S]*?)<\/Section>/.exec(chunk)?.[1] ?? '';
    const sizes = characterSection.split('<Row IX=').slice(1).reduce<Record<string, number>>(
      (acc, row) => {
        const ix = /^"(\d+)"/.exec(row)?.[1];
        const size = /<Cell N="Size" V="([\d.]+)"/.exec(row)?.[1];
        return ix && size ? { ...acc, [ix]: +size } : acc;
      },
      {},
    );
    const textBody = unescapeXml(/<Text>([\s\S]*?)<\/Text>/.exec(chunk)?.[1] ?? '');
    // A tile with no sub-line emits its name as bare text, with no `<cp>` run
    // marker at all - so a rule that keys on the marker measures nothing on
    // exactly the tiles that carry only a name.
    const marked = [...textBody.matchAll(/<cp IX="(\d+)"\/>([\s\S]*?)(?=<cp IX="\d+"\/>|$)/g)]
      .map((run) => ({ ix: run[1], text: run[2] }));
    const runs = marked.length > 0 ? marked : [{ ix: '0', text: textBody }];
    if (runs.length === 0 || txtW <= 0 || tileH <= 0) continue;
    let inkH = 0;
    let lineCount = 0;
    let widestGlyph = 0;
    runs.forEach((run, i) => {
      const fontIn = sizes[run.ix] ?? 0;
      const text = run.text.replace(/<[^>]*>/g, '').trim();
      if (!text || fontIn <= 0) return;
      // The sub-line's own multiple, which is looser than the name's.
      const multiple = i === 0 ? 1.3 : 1.4;
      const lines = text.split('\n').reduce(
        (sum, para) => sum + measuredWrappedLines(para, txtW, fontIn * 72),
        0,
      );
      // The horizontal failure, which neither `over` nor `short` can see. A
      // long WORD is not it - Visio breaks a word that outgrows its column
      // between glyphs, which is what the packing counter models. A single
      // GLYPH wider than the column is: there is no break inside it, so the
      // renderer centres it and paints it out through both sides.
      for (const glyph of text) {
        if (/\s/.test(glyph)) continue;
        widestGlyph = Math.max(widestGlyph, measuredTextWidthIn(glyph, fontIn * 72));
      }
      lineCount += lines;
      inkH += lines * fontIn * multiple;
    });
    if (inkH <= 0) continue;
    measuredChunks += 1;
    if (widestGlyph > txtW + 0.01) {
      issues.push(
        `Visio tile "${label}" draws a glyph ${widestGlyph.toFixed(3)}in wide in a `
        + `${txtW.toFixed(3)}in column - no break fits inside a glyph, so it paints out both sides`,
      );
    }
    // Two failures, both fatal, and the second is the one a shape can hide.
    //
    // The block is centred on `TxtPinY`, so the ink runs half its height
    // either side of it and anything past the shape is painted on the drawing.
    // But a tall tile has headroom, and there an under-counted block does not
    // escape the shape at all - it just paints over the icon and the tile
    // chrome inside it, which is why the `ceil(width / column)` mutation
    // survived the whole corpus. So the declared block is checked against the
    // ink as well: `TxtHeight` is the exporter's own answer to "how tall is
    // this text?", and a text block shorter than its own text is wrong
    // wherever it sits.
    const over = Math.max(txtPinY + inkH / 2 - tileH, inkH / 2 - txtPinY);
    const short = inkH - txtH;
    if (over > 0.01 || short > 0.02) {
      const drawn = textBody.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      issues.push(
        `Visio tile "${label}" draws "${drawn.slice(0, 28)}" in a ${txtW.toFixed(3)}in column `
        + `- ${lineCount} line(s) need ${inkH.toFixed(3)}in, the block is declared `
        + `${txtH.toFixed(3)}in`
        + (short > 0.02 ? ` (${short.toFixed(3)}in short of its own text)` : '')
        + (over > 0.01
          ? `, and ${over.toFixed(3)}in is painted outside a ${tileH.toFixed(3)}in shape`
          : ` on a ${tileH.toFixed(3)}in shape`),
      );
    }
  }
  // A named tile whose text this rule could not reach is a hole in the rule,
  // not a clean tile. `drawnNames` is built by a separate scan, so it is an
  // independent count of how many tiles really carry text.
  const namedTiles = new Set(
    [...xml.matchAll(/NameU="Service\.\d+" Name="([^"]*)"[\s\S]*?<Text>([\s\S]*?)<\/Text>/g)]
      .filter((m) => m[2].replace(/<[^>]*>/g, '').trim() !== '')
      .map((m) => m[1]),
  ).size;
  if (measuredChunks < namedTiles) {
    issues.push(
      `Visio containment rule measured ${measuredChunks} of ${serviceChunks} service shape(s) `
      + `but ${namedTiles} draw text - ${namedTiles - measuredChunks} named tile(s) went unchecked`,
    );
  }

  // The legibility floor both exporters clamp to, in points.
  const VISIO_FLOOR_PT = 7;

  // The coverage oracle, which for its whole life ran on the deck alone.
  //
  // The deck and the sheet share one width model, so a character neither of
  // them can measure is a guess in BOTH - but the rule that says so was inside
  // `auditPptx` and had no counterpart here. That is the wrong way round. When
  // a name is too wide for its column the deck cuts it and keeps the rest on
  // the index slide, while Visio draws nothing at all and the name survives
  // only in a Name= attribute that is metadata and never printed. So the
  // format with no recovery path was the format with no oracle: a Cyrillic
  // fixture that emptied every tile on the sheet reported PASS here while the
  // deck, which had at least drawn something, reported the guess.
  const unmeasuredInk = new Map<string, number>();
  for (const match of xml.matchAll(/<Text>([\s\S]*?)<\/Text>/g)) {
    // By cluster, for the reason the deck's copy of this rule gives.
    for (const cluster of auditClusters(unescapeXml(match[1].replace(/<[^>]*>/g, '')))) {
      if (cluster.measured) continue;
      unmeasuredInk.set(cluster.text, (unmeasuredInk.get(cluster.text) ?? 0) + 1);
    }
  }
  if (unmeasuredInk.size > 0) {
    const worst = [...unmeasuredInk.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6);
    issues.push(
      `the sheet draws ${unmeasuredInk.size} character(s) with no measured advance, so every `
      + `width and wrap that touches them is a guess: `
      + worst.map(([character, n]) => `U+${(character.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')} x${n}`).join(', '),
    );
  }

  // A name the sheet refused to draw is not recoverable the way a cut name is.
  //
  // Visio has one page. There is no index behind it, so when `drawsName` says
  // no the authored name exists nowhere a reader can see - it is left in the
  // shape's Name attribute, which is a handle for automation and is never put
  // on paper. The deck's version of this rule can afford to ask only whether
  // the name FITS; this one has to ask whether it was DROPPED.
  for (const chunk of xml.split('<Shape ID=')) {
    const head = chunk.slice(0, 400);
    if (!/NameU="Service\.\d+"/.test(head)) continue;
    const named = /\sName="([^"]*)"/.exec(head);
    const authored = named ? unescapeXml(named[1]) : '';
    if (!authored) continue;
    const drawn = [...chunk.matchAll(/<Text>([\s\S]*?)<\/Text>/g)]
      .map((m) => unescapeXml(m[1].replace(/<[^>]*>/g, '')).trim())
      .join('');
    if (drawn !== '') continue;
    const widthCell = /<Cell N="TxtWidth" V="([-\d.eE]+)"/.exec(chunk);
    const column = widthCell ? Number(widthCell[1]) : 0;
    if (!(column > 0)) continue;
    // What the exporter would actually put in the shape: one real character
    // and an ellipsis. Asking this of the WHOLE authored name is the wrong
    // question - the mean advance of thirty glyphs is cheaper than the
    // ellipsis, so the rule claimed a name was available on a 0.104in column
    // that holds one ellipsis and nothing else. A tile drawing only "..."
    // carries no information, and the exporter is right to leave it blank.
    const shortest = `${[...authored][0] ?? ''}\u2026`;
    if (measuredDrawableInColumn(shortest, VISIO_FLOOR_PT, column)) {
      issues.push(
        `Visio tile "${authored}" draws no text at all, and the name is nowhere on the page - `
        + `its ${column.toFixed(4)}in column holds the widest glyph `
        + `(${measuredWidestGlyphIn(authored, VISIO_FLOOR_PT).toFixed(4)}in) and two typical ones at `
        + `the ${VISIO_FLOOR_PT}pt floor, so a truncated name was available and was not drawn`,
      );
    }
  }

  // Two shapes, not one. Every other rule in this file measures INSIDE a block -
  // line counts against TxtHeight, containment against the shape - so a picture
  // printed across the name it labels is invisible to all of them.
  //
  // The tile positions its text band `0.06in` above its floor and its icon
  // `0.07in` below its ceiling. Both were flat while every dimension around
  // them scaled, so when the icon is room-limited the two blocks overlap by
  // exactly `0.13 - 0.19 * scale` inches: below scale 0.6842 the icon is drawn
  // ON TOP OF the name. A 260-service pipeline scales to 0.436 and overlapped
  // on 100% of its tiles; at 440 the icon was wholly inside the text block on
  // 340 of them. A grid never triggers it - the page grows instead - so only a
  // wide shallow drawing reaches it, which is the shape of a pipeline.
  const overlaps: string[] = [];
  // The icon is a CHILD shape, so splitting on `<Shape ID=` puts it in its OWN
  // chunk, never in the parent's. The first version of this rule looked for
  // `NameU="Icon."` inside the service chunk and skipped when it was absent -
  // which was ALWAYS - so the rule ran on every corpus and could not fire once.
  // Shapes arrive in document order as Service, Tile, Icon, so pair each tile
  // with the next icon before the next service.
  const vsdxChunks = xml.split('<Shape ID=');
  for (let i = 0; i < vsdxChunks.length; i += 1) {
    const head = vsdxChunks[i].slice(0, 400);
    if (!/NameU="Service\.\d+"/.test(head)) continue;
    const named = /\sName="([^"]*)"/.exec(head);
    const authored = named ? unescapeXml(named[1]) : '';
    const cellIn = (src: string, name: string): number | null => {
      const hit = new RegExp(`<Cell N="${name}" V="([-\\d.eE]+)"`).exec(src);
      return hit ? Number(hit[1]) : null;
    };
    const txtPinY = cellIn(vsdxChunks[i], 'TxtPinY');
    const txtLocPinY = cellIn(vsdxChunks[i], 'TxtLocPinY');
    const txtHeight = cellIn(vsdxChunks[i], 'TxtHeight');
    if (txtPinY === null || txtLocPinY === null || txtHeight === null) continue;
    if (!(txtHeight > 0)) continue;
    let iconChunk: string | null = null;
    for (let j = i + 1; j < vsdxChunks.length; j += 1) {
      const ahead = vsdxChunks[j].slice(0, 400);
      if (/NameU="Service\.\d+"/.test(ahead)) break;
      if (ahead.includes('NameU="Icon.')) { iconChunk = vsdxChunks[j]; break; }
    }
    if (!iconChunk) continue;
    const iconPinY = cellIn(iconChunk, 'PinY');
    const iconHeight = cellIn(iconChunk, 'Height');
    if (iconPinY === null || iconHeight === null) continue;
    // Both are in the parent group's local frame, and in Visio y grows upward,
    // so the band's ceiling is its pin plus what sits above the pin, and the
    // icon's floor is its pin less half its height.
    const bandTop = txtPinY - txtLocPinY + txtHeight;
    const iconBottom = iconPinY - iconHeight / 2;
    const over = bandTop - iconBottom;
    if (over > 1e-6) {
      overlaps.push(`"${authored}" by ${over.toFixed(4)}in (${((over / txtHeight) * 100).toFixed(0)}% of its band)`);
    }
  }
  if (overlaps.length > 0) {
    issues.push(
      `${overlaps.length} Visio service tile(s) draw the icon on top of the name: `
      + `${overlaps.slice(0, 3).join('; ')}`,
    );
  }

  // FURNITURE IS OPAQUE AND FURNITURE IS DRAWN LAST, so anything it lands on is
  // gone. This has now happened twice for the same reason: the legend was
  // pinned to the bottom-left corner and painted over the tiles until it was
  // given a reserved strip, and the service-name panel repeated it exactly -
  // solid white fill, emitted after every service, and "reserved" only in
  // `furnitureRects`, which keeps CONNECTOR LABELS off it and has never had
  // anything to do with where the tiles go. On a 48-tile grid it covered 20
  // tiles completely.
  //
  // Every rule before this one measured a shape against itself or against its
  // own children, so a shape sitting on top of an unrelated shape was invisible
  // to all of them. This is the general form: a service tile may not intersect
  // any piece of page furniture, whichever piece it is.
  const furnitureRect: Array<{ name: string; x0: number; y0: number; x1: number; y1: number }> = [];
  const tileRect: Array<{ name: string; x0: number; y0: number; x1: number; y1: number }> = [];
  for (const chunk of xml.split('<Shape ID=')) {
    const head = chunk.slice(0, 400);
    const nameU = /NameU="([^"]+)"/.exec(head)?.[1] ?? '';
    const isTile = /^Service\.\d+$/.test(nameU);
    const isFurniture = /^(Legend|Workflow|ServiceNames)\./.test(nameU);
    if (!isTile && !isFurniture) continue;
    const num = (cellName: string): number | null => {
      const hit = new RegExp(`<Cell N="${cellName}" V="([-\\d.eE]+)"`).exec(chunk);
      return hit ? Number(hit[1]) : null;
    };
    const pinX = num('PinX'); const pinY = num('PinY');
    const w = num('Width'); const h = num('Height');
    const locX = num('LocPinX') ?? (w === null ? null : w / 2);
    const locY = num('LocPinY') ?? (h === null ? null : h / 2);
    if (pinX === null || pinY === null || w === null || h === null || locX === null || locY === null) continue;
    const rect = { name: nameU, x0: pinX - locX, y0: pinY - locY, x1: pinX - locX + w, y1: pinY - locY + h };
    (isTile ? tileRect : furnitureRect).push(rect);
  }
  const buriedTiles: string[] = [];
  for (const tile of tileRect) {
    for (const furniture of furnitureRect) {
      const ox = Math.min(tile.x1, furniture.x1) - Math.max(tile.x0, furniture.x0);
      const oy = Math.min(tile.y1, furniture.y1) - Math.max(tile.y0, furniture.y0);
      if (ox <= 1e-6 || oy <= 1e-6) continue;
      const tileArea = Math.max(1e-9, (tile.x1 - tile.x0) * (tile.y1 - tile.y0));
      buriedTiles.push(`${tile.name} is ${((ox * oy / tileArea) * 100).toFixed(0)}% under ${furniture.name}`);
      break;
    }
  }
  if (buriedTiles.length > 0) {
    issues.push(
      `${buriedTiles.length} Visio service tile(s) are drawn underneath opaque page furniture, which is `
      + `emitted last and so paints over them: ${buriedTiles.slice(0, 3).join('; ')}`,
    );
  }

  // A COLUMN TOO NARROW TO READ, asked TWO ways.
  //
  // The band picks its column count by minimising stack height and nothing
  // else, so on brief descriptions - where every split shortens the stack - it
  // ran to its 12 column cap and set each sentence in 0.2583in of text column,
  // about two characters wide. Every rule before this one asked whether the
  // text FITS; fitting in two characters is not the same as being readable.
  //
  // The width floor alone is not enough, and neither is the line bound alone,
  // because they fail on different inputs. Twelve one-word sentences shred to a
  // hairline column that each still sets in ONE line, so no line-count bound
  // can see it - only the width can. Twelve long sentences keep a column wide
  // enough to clear any width floor while every one of them wraps to eight or
  // ten lines, which is a stack of ribbons no reader follows - only the line
  // count can see that. So both, and the fixtures below prove neither covers
  // for the other.
  const cramped: string[] = [];
  const shredded: number[] = [];
  const bodies: string[] = [];
  const columnXs = new Set<string>();
  let workflowPt = 0;
  let workflowColW = 0;
  for (const chunk of xml.split('<Shape ID=')) {
    if (!/Name="workflow-text-\d+"/.test(chunk.slice(0, 400))) continue;
    const w = /<Cell N="Width" V="([\d.]+)"/.exec(chunk);
    const text = /<Text>([\s\S]*?)<\/Text>/.exec(chunk);
    if (!w) continue;
    const widthIn = Number(w[1]);
    const pinX = /<Cell N="PinX" V="([\d.]+)"/.exec(chunk);
    if (pinX) columnXs.add(Number(pinX[1]).toFixed(2));
    workflowColW = Math.max(workflowColW, widthIn);
    const body = unescapeXml((text?.[1] ?? '').replace(/<[^>]*>/g, '')).trim();
    if (widthIn < 0.9) {
      cramped.push(`${widthIn.toFixed(4)}in for ${JSON.stringify(body.slice(0, 28))}`);
    }
    // Visio sets type in inches; the size cell is the em height.
    const size = /<Cell N="Size" V="([\d.]+)"/.exec(chunk);
    const pt = size ? Number(size[1]) * 72 : 7.2;
    workflowPt = pt;
    // The width cell IS the text column on the Visio side: the band already
    // subtracted the badge gutter before emitting the shape, and no margin
    // cells are written. Subtracting PowerPoint's 0.1in-a-side inset here
    // charged the sentence a column it never lost and reported four lines for
    // a row the sheet sets in three.
    if (body) {
      shredded.push(measuredWrappedLines(body, Math.max(0.01, widthIn), pt));
      bodies.push(body);
    }
  }
  if (cramped.length > 0) {
    issues.push(
      `${cramped.length} workflow sentence(s) are set in a column too narrow to read: `
      + `${cramped.slice(0, 3).join('; ')}`,
    );
  }
  // The MEDIAN, not the worst. One long sentence among twelve is an author's
  // sentence, not a column defect, and failing on the maximum would report
  // every band that carries one. When HALF the band wraps past three lines the
  // column itself is the problem.
  //
  // And RELATIVE to the unsplit band, not an absolute three. An 800-character
  // sentence takes six lines in the widest column this page has to give, so an
  // absolute bound accuses the exporter of a defect it cannot repair and the
  // only way to satisfy it would be to edit the author's prose. The question
  // a reader actually cares about is whether SPLITTING the band shredded the
  // sentence: the exporter chose the column count, so that is the part it owns.
  if (shredded.length >= 4) {
    const sorted = [...shredded].sort((a, b) => a - b);
    const median = sorted[Math.floor(sorted.length / 2)];
    const cols = Math.max(1, columnXs.size);
    // What ONE column would have given: the band's full width, less the one
    // badge gutter that column still pays - and CAPPED at the 7.5in the
    // exporter caps a single-column panel at, or the comparison credits one
    // column with a width the sheet would never have drawn.
    const unsplitW = (cols > 1 ? Math.min((workflowColW + 0.6) * cols, 7.5) : workflowColW + 0.6) - 0.6;
    const unsplitPer = bodies
      .map((b) => measuredWrappedLines(b, Math.max(0.01, unsplitW), workflowPt));
    const unsplit = [...unsplitPer].sort((a, b) => a - b);
    const floorLines = Math.max(3, unsplit[Math.floor(unsplit.length / 2)]);
    if (median > floorLines) {
      issues.push(
        `the workflow band sets half its ${shredded.length} sentence(s) in ${median} or more `
        + `wrapped lines at ${workflowPt.toFixed(1)}pt, so the ${cols}-column split shreds the `
        + `prose into ribbons a single column would have set in `
        + `${unsplit[Math.floor(unsplit.length / 2)]}: the widest sentence takes `
        + `${sorted[sorted.length - 1]} lines`,
      );
    }
    // AND THE TAIL, because a median cannot see the rows the deck exists to
    // carry. A workflow of 48 one-word acknowledgements and 3 real paragraphs
    // has a median of ONE - the terse steps were never going to wrap at any
    // width, so they are not evidence about the column - while the 3 sentences
    // that carry the architecture were set in 9 lines of a 1.8125in column
    // where one column would have set them in 3. The median said nothing and
    // the band was three times the reviewer's own "four lines is where the eye
    // loses its place".
    //
    // Per ENTRY and relative to that same entry unsplit, so a long sentence is
    // measured against what it would have been rather than against a constant.
    // Twice is the bar: one extra fold is the ordinary price of a column, three
    // times over is a ribbon.
    const TAIL_MULTIPLE = 2;
    const tail: string[] = [];
    shredded.forEach((lines, i) => {
      const alone = unsplitPer[i];
      if (lines > Math.max(3, alone * TAIL_MULTIPLE)) {
        tail.push(
          `${lines} lines against ${alone} in one column: `
          + `${JSON.stringify(bodies[i].slice(0, 36))}`,
        );
      }
    });
    if (cols > 1 && tail.length > 0) {
      issues.push(
        `${tail.length} workflow sentence(s) are folded more than ${TAIL_MULTIPLE}x by the `
        + `${cols}-column split, so the rows the drawing exists to explain are set as ribbons `
        + `while the median stays quiet: ${tail.slice(0, 3).join('; ')}`,
      );
    }
  }

  // EVERY PAGE THE EXPORTER EMITS MUST BE MEASURED BY SOME RULE. Page 2 was
  // shipped, opened cleanly, and looked at by exactly one rule - `namedInIndex`,
  // which reads its TEXT and never its GEOMETRY. So the index sized its column
  // to a constant 3.4in while its rows were the longest strings in the file,
  // wrapped to two lines of 0.135in inside a 0.2in box, and overprinted its
  // neighbour by 35% of the type size. Moving furniture to a new page moved it
  // out of the auditor's field of view, which is the same defect as a rule that
  // cannot fire: a shipped artefact nothing looks at.
  const INDEX_FONT_PT = 7.2;
  const indexSpill: string[] = [];
  const indexRows: Array<{ name: string; left: number; right: number; top: number; bottom: number }> = [];
  for (const chunk of indexXml.split('<Shape ID=')) {
    const named = /Name="(service-name-\d+)"/.exec(chunk.slice(0, 400));
    if (!named) continue;
    const w = /<Cell N="Width" V="([\d.]+)"/.exec(chunk);
    const h = /<Cell N="Height" V="([\d.]+)"/.exec(chunk);
    const pinY = /<Cell N="PinY" V="([-\d.]+)"/.exec(chunk);
    const text = /<Text>([\s\S]*?)<\/Text>/.exec(chunk);
    if (!w || !h || !text) continue;
    const boxW = Number(w[1]);
    const boxH = Number(h[1]);
    const body = unescapeXml(text[1].replace(/<[^>]*>/g, '')).trim();
    const lines = auditWrappedLines(body, boxW, INDEX_FONT_PT);
    const needsIn = lines * (INDEX_FONT_PT / 72) * 1.35;
    if (needsIn > boxH + 1e-6) {
      indexSpill.push(
        `${named[1]} needs ${needsIn.toFixed(4)}in (${lines} lines) in a ${boxH.toFixed(4)}in box`,
      );
    }
    if (pinY) {
      const pinX = /<Cell N="PinX" V="([-\d.]+)"/.exec(chunk);
      const centre = Number(pinY[1]);
      indexRows.push({
        name: named[1],
        left: pinX ? Number(pinX[1]) - boxW / 2 : 0,
        right: pinX ? Number(pinX[1]) + boxW / 2 : boxW,
        top: centre + needsIn / 2,
        bottom: centre - needsIn / 2,
      });
    }
  }
  if (indexSpill.length > 0) {
    issues.push(
      `${indexSpill.length} index row(s) print more lines than their box holds, so they overprint the row `
      + `above or below: ${indexSpill.slice(0, 3).join('; ')}`,
    );
  }
  const collidingRows: string[] = [];
  for (let i = 0; i < indexRows.length; i += 1) {
    for (let j = i + 1; j < indexRows.length; j += 1) {
      const a = indexRows[i];
      const b = indexRows[j];
      // Rows in DIFFERENT COLUMNS share a Y by design, so a bare Y comparison
      // would report the whole index as broken the moment it needs two columns.
      const across = Math.min(a.right, b.right) - Math.max(a.left, b.left);
      if (across <= 1e-6) continue;
      const overlap = Math.min(a.top, b.top) - Math.max(a.bottom, b.bottom);
      if (overlap > 1e-6) {
        collidingRows.push(`${a.name} and ${b.name} overlap by ${overlap.toFixed(4)}in`);
      }
    }
  }
  if (collidingRows.length > 0) {
    issues.push(
      `${collidingRows.length} pair(s) of index rows are drawn through each other: `
      + `${collidingRows.slice(0, 3).join('; ')}`,
    );
  }

  // ASK-60-A: COUNT COLLISIONS, NOT CHARACTERS. A stub is a lookup key into an
  // index that lives on ANOTHER PAGE, so the reader cannot see both at once and
  // the drawing has to be self-consistent alone. "N..." is a perfectly good key
  // until a second tile draws "N..." too, at which point neither row in the
  // index is matchable and the two tiles are indistinguishable. The bar is
  // uniqueness over the strings actually DRAWN, shortened or not - a stub
  // colliding with a name another tile drew in full is the same ambiguity.
  const drawnStrings = new Map<string, Set<string>>();
  for (const chunk of xml.split('<Shape ID=')) {
    const head = chunk.slice(0, 400);
    if (!/NameU="Service\.\d+"/.test(head)) continue;
    const authoredAttr = /NameU="Service\.\d+" Name="([^"]*)"/.exec(head);
    const text = /<Text>([\s\S]*?)<\/Text>/.exec(chunk);
    if (!text) continue;
    const body = unescapeXml(text[1].replace(/<[^>]*>/g, '')).split('\n')[0].trim();
    if (!body) continue;
    const authored = unescapeXml(authoredAttr?.[1] ?? '');
    if (!drawnStrings.has(body)) drawnStrings.set(body, new Set());
    drawnStrings.get(body)!.add(authored);
  }
  // Two tiles of the SAME service drawing the same string is not ambiguity, it
  // is the truth: a diagram with twenty Copilot Studio nodes should say so
  // twenty times. The defect is two DIFFERENTLY-named services collapsing onto
  // one string, which is what makes the index rows unmatchable.
  const ambiguous = [...drawnStrings].filter(([, names]) => names.size > 1);
  if (ambiguous.length > 0) {
    issues.push(
      `${ambiguous.length} string(s) are drawn for more than one differently-named Visio service, so neither `
      + `the tiles nor their index rows can be told apart: `
      + `${ambiguous.slice(0, 3).map(([s, n]) => `${JSON.stringify(s)} for ${n.size} names`).join('; ')}`,
    );
  }

  // The other half of that bargain, and the rule the scaler actually broke: the
  // type has to stay in proportion to the tile it labels. Visio wraps a name
  // inside its shape, so holding the point size fixed while the shape shrinks
  // does not spill it sideways — it forces more and more lines into a text
  // block that is itself shrinking, until the name is clipped to its first
  // syllable and the icon is squeezed out entirely. A tile drawn at 1.5625in
  // carries 0.105in type; that ratio is what "fits" means here, and it must
  // survive any scaling the page limit forces.
  const NATURAL_TILE_IN = 150 / PX_PER_IN;
  const NATURAL_LABEL_IN = 0.105;
  for (const chunk of xml.split('<Shape ID=')) {
    if (!/NameU="Service\.\d+"/.test(chunk.slice(0, 200))) continue;
    const cellOf = (name: string): number => {
      const hit = new RegExp(`<Cell N="${name}" V="([\\d.-]+)"`).exec(chunk);
      return hit ? +hit[1] : 0;
    };
    const label = /Name="([^"]*)"/.exec(chunk)?.[1] ?? '';
    const tileIn = cellOf('Width');
    const tileH = cellOf('Height');
    const fontIn = +(/<Section N="Character">[\s\S]*?<Cell N="Size" V="([\d.]+)"/.exec(chunk)?.[1] ?? 0);
    // The DRAWN text, not the `Name` attribute. The attribute deliberately
    // carries the whole name whatever the tile does with it — that is how a cut
    // name stays findable in Drawing Explorer — so measuring it reported a
    // ratio for every shape including the ones that draw no text at all.
    const drawn = unescapeXml(/<Text>([\s\S]*?)<\/Text>/.exec(chunk)?.[1] ?? '')
      .replace(/<[^>]*>/g, '').trim();
    if (!label || !drawn || tileIn <= 0 || fontIn <= 0) continue;
    const ratio = fontIn / tileIn;
    const natural = NATURAL_LABEL_IN / NATURAL_TILE_IN;
    // The ratio is a proxy for "the name wraps past the room the tile has",
    // and on a tall narrow tile the proxy is simply wrong: it demands the type
    // scale with the WIDTH while the room is the AREA. A 0.78 x 3.13in tile
    // reported 2.0x while its name wrapped to three lines of a shape with room
    // for twenty-nine, and satisfying the proxy there would have meant 3.8pt
    // type — below every legibility floor in this file, so the rule was
    // unsatisfiable rather than strict. Measure the thing the proxy stands for
    // and only fall back on the ratio when the real measurement is unavailable.
    //
    // The column is READ, not guessed. `tileIn - 0.1` is a fixed inset
    // subtracted from a variable tile, so on the 0.23in tiles this rule exists
    // to police it under-stated the column by 43% and reported an 11-line
    // overflow on a name the shape sets in seven.
    const columnIn = cellOf('TxtWidth') > 0 ? cellOf('TxtWidth') : Math.max(0.01, tileIn - 0.1);
    const wrapped = measuredWrappedLines(drawn, columnIn, fontIn * 72);
    const neededIn = wrapped * fontIn * 1.2;
    if (tileH > 0 && neededIn <= tileH * 0.9) continue;
    if (ratio > natural * 1.05) {
      issues.push(`service name "${label}" is set at ${(fontIn * 72).toFixed(2)}pt on a ${tileIn.toFixed(2)}in tile `
        + `— ${(ratio / natural).toFixed(1)}x the type-to-tile ratio the sheet draws at full size, `
        + `so the name wraps to ${wrapped} line(s) needing ${neededIn.toFixed(2)}in of a ${tileH.toFixed(2)}in shape`);
    }
  }

  // The same bargain for the two pieces of drawing furniture that are not
  // tiles. A zone caption and a numbered step badge sit among the tiles and
  // scale with them, and both were invisible to the rule above because it only
  // ever matched `Service.n` — a caption held at natural size on a deeply
  // scaled sheet was 4.4x the service names beside it and overflowed the zone
  // onto the tiles inside it, and a badge held at 0.24in was wider than a
  // whole service.
  for (const zone of xml.matchAll(/NameU="Zone\.\d+" Name="([^"]*)"[\s\S]*?<Cell N="Width" V="([\d.]+)"[\s\S]*?<Cell N="Height" V="([\d.]+)"[\s\S]*?<Cell N="Size" V="([\d.]+)"/g)) {
    const zoneW = +zone[2];
    const zoneH = +zone[3];
    const fontIn = +zone[4];
    if (!zone[1] || zoneW <= 0 || zoneH <= 0 || fontIn <= 0) continue;
    const lines = Math.max(1, Math.ceil(textWidthIn(zone[1], fontIn * 72) / Math.max(zoneW * 0.92, 0.02)));
    const blockIn = lines * fontIn * 1.3;
    if (blockIn > zoneH * 0.6) {
      issues.push(`zone caption "${zone[1]}" needs ${blockIn.toFixed(3)}in of type `
        + `on a ${zoneW.toFixed(3)} x ${zoneH.toFixed(3)}in zone `
        + `— ${((blockIn / zoneH) * 100).toFixed(0)}% of the box it names, so it covers what is inside it`);
    }
  }
  // Both badge rules measure a disc against the two tiles it is drawn BETWEEN,
  // resolved by identity, not by position and not by any statistic.
  //
  // Four attempts got here. Minimum over the sheet let a parked 14px sliver
  // decide that a 0.240in disc between 1.56in tiles was 165% oversized. Median
  // over the sheet failed the parity flip. Median over badged endpoints failed
  // as soon as a sheet carried two numbered chains. Nearest tile by position
  // failed an ordinary three-tier diagram, because a hop is 3in long and any
  // small node dropped in the middle of the page is nearer to the badge than
  // either tile the badge belongs to - it reported the same "165% on a 0.146in
  // tile" for a Private DNS zone node that had nothing to do with the arrow.
  //
  // The identity was in the file the whole time: `stepBadgeXml` writes
  // `Name="step-<edgeId>"` and the scenario names that edge's endpoints.
  const shapeGeom = (nameU: RegExp): Array<{ name: string; w: number }> => {
    const out: Array<{ name: string; w: number }> = [];
    for (const chunk of xml.split('<Shape ID=')) {
      const head = chunk.slice(0, 400);
      if (!nameU.test(head)) continue;
      const w = /<Cell N="Width" V="([\d.]+)"/.exec(chunk);
      // The tile's own node id, off its shape data. Resolving by the visible
      // NAME instead skipped every badge on a sheet where two nodes carry the
      // same service - which is not an edge case, it is the ordinary case,
      // because `box.label` falls back to the service when a node has no
      // distinct label. Measured over the corpus, name matching left 732 of
      // 2258 badges (32.4%) unmeasured and sixteen scenarios at zero, among
      // them `squeezed-badges`, whose entire purpose is badge sizing. The id
      // has been in the file all along, as `<Row N="NodeId">`.
      const id = /<Row N="NodeId">[\s\S]*?<Cell N="Value" V="([^"]*)"/.exec(chunk);
      const name = id ?? /\sName="([^"]*)"/.exec(head);
      if (w) out.push({ name: unescapeXml(name?.[1] ?? ''), w: +w[1] });
    }
    return out;
  };
  const badgeGeom = shapeGeom(/NameU="StepBadge\.\d+"/);
  const tileGeom = shapeGeom(/NameU="Service\.\d+"/);
  const tileByName = new Map<string, number>();
  for (const t of tileGeom) tileByName.set(t.name, t.w);
  const tileFor = (nodeId: string): number | undefined => tileByName.get(nodeId);
  const endsOfEdge = new Map<string, [string, string]>();
  const stepOfEdge = new Map<string, number>();
  for (const edge of scenario.edges) {
    endsOfEdge.set(edge.id, [edge.source, edge.target]);
    const step = (edge as unknown as { data?: { stepNumber?: number | string } }).data?.stepNumber;
    stepOfEdge.set(edge.id, Number(step) || 1);
  }
  // The exporter's own documented floor, `badgeMinDiameterIn`, at the largest
  // font floor it can use. A sheet that has been scaled down uses a smaller
  // font and therefore a smaller floor, so this over-estimates on those - and
  // it is the exemption side, so over-estimating only ever costs a report,
  // never invents one.
  const badgeFloorIn = (stepNumber: number): number => {
    const pt = 0.0973;
    const digits = String(Math.max(1, Math.abs(Math.trunc(stepNumber)))).length;
    return Math.max(pt * 1.15, Math.hypot(digits * pt * 0.66, pt * 0.7) / 0.9);
  };
  // The exporter's `STEP_BADGE_IN * fonts.scale`, recovered from the sheet.
  //
  // Every box on a Visio page shares one scale factor, and `fontsForScale`
  // returns natural fonts at or above 0.999, so `scale = min(1, drawn * 96 /
  // authored)` on any endpoint whose authored width is known. Returns NaN when
  // it is not, which fails the comparison and leaves the rule armed.
  const authoredVisioW = new Map<string, number>();
  for (const node of scenario.nodes) {
    const w = Number((node as unknown as { width?: number }).width);
    if (Number.isFinite(w) && w > 0) authoredVisioW.set(String(node.id), w);
  }
  const naturalBadgeIn = (
    ends: readonly string[],
    widthOf: (id: string) => number | undefined,
  ): number => {
    for (const id of ends) {
      const authored = authoredVisioW.get(id);
      const drawnIn = widthOf(id);
      if (!authored || !drawnIn) continue;
      return 0.24 * Math.min(1, (drawnIn * 96) / authored);
    }
    return NaN;
  };
  if (badgeGeom.length > 0 && tileGeom.length > 0) {
    const ratios: Array<{ ratio: number; floored: boolean }> = [];
    // The same empty intersection PowerPoint reports, which this side had been
    // passing over in silence. Round 73 measured one drawing at 66% in the deck
    // and 156% on the sheet, from identical input, and only the deck said so -
    // and the sheet is the format that cannot split its way out of it.
    const vsdxConflicts: Array<{ tile: number; floor: number; ratio: number }> = [];
    // Visio has no windows to spend, so its way to a tile that can carry a
    // callout is paper: `magnifiedForCallouts` prints the whole drawing bigger
    // until the median tile clears the bar. Exempt only where that ran out of
    // paper at `MAX_USEFUL_PAGE_IN`, which is the sheet's equivalent of the
    // deck's bleed asymptote - past it there is no move left.
    const magnified = calloutMagnificationFor(scenario.nodes, scenario.edges);
    const paperBound = magnified.paperBound;
    for (const badge of badgeGeom) {
      const edgeId = /^step-(.*)$/.exec(badge.name)?.[1];
      const ends = edgeId ? endsOfEdge.get(edgeId) : undefined;
      if (!ends || !edgeId) continue;
      const widths = ends.map(tileFor).filter((w): w is number => !!w && w > 0);
      // BOTH ends or neither. Measuring against the one end that resolved was
      // worse than skipping: on a fan whose four discs are all 77% of their
      // small end, the one hop with a resolvable large end read 7.2% against
      // the tile it is not dwarfing, and was then exempted as floored.
      if (widths.length < ends.length) continue;
      // The narrower end, matching the exporter's own ceiling: a disc that is
      // proportionate to one tile and swamps the other is still swamping one.
      const tile = Math.min(...widths);
      const floor = badgeFloorIn(stepOfEdge.get(edgeId) ?? 1);
      // Two denominators, because the two rules ask different questions.
      //
      // The CEILING asks whether the disc swamps a service, and a disc that is
      // proportionate to one end and swamps the other is still swamping one,
      // so that is the narrower end. The undersize rule asks whether a reader
      // can FIND the disc, and a reader hunts it against the largest thing
      // beside it: on a hub with narrow spokes, 55% of the spoke is 5.5% of
      // the hub, and the same disc is proportionate by one measure and a speck
      // by the other. Measured on `probe-hub-spoke` with 14px spokes, the disc
      // is 76.7% of the spoke and 2.7% of the hub.
      ratios.push({
        ratio: badge.w / Math.max(...widths),
        // Exempt at every value the sizing model can produce, which is all
        // three of `max(minDiameter, min(natural, ceiling))`.
        //
        // On the floor that is the placement search squeezing a disc into a
        // gap so the number sits on its own arrow; on the ceiling there is no
        // larger diameter that leaves the narrow end unswamped; at natural
        // size nothing has happened to the disc at all. Round 73 shipped this
        // with the third missing and `probe-wide-hub` reported its natural
        // 0.24in disc at 7.2% because one other shape on the sheet is 320
        // authored px wide.
        //
        // What is left is a disc cut by something that is not its own hop,
        // which is the fault this rule was written for: a sheet-wide ceiling
        // caps at 0.1375in with a 0.24in natural size and a 0.1119in floor.
        floored: badge.w <= floor + 1e-4
          || badge.w >= tile * 0.55 - 1e-4
          || Math.abs(badge.w - naturalBadgeIn(ends, tileFor)) <= 1e-3,
      });
      // A hair of slack, because the ceiling IS `tile * 0.55`, so a capped
      // badge sits exactly on this bar and the verdict would otherwise be
      // decided by the rounding of the decimals written into the XML -
      // measured, the same drawing fired at 20px, passed at 24px and fired
      // again at 30px, printing 55.0% every time. Both numbers are written to
      // four decimals, so the tile can round down by 5e-5 and the badge up by
      // the same, and 0.2177000 on 0.3958000 cleared a 1e-6 bar by nine
      // millionths of an inch while the true figures were 0.2177083 on
      // 0.3958333 - exactly the ceiling. 2e-4 covers both roundings with room
      // and is four hundred times smaller than the smallest real breach on
      // record.
      //
      // And only where a proportionate disc could also have been a legible
      // one. Under about a 0.20in tile there is no such diameter: 55% of it is
      // narrower than the smallest circle that holds a digit, so the exporter
      // draws the floor and the disc is disproportionate no matter what it
      // chooses. Reporting that as a badge defect points at the wrong thing -
      // the only fix is a bigger tile, which is a planner trade - and it would
      // fire on every hop of such a drawing at once.
      if (tile * 0.55 >= floor - 1e-6 && badge.w > tile * 0.55 + 2e-4) {
        issues.push(`a step badge is ${badge.w.toFixed(3)}in across on a ${tile.toFixed(3)}in tile `
          + `— ${((badge.w / tile) * 100).toFixed(0)}% of the service it is calling out`);
        break;
      }
      // Same rounding slack as the ceiling test above, and for the same reason
      // with more force: `magnifiedForCallouts` aims the median tile at exactly
      // `floor / 0.55`, so a magnified sheet lands ON this bar by construction
      // and the verdict would be decided by the fourth decimal the XML carries.
      // Measured, 20 hops read "needs 0.2270in, permits 0.2270in".
      //
      // Measured against the DRAWN disc and not against `floor` alone, because
      // `floor` here is deliberately an over-estimate - it assumes the largest
      // font the exporter can use, and a sheet scaled down uses a smaller one.
      // That was safe while it only ever exempted; used to report, it invented
      // 380 of them on `scaled-zone-row`, where the discs sit at 16% of their
      // tiles and are not swamping anything. The fault this rule is for is a
      // disc that had to breach proportionality to stay legible, so the drawn
      // breach is the evidence and the floor is only the explanation.
      if (tile * 0.55 < floor - 2e-4 && badge.w > tile * 0.55 + 2e-4 && !paperBound) {
        vsdxConflicts.push({ tile, floor, ratio: badge.w / tile });
      }
    }
    if (vsdxConflicts.length > 0) {
      const worst = vsdxConflicts.reduce((a, b) => (b.ratio > a.ratio ? b : a));
      issues.push(`${vsdxConflicts.length} step callout(s) sit on a tile too small to carry one: `
        + `a callout needs at least ${worst.floor.toFixed(4)}in to stay readable and at most `
        + `${(worst.tile * 0.55).toFixed(4)}in to stay proportionate on its ${worst.tile.toFixed(4)}in `
        + `tile, so it draws at ${(worst.ratio * 100).toFixed(0)}% of the service it is calling out`);
    }
    // And the converse, which nothing measured. A badge that has been cut to a
    // speck beside the tile it numbers is one the reader has to hunt for
    // before the numbered order means anything.
    //
    // Per badge, now that each one is measured against its own endpoints. The
    // sheet-wide form of this test could not see the fault it was written for:
    // when a ceiling binds, the badges beside the tiles that set it are at
    // exactly 55% by construction, so the widest ratio reads 55.0% while the
    // badges on the OTHER chain sit at 8.8% - and numbering six sensors cut an
    // unrelated pipeline's discs by 43% with the gate silent.
    //
    // Badges sitting ON the documented minimum are exempt, because that is the
    // placement search doing its job: squeezing a disc into a gap so it can
    // sit on its own arrow rather than being pushed onto the next hop's. Four
    // ordinary sheets have such a badge. A foreign ceiling does not land on
    // that floor - the two-chain fault caps at 0.1375in, well clear of the
    // 0.1119in floor - so the exemption costs this rule nothing it was for.
    for (const [at, seen] of ratios.entries()) {
      if (seen.ratio >= 0.1 || seen.floored) continue;
      issues.push(`step badge ${at + 1} of ${ratios.length} is only ${(seen.ratio * 100).toFixed(1)}% `
        + 'of the widest tile it sits beside — too small to find');
      break;
    }
  }
  // A callout is a white number on a dark disc, and the disc is the only thing
  // making it readable. Every other piece of type on the sheet is measured
  // against the shape that has to hold it — the service name against its tile,
  // the zone caption against its zone — but the badge was measured only by its
  // diameter against a tile, which says nothing about whether the number fits
  // the disc. Overflow here is not untidy, it is white ink on white paper: the
  // reader sees a dark speck with an invisible smear across it, and the digits,
  // which are the one thing muting a label into the workflow band is supposed
  // to preserve, carry no information at all.
  for (const badge of xml.matchAll(/NameU="StepBadge\.\d+"[\s\S]*?<Cell N="Width" V="([\d.]+)"[\s\S]*?<Cell N="Size" V="([\d.]+)"[\s\S]*?<Text>([^<]*)<\/Text>/g)) {
    const discIn = +badge[1];
    const fontIn = +badge[2];
    const digits = badge[3].trim();
    if (!digits || discIn <= 0 || fontIn <= 0) continue;
    const needIn = textWidthIn(digits, fontIn * 72);
    // On the diagonal. The number is centred in the disc, so it occupies a
    // chord rather than the diameter, and the half-chord at the height of the
    // glyphs is shorter than the radius. Testing width against diameter passes
    // a badge whose first and last digit are outside the circle — which is not
    // untidy, it is white ink on white paper, and the digits are the one thing
    // muting a label into the workflow band is supposed to preserve.
    const diagonalIn = Math.hypot(needIn, fontIn * 0.7);
    // A tenth of the disc is kept as a ring. A badge is a white number on a
    // dark disc and the disc is the only thing making it readable, so digits
    // that run to the edge stop being backed by it — and a disc solved for
    // exactly the number it holds clears a bare containment test by 0.2%, which
    // is not a margin, it is a rounding error. The natural badge sits at 0.60.
    if (diagonalIn > discIn * 0.9) {
      issues.push(`step badge "${digits}" needs ${diagonalIn.toFixed(4)}in across the disc on a `
        + `${discIn.toFixed(4)}in disc — ${(diagonalIn / discIn * 100).toFixed(0)}% of the disc that `
        + `backs it, so the number runs to the rim`);
    }
  }

  // Every sentence the author wrote has to survive somewhere a reader can find
  // it. The sheet drops a label it cannot write anywhere legible and hands the
  // wording to the workflow band, which is a fair trade only for as long as the
  // band actually says it — and this is the one rule that cannot be satisfied
  // by drawing less, because deleting the label is exactly what it checks for.
  //
  // Counted rather than merely looked for, because a drawing repeats its
  // wording: eight parallel hops carrying one sentence are eight sentences the
  // reader has to be able to account for, and a rule that stops at the first
  // surviving copy is passed by muting the other seven.
  const foldVsdx = (s: string): string => s
    .toLowerCase()
    .replace(/[\s\u3000]+/g, '')
    .replace(/[.,;:!?、。（）()[\]「」"'`´’‘“”\-…]/g, '');
  const textOf = (namePrefix: string): string => [
    ...xml.matchAll(new RegExp(`<Shape [^>]*NameU="${namePrefix}\\.\\d+"[\\s\\S]*?<\\/Shape>`, 'g')),
  ]
    .map((m) => unescapeXml(/<Text>([\s\S]*?)<\/Text>/.exec(m[0])?.[1] ?? ''))
    .join('\u0000');
  // Connector text and workflow prose only. A service happening to be named
  // after a verb in somebody's sentence is not that sentence surviving.
  const spoken = foldVsdx(`${textOf('Connector')}\u0000${textOf('LegendText')}`);
  const occurrences = (stem: string): number => {
    if (!stem) return 0;
    let count = 0;
    for (let at = spoken.indexOf(stem); at >= 0; at = spoken.indexOf(stem, at + 1)) count += 1;
    return count;
  };
  const wanted = new Map<string, { need: number; sample: string }>();
  for (const edge of scenario.edges) {
    const label = auditStrip(readEdgeLabel(edge)).trim();
    // Truncation is a different rule's business, so compare on a stem short
    // enough that the exporter is always allowed to keep it.
    const stem = foldVsdx(label).slice(0, 12);
    if (!stem) continue;
    const seen = wanted.get(stem);
    if (seen) seen.need += 1; else wanted.set(stem, { need: 1, sample: label });
  }
  const lost: string[] = [];
  for (const [stem, { need, sample }] of wanted) {
    const found = occurrences(stem);
    if (found < need) lost.push(`"${sample}" x${need - found}`);
  }
  if (lost.length > 0) {
    issues.push(`the Visio sheet has lost connector wording: ${lost.slice(0, 3).join(', ')}`);
  }

  // Visio does not clip a text block — it draws the overflow past both edges,
  // straight through whatever is above and below. The workflow band drew every
  // step in a fixed 0.18in block at a fixed 0.26in pitch, so a sentence that
  // wrapped ran through the row beneath it: a 76-character step (which is
  // ordinary Architecture Center prose) is three lines on an 11in page, and
  // every row in the band overran the next, all the way down. Measured against
  // the box the exporter actually wrote, not against the pitch it intended.
  const workflowRows = [...xml.matchAll(
    /NameU="LegendText\.\d+" Name="workflow-text-(\d+)"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>[\s\S]*?<Cell N="Size" V="([\d.-]+)"\/>[\s\S]*?<Text>([\s\S]*?)<\/Text>/g,
  )].map((m) => ({
    step: m[1],
    x: +m[2],
    y: +m[3],
    w: +m[4],
    h: +m[5],
    pt: +m[6] * 72,
    text: unescapeXml(m[7]),
  }));
  const spilling = workflowRows
    .map((row) => {
      // The exporter's own estimator is deliberately not reused: a guard that
      // imports the estimator agrees with the bug by construction. This is an
      // independent greedy count against the same 1.35 line multiple the rest
      // of the deck is measured with — the 1.22 used before was smaller than
      // the exporter's own multiple, so the guard granted the exporter more
      // room than the exporter granted itself and fired on nothing.
      const lines = auditWrappedLines(row.text.trim(), row.w, row.pt);
      return { row, needed: (lines * row.pt * 1.35) / 72, lines };
    })
    .filter((r) => r.needed > r.row.h + 0.01);
  if (spilling.length > 0) {
    const worst = spilling.sort((a, b) => b.needed / b.row.h - a.needed / a.row.h)[0];
    issues.push(`${spilling.length} Visio workflow row(s) overrun their neighbours — step ${worst.row.step} needs ${worst.needed.toFixed(2)}in (${worst.lines} lines at ${worst.row.pt.toFixed(1)}pt) in a ${worst.row.h.toFixed(2)}in row`);
  }
  // The rows have to be inside the panel that frames them, and the panel has to
  // be the size of its rows: a band reserved larger than its contents steals
  // the page from the drawing just as surely as one drawn too small spills.
  const panel = /NameU="Workflow\.\d+"[\s\S]*?<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/.exec(xml);
  if (panel && workflowRows.length > 0) {
    const panelTop = +panel[1] + +panel[3] / 2;
    const panelBottom = +panel[1] - +panel[3] / 2;
    const outside = workflowRows.filter((r) => r.y + r.h / 2 > panelTop + 0.01 || r.y - r.h / 2 < panelBottom - 0.01);
    if (outside.length > 0) {
      issues.push(`${outside.length} Visio workflow row(s) are drawn outside the ${(+panel[3]).toFixed(2)}in band that frames them, starting at step ${outside[0].step}`);
    }
    const lowest = Math.min(...workflowRows.map((r) => r.y - r.h / 2));
    const dead = lowest - panelBottom;
    if (dead > 0.6) {
      issues.push(`the Visio workflow band reserves ${dead.toFixed(2)}in below its last row — the page it takes has to be the page it uses`);
    }  }

  // Workflow numbering must survive into Visio too, or the same drawing tells
  // a different story in PowerPoint and in Visio. Measured against the repaired
  // edges, which is what both exporters draw from.
  const numberedEdges = narrateEdgeCallouts(scenario.edges).filter(
    (e) => readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber) !== undefined,
  );
  const badgeBlocks = [...xml.matchAll(/<Shape [^>]*NameU="StepBadge\.\d+"[\s\S]*?<\/Shape>/g)].map((m) => m[0]);
  if (badgeBlocks.length !== numberedEdges.length) {
    issues.push(`${badgeBlocks.length} Visio step badges for ${numberedEdges.length} numbered connectors`);
  }
  const expectedNumbers = new Set(
    numberedEdges.map((e) => String(readStepNumber((e.data as { stepNumber?: unknown } | undefined)?.stepNumber))),
  );
  // Service boxes in page coordinates, so a badge that lands on one is caught.
  const serviceBoxes: Array<{ x: number; y: number; w: number; h: number; name?: string }> = [];
  for (const m of xml.matchAll(
    /NameU="(Service\.\d+)"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )) {
    const [, name, pinX, pinY, w, h] = m;
    serviceBoxes.push({ x: +pinX - +w / 2, y: +pinY - +h / 2, w: +w, h: +h, name });
  }

  // The tile's *text*, measured — not the height the tile says it reserved.
  //
  // `serviceBoxes` above reads `Width`/`Height` and feeds one buried-tile rule;
  // nothing here had ever looked at what the tile writes in itself. The band is
  // sized by a clamp, and Visio does not clip text to a text block, so a name
  // too long for the clamp is drawn straight out of the shape — through the
  // icon and off the tile. This is the connector chip's defect on the tile, and
  // it was reachable at the app's own default node size.
  //
  // Two rules, deliberately of different kinds:
  //
  //   A. A line of type cannot be shorter than its own em. That is not a model
  //      of a renderer, it is what "font size" means, so `lines x size` is a
  //      floor no line spacing can argue with. The exporter reserves at 1.3 em
  //      and this asserts at 1.0, which means the guard shares *no* constant
  //      with the thing it is checking — the exporter cannot satisfy it by
  //      agreeing with it.
  //   B. Ink drawn at the exporter's own spacing must still land on the tile.
  //      A is about the declared band; B is about the catastrophe the reader
  //      actually sees, which is words printed over the icon and off the shape.
  const tileTextIssues: string[] = [];
  const tileEscapes: string[] = [];
  for (const m of xml.matchAll(
    /<Shape [^>]*NameU="Service\.\d+"[\s\S]*?<\/Shape>\s*<\/Shapes>\s*<\/Shape>/g,
  )) {
    const shape = m[0];
    const cell = (n: string): number => {
      const hit = new RegExp(`<Cell N="${n}" V="([^"]*)"`).exec(shape);
      return hit ? Number(hit[1]) : NaN;
    };
    const body = unescapeXml(/<Text>([\s\S]*?)<\/Text>/.exec(shape)?.[1] ?? '');
    // Written out rather than reusing the `<Text>([^<]*)` scrape below, which
    // returns '' for any body that opens with a `<cp>` run marker — as every
    // tile carrying a sub-line does.
    const drawn = body
      .replace(/<cp IX="\d+"\/>/g, '')
      .split('\n')[0]
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&#10;/g, '\n').replace(/&amp;/g, '&');
    const sizeIn = Number(
      /<Section N="Character">\s*<Row IX="0">(?:<Cell [^>]*\/>)*?<Cell N="Size" V="([^"]*)"/.exec(shape)?.[1] ?? '0',
    );
    const txtW = cell('TxtWidth');
    const txtH = cell('TxtHeight');
    const txtPinY = cell('TxtPinY');
    const tileH = cell('Height');
    if (!drawn || !(sizeIn > 0) || !(txtW > 0) || Number.isNaN(txtH) || Number.isNaN(txtPinY)) continue;
    const lines = auditWrappedLines(drawn, txtW, sizeIn * 72);
    const emFloor = lines * sizeIn;
    if (emFloor > txtH + 0.005) {
      tileTextIssues.push(
        `"${drawn.slice(0, 24)}" needs ${emFloor.toFixed(3)}in for ${lines} line(s) at ${(sizeIn * 72).toFixed(2)}pt in a ${txtH.toFixed(3)}in band`,
      );
    }
    const inkH = lines * sizeIn * 1.3;
    const overshoot = Math.max(inkH / 2 - txtPinY, txtPinY + inkH / 2 - tileH);
    if (overshoot > 0.01) {
      tileEscapes.push(`"${drawn.slice(0, 24)}" by ${overshoot.toFixed(3)}in`);
    }
  }
  if (tileTextIssues.length > 0) {
    issues.push(
      `${tileTextIssues.length} Visio service tile(s) write more text than the band they declare: ${tileTextIssues.slice(0, 3).join('; ')}`,
    );
  }
  if (tileEscapes.length > 0) {
    issues.push(
      `${tileEscapes.length} Visio service tile(s) draw their name outside the tile: ${tileEscapes.slice(0, 3).join('; ')}`,
    );
  }

  // Glue. A .vsdx whose connectors are not attached to the shapes they join is
  // a picture, not a diagram, and being editable is the whole reason to export
  // Visio: drag a service in an unglued drawing and the arrows stay behind.
  // Two halves have to agree. The `<Connects>` table has to name both ends,
  // and the geometry has to start and finish on the shapes the table names —
  // a line glued to a box it does not touch is snapped across the page the
  // first time Visio reroutes it, and the reader's layout jumps.
  const shapeBoxById = new Map<string, { x: number; y: number; w: number; h: number }>();
  for (const m of xml.matchAll(
    /<Shape ID="(\d+)" NameU="Service\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )) {
    shapeBoxById.set(m[1], { x: +m[2] - +m[4] / 2, y: +m[3] - +m[5] / 2, w: +m[4], h: +m[5] });
  }
  const glue = new Map<string, { begin?: string; end?: string }>();
  for (const m of xml.matchAll(/<Connect FromSheet="(\d+)" FromCell="(BeginX|EndX)"[^>]*ToSheet="(\d+)"/g)) {
    const entry = glue.get(m[1]) ?? {};
    if (m[2] === 'BeginX') entry.begin = m[3]; else entry.end = m[3];
    glue.set(m[1], entry);
  }
  let unglued = 0;
  let detached = 0;
  for (const block of xml.matchAll(/<Shape ID="(\d+)" NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const id = block[1];
    const ends = glue.get(id);
    if (!ends?.begin || !ends?.end) { unglued += 1; continue; }
    const at = (cell: string): number => +(new RegExp(`<Cell N="${cell}" V="([\\d.-]+)"/>`).exec(block[0])?.[1] ?? NaN);
    const pairs: Array<[string, number, number]> = [
      [ends.begin, at('BeginX'), at('BeginY')],
      [ends.end, at('EndX'), at('EndY')],
    ];
    for (const [sheet, x, y] of pairs) {
      const box = shapeBoxById.get(sheet);
      if (!box || !Number.isFinite(x) || !Number.isFinite(y)) continue;
      // A stub jog leaves the endpoint a little outside the tile, so the bar is
      // the gap the exporter itself uses rather than exact containment.
      const gap = Math.max(box.x - x, x - (box.x + box.w), box.y - y, y - (box.y + box.h));
      if (gap > 0.2) detached += 1;
    }
  }
  if (unglued > 0) issues.push(`${unglued} Visio connector(s) are not glued to the shapes they join`);
  if (detached > 0) issues.push(`${detached} Visio connector end(s) are glued to a shape they do not touch`);

  // A connector whose two ends are the same point is not a short arrow, it is
  // no arrow: Visio draws nothing, the relationship is absent from the sheet,
  // and the step number that belongs to it is stranded on whatever tile it
  // landed on. It happens when the fit squeezes two tiles flush, so the hop
  // between them runs from a shared edge to itself. The bar is an arrowhead,
  // because a line shorter than its own head cannot show a direction either.
  let tooShort = 0;
  for (const block of xml.matchAll(/<Shape ID="\d+" NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const at = (cell: string): number => +(new RegExp(`<Cell N="${cell}" V="([\\d.-]+)"/>`).exec(block[0])?.[1] ?? NaN);
    const dx = at('EndX') - at('BeginX');
    const dy = at('EndY') - at('BeginY');
    if (!Number.isFinite(dx) || !Number.isFinite(dy)) continue;
    if (Math.hypot(dx, dy) < 0.04) tooShort += 1;
  }
  if (tooShort > 0) {
    issues.push(`${tooShort} Visio connector(s) are shorter than an arrowhead and draw nothing`);
  }

  // Arrows must not be drawn through services. PowerPoint has had this rule for
  // several rounds; Visio shares the router but had no geometry rule of any
  // kind, so a routing regression could ship in the .vsdx while the deck stayed
  // clean. Geometry rows are in the connector's own rotated frame, measured
  // from its begin point, so they are carried back to the page before judging.
  for (const block of xml.matchAll(/<Shape ID="\d+" NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const shape = block[0];
    const num = (cell: string): number => +(new RegExp(`<Cell N="${cell}" V="([\\d.-]+)"/>`).exec(shape)?.[1] ?? NaN);
    const bx = num('BeginX');
    const by = num('BeginY');
    const theta = num('Angle');
    if (!Number.isFinite(bx) || !Number.isFinite(by) || !Number.isFinite(theta)) continue;
    const cos = Math.cos(theta);
    const sin = Math.sin(theta);
    const pts = [...shape.matchAll(/<Cell N="X" V="([\d.-]+)"\/><Cell N="Y" V="([\d.-]+)"\/>/g)]
      .map((p) => ({ x: bx + +p[1] * cos - +p[2] * sin, y: by + +p[1] * sin + +p[2] * cos }));
    if (pts.length < 2) continue;
    const ownEnds = [pts[0], pts[pts.length - 1]];
    let through = 0;
    let crossed = '';
    for (let i = 1; i < pts.length; i += 1) {
      const a = pts[i - 1];
      const b = pts[i];
      const len = Math.hypot(b.x - a.x, b.y - a.y);
      const steps = Math.max(2, Math.ceil(len / 0.02));
      for (let s = 0; s < steps; s += 1) {
        const t = (s + 0.5) / steps;
        const at = { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
        // Its own endpoints sit on their tiles by design; only a third shape
        // being crossed is a defect.
        if (ownEnds.some((e) => Math.hypot(e.x - at.x, e.y - at.y) < 0.35)) continue;
        const inside = serviceBoxes.find(
          (box) => at.x > box.x + 0.02 && at.x < box.x + box.w - 0.02 && at.y > box.y + 0.02 && at.y < box.y + box.h - 0.02,
        );
        if (inside) {
          through += len / steps;
          crossed = inside.name ?? crossed;
        }
      }
    }
    if (through > 0.2) {
      const name = /NameU="(Connector\.\d+)"/.exec(shape)?.[1] ?? 'connector';
      const ends = `(${pts[0].x.toFixed(2)},${pts[0].y.toFixed(2)})->(${pts[pts.length - 1].x.toFixed(2)},${pts[pts.length - 1].y.toFixed(2)})`;
      issues.push(`Visio ${name} ${ends} is drawn through ${crossed || 'a service'} for ${through.toFixed(2)}in`);
    }
  }
  for (const block of badgeBlocks) {
    const shown = unescapeXml(/<Text>([^<]*)<\/Text>/.exec(block)?.[1] ?? '');
    if (!expectedNumbers.has(shown)) {
      issues.push(`Visio step badge shows "${shown}", which is not a workflow step number`);
    }
    if (!/<Row T="Ellipse"/.test(block)) {
      issues.push('Visio step badge is not drawn as an ellipse');
    }
    const geo = /<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/.exec(block);
    if (!geo) continue;
    const badge = { x: +geo[1] - +geo[3] / 2, y: +geo[2] - +geo[4] / 2, w: +geo[3], h: +geo[4] };
    if (badge.x < -0.01 || badge.y < -0.01
      || badge.x + badge.w > pkg.pageWidthIn + 0.01 || badge.y + badge.h > pkg.pageHeightIn + 0.01) {
      issues.push(`Visio step badge "${shown}" sits outside the page`);
    }
    for (const box of serviceBoxes) {
      const ow = Math.min(badge.x + badge.w, box.x + box.w) - Math.max(badge.x, box.x);
      const oh = Math.min(badge.y + badge.h, box.y + box.h) - Math.max(badge.y, box.y);
      if (ow > 0 && oh > 0 && ow * oh > 0.25 * badge.w * badge.h) {
        issues.push(`Visio step badge "${shown}" covers a service shape`);
        break;
      }
    }
  }

  // A connector's text is a block on the page like any other. It carries the
  // sentence the arrow exists to say, so two of them on the same spot is the
  // same defect as two chips on the same spot in PowerPoint — and until the
  // exporter emitted an explicit text position, a fan of parallel hops wrote
  // every one of its sentences at the identical midpoint.
  const labelBoxes: Array<{
    text: string; edge: string; x: number; y: number; w: number; h: number;
    /** Painted extent, which on a Visio chip is not the declared box. */
    inkW: number; inkH: number; cx: number; cy: number;
  }> = [];
  for (const block of xml.matchAll(/<Shape [^>]*NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const shape = block[0];
    const shown = unescapeXml(/<Text>([^<]*)<\/Text>/.exec(shape)?.[1] ?? '');
    if (!shown.trim()) continue;
    const pin = /<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>/.exec(shape);
    const angle = /<Cell N="Angle" V="([\d.-]+)"\/>/.exec(shape);
    const txt = /<Cell N="TxtPinX" V="([\d.-]+)"\/>\s*<Cell N="TxtPinY" V="([\d.-]+)"\/>\s*<Cell N="TxtWidth" V="([\d.-]+)"\/>\s*<Cell N="TxtHeight" V="([\d.-]+)"\/>/.exec(shape);
    if (!pin) continue;
    if (!txt) {
      issues.push(`Visio connector text "${shown.slice(0, 18)}" has no explicit position, so Visio centres it on the line`);
      continue;
    }
    // TxtPin is in the connector's own rotated frame, measured from its begin
    // point, while PinX/PinY is the centre of the line.
    const theta = angle ? +angle[1] : 0;
    const length = +(/<Cell N="Width" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0);
    const lx = +txt[1] - length / 2;
    const ly = +txt[2];
    const cx = +pin[1] + lx * Math.cos(theta) - ly * Math.sin(theta);
    const cy = +pin[2] + lx * Math.sin(theta) + ly * Math.cos(theta);
    // The ink, not the box the exporter declared for it.
    //
    // `TxtHeight` is the exporter's own answer to "how tall is this text?", so
    // reading it back asks the suspect to testify: a chip whose height was
    // sized by `ceil(width / column)` declared 0.180in while Visio drew 0.440in
    // of it, and two such chips 0.490in apart never overlapped on paper no
    // matter what was written through what. Measure the sentence instead, the
    // way the PowerPoint side has measured its own since the painted-ink rule.
    const sizePt = +(/<Cell N="Size" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0) * 72 || 7.2;
    const inkH = (auditWrappedLines(shown, Math.max(0.1, +txt[3] - 0.08), sizePt) * sizePt * 1.35) / 72;
    const boxH = Math.max(+txt[4], inkH);
    const inkW = Math.max(0, ...auditLineWidths(shown, Math.max(0.1, +txt[3] - 0.08), sizePt));
    labelBoxes.push({
      text: shown,
      edge: /Name="edge-([^"]*)"/.exec(shape)?.[1] ?? '',
      x: cx - +txt[3] / 2,
      y: cy - boxH / 2,
      w: +txt[3],
      h: boxH,
      inkW,
      inkH,
      cx,
      cy,
    });
  }
  const overlap = (a: { x: number; y: number; w: number; h: number }, b: { x: number; y: number; w: number; h: number }): number => {
    const ow = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
    const oh = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return ow > 0 && oh > 0 ? ow * oh : 0;
  };
  let stacked = 0;
  const piles: string[] = [];
  for (let i = 0; i < labelBoxes.length; i += 1) {
    for (let j = i + 1; j < labelBoxes.length; j += 1) {
      const hit = overlap(labelBoxes[i], labelBoxes[j]);
      if (hit > 0.25 * Math.min(labelBoxes[i].w * labelBoxes[i].h, labelBoxes[j].w * labelBoxes[j].h)) {
        stacked += 1;
        piles.push(`"${labelBoxes[i].text.slice(0, 12)}"/"${labelBoxes[j].text.slice(0, 12)}" at ${labelBoxes[i].x.toFixed(2)},${labelBoxes[i].y.toFixed(2)}`);
      }
    }
  }
  if (stacked > 0) {
    issues.push(`${stacked} pair(s) of Visio connector labels are written on top of each other: ${piles.slice(0, 3).join('; ')}`);
  }

  // The PowerPoint semantic, on the sheet. The rule above wants a quarter of a
  // box's area before it will speak, which is right for *declared* boxes —
  // those carry padding and a touch is normal — but ink written through ink is
  // never legitimate, and a fan of four hops whose rung pitch was frozen at
  // 0.490in against 0.570in of text overlapped by 0.080in: a tenth of the area
  // bar, and every sentence unreadable. This is what the deck's painted-ink
  // rule has caught since it was written, and what the sheet could not.
  //
  // This compared *boxes* and excused a share of the narrower one — 0.25 of it,
  // a number taken from the single scenario that was grazing. Measured, that
  // scenario's chips carry 0.094in of dead padding a side, so the derivable bar
  // was 0.188in and the constant was 0.425in: everything between them was glyph
  // on glyph and silent, a 0.237in window, about six characters at 7.2pt. The
  // padding is not a constant either — it is whatever the widest drawn line
  // leaves over — so the rule now builds the ink rectangle each chip actually
  // paints, centred the way the text is centred, and asks whether those
  // intersect. No share, no constant, and it subsumes the height arm.
  const bled: string[] = [];
  for (let i = 0; i < labelBoxes.length; i += 1) {
    for (let j = i + 1; j < labelBoxes.length; j += 1) {
      const a = labelBoxes[i];
      const b = labelBoxes[j];
      const ink = (c: typeof a) => ({
        x: c.cx - c.inkW / 2,
        y: c.cy - c.inkH / 2,
        w: c.inkW,
        h: c.inkH,
      });
      const ia = ink(a);
      const ib = ink(b);
      const ow = Math.min(ia.x + ia.w, ib.x + ib.w) - Math.max(ia.x, ib.x);
      const oh = Math.min(ia.y + ia.h, ib.y + ib.h) - Math.max(ia.y, ib.y);
      if (ow > 0.02 && oh > 0.02) {
        bled.push(
          `"${a.text.split('\n')[0].slice(0, 12)}"/"${b.text.split('\n')[0].slice(0, 12)}" by ${ow.toFixed(3)}x${oh.toFixed(3)}in of ink`,
        );
      }
    }
  }
  if (bled.length > 0) {
    issues.push(`${bled.length} pair(s) of Visio connector labels paint over each other: ${bled.slice(0, 3).join('; ')}`);
  }

  // A callout has to be readable AS the label of the arrow it belongs to. One
  // parked beside a different hop is worse than one overlapping a tile: the
  // reader matches it to the wrong arrow and never knows they did. The deck has
  // said this since the parallel-edge work; the sheet never had, because every
  // shape in the .vsdx was named `Connector.41` and nothing could tell which
  // arrow a number belonged to. Both now carry the edge they draw as their
  // shape name, which is also what Visio's Drawing Explorer lists.
  const arrowPaths: Array<{ edge: string; bundle: string; pts: Array<{ x: number; y: number }> }> = [];
  for (const block of xml.matchAll(/<Shape [^>]*NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const shape = block[0];
    const edge = /Name="edge-([^"]*)"/.exec(shape)?.[1];
    const pin = /<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>/.exec(shape);
    if (!edge || !pin) continue;
    const theta = +(/<Cell N="Angle" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0);
    const length = +(/<Cell N="Width" V="([\d.-]+)"\/>/.exec(shape)?.[1] ?? 0);
    // Geometry rows are in the arrow's own rotated frame, measured from its
    // begin point, while the pin is the centre of the begin→end chord.
    const pts = Array.from(shape.matchAll(/<Row T="(?:MoveTo|LineTo)" IX="\d+"><Cell N="X" V="([\d.-]+)"\/><Cell N="Y" V="([\d.-]+)"\/><\/Row>/g))
      .map((row) => {
        const lx = +row[1] - length / 2;
        const ly = +row[2];
        return {
          x: +pin[1] + lx * Math.cos(theta) - ly * Math.sin(theta),
          y: +pin[2] + lx * Math.sin(theta) + ly * Math.cos(theta),
        };
      });
    if (pts.length < 2) continue;
    const model = scenario.edges.find((e) => String(e.id) === edge);
    arrowPaths.push({
      edge,
      bundle: model ? [String(model.source), String(model.target)].sort().join('|') : edge,
      pts,
    });
  }
  if (arrowPaths.length > 1) {
    const gapTo = (arrow: { pts: Array<{ x: number; y: number }> }, at: { x: number; y: number }): number => {
      let best = Infinity;
      for (let i = 1; i < arrow.pts.length; i += 1) {
        const a = arrow.pts[i - 1];
        const b = arrow.pts[i];
        const vx = b.x - a.x;
        const vy = b.y - a.y;
        const len2 = vx * vx + vy * vy;
        const t = len2 > 0 ? Math.min(1, Math.max(0, ((at.x - a.x) * vx + (at.y - a.y) * vy) / len2)) : 0;
        best = Math.min(best, Math.hypot(at.x - (a.x + vx * t), at.y - (a.y + vy * t)));
      }
      return best;
    };
    const stray = (
      what: string,
      items: Array<{ id: string; edge: string; at: { x: number; y: number } }>,
      crossBundleOnly: boolean,
    ): void => {
      const reports: string[] = [];
      for (const item of items) {
        const own = arrowPaths.find((a) => a.edge === item.edge);
        if (!own) continue;
        // Fan siblings are exempt for the numbers: a bundle of parallel edges
        // between one pair of services is a single object to the reader, so a
        // rung nearer sibling 5 than sibling 6 misleads nobody.
        const others = arrowPaths.filter((a) => (crossBundleOnly ? a.bundle !== own.bundle : a.edge !== own.edge));
        if (others.length === 0) continue;
        const mine = gapTo(own, item.at);
        const nearest = others.reduce((best, a) => (gapTo(a, item.at) < gapTo(best, item.at) ? a : best), others[0]);
        const theirs = gapTo(nearest, item.at);
        if (theirs < mine - 0.25) {
          reports.push(`"${item.id.slice(0, 20)}" is ${theirs.toFixed(2)}in from ${nearest.edge} but ${mine.toFixed(2)}in from its own arrow`);
        }
      }
      if (reports.length > 0) {
        issues.push(`${reports.length} Visio ${what} nearer another hop than their own: ${reports.slice(0, 3).join('; ')}`);
      }
    };
    const labelItems = labelBoxes
      .filter((box) => box.edge !== '')
      .map((box) => ({ id: box.text, edge: box.edge, at: { x: box.x + box.w / 2, y: box.y + box.h / 2 } }));
    stray('connector label(s)', labelItems, false);
    const badgeItems: Array<{ id: string; edge: string; at: { x: number; y: number } }> = [];
    for (const m of xml.matchAll(
      /NameU="StepBadge\.\d+" Name="step-([^"]*)"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>[\s\S]*?<Text>([\s\S]*?)<\/Text>/g,
    )) {
      badgeItems.push({ id: `callout ${unescapeXml(m[4]).trim()}`, edge: m[1], at: { x: +m[2], y: +m[3] } });
    }
    stray('numbered callout(s)', badgeItems, true);
  }

  // The workflow band and the connection legend are opaque white panels drawn
  // last, over everything. Every other rule about them asks whether the band is
  // well-formed — its rows fit, its rows are inside it, it reserves no dead air
  // — and every one of them passes while the panel sits on top of six of the
  // nine services in the drawing. This asks the only question the reader asks:
  // is anything underneath it? Nothing else in the corpus could see the band
  // paint out a tile, or the label search park a ladder in the band's strip
  // because that strip held no service and no other label.
  const panelRects: Array<{ name: string; x: number; y: number; w: number; h: number }> = [];
  for (const m of xml.matchAll(
    /NameU="(Workflow|Legend)\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )) {
    const [, name, pinX, pinY, w, h] = m;
    panelRects.push({ name: name === 'Workflow' ? 'workflow band' : 'connection legend', x: +pinX - +w / 2, y: +pinY - +h / 2, w: +w, h: +h });
  }
  if (panelRects.length > 0) {
    const badgeRects: Array<{ text: string; x: number; y: number; w: number; h: number }> = [];
    for (const m of xml.matchAll(
      /NameU="StepBadge\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>[\s\S]*?<Text>([\s\S]*?)<\/Text>/g,
    )) {
      const [, pinX, pinY, w, h, text] = m;
      badgeRects.push({ text: text.trim(), x: +pinX - +w / 2, y: +pinY - +h / 2, w: +w, h: +h });
    }
    const buried = (
      what: string,
      boxes: Array<{ text?: string; name?: string; x: number; y: number; w: number; h: number }>,
      bar: number,
    ): void => {
      const hidden: string[] = [];
      for (const box of boxes) {
        const own = Math.max(box.w * box.h, 1e-9);
        for (const p of panelRects) {
          if (overlap(box, p) > bar * own) {
            hidden.push(`"${(box.text ?? box.name ?? '?').slice(0, 16)}" under the ${p.name}`);
            break;
          }
        }
      }
      if (hidden.length > 0) {
        issues.push(`${hidden.length} ${what} drawn under an opaque panel: ${hidden.slice(0, 3).join(', ')}`);
      }
    };
    // The band's page reservation is measured before the drawing is laid out,
    // from sentences that can still grow when a muted label hands its wording
    // over. It is deliberately an over-estimate, because under-reserving paints
    // the panel across the drawing — but an over-estimate is blank paper
    // between the drawing and the band, and on a 21.5in sheet it was 2.5in of
    // it. Measured as the asymmetry of the drawing's margins rather than the
    // gap itself: the drawing is centred between the two panels, so a small
    // architecture on a minimum-size sheet has wide margins for a legitimate
    // reason, and only the reservation the band did not use pushes the top
    // margin past the bottom one.
    const band = panelRects.find((p) => p.name === 'workflow band');
    if (band && serviceBoxes.length > 0) {
      const floor = panelRects.filter((p) => p.name === 'connection legend').reduce((lo, p) => Math.max(lo, p.y + p.h), 0);
      const above = band.y - Math.max(...serviceBoxes.map((s) => s.y + s.h));
      const below = Math.min(...serviceBoxes.map((s) => s.y)) - floor;
      if (above - below > BAND_RESERVE_SLACK_IN) {
        issues.push(`${(above - below).toFixed(2)}in of blank paper between the drawing and the workflow band — the band reserved page it did not use`);
      }
    }
    // Two opaque panels that overlap each other are a hole in every rule above:
    // the "nothing may be drawn under a panel" test rescues a badge by stepping
    // it out of the one it is under, and if the panels intersect, the seat it
    // steps to is inside the other one. It cannot happen today — the legend
    // reserves 0.24n + 0.79in of page while its rectangle only reaches
    // 0.24n + 0.69in, and the page height always carries both reservations plus
    // the padding — but that is an argument about three constants in a file
    // nobody reads while changing a fourth. This is the same statement, checked.
    for (let i = 0; i < panelRects.length; i += 1) {
      for (let j = i + 1; j < panelRects.length; j += 1) {
        if (panelRects[i].name === panelRects[j].name) continue;
        if (overlapArea(panelRects[i], panelRects[j]) > 0) {
          issues.push(`the ${panelRects[i].name} and the ${panelRects[j].name} overlap each other, so stepping a shape out of one puts it inside the other`);
        }
      }
    }
    // A service is a picture: any of it lost is a service the reader cannot
    // identify. Text is gone the moment enough of it is covered to stop it
    // reading, which is the same 1% bar the exporter's own muting pass uses.
    buried('Visio service tile(s)', serviceBoxes, 0.01);
    buried('Visio step badge(s)', badgeRects, 0.01);
    buried('Visio connector label(s)', labelBoxes, 0.01);
  }
  let onService = 0;
  const buried: string[] = [];
  for (const label of labelBoxes) {
    if (serviceBoxes.some((box) => overlap(label, box) > 0.4 * label.w * label.h)) {
      onService += 1;
      buried.push(`"${label.text.slice(0, 14)}" at ${label.x.toFixed(2)},${label.y.toFixed(2)}`);
    }
  }
  if (onService > 0) {
    issues.push(`${onService} Visio connector label(s) are buried under a service shape: ${buried.slice(0, 3).join('; ')}`);
  }
  const offSheet = labelBoxes.filter((label) => label.x < -0.01 || label.y < -0.01
    || label.x + label.w > pkg.pageWidthIn + 0.01 || label.y + label.h > pkg.pageHeightIn + 0.01).length;
  if (offSheet > 0) issues.push(`${offSheet} Visio connector label(s) run off the sheet`);

  // A sparse page is the outlier symptom: a huge sheet holding a small drawing.
  const shapeArea = [...xml.matchAll(/NameU="Service\.\d+"[\s\S]*?<Cell N="Width" V="([\d.]+)"\/>\s*<Cell N="Height" V="([\d.]+)"\/>/g)]
    .reduce((sum, m) => sum + +m[1] * +m[2], 0);
  const density = shapeArea / (pkg.pageWidthIn * pkg.pageHeightIn);
  // Sparse is the symptom; outliers are the disease, and only the disease is
  // the exporter's to cure. A drawing whose services are spread evenly across
  // the sheet — a long cascade, a wide bus, a timeline — is thin everywhere,
  // and Visio reproducing it at 1:1 is the tool working correctly; there is
  // nothing to trim, and reporting it only teaches the gate to be ignored.
  // What a stray actually looks like is a sheet whose span collapses once the
  // few boxes at the extremes are set aside.
  const centres = [
    ...xml.matchAll(new RegExp('<Shape [^>]*NameU="Service\\.\\d+"[\\s\\S]*?<\\/Shape>', 'g')),
  ].map((m) => rectOf(m[0]))
    .filter((r): r is { x: number; y: number; w: number; h: number } => r !== null)
    .map((r) => ({ x: r.x + r.w / 2, y: r.y + r.h / 2 }));
  const lopsided = (pick: (c: { x: number; y: number }) => number): number => {
    const sorted = centres.map(pick).sort((a, b) => a - b);
    const full = sorted[sorted.length - 1] - sorted[0];
    if (full <= 0) return 1;
    const cut = Math.floor(sorted.length * 0.1);
    const core = sorted[sorted.length - 1 - cut] - sorted[cut];
    return core / full;
  };
  const outlierDriven = centres.length >= 4 && Math.min(lopsided((c) => c.x), lopsided((c) => c.y)) < 0.6;
  if (serviceCount >= 4 && density < 0.005 && outlierDriven) {
    issues.push(`page is ${(density * 100).toFixed(2)}% full — a stray node blew the sheet up`);
  }

  // Density alone lets a two-region drawing through: sixteen services on a 73in
  // sheet is 3% full, well clear of the floor, and 52in of that sheet is one
  // continuous band with nothing in it. Whitespace is not content — it costs
  // the rest of the drawing its scale, and on the fixed-size deck it costs it
  // in font size.
  //
  // Measured across the services and the corridor labels, for the same reason
  // the exporter closes voids by them: one rectangle drawn around the whole
  // architecture spans every band there is, and counting it as content let a
  // five-region drawing report 0.0in of void while carrying 256in of it — but
  // a childless box *is* the content of the band it names, and reporting the
  // band it deliberately occupies would demand the exporter delete it.
  const parentedZones = new Set(scenario.nodes.map((n) => n.parentNode).filter((id): id is string => !!id));
  const corridorRects = scenario.nodes
    .filter((n) => n.type === 'groupNode' && !parentedZones.has(n.id))
    .map((n) => zoneRects.get(escAttr(String(n.data?.label ?? ''))))
    .filter((r): r is { x: number; y: number; w: number; h: number } => !!r);
  type VoidRect = { x: number; y: number; w: number; h: number };
  const tileRects = [...xml.matchAll(new RegExp('<Shape [^>]*NameU="Service\\.\\d+"[\\s\\S]*?<\\/Shape>', 'g'))]
    .map((m) => rectOf(m[0])).filter((r): r is VoidRect => r !== null);
  const widestVoid = (start: (r: VoidRect) => number, size: (r: VoidRect) => number): number => {
    // Per-axis, and for the same reason the exporter closes voids per-axis: a
    // childless box standing *between* two clusters is the content of the band
    // it names, but one stretched *over* the drawing is not. Counting the
    // second as content is what let a sovereign caption across the top report
    // 0.0in of void on a sheet carrying 56.8in of it — the audit was blinded by
    // the identical rectangle that blinded the exporter, so the gate that
    // should have caught the defect passed it clean.
    const standsBetween = (zone: VoidRect): boolean => !tileRects.some((r) => {
      const over = Math.min(start(r) + size(r), start(zone) + size(zone)) - Math.max(start(r), start(zone));
      return over > size(r) / 2;
    });
    const spans = [...tileRects, ...corridorRects.filter(standsBetween)]
      .map((r) => [start(r), start(r) + size(r)] as [number, number])
      .sort((a, b) => a[0] - b[0]);
    if (spans.length === 0) return 0;
    let reach = spans[0][1];
    let widest = 0;
    for (const [from, to] of spans) {
      widest = Math.max(widest, from - reach);
      reach = Math.max(reach, to);
    }
    return widest;
  };
  const voidW = widestVoid((r) => r.x, (r) => r.w);
  const voidH = widestVoid((r) => r.y, (r) => r.h);
  if (Math.max(voidW, voidH) > 16) {
    issues.push(`the drawing contains a ${Math.max(voidW, voidH).toFixed(1)}in band with nothing in it — empty space must be closed, not exported`);
  }

  return {
    scenario: scenario.id,
    format: 'vsdx',
    issues,
    drawnNames: (() => {
      const unesc = (s: string): string => unescapeXml(s.replace(/<[^>]*>/g, '')).trim();
      const best = new Map<string, string>();
      for (const m of xml.matchAll(/NameU="Service\.\d+" Name="([^"]*)"[\s\S]*?<Text>([\s\S]*?)<\/Text>/g)) {
        const drawn = unesc(m[2]);
        if (drawn === '') continue;
        const authored = unesc(m[1]);
        if ((best.get(authored)?.length ?? -1) < drawn.length) best.set(authored, drawn);
      }
      // The drawing's own index page names the service too. Mirrors the deck
      // side exactly - a name printed in the index is a name a reader can see,
      // and refusing to count it would report the sheet as having lost a
      // service it prints in full. Read from PAGE 2: the index is not on the
      // drawing, which is the whole point of it.
      //
      // Matched by SUFFIX, not by splitting the row on its separator. A row is
      // "<stub>  =  <full name>", and a name that was not drawn at all leaves an
      // empty stub that the XML text trims away - so the separator is not always
      // where the format says it is, and a name can contain the separator itself.
      const panelRows = [...indexXml.matchAll(/Name="service-name-\d+"[\s\S]*?<Text>([\s\S]*?)<\/Text>/g)]
        .map((m) => unesc(m[1]));
      const namedInIndex = (authored: string): boolean => panelRows.some(
        (row) => row === authored || row.endsWith(`  =  ${authored}`) || row.endsWith(`=  ${authored}`),
      );
      for (const m of xml.matchAll(/NameU="Service\.\d+" Name="([^"]*)"/g)) {
        const authored = unesc(m[1]);
        if (!authored || !namedInIndex(authored)) continue;
        // Longest rendition wins, the same rule the tiles use above: the panel
        // prints the name IN FULL, so a tile that could only fit "N..." has not
        // cost the reader the name.
        if ((best.get(authored)?.length ?? -1) < authored.length) best.set(authored, authored);
      }
      return [...best]
        .map(([authored, drawn]) => ({ authored, drawn }))
        .sort((a, b) => a.authored.localeCompare(b.authored));
    })(),
    metrics: {
      pageWidthIn: +pkg.pageWidthIn.toFixed(2),
      pageHeightIn: +pkg.pageHeightIn.toFixed(2),
      mediaParts: media.length,
      textBlocks: textCount,
      minFontPt,
      fillPct: +(density * 100).toFixed(2),
      stepBadges: badgeBlocks.length,
    },
  };
}

/**
 * Adding a service must not make the deck shorter *and* its type smaller.
 *
 * Coarsening the window grid toward a square costs scale on whichever axis is
 * coarsened, and the reader gets the smaller of the two axes' scales — so
 * spending the cost on the axis that already binds shrinks the type and buys
 * nothing. A diagonal cascade is long in one axis by construction, and it used
 * to lose exactly that axis: fifty-two services came out at 6.0pt on *fewer*
 * slides than fifty-one, which means adding a service to the diagram made the
 * deck both shorter and less readable.
 *
 * Every rule in this file judges one export on its own, and no single export of
 * that cascade looks wrong — 6.0pt on 30 slides is a perfectly ordinary deck.
 * The defect is only visible as a discontinuity across the family, so this
 * walks consecutive sizes and compares them.
 */
async function auditDeckGrowth(): Promise<Report> {
  const issues: string[] = [];
  const seen: Array<{ n: number; slides: number; font: number }> = [];
  for (const n of [118, 119, 120, 121, 122]) {
    const report = await auditPptx(diagonalCascadeScenario(n, `deck-growth-${n}`));
    seen.push({
      n,
      slides: Number(report.metrics.slides ?? 0),
      font: Number(report.metrics.minFontPt ?? 0),
    });
  }
  for (let i = 1; i < seen.length; i += 1) {
    const prev = seen[i - 1];
    const cur = seen[i];
    if (cur.slides < prev.slides && cur.font < prev.font - 0.01) {
      issues.push(
        `a ${cur.n}-service cascade is ${prev.slides - cur.slides} slide(s) shorter than a ${prev.n}-service one *and* ${(prev.font - cur.font).toFixed(2)}pt smaller (${prev.slides} slides at ${prev.font}pt, then ${cur.slides} at ${cur.font}pt) — adding a service made the deck worse in both directions`,
      );
    }
  }
  return {
    scenario: 'deck-growth',
    format: 'pptx',
    issues,
    // The type says numbers, and `${slides}sl/${font}pt` is a string that
    // happened to type-check nowhere because nothing ever compiled this file.
    // Two numeric keys carry the same information and survive a diff.
    metrics: Object.fromEntries(
      seen.flatMap((s) => [[`n${s.n}-slides`, s.slides], [`n${s.n}-pt`, s.font]] as Array<[string, number]>),
    ),
  };
}

/**
 * A watchdog that survives a blocked event loop.
 *
 * An exporter that spins forever does it SYNCHRONOUSLY: the tile-name fitter
 * that hung the browser this round also hung this script, which stopped
 * writing artefacts and sat there for fourteen minutes with no output and no
 * non-zero exit. A `setTimeout` cannot fire in that state and neither can
 * `Promise.race` — the timer never gets a turn. So the clock has to live in
 * another thread: the main thread stamps a shared buffer before each scenario,
 * a worker polls it, and if the stamp goes stale the worker names the scenario
 * and signals the OS process, which is the one thing that still works when the
 * main thread will not yield.
 *
 * Without this a non-terminating exporter turns `npm test` into an infinite
 * wait rather than a red build, which is the difference between a regression
 * test and no test at all.
 */
const SCENARIO_BUDGET_MS = 240_000;
const heartbeatBuffer = new SharedArrayBuffer(8);
const heartbeat = new Int32Array(heartbeatBuffer);
const auditStartedAt = Date.now();
let watchdog: Worker | null = null;

function beat(id: string): void {
  watchdog?.postMessage(id);
  Atomics.store(heartbeat, 0, Math.floor((Date.now() - auditStartedAt) / 100));
  Atomics.store(heartbeat, 1, 1);
}

function startWatchdog(): void {
  const source = `
    const { parentPort, workerData } = require('node:worker_threads');
    const view = new Int32Array(workerData.buffer);
    const budget = workerData.budget;
    let name = 'startup';
    parentPort.on('message', (m) => { name = m; });
    setInterval(() => {
      if (Atomics.load(view, 1) === 0) return;
      const stampMs = Atomics.load(view, 0) * 100;
      const stalled = (Date.now() - workerData.startedAt) - stampMs;
      if (stalled > budget) {
        // writeSync, not console.error: the process is about to be killed and
        // a buffered stream never gets flushed, so the one diagnostic that
        // explains the failure would be lost exactly when it is needed.
        require('node:fs').writeSync(2, '\\nWATCHDOG: scenario "' + name + '" has made no progress for '
          + (stalled / 1000).toFixed(0) + 's. The exporter is not terminating - '
          + 'a synchronous loop cannot be interrupted, so the run is being killed.\\n');
        process.kill(process.pid, 'SIGKILL');
      }
    }, 2000).unref();
  `;
  watchdog = new Worker(source, {
    eval: true,
    workerData: { buffer: heartbeatBuffer, budget: SCENARIO_BUDGET_MS, startedAt: auditStartedAt },
  });
  watchdog.unref();
}

function stopWatchdog(): void {
  Atomics.store(heartbeat, 1, 0);
  void watchdog?.terminate();
  watchdog = null;
}

async function main(): Promise<void> {
  mkdirSync(OUT, { recursive: true });
  const base = [
    compactScenario(), wideScenario(), oversizeScenario(), outlierScenario(),
    bandedScenario(), narrativeScenario(), barbellScenario(), hubFanScenario(), sharedServiceScenario(), tightGridScenario(), bandedTwoStraysScenario(), wideChainScenario(), grid5x5TightScenario(), parallelScenario(),
    oppositeStraysScenario(), cornerStraysScenario(), symmetricStraysScenario(),
    hubSpokeScenario(), scopeZoneScenario(), strayZonePairScenario(), zoneStrayScenario(),
    boundaryVoidScenario(), stackedSubnetsScenario(), tightSubnetsScenario(), flushSubnetsScenario(), diagonalCascadeScenario(),
    diagonalCascadeScenario(27, 'diagonal-cascade-27'),
    // Past the deck ceiling. A drawing this sparse needs one window per service
    // to reach seven points, so it is the shape that used to be coarsened until
    // the ceiling was satisfied and the type was not — 52 services came out at
    // 6.31pt, and 90 at 4.00pt, on exactly forty-eight slides either way.
    diagonalCascadeScenario(52, 'diagonal-cascade-52'),
    bandAboveScenario(), framedCascadeScenario(), tightSeamScenario(), overRowScenario(),
    // Past where the type floor used to stop tracking the drawing: the ratio
    // rule below was unsatisfiable by construction from about 24% down.
    overRowScenario(700, 'over-row-700'),
    scaledZoneRowScenario(),
    midZoneRowScenario(),
    stringStepPromotionScenario(),
    unlabelledStepInflationScenario(),
    dataLabelPromotionScenario(),
  hairlineTilesScenario(),
  probeArrowScenario(),
  probeAccentScenario(),
  probeAmpScenario(),
  probeScriptScenario(),
  scaleDownPipelineScenario(),
  whitespaceLabelsScenario(),
  panelBurialScenario(),
  briefWorkflowStepsScenario(),
  longIndexRowsScenario(),
  overlongIndexRowsScenario(),
  shrinkableIndexRowsScenario(),
  mixedLengthIndexScenario(),
  bimodalWorkflowScenario(),
  bandGapScenario(),
  bandFillScenario(),
  decomposedNameScenario(),
  normFormScenario(),
  tinyTileSpreadScenario(),
  widthCliffScenario(),
  slaveredBadgeScenario(),
  threeTierBadgeScenario(),
  twoChainBadgeScenario(),
  bimodalSidecarScenario(),
  numberedSpreadScenario(),
  numberedMidSpreadScenario(),
  duplicateLabelFanScenario(),
  hubSpokeCalloutScenario(),
  numberedEstateScenario(),
  wideHubCalloutScenario(),
  sliverEstateScenario(),
  bandEstateScenario(15),
  bandEstateScenario(16),
  blindSliverScenario(),
  mixedSliverScenario(),
  gutterRegionScenario(),
  magnifiedGutterScenario(),
  spreadGlyphScenario('probe-tight', 200, 300, 200),
  spreadGlyphScenario('probe-spread', 290, 440, 200),
  spreadGlyphScenario('probe-offrow', 290, 440, 400),
  glyphChainScenario('probe-glyph16', 16),
  glyphChainScenario('probe-glyph12', 12),
  touchingBadgeScenario('probe-touching', 30),
  mixCountScenario(4),
  mixCountScenario(5),
  mixCountScenario(6),
  zoneMedianScenario(0),
  zoneMedianScenario(2),
  halfTailScenario(),
  refusedRaiseScenario(),
  longTitleScenario(20),
  longTitleScenario(70),
  longTitleScenario(95),
  longTitleScenario(130),
  collidingStubsScenario(),
  hairlineStubsScenario(),
  emojiClusterScenario(),
  briefWorkflowScenario(),
  shreddedWorkflowScenario(),
  tallNarrowTilesScenario(),
    corridorZoneScenario(),
    ladderInGridScenario(), twinLaddersScenario(), strayLadderScenario(), legendCornerScenario(), duplicateStepsScenario(), denseZoneScenario(),
    metaChipScenario(), gridFanScenario(), gridFan3Scenario(), fan8Tight5x5Scenario(), metaSublineScenario(), grid5x5CaptionScenario(), longNameGridScenario(), longLabelGridScenario(), metaTightScenario(),     longNameFanScenario(), estateChainScenario(), chain24Scenario(), tripleMutedScenario(), estate72Scenario(),     workflowProseScenario(), workflowLongProseScenario(), workflowFanScenario(), workflowWideBandScenario(), allCategoriesScenario(), controlCharScenario(), shortServiceGridScenario(),
    cascadeScenario(),
    sharedPrefixEstateScenario(),
    shortTileEstateScenario(),
    compactEstateScenario(),
    wrappedInventoryScenario(),
    tokenWrapInventoryScenario(),
    longWorkflowScenario(),
    visioTokenWorkflowScenario(),
    hardBreakInventoryScenario(),
    visioBrokenLabelFanScenario(),
    visioDefaultTileNamesScenario(),
    flushTopZoneScenario(),
    zoneCaptionCorridorScenario(),
    tileNameWithMetaScenario(),
    zoneCaptionWideEstateScenario(),
    squeezedBadgeScenario(),
    await generatedScenario(), await groupedGeneratedScenario(),
  ];
  // Dark twins. Adding a `dark` flag was not enough on its own: nothing set it,
  // so the dark palette stayed exactly as unmeasured as it had always been and
  // the contrast failures found so far were all light-theme ones. Every colour
  // the deck picks is theme-dependent — panel fills, zone tints, callout
  // accents, the workflow band, the footer — so the twins carry the scenarios
  // that between them draw every one of those, not just a dense grid.
  const darkTwins = ['dense-zone', 'narrative', 'legend-corner', 'meta-subline', 'grid5x5-captions', 'generated']
    .map((id) => base.find((s) => s.id === id))
    .filter((s): s is Scenario => s !== undefined)
    .map((s) => ({ ...s, id: `${s.id}-dark`, dark: true }));
  const scenarios = [...base, ...darkTwins];
  const reports: Report[] = [];
  // A single scenario name (or comma-separated list) narrows the run. The full
  // corpus takes minutes, which makes an iterate-on-one-fixture loop painful;
  // CI and `npm test` pass no argument and so always run everything.
  const only = process.argv.slice(2).filter((a) => !a.startsWith('-')).flatMap((a) => a.split(','));
  const selected = only.length > 0 ? scenarios.filter((s) => only.includes(s.id)) : scenarios;
  // `deck-growth` is a family comparison, not a scenario: it exports the same
  // drawing at several sizes and judges the differences between them, so it has
  // no entry in `scenarios` and has to be named explicitly to be selectable.
  const growth = only.length === 0 || only.includes('deck-growth');
  if (only.length > 0 && selected.length === 0 && !growth) {
    throw new Error(`no scenario matched ${only.join(', ')}; known: deck-growth, ${scenarios.map((s) => s.id).join(', ')}`);
  }
  for (const scenario of selected) {
    beat(scenario.id);
    reports.push(await auditPptx(scenario));
    reports.push(await auditVsdx(scenario));
  }
  beat('deck-growth');
  if (growth) reports.push(await auditDeckGrowth());
  stopWatchdog();
  // The two files, against each other.
  //
  // Every rule above reads one format's own output, so a name that one
  // exporter drops and the other keeps is invisible to both by construction —
  // and that is exactly what happened twice running, in opposite directions.
  // A user who exports the same diagram to PowerPoint and to Visio and finds
  // different services named in each has been handed two drawings, not two
  // renderings of one, and neither file can be checked against the other from
  // the inside.
  //
  // Compared as SETS, not as counts.
  //
  // The two formats draw the same diagram at deliberately different scales —
  // Visio exports one large sheet, PowerPoint a 13.33in slide — so the same
  // tile is a different number of inches wide in each, and a tile that clears
  // the legibility bar in one may honestly miss it in the other. Counting
  // tiles reported that as a defect: nine identically-named tiles were named
  // five times in the deck and six times on the sheet, purely because the
  // sheet is 1.1x larger. What is never legitimate is a service that is named
  // in one file and appears nowhere in the other, which is exactly what
  // happened in both directions in consecutive rounds.
  for (const scenario of selected) {
    const p = reports.find((r) => r.scenario === scenario.id && r.format === 'pptx');
    const v = reports.find((r) => r.scenario === scenario.id && r.format === 'vsdx');
    if (!p?.drawnNames || !v?.drawnNames) continue;
    // Both sides through the same normaliser, and specifically through the
    // same one the exporters use. A name carrying a vertical tab or a lone
    // surrogate reaches both files with that code point replaced by a space,
    // while the deck's key here was taken from the authored label and still
    // held the raw character — so the two keys differed for 41 names that both
    // files had drawn perfectly well. That is a difference between the
    // comparison's own two spellings, not between the drawings.
    //
    // Canonical form is the same mistake on a second axis. The exporters now
    // compose every drawn string, so a name authored decomposed arrives in
    // both files composed and identical - and this key, taken from the
    // authored label, still held the decomposed spelling. Seven Turkish and
    // Vietnamese names that both files had drawn correctly were reported as
    // missing from one of them.
    const key = (s: string): string => stripXmlForbidden(s).replace(/\s+/g, ' ').trim().normalize('NFC');
    const pt = new Map(p.drawnNames.map((n) => [key(n.authored), key(n.drawn)]));
    const vt = new Map(v.drawnNames.map((n) => [key(n.authored), key(n.drawn)]));
    for (const name of new Set([...pt.keys(), ...vt.keys()])) {
      if (pt.has(name) !== vt.has(name)) {
        const missing = pt.has(name) ? 'the Visio drawing' : 'the PowerPoint deck';
        p.issues.push(
          `"${name}" is named in the .${pt.has(name) ? 'pptx' : 'vsdx'} and on no shape at all in `
          + `${missing} — the two exports of one diagram do not name the same services`,
        );
        continue;
      }
      // Presence alone is too weak. Both formats can name a service and still
      // disagree about it completely: one keeps the whole name, the other cuts
      // it to a four-character stub. Compare a DIMENSIONLESS quantity - the
      // fraction of the authored name that survives - so the comparison stays
      // valid across the two formats' deliberately different page scales.
      const authoredLen = [...name].length;
      if (authoredLen === 0) continue;
      const kept = (drawn: string): number => {
        const withoutEllipsis = drawn.replace(/[\u2026.]+$/, '');
        return Math.min(1, [...withoutEllipsis].length / authoredLen);
      };
      const kp = kept(pt.get(name) ?? '');
      const kv = kept(vt.get(name) ?? '');
      if (Math.abs(kp - kv) > 0.4) {
        const fuller = kp > kv ? 'deck' : 'sheet';
        p.issues.push(
          `"${name}" survives ${(Math.max(kp, kv) * 100).toFixed(0)}% in the ${fuller} but only `
          + `${(Math.min(kp, kv) * 100).toFixed(0)}% in the other — `
          + `"${pt.get(name)}" against "${vt.get(name)}"`,
        );
      }
    }
  }
  for (const report of reports) {
    console.log(`\n=== ${report.scenario} / ${report.format} ===`);
    console.log('metrics:', JSON.stringify(report.metrics));

    if (report.issues.length === 0) console.log('  PASS - no issues');
    else report.issues.slice(0, 14).forEach((i) => console.log('  ISSUE:', i));
    if (report.issues.length > 14) console.log(`  ...and ${report.issues.length - 14} more`);
  }
  writeFileSync(path.join(OUT, 'report.json'), JSON.stringify(reports, null, 2));
  const total = reports.reduce((sum, r) => sum + r.issues.length, 0);
  console.log(`\nTOTAL ISSUES: ${total}`);
  if (total > 0) process.exitCode = 1;
}

startWatchdog();

main().catch((error) => {
  console.error(error);
  process.exit(1);
});





