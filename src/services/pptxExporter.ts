// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * PowerPoint Export Service
 *
 * Generates a single widescreen (16:9) .pptx slide from the current diagram
 * canvas image.  The slide theme (dark / light) matches the app's current
 * colour mode so the exported slide looks exactly like what the user sees.
 *
 * Library: PptxGenJS v4 (client-side, no backend required)
 */

import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';

/**
 * Interop guard: bundlers hand back the class directly, while Node resolving
 * the CommonJS build hands back `{ default: PptxGenJS }`. Unit tests import
 * this module under Node, so normalise the constructor once.
 */
const PptxCtor = (PptxGenJS as unknown as { default?: typeof PptxGenJS }).default ?? PptxGenJS;

/**
 * A deck whose every slide sanitises the text put on it.
 *
 * XML 1.0 forbids the C0 control characters outright, and no escaping helps —
 * `&#11;` is exactly as illegal as a raw U+000B. A single one anywhere in
 * `ppt/slides/*.xml` makes PowerPoint refuse to open the file, and the export
 * itself succeeds silently, so the first anyone hears of it is the recipient
 * reporting a corrupt deck. They are not exotic: U+000B is Word and
 * PowerPoint's own manual line break, so it arrives by copy-paste, and it is a
 * legal JSON escape, so it survives an IaC or prototype import untouched.
 *
 * Wrapped at the slide factory rather than at the forty-odd `addText` calls,
 * because the interesting failure is the call site nobody remembered.
 */
function newDeck(): PptxGenJS {
  const pptx = new PptxCtor();
  const addSlide = pptx.addSlide.bind(pptx);
  pptx.addSlide = ((...args: Parameters<typeof pptx.addSlide>) => {
    const slide = addSlide(...args);
    // Every writer, and every argument of it. Sanitising only the text argument
    // of `addText` left two ways through: the options bag was passed on
    // untouched, and `objectName` in it becomes the `name` attribute of
    // `<p:cNvPr>` carrying a node, group or edge id straight from the diagram;
    // and `addShape`/`addImage` were not wrapped at all, which is where the
    // icon and tile ids go. A shape name is as fatal to the parse as a caption.
    for (const key of ['addText', 'addTable', 'addShape', 'addImage', 'addChart', 'addMedia', 'addNotes'] as const) {
      const fn = (slide as unknown as Record<string, unknown>)[key];
      if (typeof fn !== 'function') continue;
      const bound = (fn as (...a: unknown[]) => unknown).bind(slide);
      (slide as unknown as Record<string, unknown>)[key] = (...args: unknown[]) => bound(...args.map(cleanText));
    }
    return slide;
  }) as typeof pptx.addSlide;

  // `docProps/core.xml` is written from these, and a deck whose metadata is
  // ill-formed is just as unopenable as one whose slides are — the caller
  // passes the diagram name and the author's name straight through, and both
  // are free text a user typed or pasted.
  for (const key of ['author', 'company', 'revision', 'subject', 'title'] as const) {
    let holder: object | null = pptx;
    let desc: PropertyDescriptor | undefined;
    while (holder && !desc) {
      desc = Object.getOwnPropertyDescriptor(holder, key);
      holder = Object.getPrototypeOf(holder) as object | null;
    }
    if (!desc?.set) continue;
    const { get, set } = desc;
    Object.defineProperty(pptx, key, {
      configurable: true,
      get: get ? () => get.call(pptx) : undefined,
      set: (value: unknown) => set.call(pptx, cleanText(value)),
    });
  }
  return pptx;
}

/** Strip XML-forbidden code points wherever text hides in a pptxgenjs argument. */
function cleanText(value: unknown): unknown {
  if (typeof value === 'string') return stripXmlForbidden(value);
  if (Array.isArray(value)) return value.map(cleanText);
  if (value && typeof value === 'object') {
    const source = value as Record<string, unknown>;
    // `text` is the only string pptxgenjs renders, but it is not the only one
    // it writes. `objectName` becomes the `name` attribute of `<p:cNvPr>`, and
    // it is built from diagram ids rather than typed prose, which is exactly
    // why it was missed: ids look like they came from us. They did not — an
    // imported template or a model response names its own nodes.
    //
    // Everything else in an options bag is a colour, a size, a font name or a
    // base64 image, and rewriting those would be a different kind of bug, so
    // this deliberately does not recurse into values it was not asked about.
    const hasText = 'text' in source;
    const hasName = typeof source.objectName === 'string';
    if (!hasText && !hasName) return value;
    const out = { ...source };
    if (hasText) out.text = cleanText(source.text);
    if (hasName) out.objectName = stripXmlForbidden(source.objectName as string);
    return out;
  }
  return value;
}

import { generateModelFilename } from '../utils/modelNaming';
import { rasterizeIcons, type RasterizedIcon } from '../utils/exportIconRaster';
import { stripXmlForbidden } from '../utils/xmlText';
import { nativizePackage } from './pptxNativeShapes';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  compactEmptyGutters,
  clampedBoxes,
  computeBounds,
  computeContentBounds,
  computeFitTransform,
  fitLabelToLines,
  fitLabelToWidth,
  metaSubline,
  partitionBoxes,
  stripHash,
  truncateLabel,
  widestGlyphIn,
  widestGlyphUpperIn,
  drawableInColumn,
  advanceWidthIn,
  trailingWhitespaceIn,
  usedConnectionLegend,
  workflowListFromEdges,
  narrateEdgeCallouts,
  readEdgeLabel,
  zoneStyleFor,
  readableTextOn,
  carriesWording,
  singleLineName,
  type Bounds,
  type ExportBox,
  type ExportRoute,
  type FitTransform,
  type Point,
} from './diagramExportGeometry';
import { readStepNumber as readStepValue } from '../utils/workflowStepMapping';

// ─── Theme palettes ───────────────────────────────────────────────────────────

interface SlideTheme {
  bg: string;
  headerBg: string;
  accent: string;
  titleText: string;
  metaText: string;
  footerText: string;
}

const DARK_THEME: SlideTheme = {
  bg: '1e293b',       // slate-800
  headerBg: '0f172a', // slate-900
  accent: '0078d4',   // Azure blue
  titleText: 'ffffff',
  metaText: '94a3b8', // slate-400
  // slate-400: slate-600 read at 1.93:1 on this background, so the attribution
  // line was effectively invisible in the exported deck.
  footerText: '94a3b8',
};

const LIGHT_THEME: SlideTheme = {
  bg: 'f8fafc',       // slate-50
  headerBg: 'e2e8f0', // slate-200
  accent: '0078d4',   // Azure blue
  titleText: '0f172a', // slate-900
  metaText: '475569',  // slate-600
  // slate-500: slate-400 read at 2.45:1 here, well under the WCAG AA bar.
  footerText: '64748b',
};

// ─── Slide layout (inches) ───────────────────────────────────────────────────
//
// The deck is normally 16:9 (13.333" × 7.5"). A diagram that would have to be
// squeezed below legible size on that canvas gets a larger custom slide
// instead — PowerPoint accepts any page up to 56", and a bigger page keeps the
// shapes at their true 96 dpi size so labels stay readable and editable.

const PX_PER_IN = 96;
/**
 * Type below this is grey ink rather than small words: a reader cannot resolve
 * it on a projected slide, and it makes the drawing harder to read, not more
 * informative. Only the overview thumbnail is allowed to reach it, and there it
 * drops the wording instead, because every one of those strings is legible on
 * the detail slide that follows.
 */
const OVERVIEW_LEGIBLE_PT = 6;
/**
 * Inset a connector chip reserves on each side, in INCHES.
 *
 * The sizer has always modelled 0.06in a side, but the shape was emitted with
 * pptxgenjs margin: 0.02 — and that option is in POINTS, so the file carried
 * a 0.0003in inset against a model reserving 0.12in in total. The model was
 * not wrong, the file was: every wrap decision here assumes this much padding,
 * so emit it ( * 72 to convert) rather than weaken the model to match a
 * typo. Text now wraps in PowerPoint where the sizer says it wraps.
 */
const CHIP_INSET_IN = 0.06;
/**
 * The column a chip's words actually get. Floored at one hair rather than at a
 * constant 0.05in: a fixed floor on a box that shrinks with the drawing is how
 * a 0.009in chip was told it had 0.05in of line, counted five of them, and was
 * emitted 0.6in tall — sixty times the tile it labelled.
 */
const chipColumn = (width: number): number => Math.max(0.001, width - CHIP_INSET_IN * 2);
const BASE_W = 13.333;
const BASE_H = 7.5;
const MAX_SLIDE_IN = 56; // PowerPoint's hard page-size limit
const ACCENT_H = 0.07;
const HEADER_H = 0.83;
const HEADER_END = ACCENT_H + HEADER_H;   // 0.9"
const SEP_H = 0.04;
const IMAGE_Y = HEADER_END + SEP_H + 0.06; // ~1.0"
const IMAGE_X = 0.2;
const FOOTER_H = 0.28;

// Fixed 16:9 geometry for the multi-slide architecture deck (title, services,
// validation, cost). Only the single-slide diagram export grows its page.
const W = BASE_W;
const IMAGE_W = W - IMAGE_X * 2;
const FOOTER_Y = BASE_H - FOOTER_H - 0.08;
const IMAGE_H = FOOTER_Y - IMAGE_Y - 0.1;

interface SlideGeometry {
  w: number;
  h: number;
  frame: DiagramFrame;
  footerY: number;
  /** True when the drawing is wider/taller than PowerPoint's 56" page limit. */
  overflow: boolean;
  /** True when far-placed nodes were pulled back onto the page to stay visible. */
  outliersClamped: boolean;
  /**
   * True when the callout bar this drawing's step numbers ask for is out of
   * reach in this frame, so the planner fell back to {@link MARKABLE_TILE_W_IN}.
   * See {@link calloutBarReachable}.
   */
  calloutBarClamped: boolean;
  /**
   * The median authored service width the planner aimed at, in pixels.
   *
   * Exported alongside the clamp for the same reason it is: the gate's own
   * median spanned every node in the scenario, groups included, while this one
   * comes from `partitionBoxes(...).services`. A zone rectangle is wider than
   * a service, so two of them dragged the gate's median above every tile on the
   * drawing and switched the callout rule off deck-wide - and on a landing zone
   * diagram, where subscription, VNet and subnet frames outnumber the services
   * inside them, the gate's median was a zone width and every service was
   * "below the median".
   */
  medianServiceW: number;
  /** See {@link calloutPlanFor}. The width the winning plan served, in pixels. */
  servedTileW: number;
  /**
   * See {@link calloutPlanFor}. The narrowest authored tile this frame can
   * carry a proportionate callout on at any grid, in pixels.
   */
  reachableTileW: number;
  /** See {@link calloutPlanFor}. Whether a finer plan was available and affordable. */
  chaseAffordable: boolean;
  /**
   * Tiles of the drawing, one per slide. A single entry (the usual case) means
   * the whole architecture fits on one legible page. More than one means the
   * drawing was too large to stay readable and was split the way a printed
   * Azure Architecture Center diagram is continued across pages.
   */
  windows: DiagramWindow[];
}

/**
 * One tile of a split drawing.
 *
 * `fit` sizes the page and is the same shape on every part, so the whole deck
 * renders at one scale. `own` says which shapes this part is responsible for
 * and covers the drawing exactly once between all the parts — it is wider than
 * `fit` wherever a neighbouring cell held no service and was not worth its own
 * slide, so nothing anchored there is left belonging to no part at all.
 */
interface DiagramWindow {
  fit: Bounds;
  own: Bounds;
}

/**
 * Smallest tile label PowerPoint may render. `addNodeShape` derives the label
 * size from the tile height (`h * 12`), so a legible tile needs at least
 * `LEGIBLE_TILE_PT / 12` inches of height. Anything below this is unreadable on
 * a projector and forces the recipient to redraw the deck by hand.
 */
// What a chip that reads as somebody else's label costs, in the same units the
// placement walk scores covered tile area in. A misread chip is worse than an
// untidy one: the reader silently attaches the wording to the wrong arrow and
// never finds out, so it is priced above any overlap a walk is likely to see.
const LEGIBLE_TILE_PT = 7;
/**
 * The smallest tile that can carry an identifying mark.
 *
 * The tile's text column is `w - 0.06`, and the legibility rule the gate and the
 * renderer share asks for two of the string's widest glyph. A digit at
 * `LEGIBLE_TILE_PT` measures 0.0524in, so the column must reach 0.1048in and the
 * tile 0.1648in, rounded up once to leave the arithmetic room to move.
 *
 * WIDTH ONLY, deliberately. The analogous height bar - about 0.375in, the room
 * one 7pt line needs in a band that is roughly a third of the tile - is reached
 * by any node under 38px tall, which is an ordinary size rather than a
 * pathological one, and demanding it moved the transform under scenarios the
 * planner had already sized correctly. 19.2px wide is not an ordinary size.
 */
const MARKABLE_TILE_W_IN = 0.2;
/**
 * Floor for the SKU / region / price sub-line. Below this the string is there
 * but nobody can read it, which is worse than an honest ellipsis.
 */
const META_LEGIBLE_PT = 5;

/**
 * The label size tiling actually aims for.
 *
 * Splitting to exactly the floor leaves the deck permanently at its worst
 * acceptable size — 7pt labels on 1.19" tiles, which is "not an issue" and
 * still not something anyone wants to present. Aim higher and fall back to the
 * floor only when the extra sheets would blow the slide budget.
 */
const COMFORTABLE_TILE_PT = 8.6;

/**
 * Below this, an extra sheet buys size at the cost of the reader losing the
 * architecture: a slide carrying two services out of twelve shows no system,
 * only a fragment.
 */
const MIN_SERVICES_PER_SLIDE = 6;

/**
 * Never explode one architecture into an unreviewable pile of slides.
 *
 * A 3 x 3 grid is where splitting stops being obviously better than one larger
 * sheet, so it is the point past which the *comfortable* grid is not worth
 * taking. It is not the hard limit: PowerPoint has exactly one page size per
 * deck, so growing past 13.333" drags the title and workflow slides onto a
 * plotter sheet nobody can open, and a 56" page tiled into four 56" parts is
 * no better than one. Beyond the comfortable grid, more standard slides is
 * still the lesser evil.
 */
const MAX_DIAGRAM_SLIDES = 9;

/**
 * Hard ceiling on the tiled deck. A drawing that needs more sheets than this
 * to stay above the legibility floor is genuinely a plotter drawing, and the
 * page grows for it.
 */
const MAX_TILED_SLIDES = 24;

/**
 * The same ceiling for a deck whose page size is fixed, where the alternative
 * to another slide is not a bigger sheet but smaller type.
 *
 * Counted in slides actually emitted, not grid cells: a sparse drawing needs a
 * fine grid to reach seven points but fills few of its cells, and charging it
 * for the empty ones is what refused a readable thirteen-slide deck in favour
 * of an unreadable one.
 *
 * Deliberately high. Nobody wants forty slides of one diagram, but everybody
 * would rather have forty they can read than twenty-four they cannot, and the
 * only drawings that get anywhere near this are the ones that would otherwise
 * have shipped at four points.
 */
const MAX_LEGIBLE_TILED_SLIDES = 48;

/**
 * And the length past which a deck has stopped being a deck at all, applied
 * only once legibility has already been given its way.
 *
 * The ceiling above is a preference — it yields to type that reads. This one
 * does not, because at some size a drawing is simply too large for a fixed
 * page and the honest answer is the most readable deck of a finite length,
 * not an unbounded one. Set where a reader would abandon the deck rather than
 * where a designer would.
 */
const MAX_FIXED_PAGE_SLIDES = 120;

/**
 * And a bound on the grid itself, so the search for a grid that fits the deck
 * ceiling cannot walk a pathological one.
 *
 * This is a bound on the *search*, not a judgement about the drawing, and
 * setting it as if it were the latter is what pinned a 27-service cascade to a
 * 56in page: a diagonal occupies one cell per service, so it needs an n x n
 * grid to be read at all, and 400 cells ran out at 20. Emitted slides are what
 * a reader counts and `MAX_LEGIBLE_TILED_SLIDES` is what limits them; empty
 * cells cost nothing but the cost of enumerating them, which is what this
 * number is actually for.
 */
const MAX_TILED_CELLS = 22500;

/**
 * A sheet has to carry a piece of the architecture, not a lone tile floating
 * on white. Below this the split has stopped adding information.
 */
const MIN_SERVICES_PER_TILED_SLIDE = 1.5;

/**
 * How far past its own window a slide keeps drawing.
 *
 * A connector chip is anchored at the middle of its arrow but is drawn around
 * that point, so a chip near a seam is half outside the window that owns it.
 * Clamping it back inside slid it on top of a service tile. The window that
 * owns a shape therefore renders a margin beyond its bounds — wide enough for
 * half a chip — while ownership itself stays exact, so nothing is drawn twice.
 */
const WINDOW_BLEED_PX = 100;

/**
 * What an inch of travel away from a chip's own arrow costs, in square inches
 * of overlap it is allowed to avoid by moving. This is only a tie-breaker that
 * keeps a chip from drifting for no gain — the hard bounds on the walk are what
 * keep a label attributable — so it is priced well under the area of the tile a
 * chip would otherwise be sitting on.
 */
const DRIFT_COST_PER_IN = 0.02;

/**
 * How much worse it is to cover another annotation than to cover a service
 * tile. A tile under a chip is untidy and still recognisable; a label under a
 * chip is simply lost, and so is the numbered callout that hangs off it.
 */
const ANNOTATION_WEIGHT = 24;

/**
 * How much worse it is to lie across somebody else's arrow than across a tile.
 * A chip on a foreign line is read as that line's label — the reader matches it
 * to the wrong hop and never finds out — but the line is thin, so per square
 * inch it has to cost more than a tile or the walk would always prefer it.
 */
const RIBBON_WEIGHT = 4;

/**
 * Something a connector label has to keep clear of. `annotation` marks the
 * blocks belonging to other labels, which are worth far more than the service
 * tiles they sit between.
 */
interface Obstacle {
  x: number;
  y: number;
  w: number;
  h: number;
  annotation?: boolean;
  /** A service name. Covering one is a different kind of damage from a tile. */
  caption?: boolean;
  /** Bundle this obstacle belongs to; its own labels ignore it. */
  owner?: string;
  /** Cost multiplier per square inch covered; defaults to a tile's 1. */
  weight?: number;
  /** Service this obstacle is part of, so a label can tell whose tile it is. */
  node?: string;
}

/**
 * Shortest workflow row that still fits a 12 pt sentence next to a badge. Rows
 * stop shrinking here and the list continues on another slide instead.
 */
const MIN_WORKFLOW_ROW_IN = 0.34;
// The workflow list's preferred and floor type sizes. A step sentence shrinks
// between them to fit its row rather than spilling out of it.
const WORKFLOW_ROW_PT = 12;
const WORKFLOW_MIN_PT = 9;

/**
 * How much of a chip may end up over a service name or another callout before
 * the chip is dropped in favour of a numbered callout plus a step-list row.
 *
 * Set just under the smallest overlap the export audit reports (0.018 sq in),
 * so a chip is handed over before it is drawn on top of something the reader
 * needs rather than after.
 */
const SPOILED_CHIP_SQ_IN = 0.015;

/**
 * Whether a chip at `block` stands on something the reader needs — a name, a
 * numbered callout, or a service that is not at either end of its own arrow.
 *
 * One implementation, called from every place a chip is seated. It used to be
 * written inline at the first seat only, and the two repair passes that move a
 * chip afterwards re-ran the placement search without it: a chip rejected for
 * covering a caption could be moved onto a different caption and kept, because
 * the only thing the repair asked about the new seat was whether it hit
 * another chip. That is how a chip came to cover 32% of a tile name on a slide
 * whose first-pass placement had been correctly refused.
 *
 * A chip standing on a service that is not at either end of its own arrow is
 * the same failure by a different route: the reader takes it for that
 * service's caption and the hop it actually describes goes unlabelled. The
 * walk cannot always avoid it — on `meta-subline` and `workflow-prose` every
 * slot within reach laps a bystander, and weighting them twelve times a tile
 * moved the chip not at all — so the wording is handed to the step list
 * exactly as it is when a caption is in the way.
 *
 * Priced separately and much more loosely than the caption bar. A chip may
 * brush a bystander's rim; the export audit allows a fiftieth of a tile, so
 * this bar sits at exactly that and hands the wording over the moment the
 * drawing would fail the gate. Priced as a FRACTION of the bystander, exactly
 * as the audit prices it, so the exporter and the gate can never disagree
 * about what counts as standing on a stranger. A flat area bar cannot match
 * it: a fiftieth of a small tile and of a large one are different numbers of
 * square inches, and the difference is what let a 2% lap ship while a 13% one
 * was muted.
 */
function chipSpoils(
  block: { x: number; y: number; w: number; h: number },
  route: { sourceId: string; targetId: string; bundleKey: string },
  obstacles: readonly Obstacle[],
): boolean {
  const STRANGER_TILE_FRACTION = 0.02;
  // A label may lean on the two services its own arrow connects — the reader
  // attributes it correctly, and on a hop shorter than the label there is
  // nowhere else for it to go. "Lean on" is not "cover": this test excluded
  // endpoint tiles ENTIRELY, so a chip could bury the icon it was pointing at
  // and the walk would not move it. Same tenth of the tile the gate allows, so
  // the placer and the thing that checks it cannot disagree.
  const OWN_TILE_FRACTION = 0.1;
  // A caption is judged by how much of ITSELF is gone, not by how many square
  // inches the overlap is. A sub-line is a 0.02in² sliver: burying 93% of
  // "Standard_D4s_v5 · japaneast" costs less area than `SPOILED_CHIP_SQ_IN`
  // allows, so an absolute budget scored total destruction as free — and a
  // chip that grew a line taller took exactly that free ride, across 45 tiles
  // in 3 decks.
  const SPOILED_CAPTION_FRACTION = 0.5;
  const ownEnds = new Set([route.sourceId, route.targetId]);
  let spoiled = 0;
  let onStrangers = 0;
  let onOwnEnds = 0;
  let onCaptions = 0;
  for (const o of obstacles) {
    const tile = !o.annotation && !o.caption && o.node !== undefined;
    const stranger = tile && !ownEnds.has(o.node as string);
    if (!o.annotation && !o.caption && !tile) continue;
    if (o.owner !== undefined && o.owner === route.bundleKey) continue;
    const dx = Math.min(block.x + block.w, o.x + o.w) - Math.max(block.x, o.x);
    const dy = Math.min(block.y + block.h, o.y + o.h) - Math.max(block.y, o.y);
    if (dx > 0 && dy > 0) {
      const fraction = (dx * dy) / Math.max(1e-6, o.w * o.h);
      if (stranger) onStrangers = Math.max(onStrangers, fraction);
      else if (tile) onOwnEnds = Math.max(onOwnEnds, fraction);
      else {
        spoiled += dx * dy;
        if (o.caption) onCaptions = Math.max(onCaptions, fraction);
      }
    }
  }
  return spoiled > SPOILED_CHIP_SQ_IN
    || onStrangers > STRANGER_TILE_FRACTION
    || onOwnEnds > OWN_TILE_FRACTION
    || onCaptions > SPOILED_CAPTION_FRACTION;
}

/**
 * The largest inches-per-pixel the renderer will actually draw a window at.
 *
 * The natural cap is one authored pixel to one screen pixel: a two-node
 * diagram must not be blown up to absurd tiles. But a density is the wrong
 * thing to cap when the drawing itself was authored small, so a drawing whose
 * narrowest tile cannot carry a mark at all raises the cap far enough to reach
 * that bar and no further. Every drawing already above it keeps the identical
 * transform.
 *
 * One function because there were two copies of this expression and they had
 * already diverged: the planner still capped at `1 / PX_PER_IN` and its comment
 * asserted the renderer did too, while the renderer had been raising it for
 * sub-19.2px tiles. `Infinity` for "no tile to protect" - the caller has no
 * services - reduces to the natural cap.
 */
/**
 * The narrowest tile that will actually be drawn, in authored pixels.
 *
 * Written once and called from both the planner and the renderer because
 * they have to agree: the ceiling one of them raises is the ceiling the other
 * one plans against, and two spellings of "narrowest" is exactly how they
 * drifted apart the first time. Zero-width boxes are skipped rather than
 * treated as the minimum - a box with no width is not a tile a mark has to
 * fit inside, and counting it would pin the ceiling to a shape nobody sees.
 *
 * A loop, not `Math.min(...boxes)`: spreading an array into a call is limited
 * by the engine's argument count, and this list is one per service on the
 * diagram, so a large estate would throw where a small one worked.
 */
export function narrowestBoxW(boxes: readonly { w: number }[]): number {
  let narrowest = Infinity;
  for (const box of boxes) {
    if (box.w > 0 && box.w < narrowest) narrowest = box.w;
  }
  return narrowest;
}

function rendererMaxScale(minBoxW: number, markIn: number = MARKABLE_TILE_W_IN): number {
  return Math.max(
    1 / PX_PER_IN,
    Number.isFinite(minBoxW) && minBoxW > 0 ? markIn / minBoxW : 0,
  );
}

/**
 * The median width and height of a set of tiles, in authored pixels.
 *
 * One expression, because the planner's two targets and the reachability
 * predicate below all need the same statistic and a second copy of it is what
 * let the gate mirror the planner with a MINIMUM while the planner used a
 * median - a disagreement that made the gate accuse a correct 53 slide plan of
 * chasing a bar it had already reached.
 */
function medianExtent(boxes: readonly { w: number; h: number }[]): { w: number; h: number } {
  const widths = boxes.map((b) => b.w).filter((w) => w > 0).sort((a, b) => a - b);
  const heights = boxes.map((b) => b.h).filter((h) => h > 0).sort((a, b) => a - b);
  return {
    w: widths[Math.floor(widths.length * 0.5)] ?? Infinity,
    h: heights[Math.floor(heights.length * 0.5)] ?? Infinity,
  };
}

/**
 * Whether `markIn` is a tile width some grid in this frame can actually reach.
 *
 * `legibleScaleFor` returns at most `finestPerIn`, so a demand above it is
 * unreachable by that function's own return value, whatever grid is tried.
 * This is the exact bound and not a bound on it: an earlier form compared
 * against the same expression with the `target` term dropped, which is larger
 * by `(2 * WINDOW_BLEED_PX + target) / (2 * WINDOW_BLEED_PX)`, and left a band
 * about one authored pixel wide in which the bar is unreachable and unclamped.
 * Measured at every digit count - 10px, 13px and 16px - and the three-digit
 * case tripled a deck: 15px planned 21 slides, 16px planned 64, and the tile
 * 16px bought was 0.4274in against the 0.4457in bar it was spent chasing.
 */
function calloutBarReachable(
  target: number,
  frame: { w: number; h: number },
  minBoxW: number,
  markIn: number,
): boolean {
  const finestPerIn = Math.min(frame.w, frame.h) / (WINDOW_BLEED_PX * 2 + Math.max(1, target));
  return rendererMaxScale(minBoxW, markIn) <= finestPerIn;
}

/**
 * The tile width the planner should aim for, given what the drawing carries.
 *
 * `MARKABLE_TILE_W_IN` is the bar for a tile that can hold an identifying
 * mark, and for an unnumbered drawing that is the whole requirement. A
 * NUMBERED drawing has a second one: every hop carries a disc that must be
 * readable and must not swamp the service it points at, and those two
 * demands only intersect above `floor / 0.55`. Below it the exporter has no
 * move left - it draws the floor and the disc is disproportionate whatever it
 * chooses - so the planner's own success condition was the badge rule's
 * failure condition, and a deck could sit exactly on the markable bar with
 * every callout at 97% of its tile and pass.
 *
 * Taken from the LARGEST step number, because a three-digit disc is 58% wider
 * than a one-digit one and the bar has to hold for the widest number drawn.
 *
 * This only binds on a drawing authored small: `rendererMaxScale` is the
 * larger of this and natural size, so it moves nothing until the median tile
 * is under `96 * markIn` pixels - about 34px for a two-digit deck, against
 * 19.2px before.
 */
function markableTileWIn(steps: Iterable<number | undefined | null>): number {
  let widest = 0;
  for (const raw of steps) {
    const step = Number(raw);
    if (Number.isFinite(step) && step > widest) widest = step;
  }
  if (widest <= 0) return MARKABLE_TILE_W_IN;
  return Math.max(
    MARKABLE_TILE_W_IN,
    badgeFloorIn(widest, BADGE_LEGIBLE_PT) / BADGE_TILE_SHARE,
  );
}

/** The step numbers a drawing's edges carry, for {@link markableTileWIn}. */
function stepNumbersOf(
  edges: readonly { data?: { stepNumber?: unknown } }[],
): (number | undefined)[] {
  return edges.map((edge) => {
    const step = Number(edge?.data?.stepNumber);
    return Number.isFinite(step) ? step : undefined;
  });
}

/**
 * Inches-per-pixel worth splitting the drawing to reach, for a `target`-pixel
 * tile in a `frame`-inch window.
 *
 * Two ceilings, and getting either wrong has produced the same catastrophe from
 * opposite directions, because both coarsening loops break on
 * `scaleOf(c, r) >= legibleScale && scaleOf(next) < legibleScale`. If nothing
 * reachable ever reaches `legibleScale`, the left conjunct is false at every
 * step, the break never fires, and the loop walks past every grid that reads —
 * so a demand that cannot be met is not a legibility floor, it is a way of
 * switching the floor off.
 *
 * `finestPerIn` is what this *frame* can deliver: `gridFor` returns null the
 * moment the bleed alone fills the window, and a null grid sends the planner to
 * `capped(150, 150)`. Missing it, 400 services came out as 49 slides on which
 * every tile read "Azure…".
 *
 * `1 / PX_PER_IN` is what the *renderer* will: every window is drawn through
 * `computeFitTransform(..., { maxScale: 1 / PX_PER_IN })`, so a tile can never
 * be larger than the size it was authored at, while `LEGIBLE_TILE_PT / 12 /
 * target` exceeds that for any tile under 56px. Missing it, 60 services
 * authored 20px tall came out as 61 slides carrying one tile each on a page
 * 0.3% inked — with tiles no wider, type no larger and no name any more
 * complete than the 25 slides they needed.
 *
 * That ceiling is not a constant, and the comment above asserting it was one
 * went stale the day the renderer started raising it for drawings authored too
 * small to carry a mark. The planner then stopped splitting at a scale the
 * renderer would have exceeded, which is the same defect in the opposite
 * direction: tiles under the markable bar on a deck that had slides to spare.
 * Both now read `rendererMaxScale`, so there is one expression and no second
 * place to update.
 *
 * Exported so the invariant can be asserted directly. The end-to-end audit can
 * only see the catastrophic end of this: between 40 and 55 authored pixels the
 * deck over-tiles by 24-48% while every window still carries two tiles, and no
 * property of the emitted file distinguishes that from a small correct deck.
 * The distinguishing fact is a counterfactual — a coarser split would have
 * produced identical tiles — so it has to be checked here, on the function.
 */
export function legibleScaleFor(
  target: number,
  frame: { w: number; h: number },
  minBoxW: number = Infinity,
  markIn: number = MARKABLE_TILE_W_IN,
): number {
  const finestPerIn = Math.min(frame.w, frame.h) / (WINDOW_BLEED_PX * 2 + Math.max(1, target));
  // Chase the callout bar only while the frame can actually deliver it.
  //
  // `finestPerIn` is what this frame can return, so a demand above it is met by
  // no grid at all - and since both coarsening loops break on
  // `scaleOf(c, r) >= legibleScale`, a demand that is never met is not a floor,
  // it is a way of switching the floor off: the loop walks past every grid that
  // reads and stops only when it runs out of columns.
  //
  // Measured, this is what made numbering cost 4.6x the slides. One drawing of
  // 14px tiles needed 12 windows unnumbered, 33 at one digit, 44 at two and 55
  // at three - the architecture never changed, the step numbers did - and at
  // the end of all that spending the tile was 0.3740in against a 0.4457in bar,
  // so the deck paid 43 extra windows for a target it was never going to reach,
  // and missed it anyway.
  //
  // When the callout bar is out of reach, fall back to the bar that predates
  // numbering: a tile wide enough to carry an identifying mark. That is
  // reachable, so the break fires, and the deck spends what an unnumbered
  // drawing of the same architecture spends. The disc is then disproportionate
  // - there is no scale at which it is not - but the reader gets the same
  // number of legible sheets instead of 4.6 times as many illegible ones.
  const barIn = calloutBarReachable(target, frame, minBoxW, markIn)
    ? markIn
    : MARKABLE_TILE_W_IN;
  const demand = rendererMaxScale(minBoxW, barIn);
  // The type ceiling does not get to cancel the markable bar: they measure
  // different axes.
  //
  // `LEGIBLE_TILE_PT / 12 / target` is derived from the median tile's HEIGHT
  // and says only that a taller tile's caption gains nothing from a finer
  // grid. The markable demand is derived from its WIDTH and says the tile
  // cannot carry a mark at all. Combining them with `min` made the height term
  // veto the width one on every tall narrow tile - and that is precisely the
  // shape the markable raise exists for. Measured on a deck of 24x96 sensors,
  // `min` returned 0.006076 in/px for a 0.2829in bar and for a 0.2000in bar
  // alike, an identical answer to two different questions: the tile drew
  // 0.1458in wide and carried a disc at 89% of itself, while the very same
  // 24px tile reached exactly 0.2829in on decks whose tiles were not tall.
  // The renderer's ceiling had the room the whole time; the `min` threw it out.
  //
  // Still bounded above by `demand` and `finestPerIn`, which are true ceilings,
  // so nothing here can chase a scale the renderer or the frame would refuse.
  // And narrow by construction, in two independent ways. It is zero unless the
  // renderer actually RAISED its ceiling for this tile - a tile already wide
  // enough to carry a mark asks for nothing here, so passing `minBoxW` in can
  // never move an answer the planner would have given without it, which
  // `pptxSlideBanding` asserts directly across a 400-target sweep. And even
  // when it is non-zero the max only binds on a tile roughly three times taller
  // than it is wide, so ordinary drawings plan exactly as before.
  const markableDemand = Number.isFinite(minBoxW)
    && minBoxW > 0
    && minBoxW < barIn * PX_PER_IN - 1e-9
    ? barIn / minBoxW
    : 0;
  const typeCeiling = LEGIBLE_TILE_PT / 12 / Math.max(1, target);
  return Math.min(finestPerIn, demand, Math.max(typeCeiling, markableDemand));
}
/**
 * Split the drawing into as few standard-slide windows as keep tiles legible.
 *
 * PowerPoint allows exactly one page size per deck, so every window shares the
 * same slide geometry and the reader moves through them in reading order:
 * left to right, then down. Returns a single full window whenever the diagram
 * already fits, which keeps the common path byte-identical.
 *
 * `legible` says whether the frame can show the drawing at a readable size at
 * all — either whole or tiled. When it cannot, the caller grows the page
 * instead, because splitting further only multiplies slides.
 */
/**
 * Plan the windows, then check that the raise paid for itself.
 *
 * The ceiling is raised so that a drawing of small tiles can be split until
 * its tiles reach the markable bar. What it must not do is split a drawing
 * until each slide holds one tile: `probe-whitespace` chases four slivers and
 * drags their 160px neighbours to 2.3in each, ending with six of eight slides
 * carrying a single service. Both harms are real, and only one of them was
 * ever visible to the gate, so the earlier attempt to tell them apart by a
 * max-over-min ratio chose the invisible one - and a sweep found the ratio has
 * a one pixel cliff, where widening one node of sixty from 56px to 57px took
 * the deck from sixty named services to one.
 *
 * So the raise is not predicted from a proxy, it is measured on the plan it
 * produces. If the raised grid averages fewer than `MIN_SERVICES_PER_TILED_SLIDE`
 * services per window - the constant the fixed-page path already uses for
 * exactly this judgement - the raise is refused and the unraised plan stands.
 */
function planDiagramWindows(
  bounds: Bounds,
  services: ExportBox[],
  frame: DiagramFrame,
  options: { mustTile?: boolean; markIn?: number; serveW?: number } = {},
): { windows: DiagramWindow[]; legible: boolean; servedW: number; chaseAffordable: boolean } {
  const affordable = (plan: { windows: DiagramWindow[] }): boolean => plan.windows.length > 0
    && services.length / plan.windows.length >= MIN_SERVICES_PER_TILED_SLIDE;
  const widths = services.map((box) => box.w).filter((w) => w > 0).sort((a, b) => a - b);
  const medianW = widths[Math.floor(widths.length * 0.5)] ?? Infinity;
  // Serve the NARROWEST numbered tile first, when the deck can afford to.
  //
  // The median is the right target for a drawing at large - an extremum has a
  // neighbour, and the paragraphs in `planWindowsAtCeiling` are the record of
  // what taking one costs. But "the planner declined to serve this tile" was
  // being used to excuse every disproportionate callout below the median, and
  // `sorted[floor(n/2)]` puts up to HALF the drawing below that line by
  // construction: on four services of 150, 24, 150, 24 authored px it excused
  // two of the four and all three hops, when serving the 24px tiles cost
  // exactly one extra window and widened every other tile on the deck by 39%.
  //
  // So the cost is measured rather than assumed, on the same density floor that
  // already decides whether the median raise was worth its slides. Where the
  // finer plan clears the floor the deck pays and the callouts fit; where it
  // does not - `probe-whitespace` numbered, six services over six windows - the
  // refusal stands, and now stands on a measurement.
  // The scale a plan actually delivers, read off the plan rather than
  // recomputed from a grid this function never sees.
  //
  // Every window renders `fit` plus the same bleed on all four sides and the
  // whole deck shares one scale, so the binding window is the largest one and
  // the binding axis is whichever fits worst - the same expression
  // `planWindowsAtCeiling` uses internally, evaluated on its output.
  const perInOf = (plan: { windows: DiagramWindow[] }): number => {    if (plan.windows.length === 0) {
      return Math.min(
        frame.w / Math.max(1, bounds.maxX - bounds.minX),
        frame.h / Math.max(1, bounds.maxY - bounds.minY),
      );
    }
    const spanW = Math.max(...plan.windows.map((w) => w.fit.maxX - w.fit.minX));
    const spanH = Math.max(...plan.windows.map((w) => w.fit.maxY - w.fit.minY));
    return Math.min(
      frame.w / (spanW + WINDOW_BLEED_PX * 2),
      frame.h / (spanH + WINDOW_BLEED_PX * 2),
    );
  };
  const raised = planWindowsAtCeiling(bounds, services, frame, options, true);
  // The density floor may not choose a swamped icon, even when there is
  // nothing NARROWER to chase.
  //
  // `MIN_SERVICES_PER_TILED_SLIDE` refuses a plan that carries too few services
  // a window, and that is the right call between two readable decks. On a
  // drawing whose numbered tiles are all the same width it was also deciding
  // whether the callout swamps its service, which it has nothing to say about:
  // six services over six windows is 1.0, under the floor, so the plan that
  // reaches the bar was thrown away and four discs shipped at 98% of the
  // services they number.
  //
  // Taken only when the frame can actually deliver the bar and the forced plan
  // actually reaches it, so the deck never buys windows for a target it still
  // misses. `MAX_TILED_SLIDES` still caps the result, and the requirement is
  // strictly FINER - which is what keeps this clear of the coarse bail-out that
  // took a 120 node estate from 53 windows of 0.3130in to 21 of 0.2003in.
  const bar = options.markIn ?? 0;
  const serve = options.serveW ?? 0;
  // Two lines here as well, `markIn` as the want and the floor as the
  // must-not-ship-below. The waiver read `bar` on both tests, which is the
  // third site of the same 1.82x error: a forced plan reaching 0.20793in for a
  // 10px tile against a 0.15559in disc has RESOLVED the defect and was thrown
  // away for missing an ideal it was never required to meet, shipping 0.113in
  // tiles with four of six services anonymous instead.
  // How many services can draw a MARK at this plan's scale.
  //
  // The escapes below used to compare disc proportion, which is a fact about a
  // callout, to decide a question whose visible consequence is whether services
  // have names at all. On six services with one 14px icon the unraised plan
  // cleared the disc line by 0.0028in - 1.8% - and four services went anonymous
  // on that margin, while the raised plan named all six. The mark bar for a
  // one-character key sits between a 0.1584in tile and a 0.1697in one, so the
  // deck needed 7% more scale and the only alternative on the menu was 84%
  // more; nothing ever asked for the coarsest grid that clears the bar. This
  // asks the machinery that actually decides - the same `drawableInColumn`
  // against the same column inset the tile itself uses - so there is no new
  // constant and no third copy of the bar.
  const markableCountAt = (perIn: number): number => services.filter((s) => drawableInColumn(
    '1',
    LEGIBLE_TILE_PT,
    Math.max(0.05, s.w * perIn - 0.06),
  )).length;
  const markableAt = (plan: { windows: DiagramWindow[] }): number => markableCountAt(perInOf(plan));
  // The coarsest plan that still names as many services as the fine one.
  //
  // The two candidates on the menu are extremes - measured at 0.011314 and
  // 0.020793 in/px, 84% apart - and the deck that lost four names to the gap
  // needed 7.1%. Nothing ever asked for the scale in between, so "name the
  // services" and "do not shred the deck into one tile a slide" looked like
  // opposites: taking the fine plan named all six and cost 9 slides, 6 of them
  // single-tile, 5 oversized edge chips and a label cut to "Ze...". They are
  // not opposites. `perIn` is continuous and the mark bar is a threshold on it,
  // so the cheapest plan that clears the threshold is found by bisecting the
  // scale, not by choosing an end.
  const coarsestNaming = (
    fine: { windows: DiagramWindow[]; legible: boolean },
    floorPlan: { windows: DiagramWindow[] },
  ): { windows: DiagramWindow[]; legible: boolean } => {
    const want = markableAt(fine);
    let lo = perInOf(floorPlan);
    let hi = perInOf(fine);
    // No badge in the drawing is not a reason to buy the expensive plan. The
    // first draft asked the planner for a tile WIDTH, which had to be derived
    // from the callout bar, so it stood down on a deck with no callouts at all
    // - and that deck went on shipping 8 slides, 6 of them a single tile, for
    // a naming gain it could have had on four. The target is a scale now, and
    // a scale is well defined whether or not anything is numbered.
    if (!(hi > lo)) return fine;
    if (markableCountAt(lo) < want) {
      for (let i = 0; i < 24; i += 1) {
        const mid = (lo + hi) / 2;
        if (markableCountAt(mid) >= want) hi = mid; else lo = mid;
      }
    } else {
      hi = lo;
    }
    if (hi >= perInOf(fine) - 1e-9) return fine;
    // Built from the PLAIN side and raised to the target, not from the fine
    // side and relaxed toward it. Asking the ceiling-raised planner for a
    // coarser grid gets the raised grid back unchanged - it is already the
    // finest rung - so the whole bisection returned the extreme it was written
    // to avoid. The cheap plan is the one that needs 4% more scale.
    const eased = planWindowsAtCeiling(bounds, services, frame, { ...options, namePerIn: hi }, false);
    // Accepted only when it keeps every name AND costs less than the extreme.
    // A bisection on scale says nothing about what grid the planner can build
    // at that scale, so the plan it returns has to be re-measured, not assumed.
    return eased.windows.length > 0
      && eased.windows.length <= fine.windows.length
      && markableAt(eased) >= want
      ? eased
      : fine;
  };
  const defectLine = bar * BADGE_TILE_SHARE;
  // A plan that puts most of its slides on one service each has stopped buying
  // anything: splitting enlarges a tile by giving it more of the frame, and a
  // tile alone on a slide is already at its natural width. Waiving the density
  // floor removes the only thing that was watching for this, so the waiver
  // carries its own bound - measured, not assumed: the unbounded form took a
  // 21 service deck to 21 windows with 14 of them carrying a single tile.
  //
  // Loneliness alone is not the defect: a window holding one service is exactly
  // how a name gets bought when the tiles are too small to carry one. What
  // makes the loneliness wasted is that the drawing has ALREADY reached the
  // width a mark needs, because past `MARKABLE_TILE_W_IN` there is no further
  // mark to win at any scale and the extra slides are pagination. Measured, the
  // two populations do not overlap: the deck that spent thirteen extra windows
  // to rescue one name sat at 0.204in, and the deck that rescued four sat at
  // 0.188in.
  //
  // Asked of the DRAWING, not of the window. Scale is global, so a window
  // holding one ordinary 160px service is not evidence of anything - its tile
  // is metres past the mark bar whatever the plan does, and charging the plan
  // for it took the deck that needed six windows to name six services back to
  // 0.113in tiles. The question is whether the narrowest tile on the page still
  // has something to gain.
  const narrowestServiceW = Math.min(
    ...services.map((s) => s.w).filter((w) => w > 0),
    Infinity,
  );
  const wastedShare = (plan: { windows: DiagramWindow[] }): number => {
    if (plan.windows.length === 0) return 0;
    if (!Number.isFinite(narrowestServiceW)
      || narrowestServiceW * perInOf(plan) < MARKABLE_TILE_W_IN) return 0;
    const lonely = plan.windows.filter((w) => services.filter((s) => {
      const cx = s.x + s.w / 2;
      const cy = s.y + s.h / 2;
      return cx >= w.fit.minX && cx <= w.fit.maxX && cy >= w.fit.minY && cy <= w.fit.maxY;
    }).length <= 1).length;
    return lonely / plan.windows.length;
  };
  // A lonely window is only wasted when it buys no names.
  //
  // The bound as first written rejected on raw loneliness, and so rejected a
  // plan that named all six services at 0.2829in because all six sat one to a
  // window - three services lost their names to it. `wastedShare` is what that
  // bound was reaching for: splitting past the point of enlargement is a defect
  // only where it is not the thing putting names on the canvas.
  //
  // There is no PROPORTION of names worth measuring here, in either direction.
  //
  // A share of `services.length` prices one name by how many other services
  // happen to be on the drawing, so the same rescue scored 0.2500 on a twelve
  // service deck and 0.2308 on a thirteen service one - and appending one
  // ordinary, perfectly drawn service to a twelve service diagram erased three
  // other services' names and halved every tile, 0.1896in to 0.0984in. Pricing
  // it as a share of the names AT RISK instead fixed that and then discriminated
  // nothing at all: `coarsestNaming` accepts a plan only when it keeps every
  // name the fine plan had, so the rescued share is pinned at 1, and it measured
  // exactly 1.0000 on all eight fixtures - including the deck that bought its
  // single remaining name with thirteen extra slides. A term that is constant
  // over its whole domain is not a bound.
  //
  // So the trade is judged on its two honest halves: it has to win a name, and
  // it may not spend most of the deck on slides that win nothing.
  const worthTheSplit = (
    plan: { windows: DiagramWindow[] },
    against: { windows: DiagramWindow[] },
  ): boolean => {
    if (markableAt(plan) <= markableAt(against)) return false;
    return wastedShare(plan) <= 0.5;
  };
  if (bar > 0 && serve > 0 && perInOf(raised) * serve < bar - 1e-6) {
    const forced = planWindowsAtCeiling(
      bounds, services, frame, { ...options, waiveDensity: true }, true, serve,
    );
    // Eased BEFORE the cost of the trade is judged, because `coarsestNaming`
    // changes which plan is returned and the bound belongs on the returned
    // plan. Measured on a thirteen service farm: `forced` is 10 windows at a
    // 0.700 lonely share and was refused for it, while the plan that would have
    // shipped is 8 windows at 0.250 - which clears the bound outright, and is
    // LESS lonely than the twelve service plan the same guard accepts. The
    // guard was anti-correlated with the quantity it names, and three services
    // went unnamed for it.
    //
    // Whether the chase is worth ATTEMPTING is still asked of `forced`, which
    // is the finest plan available and therefore the honest answer to "can
    // this drawing reach the callout line at all". Asking it of the eased plan
    // instead made the easing veto itself: the eased plan is by construction
    // coarser, so it fell under the line, the escape refused it, and all three
    // farms fell back to 0.10in tiles on 5 slides.
    const eased = forced.windows.length > 0 ? coarsestNaming(forced, raised) : forced;
    if (forced.windows.length > 0
      && worthTheSplit(eased, raised)
      && (perInOf(forced) * serve >= bar - 1e-6
        || perInOf(forced) * serve >= defectLine - 1e-6)) {
      return { ...eased, servedW: serve, chaseAffordable: true };
    }
  }
  // An empty window list with `legible: false` is not "the drawing fits", it is
  // "no grid in this frame reads" - the caller answers that by tiling under
  // `mustTile` or by growing the page, and a finer target belongs to that call,
  // not this one. Measured against it, a 12 window chase looked like an
  // improvement on nothing and pre-empted the 44 window forced plan that was
  // coming, taking a 120 node estate from 53 windows of 0.3130in tiles to 21
  // of 0.2003in and shipping 90 discs at 97% of the services they number.
  //
  // An empty list with `legible: true` is the opposite case and must fall
  // through: the drawing fits whole at the median, and whether it should be
  // split anyway to serve a narrower tile is exactly the question below.
  if (raised.windows.length === 0 && !raised.legible) return { ...raised, servedW: medianW, chaseAffordable: false };
  // Accepted only where it is genuinely FINER, as well as affordable. Asking
  // for a scale no grid inside the slide budget can deliver does not make
  // `planWindowsAtCeiling` try harder, it makes it bail to a coarse fallback -
  // and the density floor is delighted to accept one, because a coarse plan has
  // the best services-per-window ratio on the sheet. Measured: chasing a 14px
  // sliver across a 120 node estate bailed to 21 windows of 0.2003in tiles
  // where the median plan gave 53 windows of 0.3130in, cleared the floor at 5.7
  // services a window, and shipped 90 discs at 97% of the services they number.
  let chaseAffordable = false;
  // Attempted whenever the chase would BUY something, not when the narrowest
  // badged tile happens to beat an order statistic.
  //
  // The trigger was `serveW < medianW`, a strict comparison against
  // `sorted[floor(n/2)]`, so on any drawing where the narrowest badged tile
  // ties the median no finer plan was ever considered - and a step function on
  // an order statistic flips on a one-pixel authoring edit. Measured on the
  // same six services: widths [14,14,14,14,160,160] tie at 14 and plan 5 slides
  // at 0.158in; shaving ONE pixel off ONE icon to [13,14,...] plans 9 slides at
  // 0.283in, with five oversized edge chips and a label cut to "Ze...". Same
  // drawing to a reader, 1.8x the deck. The condition now asks whether the
  // raised plan already serves the bar for this tile, which is continuous in
  // the thing that matters and does not care where the median sits.
  const chaseWorthTrying = options.serveW !== undefined
    && options.serveW > 0
    && bar > 0
    && perInOf(raised) * options.serveW < bar * BADGE_TILE_SHARE - 1e-6;
  if (chaseWorthTrying && options.serveW !== undefined) {
    const finest = planWindowsAtCeiling(bounds, services, frame, options, true, options.serveW);
    // The density floor does not get to choose a swamped icon.
    //
    // `MIN_SERVICES_PER_TILED_SLIDE` exists to stop a deck degenerating into a
    // flip-book, and that is a preference between two readable decks. It was
    // also deciding a case it has nothing to say about: where refusing the
    // finer plan leaves the narrowest numbered tile under `markIn` - the width
    // at which a disc can be both readable and no wider than its service - the
    // refusal does not buy a denser deck, it buys a disc drawn as much as
    // twice the width of the icon it points at. Measured on twenty ordinary
    // services and one glyph on the chain: a 16px zone drew 0.0975in under a
    // 0.1556in disc, 160%, and a 12px one 213%, both refused by the density
    // floor while the frame was reaching 0.2978in and 0.2233in - roughly twice
    // what the disc needed. The page was never the constraint.
    //
    // Waived only when the finer plan actually resolves it, so the deck never
    // pays windows for a target it still misses, and only for the narrowest
    // BADGED tile, so an unnumbered sliver buys nothing.
    // Two lines, not one. `bar` is `markIn`, the width at which a disc can be
    // both readable and a proportionate 55% of its tile - what the deck WANTS.
    // `bar * BADGE_TILE_SHARE` is the floor itself, the width at which the disc
    // merely stops being wider than the service it numbers - what the deck must
    // not ship below. Keying the waiver on the first alone made it 1.82x too
    // strict and it refused every case it was written for: a 16px zone needed
    // 0.3528in of reach to satisfy `markIn` and the frame gave 0.2978in, so the
    // waiver said no - while the tile only needed 0.1556in to stop being
    // dwarfed, which that same frame cleared twice over. The identical
    // off-by-BADGE_TILE_SHARE error was live in the gate's exempt band.
    const reaches = (p: { windows: DiagramWindow[] }, line: number): boolean =>
      line > 0 && perInOf(p) * (options.serveW ?? 0) >= line - 1e-6;
    const mustChase = (!reaches(raised, bar) && reaches(finest, bar))
      || (!reaches(raised, bar * BADGE_TILE_SHARE) && reaches(finest, bar * BADGE_TILE_SHARE));
    // Recorded whether or not it is taken, because the audit has to distinguish
    // a callout the deck COULD have served from one it could not, and a flag
    // read off the plan that was chosen moves with any mutation that changes
    // the choice - which is exactly how the Visio magnifier went blind on its
    // own revert one round ago. This is the bound; servedW is the outcome.
    chaseAffordable = finest.windows.length >= raised.windows.length
      && (affordable(finest) || mustChase);
    if (chaseAffordable) return { ...finest, servedW: options.serveW, chaseAffordable };
  }
  if (raised.windows.length === 0) return { ...raised, servedW: medianW, chaseAffordable };
  if (affordable(raised)) return { ...raised, servedW: medianW, chaseAffordable };
  const plain = planWindowsAtCeiling(bounds, services, frame, options, false);
  // The density floor may not throw away a plan BECAUSE it is good.
  //
  // The escape above can only ever rescue a plan that was already under the
  // bar, so a raised plan that clears the bar walks past it - and arrives here,
  // where `affordable()` discards it for scoring 6 services over 6 windows
  // against a floor of 1.5. Measured on six services with one 14px icon: the
  // raised plan put the narrowest badged tile at 0.29110in, past the 0.28288
  // bar, and `plain` shipped at 0.1584in, 46% narrower, with four of the six
  // services drawn with no name at all. Being good was the only reason it had
  // no protection. Same shape as round 72's armed-implies-cannot-fire.
  //
  // So the comparison is made here, after affordability rather than in a guard
  // that is false whenever the raised plan is worth keeping, and only where the
  // unraised plan does NOT clear the same line - a refusal that costs nothing
  // in quality is still the density floor's to make.
  // The second escape reads the SAME two measures as the first.
  //
  // It sat after `affordable()`, so it fires only on plans already under the
  // density floor - exactly the population the first escape's bound polices -
  // and it had no bound at all, which made that bound decorative: a 6 window
  // all-lonely plan the first escape refused was handed back here unchanged.
  // It also asked the disc question, and the disc question is why a deck with
  // four anonymous services passed: the unraised plan cleared the disc line by
  // 1.8% while naming 2 of 6.
  const raisedNames = markableAt(raised);
  const plainNames = markableAt(plain);
  if (raisedNames > plainNames) {
    // Same order as the escape above: ease first, then judge what ships.
    const eased = coarsestNaming(raised, plain);
    if (worthTheSplit(eased, plain)) {
      return { ...eased, servedW: serve, chaseAffordable: true };
    }
  }
  // Only prefer the unraised plan when it is genuinely cheaper. A refusal that
  // costs the same number of slides buys nothing and loses the marks.
  // The refusal is about the RAISE, not about every hop on the deck.
  //
  // This branch used to report `servedW: Infinity`, and the audit's exemption
  // reads `authoredW >= servedTileW` - false for every finite width - so one
  // refusal switched the callout rule off for the whole deck, hops between
  // 160px tiles drawn at 1.81in included. The plain plan serves the median, so
  // that is what it reports, and the exemption covers what it is documented to
  // cover: endpoints below the tile the planner actually served.
  return plain.windows.length < raised.windows.length
    ? { ...plain, servedW: medianW, chaseAffordable }
    : { ...raised, servedW: medianW, chaseAffordable };
}

function planWindowsAtCeiling(
  bounds: Bounds,
  services: ExportBox[],
  frame: DiagramFrame,
  options: { mustTile?: boolean; markIn?: number; serveW?: number; waiveDensity?: boolean; namePerIn?: number } = {},
  raiseCeiling = true,
  serveW?: number,
): { windows: DiagramWindow[]; legible: boolean } {
  const contentW = Math.max(1, bounds.maxX - bounds.minX);
  const contentH = Math.max(1, bounds.maxY - bounds.minY);
  const whole = { windows: [] as DiagramWindow[], legible: true };
  if (services.length === 0 || frame.w <= 0 || frame.h <= 0) return whole;

  const shortest = Math.min(...services.map((box) => box.h).filter((h) => h > 0));
  if (!Number.isFinite(shortest) || shortest <= 0) return whole;

  // The strict minimum is the wrong statistic for a *target*. One sliver among
  // eighty-one ordinary tiles asks for a grid 3.75x finer than the rest of the
  // sheet needs; no grid within the slide budget delivers it, so `legible` came
  // back false at every stage and the drawing fell through to the one outcome
  // worse than either — a plotter page tiled into twenty plotter pages.
  //
  // What makes ignoring the outlier honest is that the renderer now floors a
  // window tile's type at `LEGIBLE_TILE_PT` whatever its height, so the sliver
  // reads either way. That splits one contract cleanly in two: the renderer
  // guarantees the floor, the planner optimises for it, and the planner
  // optimises for the tiles that stand to gain. When tiles are uniform, or when
  // the whole sheet is short, this is exactly the minimum and nothing moves.
  // The median, not the minimum and not a low percentile. The renderer floors a
  // window tile's type at the legibility limit however short the tile is, so
  // the planner's job is to pick the grid that serves the tiles that stand to
  // gain, and that is the typical tile. A low percentile only moved the cliff:
  // at 40 collapsed nodes in 400 the tenth percentile is still 75px and the
  // deck is ordinary, at 45 it is 12px and the deck collapsed to one string for
  // four hundred services. The median has no such neighbour — half the sheet
  // has to be collapsed before it moves, and when tiles are uniform it is the
  // minimum, so nothing moves on an ordinary drawing.
  const heights = services.map((box) => box.h).filter((h) => h > 0).sort((a, b) => a - b);
  const target = heights[Math.floor(heights.length * 0.5)] ?? shortest;

  // Inches-per-pixel worth splitting to reach for a representative tile. Both
  // ceilings and the reasoning behind them live on `legibleScaleFor`, which is
  // exported so the invariant can be asserted on the function rather than
  // inferred from the deck it produces. The narrowest tile goes in because the
  // renderer's ceiling depends on it, and a planner that stops splitting below
  // the ceiling the renderer will use leaves tiles under the markable bar with
  // slides already spent.
  //
  // I argued the other way once and the measurements refuted it. The claim was
  // that splitting cannot enlarge a tile past natural width, so the raised
  // ceiling is the renderer's business alone. Both halves are false on a
  // drawing of 14px tiles: the split windows draw at 1.034x natural, which is
  // the raised ceiling engaging and could only happen because of the split,
  // and the renderer rescues nothing on its own because frame-over-content
  // binds far below the ceiling - it would have allowed 1.371x and the planner
  // stopped asking at 1.034x. The deck that came out had 24 continuation
  // slides, 120 tiles and not one character of type on any of them, each slide
  // captioned as a readable part of a whole. The slides were spent either way;
  // they simply bought nothing.
  //
  // The raise is for a drawing that is SMALL, though, not for a drawing that
  // has a sliver in it. `probe-whitespace` is four 14px tiles beside two 160px
  // ones, and chasing the 14px one there drags the 160px ones to 2.3in each,
  // which puts one tile on a slide and spends eight of them to show a single
  // character per sliver - the same bad trade in the other direction.
  //
  // The width therefore takes the median, for the same reason the height
  // target twelve lines above it does, and the argument written there applies
  // here unchanged: an extremum has a neighbour, so one node decides the deck.
  // A max-over-min ratio is strictly worse than the low percentile that
  // paragraph rejects - measured, one node of sixty widening from 56px to 57px
  // moved the deck from sixty named services to one. Half the sheet has to be
  // small before the median moves, and on a uniform drawing the median IS the
  // minimum, so `probe-tiny-spread` is unaffected. Whether the raise was worth
  // its slides is then measured on the resulting plan, not predicted here.
  const widths = services.map((box) => box.w).filter((w) => w > 0).sort((a, b) => a - b);
  const typicalW = widths[Math.floor(widths.length * 0.5)] ?? Infinity;
  const servedW = serveW !== undefined && serveW > 0 ? Math.min(serveW, typicalW) : typicalW;
  const legibleScale = legibleScaleFor(
    target,
    frame,
    raiseCeiling ? servedW : Infinity,
    options.markIn ?? MARKABLE_TILE_W_IN,
  );
  if (Math.min(frame.w / contentW, frame.h / contentH) >= legibleScale) return whole;

  // Splitting on one axis only is why a tall drawing used to grow the page
  // without limit: a grouped architecture is large in both directions, so the
  // window has to tile in both. Each window also renders a bleed margin, which
  // comes out of its own budget so the tiles stay legible.
  const gridFor = (perIn: number): { cols: number; rows: number; slides: number } | null => {
    // Every window renders the same bleed on all four sides, including at the
    // drawing's outer edge, so that every part of the deck shares one scale.
    // The budget has to be charged the same way or the planned grid renders
    // smaller than planned and drops under the legibility floor.
    const usableW = frame.w / perIn - WINDOW_BLEED_PX * 2;
    const usableH = frame.h / perIn - WINDOW_BLEED_PX * 2;
    if (usableW <= 0 || usableH <= 0) return null;
    const cols = Math.max(1, Math.ceil(contentW / usableW));
    const rows = Math.max(1, Math.ceil(contentH / usableH));
    return { cols, rows, slides: cols * rows };
  };

  const tile = (cols: number, rows: number): DiagramWindow[] => {
    const stepX = contentW / cols;
    const stepY = contentH / rows;
    const cellX = (col: number) => ({
      minX: bounds.minX + stepX * col,
      maxX: col === cols - 1 ? bounds.maxX : bounds.minX + stepX * (col + 1),
    });
    const cellY = (row: number) => ({
      minY: bounds.minY + stepY * row,
      maxY: row === rows - 1 ? bounds.maxY : bounds.minY + stepY * (row + 1),
    });
    // `windowOwnsPoint` is a partition, so every service centre falls in
    // exactly one cell and that cell can be computed directly. Asking each of
    // the cells whether any service lands in it is the same answer for
    // `cols * rows * services` work, which is what made the search for a
    // readable grid unaffordable on the sparse drawings that need the finest
    // ones: `shrinkToFit` walks a few hundred grids, so a 150 x 150 search over
    // 140 services is hundreds of millions of point tests before a single slide
    // is written.
    const cellOf = (v: number, lo: number, step: number, n: number): number => (
      step > 0 ? Math.min(n - 1, Math.max(0, Math.floor((v - lo) / step))) : 0
    );
    const occupied = new Set<number>();
    for (const box of services) {
      const col = cellOf(box.x + box.w / 2, bounds.minX, stepX, cols);
      const row = cellOf(box.y + box.h / 2, bounds.minY, stepY, rows);
      occupied.add(row * cols + col);
    }
    const holds = (col: number, row: number): boolean => occupied.has(row * cols + col);

    // An architecture is not a filled rectangle, so a cell of the grid can own
    // no services at all — and one was being emitted as a numbered part
    // carrying two zone outlines and three arrows whose endpoints are both on
    // other slides.
    //
    // An empty cell is not simply deleted, though: ownership has to stay a
    // partition of the whole drawing, or a connector label anchored in the gap
    // an empty cell used to cover belongs to no part and is silently dropped
    // from the deck — arrow drawn, number missing, and the workflow list still
    // citing it. The surviving neighbours absorb the vacated bands instead,
    // keeping the *fitted* cell — and therefore the scale — identical on every
    // part.
    const keptRows: number[] = [];
    for (let row = 0; row < rows; row += 1) {
      if (Array.from({ length: cols }, (_, col) => col).some((col) => holds(col, row))) keptRows.push(row);
    }
    if (keptRows.length === 0) return [];

    const spans = (kept: number[], lo: (i: number) => number, hi: (i: number) => number, min: number, max: number) =>
      kept.map((index, i) => ({
        lo: i === 0 ? min : (hi(kept[i - 1]) + lo(index)) / 2,
        hi: i === kept.length - 1 ? max : (hi(index) + lo(kept[i + 1])) / 2,
      }));

    const rowOwn = spans(keptRows, (r) => cellY(r).minY, (r) => cellY(r).maxY, bounds.minY, bounds.maxY);
    const windows: DiagramWindow[] = [];
    keptRows.forEach((row, r) => {
      const keptCols: number[] = [];
      for (let col = 0; col < cols; col += 1) if (holds(col, row)) keptCols.push(col);
      const colOwn = spans(keptCols, (c) => cellX(c).minX, (c) => cellX(c).maxX, bounds.minX, bounds.maxX);
      keptCols.forEach((col, c) => {
        windows.push({
          fit: { ...cellX(col), ...cellY(row) },
          own: { minX: colOwn[c].lo, maxX: colOwn[c].hi, minY: rowOwn[r].lo, maxY: rowOwn[r].hi },
        });
      });
    });
    return windows;
  };

  // Take the coarsest grid that still reads well: aim for comfortable labels,
  // but never split so finely that a sheet stops carrying a meaningful piece
  // of the architecture. Eight sheets for a twelve-service diagram is not a
  // deck, it is a flip-book; that trade only pays on a genuinely large drawing.
  // "This frame cannot show the drawing readably at any grid" has two answers.
  // A deck that can grow its page takes that one, and every bail-out below
  // hands it the decision. A deck whose page is fixed has no such option: its
  // only alternatives are more slides or unreadable type, and more slides
  // always wins. `mustTile` callers therefore get the finest grid the ceiling
  // allows instead of an empty plan — reading `windows: []` as "it already
  // fits" is what put 4pt type on the shipping deck for every drawing sparse
  // enough to defeat the services-per-slide floors.
  const mustTile = options.mustTile === true;
  // Empty cells cost nothing. A reader counts slides, and a sparse drawing's
  // grid is mostly cells no service falls in — the diagonal cascade needs a
  // 10 x 13 grid to reach seven points and emits thirteen slides from it.
  // Capping the grid instead of the deck therefore refused a thirteen-slide
  // readable deck in favour of a twenty-four-cell one at four points.
  const slidesFor = (c: number, r: number): number => tile(c, r).length;
  // The scale a reader actually gets from a grid: each window covers
  // `content / n` of the drawing, so a finer grid is a bigger drawing, and the
  // binding axis is whichever fits worst.
  const scaleOf = (c: number, r: number): number => Math.min(
    frame.w / (contentW / c + WINDOW_BLEED_PX * 2),
    frame.h / (contentH / r + WINDOW_BLEED_PX * 2),
  );
  // Coarsening an axis costs scale along that axis alone, and the scale the
  // reader gets is the smaller of the two. Spending that cost on the axis that
  // already binds therefore shrinks the type for nothing, while the other axis
  // sits on slack it is not using — which is what stepping toward a square did
  // to every long drawing: a diagonal cascade lost the axis it was long in and
  // came out at 6.0pt on *fewer* slides than the same drawing one service
  // smaller, so adding a service made the deck shorter and less readable.
  const drop = (c: number, r: number): { c: number; r: number } => {
    if (c <= 1) return { c, r: Math.max(1, r - 1) };
    if (r <= 1) return { c: Math.max(1, c - 1), r };
    const scaleX = frame.w / (contentW / c + WINDOW_BLEED_PX * 2);
    const scaleY = frame.h / (contentH / r + WINDOW_BLEED_PX * 2);
    return scaleX > scaleY ? { c: c - 1, r } : { c, r: r - 1 };
  };
  const shrinkToFit = (cols: number, rows: number): { c: number; r: number } => {
    let c = Math.max(1, cols);
    let r = Math.max(1, rows);
    // A grid this fine is a plotter drawing however it is counted, and the
    // bound also keeps the search below from walking a pathological grid.
    while (c * r > MAX_TILED_CELLS) ({ c, r } = drop(c, r));
    // Coarsen toward the deck ceiling — but the ceiling is a preference and
    // legibility is not. Every drop buys a shorter deck with smaller type, so
    // a drawing sparse enough to need more than the ceiling's worth of windows
    // used to be coarsened right past the point where its labels stopped
    // reading: a fifty-two service cascade came out at 6.31pt and a ninety at
    // 4.00pt, on exactly forty-eight slides either way. Nobody wants ninety
    // slides of one diagram, but the alternative here is not fewer slides, it
    // is the same diagram nobody can read. Stop at the last grid that reads.
    while (c * r > 1 && slidesFor(c, r) > MAX_LEGIBLE_TILED_SLIDES) {
      const next = drop(c, r);
      if (scaleOf(c, r) >= legibleScale && scaleOf(next.c, next.r) < legibleScale) break;
      ({ c, r } = next);
    }
    // Nothing reads at any grid, or it reads and runs long. Either way the
    // deck still has to be a deck, so cap it — this is the point at which a
    // drawing is genuinely too large for a fixed page and the honest answer is
    // the most readable deck of a finite length.
    //
    // The cap is still a preference and legibility is still not. A cascade
    // needs one window per service to read, so no grid satisfies the cap at
    // all, and chasing it walked the grid down to nothing and fell through to a
    // single untiled slide: at 120 services the deck was 121 slides at 7.03pt
    // and at 121 it was one slide at 4pt. One service more turned a deck that
    // reads into a slide that does not.
    while (c * r > 1 && slidesFor(c, r) > MAX_FIXED_PAGE_SLIDES) {
      const next = drop(c, r);
      if (scaleOf(c, r) >= legibleScale && scaleOf(next.c, next.r) < legibleScale) break;
      ({ c, r } = next);
    }
    return { c, r };
  };
  const capped = (cols: number, rows: number): { windows: DiagramWindow[]; legible: boolean } | null => {
    if (!mustTile) return null;
    const { c, r } = shrinkToFit(cols, rows);
    if (c * r <= 1) return whole;
    // Report honestly whether the grid clears the legibility floor. A deck that
    // can grow its page only prefers these windows when they read; one that
    // cannot takes them either way, because its alternative is worse.
    const achieved = Math.min(
      frame.w / (contentW / c + WINDOW_BLEED_PX * 2),
      frame.h / (contentH / r + WINDOW_BLEED_PX * 2),
    );
    return { windows: tile(c, r), legible: achieved >= legibleScale };
  };

  const comfortable = gridFor(COMFORTABLE_TILE_PT / 12 / target);
  const floor = gridFor(legibleScale);
  // Comfort is a preference and legibility is not, so the comfortable grid may
  // only be preferred when it ALREADY reads.
  //
  // `comfortable` is derived from the median tile's height alone. When the
  // legibility floor demands a finer grid than comfort does - which is the
  // whole point of the markable raise, and the case every sliver drawing is in
  // - taking `comfortable` abandons `legibleScale` silently, and the function
  // then returns `{ windows: [], legible: true }`, "it fits whole", for a scale
  // it computed two lines earlier and found did not fit.
  //
  // Measured on one drawing at five, six and seven services in an identical
  // bounding box: the sixth service crossed `MIN_SERVICES_PER_SLIDE`, so the
  // short-circuit engaged, the sliver's tile went 0.2030in to 0.1906in, its
  // disc went 77% to 82% of the service it numbers, and `chaseAffordable` -
  // which the audit reads as "the deck measured the finer plan and could not
  // afford it" - flipped from true to false without any plan being measured.
  // Adding an ordinary service made the drawing worse and the gate quieter.
  const comfortableReads = !!comfortable
    && (!floor || scaleOf(comfortable.cols, comfortable.rows) >= legibleScale);
  const worthIt = comfortable
    && comfortableReads
    && comfortable.slides <= MAX_DIAGRAM_SLIDES
    && services.length / comfortable.slides >= MIN_SERVICES_PER_SLIDE;
  // A caller may ask for a SCALE, not only for a tile width to serve.
  //
  // `serveW` is clamped to `typicalW` one screen up, so it can only ever ask
  // for a finer plan than the median already gives; there was no way to say
  // "coarser than the finest, finer than the median". That is why the two
  // candidates the deck chose between measured 0.011314 and 0.020793 in/px on
  // a drawing that needed 0.011772: the 4% rung existed - `gridFor` returns it
  // - and nothing could name it. Bounded below by the grid legibility already
  // demands, so this can raise a deck's scale and never lower it.
  const naming = options.namePerIn !== undefined && options.namePerIn > 0
    ? gridFor(options.namePerIn)
    : null;
  const preferred = worthIt ? comfortable : floor;
  const grid = naming && (!preferred || naming.slides > preferred.slides) ? naming : preferred;
  if (!grid) {
    return capped(Math.ceil(Math.sqrt(MAX_TILED_CELLS)), Math.ceil(Math.sqrt(MAX_TILED_CELLS)))
      ?? { windows: [], legible: false };
  }
  const { cols, rows } = grid;
  if (cols * rows <= 1) return whole;
  // A fixed-page deck has no third option, so it never bails out here: both
  // the flip-book floor and the deck ceiling below exist to protect a deck
  // that could instead grow its page, and applying them to one that cannot is
  // what put four-point type on the shipping deck.
  if (mustTile) return capped(cols, rows) ?? { windows: [], legible: false };
  // Past the comfortable grid the choice is not "more slides or one nice page",
  // it is "more standard slides or a plotter sheet the whole deck inherits", so
  // the tiled deck is allowed to run well past the comfortable ceiling. It
  // still has to be a deck: sheets that carry barely a tile each are a
  // flip-book, and a drawing needing more than the hard ceiling really is a
  // plotter drawing.
  // Charge WINDOWS, not grid cells. `tile()` drops cells that hold nothing, so
  // a 4x3 grid over an L-shaped drawing is five slides and not twelve, and the
  // paragraph above says so in as many words: empty cells cost nothing, a
  // reader counts slides. Both bail-outs here charged cells anyway, while the
  // `mustTile` path charges `slidesFor` - so the density test on one drawing
  // divided by 12 and got 0.58 where the reader sees five windows and 1.4.
  // The cheaper the grid's shape, the likelier the deck was refused for it.
  const slides = slidesFor(cols, rows);
  if (slides > MAX_TILED_SLIDES) return { windows: [], legible: false };
  if (
    !options.waiveDensity
    && slides > MAX_DIAGRAM_SLIDES
    && services.length / slides < MIN_SERVICES_PER_TILED_SLIDE
  ) {
    return { windows: [], legible: false };
  }
  const windows = tile(cols, rows);
  // A single surviving cell means the drawing cannot be tiled at all. Saying it
  // is legible would pin an over-large drawing to one standard slide at a scale
  // that already failed the legibility test above; let the page grow instead.
  if (windows.length <= 1) return mustTile ? whole : { windows: [], legible: false };
  return { windows, legible: true };
}


/**
 * Does this window own the point, and therefore the shape centred on it?
 *
 * Windows meet exactly at a seam, so a point landing on one is inside both;
 * half-open ranges hand it to the later window and nothing is drawn twice.
 * Strays that the outlier trim pushed outside `bounds` sit in no window at all
 * under a plain range test and would vanish from the deck entirely, so the
 * outer windows claim everything beyond them.
 */
function windowOwnsPoint(window: Bounds, bounds: Bounds, x: number, y: number): boolean {
  const axis = (v: number, lo: number, hi: number, isFirst: boolean, isLast: boolean): boolean => {
    if (v < lo) return isFirst;
    if (v > hi) return isLast;
    return v < hi || isLast;
  };
  return axis(x, window.minX, window.maxX, window.minX <= bounds.minX + 0.5, window.maxX >= bounds.maxX - 0.5)
    && axis(y, window.minY, window.maxY, window.minY <= bounds.minY + 0.5, window.maxY >= bounds.maxY - 0.5);
}

/**
 * Which bounds should size the page *and* place every shape.
 *
 * Always the full bounds when they fit inside the 56" page — that keeps the
 * drawing 1 : 1 and nothing can fall off. Only when the content genuinely
 * exceeds the page does outlier trimming earn its keep: a single stray node at
 * (9000, 4000) would otherwise shrink every readable tile to a speck. In that
 * case the caller must clamp the strays back onto the page so they stay
 * visible instead of being silently drawn into the void.
 */
function chooseExportBounds(boxes: Iterable<ExportBox>): { bounds: Bounds; clamped: boolean } {
  const all = [...boxes];
  const full = computeBounds(all);
  const fullW = (full.maxX - full.minX) / PX_PER_IN;
  const fullH = (full.maxY - full.minY) / PX_PER_IN;
  if (fullW + IMAGE_X * 2 + 0.5 <= MAX_SLIDE_IN && fullH + IMAGE_Y + FOOTER_H + 0.78 <= MAX_SLIDE_IN) {
    return { bounds: full, clamped: false };
  }
  const trimmed = computeContentBounds(all);
  const trimmedW = (trimmed.maxX - trimmed.minX) / PX_PER_IN;
  const trimmedH = (trimmed.maxY - trimmed.minY) / PX_PER_IN;
  if (trimmedW < fullW * 0.8 || trimmedH < fullH * 0.8) {
    return { bounds: trimmed, clamped: true };
  }
  return { bounds: full, clamped: false };
}

/**
 * The parked layout every part of the pipeline must agree on: bounds that the
 * page is sized from, boxes the shapes and the routes are both planned from.
 *
 * Trimming far-placed nodes out of the fit is only half a decision — the strays
 * still have to be drawn somewhere, and unless the same answer reaches the page
 * sizer, the window planner, the renderer and the router, a stray tile, the
 * arrow aimed at it and the slide that claims it each pick a different one.
 */
function parkedLayout(nodes: Node[]): { boxes: Map<string, ExportBox>; bounds: Bounds; clamped: boolean } {
  // Empty space is closed before anything is measured, so the page sizer, the
  // window planner and the trim all see the drawing rather than the void
  // around it.
  const raw = compactEmptyGutters(collectExportBoxes(nodes));
  const { bounds: fitted, clamped } = chooseExportBounds(raw.values());
  const parked = clamped ? clampedBoxes(raw, fitted) : { boxes: raw, bounds: fitted };
  return { ...parked, clamped };
}

/**
 * Windows for a deck whose page size is fixed.
 *
 * The customer deck carries title, workflow, services, review and cost slides
 * all designed for a standard 16:9 page, and PowerPoint gives a deck exactly
 * one page size, so a large architecture cannot buy legibility by growing the
 * sheet the way the diagram-only deck does. It tiles instead, and it tiles even
 * when the planner reports the drawing cannot be shown legibly at any grid:
 * that verdict exists so a deck that *can* grow its page does, and reading it
 * as "no windows needed" put 4pt type on every drawing sparse enough to defeat
 * the services-per-slide floors — a hub with four spokes at 1400px, the most
 * ordinary shape in the Architecture Center, among them. Returns an empty list
 * only when the drawing already fits, which keeps the common path unchanged.
 */
function planFixedPageWindows(diagram: DiagramShapeSource, frame: DiagramFrame): DiagramWindow[] {
  const nodes = diagram.nodes ?? [];
  if (nodes.length === 0) return [];
  const { boxes, bounds } = parkedLayout(nodes);
  if (boxes.size === 0) return [];
  const { services } = partitionBoxes(boxes);
  if (services.length === 0) return [];
  return planDiagramWindows(bounds, services, frame, {
    mustTile: true,
    markIn: markableTileWIn(stepNumbersOf(diagram.edges ?? [])),
  }).windows;
}

/**
 * Pick the slide size. Grows the page (never the shrink factor) so a wide
 * architecture keeps 1 : 1 geometry; only diagrams larger than the 56" page
 * limit are scaled down, and then every dimension scales together.
 */
function planSlideGeometry(diagram?: DiagramShapeSource | null): SlideGeometry {
  const chrome = { top: IMAGE_Y, bottom: FOOTER_H + 0.18 + 0.1 };
  // The colour key is drawn last, over everything, and is 92% opaque, so
  // whatever is under it is missing from the finished deck — on a full grid
  // that was 92% of a service tile, icon and name and all, and no corner of a
  // full grid is free to move it to. Reserve its strip here rather than in the
  // renderer, so the tiler plans legibility against the height the drawing will
  // actually get: reserved later, the two disagreed and tiles slid under 7pt.
  const legendH = usedConnectionLegend(diagram?.edges ?? []).length > 0 ? 0.24 + 0.03 : 0;
  // A numbered drawing has to reach a tile that can carry its callout, not
  // merely one that can carry a mark. See `markableTileWIn`.
  const markIn = markableTileWIn(stepNumbersOf(diagram?.edges ?? []));
  const frameFor = (pageW: number, pageH: number): DiagramFrame => {
    const footer = pageH - FOOTER_H - 0.08;
    return { x: IMAGE_X, y: IMAGE_Y, w: pageW - IMAGE_X * 2, h: footer - IMAGE_Y - 0.1 - legendH };
  };
  let w = BASE_W;
  let h = BASE_H;
  let overflow = false;
  let outliersClamped = false;
  let usedFrame: DiagramFrame | null = null;
  let medians: { w: number; h: number } | null = null;
  let servedW = Infinity;
  let chaseAffordable = false;

  const nodes = diagram?.nodes ?? [];
  let windows: DiagramWindow[] = [];
  if (nodes.length > 0) {
    const parked = parkedLayout(nodes);
    if (parked.boxes.size > 0) {
      outliersClamped = parked.clamped;
      // Plan the windows against the drawing the slides will actually carry.
      // Parking a stray widens the drawing by the strip it sits in, and a
      // window plan made from the pre-parking bounds leaves that strip
      // belonging to no window at all — the stray, and every hop touching it,
      // is then drawn on no slide.
      const boxes = parked.boxes;
      const bounds = parked.bounds;
      const { services } = partitionBoxes(boxes);
      // The narrowest tile a step callout actually lands on.
      //
      // Badged, not narrowest outright, for the reason the Visio magnifier
      // takes the same filter: a sliver with no numbered arrow touching it has
      // no disc to be dwarfed by, so chasing it buys nothing and costs the
      // slides `probe-whitespace` measures.
      const badged = new Set<string>();
      for (const edge of diagram?.edges ?? []) {
        const step = Number((edge as unknown as { data?: { stepNumber?: unknown } }).data?.stepNumber);
        if (!Number.isFinite(step) || step <= 0) continue;
        badged.add(String(edge.source));
        badged.add(String(edge.target));
      }
      let narrowestBadgedW = Infinity;
      for (const [id, box] of boxes) {
        if (box.kind !== 'service' || !(box.w > 0) || !badged.has(id)) continue;
        if (box.w < narrowestBadgedW) narrowestBadgedW = box.w;
      }
      const serveW = Number.isFinite(narrowestBadgedW) ? narrowestBadgedW : undefined;
      const contentW = Math.max(1, bounds.maxX - bounds.minX) / PX_PER_IN;
      const contentH = Math.max(1, bounds.maxY - bounds.minY) / PX_PER_IN;
      // Room for the connection legend plus breathing space around the drawing.
      const wantW = contentW + IMAGE_X * 2 + 0.5;
      const wantH = contentH + chrome.top + chrome.bottom + 0.5;

      // A reader can present a deck of ordinary slides; nobody can present a
      // 28in x 16in plotter sheet, and PowerPoint gives a deck exactly one page
      // size, so the title and workflow slides inherit whatever the drawing
      // demands. Prefer standard slides: shrink onto one while the labels stay
      // above the legibility floor, tile across several when they would not,
      // and only grow the page when even the slide budget is exceeded.
      const standard = planDiagramWindows(bounds, services, frameFor(BASE_W, BASE_H), { markIn, serveW });
      usedFrame = frameFor(BASE_W, BASE_H);
      medians = medianExtent(services);
      servedW = standard.servedW;
      chaseAffordable = standard.chaseAffordable;
      // `legible: false` is a request to grow the page, not a verdict that the
      // drawing cannot be tiled. A sparse architecture — the hub-and-spoke
      // every Architecture Center reference draws — defeats the
      // services-per-slide floors purely by having whitespace between its
      // parts, and was handed a 31x32in plotter page nobody can open. Ask the
      // planner for the finest grid the slide budget allows before giving up on
      // ordinary slides; the 7pt floor still governs whether the result is
      // readable, and every part still shares one scale.
      const forced = standard.legible ? null : planDiagramWindows(bounds, services, frameFor(BASE_W, BASE_H), { mustTile: true, markIn, serveW });
      if (standard.legible) {
        windows = standard.windows;
      } else if (forced && forced.windows.length > 1) {
        // Not `forced.legible`. That test asked the planner whether the tiles
        // land at their natural size, and used the answer to decide something
        // else entirely: standard slides versus a page nobody can open. Now
        // that the renderer floors a window tile's type at the legibility limit
        // whatever the grid, the honest comparison is between a deck of
        // ordinary slides and a 56in plotter sheet — and the sheet loses every
        // time, because a reader can at least present the deck.
        windows = forced.windows;
        servedW = forced.servedW;
        chaseAffordable = forced.chaseAffordable;
      } else {
        // Only a genuinely enormous drawing gets here — one that cannot be read
        // on nine standard slides. Grow the page for it, then tile that page
        // too: a 56in sheet split into three readable parts still beats one
        // 56in sheet at 4.9pt, which is not a diagram, it is a smudge.
        overflow = wantW > MAX_SLIDE_IN || wantH > MAX_SLIDE_IN;
        w = clamp(wantW, BASE_W, MAX_SLIDE_IN);
        h = clamp(wantH, BASE_H, MAX_SLIDE_IN);
        // `mustTile` is what makes the sentence above true. Without it the
        // planner still weighs these windows against the option of growing the
        // page — and this page has already grown to the maximum, so there is no
        // such option left; the bail-outs it takes on that assumption returned
        // no windows at all and left a 200-service estate as a single 56x39.87in
        // sheet, which is the one outcome this branch exists to avoid.
        const grown = planDiagramWindows(bounds, services, frameFor(w, h), { mustTile: true, markIn, serveW });
        usedFrame = frameFor(w, h);
        windows = grown.windows;
        servedW = grown.servedW;
        chaseAffordable = grown.chaseAffordable;
        // Splitting restores legibility, so the "scaled down to fit" warning no
        // longer applies — the drawing is now at its readable size.
        if (windows.length > 1) overflow = false;
      }
    }
  }

  const footerY = h - FOOTER_H - 0.08;
  return {
    w,
    h,
    footerY,
    overflow,
    outliersClamped,
    windows,
    frame: frameFor(w, h),
    // Recorded rather than recomputed by whoever wants to know. The gate used
    // to mirror this from the scenario and got both statistics wrong: it read
    // a frame with no connection legend in it, 6.04in against the 5.77in a
    // numbered drawing actually gets, and it took the NARROWEST node where the
    // planner takes the median. The first made it fail a correctly clamped
    // 21 slide deck 98 times; the second made it accuse a correct 53 slide
    // plan of chasing a bar it had already reached, while suppressing the four
    // real conflicts on that same deck. There is one copy now, and it is the
    // copy that ran.
    calloutBarClamped: !!usedFrame && !!medians && markIn > MARKABLE_TILE_W_IN
      && !calloutBarReachable(medians.h, usedFrame, medians.w, markIn),
    medianServiceW: medians && Number.isFinite(medians.w) ? medians.w : 0,
    servedTileW: servedW,
    // The narrowest authored tile this FRAME can ever carry a proportionate
    // callout on, in pixels, whatever the plan.
    //
    // `finestPerIn` is what the finest grid returns before the bleed alone
    // fills the window, so `markIn / finestPerIn` is the authored width below
    // which no split reaches the bar. Below it the deck has no move and the
    // gate must not demand one; above it a miss is a real miss.
    //
    // Per hop and not per deck, deliberately. `calloutBarClamped` answers the
    // same question about the MEDIAN, and reading a deck-wide answer here is
    // the mistake `servedW: Infinity` made one round ago: one 14px sliver
    // would switch the rule off for hops between 160px tiles drawn at 1.81in.
    // The gate compares each hop's own endpoints against this.
    //
    // It is a bound on the frame and not a fact about the drawing, so it moves
    // only when the frame or the bleed does - which is exactly what makes it
    // the number to watch when `WINDOW_BLEED_PX` becomes proportional: at a
    // flat 100px the bar is out of reach below 14.5 authored px, and a
    // proportional bleed retires most of this exemption by construction.
    reachableTileW: usedFrame && medians && markIn > 0
      ? markIn / (Math.min(usedFrame.w, usedFrame.h)
        / (WINDOW_BLEED_PX * 2 + Math.max(1, medians.h)))
      : 0,
    chaseAffordable,
  };
}

/**
 * What the window planner decided about this drawing's callouts.
 *
 * Exported for the export audit, which has to know whether a callout that is
 * disproportionate to its tile is a defect or the documented consequence of a
 * bar no grid in the frame can reach - and, since the planner deliberately
 * serves the MEDIAN service rather than the narrowest, which tiles it declined
 * to serve at all. Both are asked for rather than replicated: the gate's own
 * copy of each disagreed with this one, in opposite directions, and shipped.
 */
export function calloutPlanFor(diagram?: DiagramShapeSource | null): {
  clamped: boolean;
  medianServiceW: number;
  /**
   * The narrowest authored width the winning plan actually served, in pixels.
   *
   * Not the median, and not the narrowest badged tile either: it is whichever
   * of the two the density floor let the deck pay for. A tile narrower than
   * this one is a tile the planner MEASURED the cost of serving and declined,
   * so a callout that is disproportionate on it has no move behind it. A tile
   * at or above it was served, so a disproportionate callout on it is a defect.
   */
  servedTileW: number;
  /**
   * The narrowest authored tile this FRAME can carry a proportionate callout
   * on at any grid, in pixels.
   *
   * A frame bound and not a plan outcome: below it no split reaches the bar,
   * so the deck has no move and a disproportionate disc is the page's fault
   * and not the planner's. Compared against each hop's OWN endpoints, so one
   * sliver cannot exempt a deck.
   */
  reachableTileW: number;
  /**
   * Whether a plan serving the narrowest badged tile was BOTH finer than the
   * chosen one and inside the density floor.
   *
   * The bound behind servedTileW, kept separately because an exemption read
   * off the plan that was chosen goes blind on any mutation that changes the
   * choice. Where this is true and a callout is still disproportionate, the
   * deck had a move and did not take it.
   */
  chaseAffordable: boolean;
} {
  const geometry = planSlideGeometry(diagram);
  return {
    clamped: geometry.calloutBarClamped,
    medianServiceW: geometry.medianServiceW,
    servedTileW: geometry.servedTileW,
    reachableTileW: geometry.reachableTileW,
    chaseAffordable: geometry.chaseAffordable,
  };
}

/**
 * Advance of the WIDEST character in `text`, in inches.
 *
 * `estimateTextWidthIn` gives every non-CJK character 0.54 em, which is Segoe
 * UI's *average lowercase* advance. An average is the right model for the width
 * of a run and completely the wrong one for `max over characters`: measured
 * against the installed Yu Gothic UI with GDI+, `@` is 0.955 em, `W` 0.934 and
 * `M` 0.898 — 77% wider than the estimate. Asking "does one letter fit?" with
 * the average answered yes for boxes that hold no capital at all, and drew a
 * 31-character word one letter per line down a 2.55in ribbon.
 *
 * The buckets are the measured maxima of each class, so this never under-states
 * a glyph. It over-states a narrow one, which is the safe direction: the cost
 * is dropping wording that would just have fitted, and the words are carried
 * elsewhere.
 *
 * Re-exported from the shared geometry module, where it now lives so the Visio
 * exporter can decide "is this still a name?" with the same rule the deck uses.
 */
export { widestGlyphIn };

/**
 * Approximate rendered width of a string in inches, from the measured Yu
 * Gothic UI advances in `diagramExportGeometry`.
 *
 * This used to charge a flat 0.54 em for everything non-CJK - the average
 * LOWERCASE advance - so a name written in title case measured about 8% narrow.
 * That is the width `wrappedLineCount` and `fitLabelToLines` divide a column
 * by, so the error surfaced as a LINE: a name cut to the five lines its tile
 * had room for really wrapped to six, and the sixth was painted 0.079in below
 * the box it was measured in.
 *
 * A class-wise bucket was tried first and rejected. A bucket is only an upper
 * bound if it carries its class maximum, and charging every lowercase letter
 * the width of `m` inflates ordinary prose by about 60%, which shrinks type and
 * cuts names that would have fitted. Measured advances are neither short nor
 * fat.
 */
export function estimateTextWidthIn(text: string, fontSizePt: number): number {
  return advanceWidthIn(text, fontSizePt);
}

/**
 * How many lines `text` takes in a box `widthIn` wide, wrapping the way
 * PowerPoint wraps.
 *
 * `ceil(width / widthIn)` is exact only for text that can break anywhere — CJK,
 * or a single token long enough that PowerPoint breaks it mid-run. Real prose
 * breaks between words and throws away the remainder of a line whenever the
 * next word will not fit, so three tokens each wider than half the box take
 * three lines where the ratio predicts two. A resource name that spells out its
 * environment and region is exactly that shape, and under-counting it is how a
 * table budgeted for the page ends up printing below it.
 */
export function wrappedLineCount(text: string, widthIn: number, fontSizePt: number): number {
  if (!text) return 1;
  // Floored at one hair, not at 0.1in. A fixed floor on a column that shrinks
  // with the drawing is the same lie the audit's own floors told and the same
  // one `chipColumn` was rewritten to stop telling: an overview chip with a
  // 0.075in column was measured against 0.1in, counted 3 lines where the text
  // takes 4, and was emitted a line short — 78 chips across 3 decks, each
  // painting its last line out through the bottom of the lozenge. Callers that
  // want a floor must apply their own; this one reports what the column holds.
  const box = Math.max(0.001, widthIn);
  // A hard line break is a line, and it is the one break every counter here
  // used to miss. `\n` survives the sanitiser — which scrubs U+000B but not
  // U+000A — and pptxgenjs turns each one into a real `<a:p>`, so the renderer
  // starts a new line where the measurement carried straight on. Splitting on
  // whitespace only *ends a run* at a newline; it never starts a line. A model
  // asked for numbered remediation steps writes them one per line, which makes
  // this the normal case rather than an exotic one, and a sixteen-row table of
  // four-line names measured 5.83in and drew 10.33in.
  return text.split(/\r\n|\r|\n/)
    .reduce((total, paragraph) => total + wrapOneLine(paragraph, box, fontSizePt), 0);
}

/** One paragraph's worth of wrapping, with no hard breaks left in it. */
function wrapOneLine(text: string, box: number, fontSizePt: number): number {
  if (!text) return 1;
  // Breaks are between words, and additionally after any full-width character,
  // which is where CJK is allowed to break.
  const runs = text.split(/(?<=[\s\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6])/);
  let lines = 1;
  let used = 0;
  for (const run of runs) {
    const w = estimateTextWidthIn(run, fontSizePt);
    // A renderer decides whether a line fits on its visible ink and lets the
    // run-final spaces hang past the column, so the fit test discounts them
    // and the advance does not. Charging a space nothing everywhere was the
    // shortcut that made this look unnecessary, and it made every interior
    // space free: 1699 emitted boxes in the corpus were a quarter em per gap
    // short of the wrap they really take.
    const visible = w - trailingWhitespaceIn(run, fontSizePt);
    if (used > 0 && used + visible > box) {
      lines += 1;
      used = 0;
    }
    // A single run wider than the whole box breaks inside itself, one
    // CHARACTER at a time.
    //
    // `ceil(w / box)` assumes the word packs exactly a boxful per line, which
    // is only true if a break may fall part-way through a glyph. Breaks fall
    // between glyphs, so every line but the last ends short of the column and
    // the ratio is a lower bound, never the count. On a 0.204in column that is
    // one whole line: a name cut to the five lines its tile has room for
    // really wrapped to six, and the sixth was painted 0.079in below the box.
    if (w > box) {
      let lineUsed = used;
      for (const glyph of run) {
        const gw = estimateTextWidthIn(glyph, fontSizePt);
        if (lineUsed > 0 && lineUsed + gw > box) {
          lines += 1;
          lineUsed = 0;
        }
        lineUsed += gw;
      }
      used = lineUsed;
      continue;
    }
    used += w;
  }
  return Math.max(1, lines);
}

/**
 * As much of `text` as will fit in `widthIn` at `fontSizePt`, with an ellipsis
 * for the rest. Used where the tile is too small for the whole name and the
 * only alternatives are unreadable type or an empty box.
 *
 * The policy itself lives in `diagramExportGeometry`, so the sheet cuts a name
 * exactly where the deck cuts it.
 */
function fitLabelToBox(rawText: string, widthIn: number, fontSizePt: number): string {
  return fitLabelToWidth(rawText, widthIn, fontSizePt / 72);
}


// ─── Public export function ───────────────────────────────────────────────────

export interface PptxExportOptions {
  diagramName: string;
  author: string;
  date: string;
  isDarkMode: boolean;
  /**
   * Canvas contents. When provided, the diagram is rendered with native
   * PowerPoint shapes (rounded rectangles, embedded icons, arrow connectors)
   * so every element stays selectable and editable inside PowerPoint. The
   * captured PNG is only used as a fallback when no shapes can be produced.
   */
  diagram?: DiagramShapeSource | null;
  /**
   * Pre-rasterised icons, bypassing the DOM-only rasteriser. `canRasterize()`
   * is false under Node, so an offline harness draws every tile with no icon
   * at all — a different tile interior, with the caption band 2.1x too tall
   * and a third of an inch out of position, which is what the chip walk and
   * every contrast composite are then tuned against. Visio already had this
   * escape hatch; PowerPoint did not.
   */
  presetIcons?: Map<string, RasterizedIcon>;
}

export interface DiagramShapeSource {
  nodes: Node[];
  edges: Edge[];
}

// ─── Native (editable) diagram rendering ─────────────────────────────────────

/** Shared category tint (bare hex for pptxgenjs) — one source of truth. */
function styleForBox(box: ExportBox): { bg: string; border: string; text: string } {
  const style = categoryStyle(box.category);
  return { bg: stripHash(style.bg), border: stripHash(style.border), text: stripHash(style.text) };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

interface DiagramFrame { x: number; y: number; w: number; h: number }

/**
 * PptxGenJS ships `custGeom` at runtime (freeform path shapes) but omits it
 * from the published `SHAPE_NAME` union, so the name is cast once here.
 */
const CUSTOM_GEOMETRY = 'custGeom' as unknown as Parameters<Slide['addShape']>[0];

function toInches(point: Point, transform: FitTransform): Point {
  return {
    x: point.x * transform.scale + transform.offsetX,
    y: point.y * transform.scale + transform.offsetY,
  };
}

/** PowerPoint dash preset for each canonical connection type. */
function pptxDashType(route: ExportRoute): 'solid' | 'dash' | 'sysDot' | 'dashDot' {
  if (!route.dashed) return 'solid';
  switch (route.connectionType) {
    case 'optional':
    case 'security':
      return 'sysDot';
    case 'telemetry':
      return 'dashDot';
    default:
      return 'dash';
  }
}

/**
 * The part of a polyline that is actually on this window's paper.
 *
 * Off-window points used to be CLAMPED onto the frame border, which turns a leg
 * running somewhere else on the drawing into a line along the edge of the
 * slide. It is not where the arrow goes, and — because every clamped route is
 * flattened onto the same border — two of them end up drawn exactly on top of
 * each other: on a 72-service estate two wrap-around hops shared 10.39in of a
 * 10.51in line.
 *
 * Clipping instead of clamping draws the hop where it really is and simply
 * stops it at the paper's edge, which is what a reader expects at a seam. A
 * route that lies wholly inside the frame comes back untouched, so the
 * overwhelming majority of arrows are byte-identical to before.
 *
 * Discarding the hop instead was tried and is wrong: nothing guarantees another
 * window draws it. A single-window clamped deck, or a hop spanning two windows
 * and contained by neither, loses the arrow from the deck ENTIRELY while its
 * chip and numbered callout stay behind pointing at blank paper.
 */
function clipToFrame(
  points: readonly { x: number; y: number }[],
  frame: DiagramFrame,
): { x: number; y: number }[] {
  const x0 = frame.x;
  const x1 = frame.x + frame.w;
  const y0 = frame.y;
  const y1 = frame.y + frame.h;
  const eps = 1e-6;
  const inside = (p: { x: number; y: number }): boolean => p.x >= x0 - eps && p.x <= x1 + eps
    && p.y >= y0 - eps && p.y <= y1 + eps;
  if (points.every(inside)) return [...points];

  // Liang-Barsky, run per segment so an axis-aligned leg keeps exact endpoints.
  const clipSegment = (
    a: { x: number; y: number },
    b: { x: number; y: number },
  ): [{ x: number; y: number }, { x: number; y: number }] | null => {
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    let t0 = 0;
    let t1 = 1;
    const edges: Array<[number, number]> = [[-dx, a.x - x0], [dx, x1 - a.x], [-dy, a.y - y0], [dy, y1 - a.y]];
    for (const [p, q] of edges) {
      if (Math.abs(p) < eps) {
        if (q < -eps) return null;
      } else {
        const r = q / p;
        if (p < 0) {
          if (r > t1) return null;
          if (r > t0) t0 = r;
        } else {
          if (r < t0) return null;
          if (r < t1) t1 = r;
        }
      }
    }
    const at = (t: number): { x: number; y: number } => (t <= 0 ? a : t >= 1 ? b : { x: a.x + dx * t, y: a.y + dy * t });
    return [at(t0), at(t1)];
  };

  const runs: Array<{ x: number; y: number }[]> = [];
  let current: { x: number; y: number }[] = [];
  const near = (a: { x: number; y: number }, b: { x: number; y: number }): boolean => Math.abs(a.x - b.x) < 1e-4 && Math.abs(a.y - b.y) < 1e-4;
  for (let i = 1; i < points.length; i += 1) {
    const piece = clipSegment(points[i - 1], points[i]);
    if (!piece) {
      if (current.length >= 2) runs.push(current);
      current = [];
      continue;
    }
    const [from, to] = piece;
    if (current.length === 0) current = [from, to];
    else if (near(current[current.length - 1], from)) current.push(to);
    else {
      if (current.length >= 2) runs.push(current);
      current = [from, to];
    }
  }
  if (current.length >= 2) runs.push(current);
  if (runs.length === 0) return [];

  // The longest surviving run: one arrow per hop per slide, and the shape that
  // carries the arrowhead should be the piece the reader actually follows.
  const length = (run: { x: number; y: number }[]): number => run.reduce(
    (sum, point, i) => (i === 0 ? 0 : sum + Math.hypot(point.x - run[i - 1].x, point.y - run[i - 1].y)),
    0,
  );
  return runs.reduce((best, run) => (length(run) > length(best) ? run : best), runs[0]);
}

function addConnector(
  pptx: PptxGenJS,
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  clampTo?: DiagramFrame,
): void {
  const raw = route.points.map((point) => toInches(point, transform));
  const points = clampTo ? clipToFrame(raw, clampTo) : raw;
  if (points.length < 2) return;

  const lineProps = {
    color: stripHash(route.color),
    width: route.connectionType === 'optional' ? 1 : 1.25,
    dashType: pptxDashType(route),
    endArrowType: 'triangle' as const,
    ...(route.bidirectional ? { beginArrowType: 'triangle' as const } : {}),
    transparency: route.opacity < 1 ? Math.round((1 - route.opacity) * 100) : undefined,
  };

  if (points.length === 2) {
    // Straight run — a preset line keeps zero-height/zero-width geometry valid
    // and stays a first-class editable connector inside PowerPoint.
    const [from, to] = points;
    slide.addShape(pptx.ShapeType.line, {
      x: Math.min(from.x, to.x),
      y: Math.min(from.y, to.y),
      w: Math.abs(to.x - from.x),
      h: Math.abs(to.y - from.y),
      flipH: to.x < from.x,
      flipV: to.y < from.y,
      line: lineProps,
      objectName: `connector-${route.id}`,
    });
    return;
  }

  const xs = points.map((point) => point.x);
  const ys = points.map((point) => point.y);
  const minX = Math.min(...xs);
  const minY = Math.min(...ys);
  const w = Math.max(...xs) - minX;
  const h = Math.max(...ys) - minY;
  slide.addShape(CUSTOM_GEOMETRY, {
    x: minX,
    y: minY,
    w: Math.max(w, 0.01),
    h: Math.max(h, 0.01),
    points: points.map((point, index) => ({
      x: point.x - minX,
      y: point.y - minY,
      ...(index === 0 ? { moveTo: true } : {}),
    })),
    fill: { type: 'none' },
    line: lineProps,
    objectName: `connector-${route.id}`,
  });
}

/**
 * Geometry of a connector's label chip, or null when the edge has no label.
 *
 * Shared with the step badge so the badge can be placed from the chip's real
 * height. Deriving it twice used to leave the badge sitting on top of tall,
 * wrapped CJK labels.
 */
function connectorLabelBox(
  route: ExportRoute,
  transform: FitTransform,
  requestedFontSize: number,
  px: number,
  clampTo?: DiagramFrame,
  obstacles: readonly Obstacle[] = [],
  bundle?: {
    count: number; rung: number; x: number; y: number;
    shift?: number; across?: number; font?: number; perCol?: number;
    /** Rung index per route id, ordered the way the arrows themselves fan. */
    rank?: ReadonlyMap<string, number>;
    /** Length of the longest hop in the bundle, so the ladder stays over it. */
    span?: number;
    /** Width the ladder has been told to fit into, when it has no clear slot. */
    maxWidth?: number;
    /** The fan has nowhere clean to stand: number the arrows, drop the chips. */
    badgesOnly?: boolean;
  },
  /** Width to wrap into, when the natural chip has nowhere clear to stand. */
  squeezeTo?: number,
  /**
   * Distance from a point to the nearest arrow that is NOT this chip's own hop.
   * Ladders price this themselves; a lone chip had no such test at all, so its
   * walk could settle it beside a stranger's arrow and every rule was happy.
   */
  foreignGap?: (x: number, y: number) => number,
): { x: number; y: number; w: number; h: number; text: string; badge: { x: number; y: number; d: number } | null; block: Obstacle; alongX: boolean; fontSize: number; stuck: number } | null {
  if (!route.label) return null;
  const anchor = toInches(route.labelAnchor, transform);
  // A fan with nowhere clean to stand carries its wording on the workflow slide
  // instead and leaves numbered callouts on the arrows. That is what the
  // Architecture Center does with a bundle of parallel flows, and it beats
  // seven chips laid across the services they run between.
  const wanted = bundle?.badgesOnly ? '' : truncateLabel(route.label, 42);

  // Size the chip from the text it actually carries, capped so it can never
  // dwarf the service tiles it sits between (a 150 px tile is 1.56" at 1 : 1).
  // The gap between the two tiles is a *preference*, not a hard cap: squeezing
  // a long label into a 190px hop turned the chip into a 0.34" ribbon several
  // inches tall, and the parallel-edge stagger — which steps by the chip's own
  // height — then flung the second and third ordinals off the top and bottom
  // of the slide, taking their step numbers with them. A chip that is a little
  // wider than the gap is fine; the obstacle walk below moves it clear.
  const first = toInches(route.points[0] ?? route.labelAnchor, transform);
  const last = toInches(route.points[route.points.length - 1] ?? route.labelAnchor, transform);
  const span = Math.max(Math.abs(last.x - first.x), Math.abs(last.y - first.y));
  const gap = span > 0 ? span - 0.08 : 1.5 * px;
  // The ladder of parallel chips runs across the edge, so its room is the
  // frame dimension perpendicular to the arrow.
  const alongX = Math.abs(last.x - first.x) >= Math.abs(last.y - first.y);
  // A bundle of parallel edges is stacked into a ladder, so every chip in it
  // gets a share of the corridor rather than its natural size. Wide-and-short
  // is what fits a ladder, so the width cap is relaxed for bundles.
  const siblings = bundle?.count ?? 1;
  // A fan deeper than the slide is tall cannot be laid out at the ordinary
  // label size; the caller works out how far the whole bundle has to shrink so
  // that every rung still fits, which keeps the text intact.
  const fontSize = bundle?.font ?? requestedFontSize;
  // A chip narrower than one letter of its own type is not a small chip, it is
  // a broken one: PowerPoint clips nothing, so it stacks the word one glyph per
  // line and paints the letters out through both sides of the lozenge. On a
  // heavily scaled overview the box shrinks with the drawing but the font is
  // held at OVERVIEW_LEGIBLE_PT, so past a certain scale the trade the sizer
  // thinks it is making is not on offer — 479 chips at 0.009in wide drew a
  // 0.6in smear of type across the middle of the first slide.
  //
  // Where there is no room for the letters, drop the wording rather than
  // scribble it. The overview is an index, not a reading surface, and the same
  // words are on the window slide that follows by exactly the route a tile name
  // too small to draw already takes.
  const widestGlyph = widestGlyphUpperIn(wanted, fontSize);
  // TWO of them, not one. "Does a single letter fit?" is a test about the
  // wrong thing: a column that holds exactly one glyph produces a chip that
  // spells its sentence vertically, one character per line, and that is the
  // artefact the guard exists to prevent — not a narrower version of it. It
  // let 97 chips through at 1.8 characters per line, the worst of them 29
  // copies of a 0.200 x 1.252in ribbon down the middle of the first slide,
  // passing by 0.0022in. Room for two of the widest glyph is the least that
  // can be called a line of text.
  const minChipW = widestGlyph * 2 + CHIP_INSET_IN * 2;
  const text = wanted;
  // Wide-and-short is the shape that fits a ladder: a fan of six chips each
  // wrapped onto four lines is taller than the slide, and clamping them into
  // the frame one by one just restacks them on the page edge. Letting a chip in
  // a fan run past the hop it labels halves its height instead.
  const prefer = siblings > 2 ? 1.4 : 0.9;
  // A ladder that has nowhere clear to stand is told to narrow itself to the
  // lane between the two services instead. Wrapping the text onto more lines is
  // a far better trade than parking seven chips on top of the tiles either
  // side, which is what a wide fan does in a dense grid.
  const maxW = bundle?.maxWidth !== undefined
    ? clamp(bundle.maxWidth, 0.34 * px, 1.5 * px)
    : squeezeTo !== undefined
      ? clamp(squeezeTo, 0.34 * px, 1.5 * px)
      : clamp(Math.max(gap, prefer * px), 0.34 * px, 1.5 * px);
  const naturalW = estimateTextWidthIn(text, fontSize) + 0.14;
  const badgeD = route.stepNumber === undefined ? 0 : stepBadgeDiameterIn(route, transform, px);  // A muted rung carries no wording, so it is exactly its callout. Reserving an
  // empty text box above the number anyway made every rung ~40% taller and a
  // third wider than the thing it draws, which is what pushed a deep fan's end
  // callouts onto the tiles it runs between.
  const bare = text === '' && badgeD > 0;
  // A hop that runs down the page has to fit its chip into the band between two
  // rows. `gap` is that band's height, but it is being used as a WIDTH cap, so
  // a long sentence is wrapped onto more lines until the chip is taller than
  // the corridor it has to sit in — and then no position clears the service it
  // labels. The band is short but wide, so widen the chip, by exactly enough to
  // shed the lines that do not fit and no more: widening further only pushes it
  // into the columns either side.
  let roomW = maxW;
  if (!alongX && !bare && bundle === undefined && squeezeTo === undefined && gap > 0) {
    const lineH0 = (fontSize * 1.3) / 72;
    const asIs = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / chipColumn(maxW)));
    // Only when the chip does not already fit the band. Widening one that does
    // buys nothing and costs a lean on the columns either side.
    if (asIs * lineH0 + 0.06 > gap) {
      const fits = Math.max(1, Math.floor((gap - 0.06) / lineH0));
      const needed = estimateTextWidthIn(text, fontSize) / fits + CHIP_INSET_IN * 2;
      roomW = clamp(Math.max(maxW, needed), 0.34 * px, 1.5 * px);
    }
  } else if (alongX && !bare && bundle === undefined && squeezeTo === undefined && gap > 0) {
    // The mirror case. The band between two columns is narrow but tall, so a
    // chip wider than the hop leans on the services either side of it. Narrow
    // it to the corridor — but only while the result stays a chip: squeezing
    // every label to the hop width is what once produced a 0.34" ribbon inches
    // tall, which `prefer` exists to prevent.
    const narrowed = clamp(gap, 0.34 * px, 1.5 * px);
    if (narrowed < maxW) {
      const lineH0 = (fontSize * 1.3) / 72;
      const lineCount = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / chipColumn(narrowed)));
      if (lineCount * lineH0 + 0.06 <= 0.9 * px) roomW = narrowed;
    }
  }
  // `roomW` is the last word on how wide this chip may be, so it is the only
  // place the "no room for a letter" test can be made. Testing against the cap
  // instead let a chip whose corridor was far narrower than the cap survive it.
  if (text !== '' && roomW + 0.002 < minChipW) {
    // The bar is about the room this column has AT THIS SIZE, and smaller type
    // needs a narrower column: `minChipW` scales with the font while `roomW`
    // is floored at `0.34 * px`, which does not. So come down half a point at
    // a time to the legibility floor before concluding that nothing can be
    // drawn here — a fan whose rungs refuse the sentence at 8.4pt sets it
    // perfectly well at 7pt, and dropping it instead loses seven labels that
    // the reader could have read. Terminates: the size strictly decreases by
    // half a point toward a fixed floor, and the floor never retries.
    // Only for a ladder. A bundle is laid out at whatever size makes its rungs
    // fit, so trading type size for room is already its contract; an ordinary
    // chip's size is the reader's floor and shrinking it to win a placement
    // moves the chip nearer a stranger's arrow than its own.
    // Floored at the READING slide's legibility floor, not the overview's.
    // `Math.min(OVERVIEW_LEGIBLE_PT, requestedFontSize)` resolved to 6 on every
    // reading slide, because the requested size there is already >= 7 — so a
    // ladder could step down to 6.0pt on a full-size slide, under the floor
    // `labelFontSize` itself enforces and under the gate's `minFont < 7` rule.
    // On the overview the requested size is the overview's own floor, so taking
    // the minimum of the two keeps that case unchanged.
    const floor = Math.min(requestedFontSize, LEGIBLE_TILE_PT);
    if (bundle && !bundle.badgesOnly && fontSize > floor + 0.01) {
      const smaller = Math.max(floor, fontSize - 0.5);
      const retry = connectorLabelBox(
        route, transform, requestedFontSize, px, clampTo, obstacles,
        { ...bundle, font: smaller }, squeezeTo, foreignGap,
      );
      if (retry) return retry;
    }
    // Nothing legible can be drawn here. A numbered hop still gets its callout
    // — the number is one glyph in a circle sized independently — and its
    // wording is on the workflow slide. An un-numbered one draws nothing at
    // all, which is what the reader would rather have than a smear.
    if (badgeD <= 0) return null;
    return connectorLabelBox(
      { ...route, label: ' ' }, transform, requestedFontSize, px, clampTo, obstacles,
      { ...(bundle ?? { count: 1, rung: 0, x: anchor.x, y: anchor.y }), badgesOnly: true },
      squeezeTo, foreignGap,
    );
  }
  const w = bare
    ? badgeD
    : clamp(
      naturalW <= roomW ? naturalW : roomW,
      // Floored against the TYPE, not against the drawing. `0.34 * px` scales
      // to nothing, so on a scaled overview it let the box shrink below one
      // letter while the font stayed at 6pt.
      Math.min(Math.max(0.34 * px, minChipW), roomW),
      roomW,
    );
  const perLine = chipColumn(w);
  const lineH = (fontSize * 1.3) / 72;

  // ceil(ink / column) is the break-anywhere ratio, and wrappedLineCount's
  // own doc comment says why it is wrong: three tokens each wider than half the
  // box take three lines where the ratio predicts two. This sizer was the one
  // consumer in the file that did not use the helper, and the under-count was
  // masked only for as long as the emitted chip gave the text 0.12in more
  // column than the model reserved. Closing that gap turned it into overflow —
  // 317 chips across 18 decks, one of them spilling 47% of its box.
  const lines = wrappedLineCount(text, perLine, fontSize);
  const h = bare ? 0 : Math.max(0.16 * px, lines * lineH + 0.06);
  const badgeGap = bare ? 0 : 0.03;

  // Slide the chip along the edge's normal, never across it, so it still reads
  // as belonging to that arrow. A rung of the ladder has to clear the chip AND
  // the numbered callout that hangs off it, or ordinal n's badge is painted
  // inside ordinal n+1's label.
  // The numbered callout hangs off the chip, so the two are placed as one
  // block: scoring the chip alone let a badge land inside a neighbour's label,
  // and the badge could not be moved without moving the chip it belongs to.
  const blockH = h + (badgeD > 0 ? badgeD + badgeGap : 0);
  // Every chip in a bundle steps by the SAME amount, measured from the tallest
  // block in it, and both the stagger and the obstacle walk move in whole
  // steps. That puts the bundle on a lattice: a chip pushed off a tile lands on
  // a free rung instead of half inside the neighbour above it.
  const natural = alongX ? blockH + 0.05 : Math.max(w / 2 + 0.06, blockH);
  // A ladder on a vertical arrow steps sideways, so its rung has to be the
  // chip's WIDTH plus a gap. Reusing the block height there — or `natural`,
  // which only ever guarantees half the width — left every rung overlapping its
  // neighbour by nearly half a chip, and only the obstacle walk hid it.
  const ladderStep = alongX ? blockH + 0.05 : w + 0.12;
  const stepOut = bundle && bundle.rung > 0 ? Math.max(bundle.rung, ladderStep) : natural;

  // De-collide parallel-edge chips: give each ordinal its own rung, centred on
  // the bundle. A fan too deep to fit even at the smallest readable size wraps
  // into a second column rather than restacking on the page edge.
  //
  // The rung is chosen by where the ARROW was fanned to, not by the ordinal.
  // `parallelOffset` alternates about the centre (0, +16, -16 …) while an
  // ordinal ladder runs straight down, so ranking by ordinal drew callout 1
  // beside the middle arrow and callout n beside the bottom one — six of seven
  // callouts against the wrong arrow at n = 7.
  const rung = bundle?.rank?.get(route.id) ?? route.ordinal;
  const perCol = Math.max(1, bundle?.perCol ?? siblings);
  const columns = Math.max(1, Math.ceil(siblings / perCol));
  const column = Math.floor(rung / perCol);
  const inColumn = Math.min(perCol, siblings - column * perCol);
  const stagger = (rung % perCol - (inColumn - 1) / 2) * stepOut;
  const columnShift = (column - (columns - 1) / 2) * (w + 0.12);

  // Clamp the whole ladder, not each rung. Holding every parallel chip inside
  // the frame on its own collapses the ladder onto the page edge and stacks the
  // ordinals right back on top of each other; shifting the anchor so the ladder
  // as a whole fits keeps the rungs their full step apart.
  // Parallel routes are already fanned apart by a fraction of a rung, so
  // stacking the stagger on top of each route's own anchor puts the chips off
  // the lattice and half inside each other. The ladder hangs off the bundle's
  // centre; only the ordinal decides which rung this chip takes.
  const base = bundle
    ? {
      x: bundle.x + (alongX ? bundle.across ?? 0 : bundle.shift ?? 0),
      y: bundle.y + (alongX ? bundle.shift ?? 0 : bundle.across ?? 0),
    }
    : anchor;
  const reach = ((perCol - 1) / 2) * stepOut;
  let ladder = 0;
  if (clampTo && reach > 0) {
    const size = alongX ? blockH : w;
    const centre = alongX ? base.y : base.x;
    const lo = alongX ? clampTo.y : clampTo.x;
    const hi = lo + (alongX ? clampTo.h : clampTo.w);
    if (centre - size / 2 - reach < lo) ladder = lo - (centre - size / 2 - reach);
    else if (centre + size / 2 + reach > hi) ladder = hi - (centre + size / 2 + reach);
  }
  // Its own hop's line is not an obstacle — a chip is supposed to sit on it.
  // Everyone else's is: a label resting on a neighbouring arrow reads as that
  // arrow's label, and the reader never finds out they misread it.
  const ownerKey = route.sourceId < route.targetId
    ? `${route.sourceId}|${route.targetId}`
    : `${route.targetId}|${route.sourceId}`;
  const seen = obstacles.filter((tile) => tile.owner === undefined || tile.owner !== ownerKey);
  const place = (offset: number, cross = 0): { x: number; y: number; clamped: boolean } => {
    const rawX = (bundle ? base.x : alongX ? anchor.x : base.x) - w / 2 + (alongX ? cross + columnShift : offset + ladder);
    const rawY = (bundle ? base.y : alongX ? base.y : anchor.y) - blockH / 2 + (alongX ? offset + ladder : cross + columnShift);
    let x = rawX;
    let y = rawY;
    if (clampTo) {
      x = clamp(x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - w));
      y = clamp(y, clampTo.y, Math.max(clampTo.y, clampTo.y + clampTo.h - blockH));
    }
    return { x, y, clamped: x !== rawX || y !== rawY };
  };
  // Where the arrow this chip labels actually runs. Declared up here, ahead of
  // the cost functions, because the numbered disc's cost depends on which end
  // of the block the disc hangs from — leaving these below `covered` worked
  // only because nothing happened to call it early enough to trip the
  // temporal dead zone.
  const ownSegments: { ax: number; ay: number; bx: number; by: number }[] = [];
  for (let i = 1; i < route.points.length; i += 1) {
    const a = toInches(route.points[i - 1], transform);
    const b = toInches(route.points[i], transform);
    ownSegments.push({ ax: a.x, ay: a.y, bx: b.x, by: b.y });
  }
  const toOwn = (cx: number, cy: number): number => {
    let bestGap = Number.POSITIVE_INFINITY;
    for (const seg of ownSegments) {
      const dx = seg.bx - seg.ax;
      const dy = seg.by - seg.ay;
      const len = dx * dx + dy * dy;
      const t = len > 0 ? Math.max(0, Math.min(1, ((cx - seg.ax) * dx + (cy - seg.ay) * dy) / len)) : 0;
      bestGap = Math.min(bestGap, Math.hypot(cx - (seg.ax + t * dx), cy - (seg.ay + t * dy)));
    }
    return bestGap;
  };
  // Which end of the block the numbered callout hangs from. It used to be
  // always the bottom, which on a chip that hangs BELOW its arrow puts the one
  // mark a reader uses to identify the hop as far from that hop as the block
  // allows — on a stack of parallel rows, nearer the next row's arrow than its
  // own. The Architecture Center draws the number against the arrow it
  // numbers, so hang it from whichever end of the block faces that arrow.
  const badgeAtTopAt = (x: number, y: number): boolean => {
    if (badgeD <= 0 || bare) return false;
    const cx = x + w / 2;
    const topY = y + badgeD / 2;
    const botY = y + blockH - badgeD / 2;
    const top = toOwn(cx, topY);
    const bot = toOwn(cx, botY);
    // A polyline that doubles back puts both ends of the block the same
    // distance from the arrow by construction — `toOwn` takes the minimum over
    // every segment, so the two tie. Neither end is then better for reading the
    // number against its own hop, and the question that still has an answer is
    // which end is furthest from somebody else's.
    if (Math.abs(top - bot) < 0.01) {
      return foreignGap ? foreignGap(cx, topY) > foreignGap(cx, botY) : false;
    }
    return top < bot;
  };
  const badgeYAt = (x: number, y: number): number => (badgeAtTopAt(x, y) ? y : y + h + badgeGap);
  const textYAt = (x: number, y: number): number => (badgeAtTopAt(x, y) ? y + badgeD + badgeGap : y);
  // Where along the edge it hangs from the disc sits.
  //
  // Nailed to the middle of the chip it is the one part of the block that
  // cannot be moved off a tile: the walk can only take the whole block with it,
  // the wording is several times wider than the disc, so the block comes to
  // rest where the WORDING fits and drags the number onto a service. A chip
  // squeezed into a corridor narrower than itself has to overlap something —
  // but the 0.21in disc almost always has somewhere clear to be along the
  // 0.67in edge it hangs from. Sliding it there keeps it touching its own chip,
  // so it still reads as that hop's number, and takes it off the icon.
  //
  // The centre is preferred outright on a tie, so every placement that was
  // already clear is unchanged.
  const badgeXAt = (at: { x: number; y: number }): number => {
    const centre = at.x + w / 2 - badgeD / 2;
    if (bare || w <= badgeD + 0.02) return centre;
    const by = badgeYAt(at.x, at.y);
    const costAt = (bx: number): number => seen.reduce((sum, tile) => {
      const dx = Math.min(bx + badgeD, tile.x + tile.w) - Math.max(bx, tile.x);
      const dy = Math.min(by + badgeD, tile.y + tile.h) - Math.max(by, tile.y);
      return dx > 0 && dy > 0 ? sum + dx * dy * (tile.annotation ? ANNOTATION_WEIGHT : tile.weight ?? 1) : sum;
    }, 0);
    let bestX = centre;
    let bestCost = costAt(centre);
    // Sliding is only ever worth it while the disc still reads as this hop's
    // number. Moved along the edge it can end up nearer a different arrow than
    // its own, which trades a hidden number for a number attached to the wrong
    // sentence — strictly the worse of the two.
    const attrAt = (bx: number): number => {
      if (!foreignGap) return Number.POSITIVE_INFINITY;
      return foreignGap(bx + badgeD / 2, by + badgeD / 2) - toOwn(bx + badgeD / 2, by + badgeD / 2);
    };
    const floor = Math.min(attrAt(centre), 0) - 0.001;
    const slots = 8;
    for (let i = 0; i <= slots && bestCost > 0; i += 1) {
      const bx = at.x + (i / slots) * (w - badgeD);
      const cost = costAt(bx);
      if (cost < bestCost - 0.0001 && attrAt(bx) >= floor) {
        bestX = bx;
        bestCost = cost;
      }
    }
    return bestX;
  };
  // The wording has an opaque chip behind it and is expected to stand in the
  // corridors between tiles, so a graze costs it little. The number does not:
  // it is a small solid disc, and dropped on a tile it hides part of an icon
  // and is the one mark the workflow list cannot be read without. Charged only
  // as block area it is cheap to sacrifice — the walk will happily park the
  // disc dead centre on a service to keep the much larger text clear, which is
  // how a wrap-around hop's callout came to sit 100% inside a stranger's tile.
  // So the disc is priced on its own footprint.
  //
  // It carried a x4 premium for several rounds. That premium is now inert:
  // annotation overlap is priced at x8 and a swallowed disc leaves its chip
  // altogether, so an A/B at 4 and at 0 across five grid pitches was
  // byte-identical on four of them and moved 0.5% of one disc's ink on the
  // fifth, whose worst burial was unchanged. An untested constant is worse
  // than no constant, so it is gone rather than gated.
  const badgeCovered = (at: { x: number; y: number }): number => {
    if (badgeD <= 0 || blockH <= h + 0.01) return 0;
    // Measured at the centre of the edge, not at the slid position. The slide
    // is a repair applied once the block has come to rest; letting the walk
    // count on it makes the cost optimistic, and the block then settles
    // somewhere the wording is worse off for the sake of a disc that had
    // somewhere to go anyway.
    const bx = at.x + w / 2 - badgeD / 2;
    const by = badgeYAt(at.x, at.y);
    return seen.reduce((sum, tile) => {
      if (tile.annotation) return sum;
      const dx = Math.min(bx + badgeD, tile.x + tile.w) - Math.max(bx, tile.x);
      const dy = Math.min(by + badgeD, tile.y + tile.h) - Math.max(by, tile.y);
      return dx > 0 && dy > 0 ? sum + dx * dy * (tile.weight ?? 1) : sum;
    }, 0);
  };
  const covered = (at: { x: number; y: number }): number => seen.reduce((sum, tile) => {
    const dx = Math.min(at.x + w, tile.x + tile.w) - Math.max(at.x, tile.x);
    const dy = Math.min(at.y + blockH, tile.y + tile.h) - Math.max(at.y, tile.y);
    return dx > 0 && dy > 0 ? sum + dx * dy * (tile.annotation ? ANNOTATION_WEIGHT : tile.weight ?? 1) : sum;
  }, 0) + badgeCovered(at);
  const onLabel = (at: { x: number; y: number }): number => seen.reduce((sum, tile) => {
    if (!tile.annotation) return sum;
    const dx = Math.min(at.x + w, tile.x + tile.w) - Math.max(at.x, tile.x);
    const dy = Math.min(at.y + blockH, tile.y + tile.h) - Math.max(at.y, tile.y);
    return dx > 0 && dy > 0 ? sum + dx * dy : sum;
  }, 0);
  // A chip is only readable as belonging to its arrow while it is near it, so
  // distance from the natural spot is part of the cost, not a free move. Left
  // unpriced, a chip squeezed out of a crowded band walked over 6in from the
  // 0.9in hop it labels and its callout came to rest beside a different arrow —
  // an unattributable number is worse than a small overlap.
  const home = place(stagger);
  // How far a chip may be moved before it stops reading as this arrow's label.
  // Roughly half a chip sideways or a couple of rows up, which on any ordinary
  // layout still leaves it nearer its own arrow than any neighbouring one. Past
  // that a small overlap is the better trade: the reviewer measured a chip
  // walking 6in from the 0.9in hop it labelled, taking its numbered callout to
  // rest beside a completely different arrow.
  const chipReach = Math.max(1.2 * blockH, 0.6 * w, 0.5);
  // Sitting on another label is the one thing worth a long walk: nothing is
  // readable there at all. Sitting on a tile is not, so that search stays
  // inside the radius where the chip still plainly belongs to its own arrow.
  const limit = onLabel(home) > 0 ? 2 * chipReach : chipReach;
  // And in no case may it leave the run of its own hop. Distance alone is a
  // poor test on a grid — a chip one row up is only an inch away but reads as
  // the label of a completely different arrow — so the walk is bounded by where
  // the arrow actually goes, not by a radius.
  //
  // Half a chip of slack past the end of the hop is only defensible when the
  // chip is too wide for the hop in the first place. When it fits, that slack
  // is simply permission to lap onto the service at the far end — measured at
  // 13% of a tile on a plain 40-service chain — so a chip that fits is held
  // entirely inside its own run.
  const runSpan = alongX ? Math.abs(last.x - first.x) : Math.abs(last.y - first.y);
  const chipSpan = alongX ? w : blockH;
  const pad = chipSpan <= runSpan ? -chipSpan / 2 : chipSpan / 2;
  const runLo = (alongX ? Math.min(first.x, last.x) : Math.min(first.y, last.y)) - pad;
  const runHi = (alongX ? Math.max(first.x, last.x) : Math.max(first.y, last.y)) + pad;
  const inReach = (at: { x: number; y: number }): boolean => {
    const centre = alongX ? at.x + w / 2 : at.y + blockH / 2;
    if (centre < runLo || centre > runHi) return false;
    const drift = alongX ? at.y - home.y : at.x - home.x;
    return Math.abs(drift) <= limit + 0.001;
  };
  // And it must still read as THIS arrow's label. The reach above is a radius,
  // which on a grid is the wrong shape: a clear slot a row away is well inside
  // it and sits right beside a different hop. So a candidate also has to stay
  // nearer its own line than anybody else's, which is exactly what a reader
  // does when they match a chip to an arrow. `ownSegments` and `toOwn` are
  // declared with the cost functions above, which need them too.
  const nearby = obstacles.filter((tile) => tile.owner !== undefined && tile.owner !== ownerKey
    && tile.x - (limit + w) <= home.x + w && tile.x + tile.w + limit + w >= home.x
    && tile.y - (limit + blockH) <= home.y + blockH && tile.y + tile.h + limit + blockH >= home.y);
  const attributable = (at: { x: number; y: number }): boolean => {
    if (ownSegments.length === 0) return true;
    const cx = at.x + w / 2;
    // The centre of the wording, and the row where the numbered callout hangs.
    // A reader attributes each of those to the arrow nearest to it, so each has
    // to be tested where it is actually drawn: the centre of the block that
    // contains both belongs to neither, and on a chip with a badge it sits a
    // quarter inch below the text it is supposed to stand for.
    const sampleYs = blockH > h + 0.01
      ? [textYAt(at.x, at.y) + h / 2, badgeYAt(at.x, at.y) + badgeD / 2]
      : [at.y + h / 2];
    for (const cy of sampleYs) {
      const mine = toOwn(cx, cy);
      if (foreignGap && foreignGap(cx, cy) < mine - 0.1) return false;
      for (const tile of nearby) {
        const gapToOther = Math.hypot(
          cx - Math.max(tile.x, Math.min(cx, tile.x + tile.w)),
          cy - Math.max(tile.y, Math.min(cy, tile.y + tile.h)),
        );
        if (gapToOther < mine - 0.35) return false;
      }
    }
    return true;
  };
  // A chip sitting where it reads as somebody else's label is a defect worth
  // about half a covered tile, which is what the ladder search already charges
  // for a misread rung. It is priced, not forbidden: the natural spot is taken
  // unconditionally whenever nothing overlaps it, so without a price a rung
  // whose home happens to be clean never looks for a better slot at all - and
  // that is how a callout came to rest 0.32in from a stranger's arrow while
  // its own ran 0.67in away. The walk still refuses to *move* to an
  // unattributable slot, so this can only ever improve a placement.
  const MISREAD_COST = 0.5;
  const score = (at: { x: number; y: number }): number => {
    const drift = Math.hypot(at.x - home.x, at.y - home.y);
    return covered(at) + DRIFT_COST_PER_IN * drift + (attributable(at) ? 0 : MISREAD_COST);
  };

  // A chip centred on its arrow lands on top of a tile whenever the free gap
  // between the two services is narrower than the label — routine on a dense
  // architecture, and unavoidable once a long CJK label meets a 0.5" hop. Walk
  // it outwards along the normal and take the first clear position, falling
  // back to the least-obscured one so a chip can never simply be dropped. The
  // walk moves in half steps and ranges well past the ladder, because on a
  // fan of parallel edges a chip pushed off a tile lands squarely in the slot
  // of the ordinal above it if it can only stop where another rung already is.
  let best = home;
  let bestScore = score(home);
  for (let step = 1; bestScore > 0 && step <= 16; step += 1) {
    for (const sign of [1, -1]) {
      // A rung keeps its slot on the ladder and dodges sideways instead: the
      // ladder has to cross the band the tiles sit in, so moving along it just
      // trades one covered tile for another, and a rung that leaves the lattice
      // lands part-way inside its neighbour.
      const out = sign * step * (stepOut / 2);
      const side = sign * step * (w / 2 + 0.06);
      const candidates = siblings > 1
        ? [place(stagger, side)]
        : [place(stagger + out), place(stagger, side)];
      for (const candidate of candidates) {
        if (siblings > 1 && candidate.clamped) continue;
        if (!inReach(candidate) || !attributable(candidate)) continue;
        const cost = score(candidate);
        if (cost < bestScore) {
          best = candidate;
          bestScore = cost;
        }
        if (bestScore <= 0) break;
      }
      if (bestScore <= 0) break;
    }
  }
  // A lone chip is free to go anywhere, so when neither axis on its own found a
  // clear slot, search the plane around the arrow. The two axes have to move
  // independently: the gap on a grid is often a short step off the line and a
  // long step along it, and a search that only walks the diagonal steps
  // straight past it onto the next hop's corner.
  if (siblings <= 1) {
    const alongStep = stepOut / 2;
    const acrossStep = w / 2 + 0.06;
    // Far enough to cross the whole drawing. `ring <= 10` was a budget in units
    // of the chip's own size, which is the same mistake the ladder search made
    // and had corrected: a chip that has already failed to place is by
    // definition surrounded, so the clear paper it needs is rarely within ten
    // of its own widths, and the walk stopped short of it every time. The
    // ladder's answer works here unchanged — reach set by the page, and a
    // coarse pass to find the clear region before a fine one finds the point,
    // because scanning every lattice point out to the far edge is quadratic.
    const reach = clampTo
      ? Math.min(48, Math.ceil(Math.max(
        clampTo.w / Math.max(alongStep, 0.05),
        clampTo.h / Math.max(acrossStep, 0.05),
      )))
      : 10;
    const rings = Math.max(10, reach);
    const pass = (stride: number, centreA: number, centreB: number, span: number): void => {
      const limit = Math.max(1, Math.ceil(span / stride));
      for (let ring = 1; bestScore > 0 && ring <= limit; ring += 1) {
        for (let a = -ring; a <= ring && bestScore > 0; a += 1) {
          for (let b = -ring; b <= ring; b += 1) {
            if (Math.max(Math.abs(a), Math.abs(b)) !== ring) continue;
            const candidate = place(
              stagger + (centreA + a * stride) * alongStep,
              (centreB + b * stride) * acrossStep,
            );
            if (!inReach(candidate) || !attributable(candidate)) continue;
            const cost = score(candidate);
            if (cost < bestScore) {
              best = candidate;
              bestScore = cost;
              bestA = centreA + a * stride;
              bestB = centreB + b * stride;
            }
            if (bestScore <= 0) break;
          }
        }
      }
    };
    let bestA = 0;
    let bestB = 0;
    pass(4, 0, 0, rings);
    if (bestScore > 0) pass(1, bestA, bestB, 4);
  }
  const badge = badgeD > 0
    ? { x: badgeXAt(best), y: badgeYAt(best.x, best.y), d: badgeD }
    : null;
  // Still standing on something. A chip is allowed to be wider than the gap it
  // labels, but when that width is the reason it has nowhere to go, wrapping it
  // into the lane between the two services is the better trade: two short lines
  // inside the gap beat one long line across a tile. Failing that, a point of
  // type buys the room the walk could not find.
  if (bestScore > 0 && squeezeTo === undefined && !bundle && text) {
    const retries: { font: number; width: number }[] = [];
    // The lane is worth trying even when it is narrow. Gating this at 0.3in
    // meant a long CJK label on a tight hop had no squeeze available at all and
    // could only stand on a tile — a wrapped chip inside the lane is the better
    // picture, and the width floor keeps it from becoming a ribbon.
    const lane = Math.max(gap, 0.34 * px);
    if (lane < w) retries.push({ font: requestedFontSize, width: lane });
    const smaller = Math.max(LEGIBLE_TILE_PT, requestedFontSize - 1);
    if (smaller < requestedFontSize - 0.01) {
      retries.push({ font: smaller, width: lane < w ? lane : maxW });
    }
    let bestBox: ReturnType<typeof connectorLabelBox> = null;
    for (const retry of retries) {
      const tighter = connectorLabelBox(
        route, transform, retry.font, px, clampTo, obstacles, bundle, retry.width, foreignGap,
      );
      if (tighter && tighter.stuck < (bestBox?.stuck ?? bestScore)) bestBox = tighter;
      if (bestBox && bestBox.stuck <= 0) break;
    }
    if (bestBox) return bestBox;
  }
  // Nowhere on the slide reads as this arrow's own label. A numbered hop can
  // give up its wording — the workflow list on the slide still carries the
  // sentence against the same number — and that is strictly better than parking
  // it where it will be read as the label of a different arrow.

  return { x: best.x, y: textYAt(best.x, best.y), w, h, text, badge, block: { x: best.x, y: best.y, w, h: blockH, annotation: true }, alongX, fontSize, stuck: bestScore };
}

function addConnectorLabel(
  slide: Slide,
  route: ExportRoute,
  fontSize: number,
  box: ReturnType<typeof connectorLabelBox>,
): void {
  if (!box || !box.text) return;

  slide.addText(box.text, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    shape: 'roundRect',
    rectRadius: 0.03,
    fill: { color: 'FEF9C3', transparency: 8 },
    line: { color: 'FDE68A', width: 0.5 },
    // Amber-800, not amber-700. The chip is a fixed light yellow in both
    // themes, but it is 8% translucent — so on a dark slide the backdrop
    // darkens it to about #ECE8B8 and amber-700 lands at 4.02:1, under the
    // WCAG AA bar. One step darker clears both composites with margin and is
    // indistinguishable on the light one.
    color: '92400E',
    fontSize,
    fontFace: 'Yu Gothic UI',
    align: 'center',
    valign: 'middle',
    margin: CHIP_INSET_IN * 72,
    wrap: true,
    objectName: `connector-label-${route.id}`,
  });
}

/**
 * The widest a numbered callout may be drawn, from the two tiles the arrow it
 * sits on runs between.
 *
 * PowerPoint sized its disc as `clamp(0.26 * px, 0.18, 0.42)`, which has a
 * floor and a ceiling in absolute inches and no reference to the tile at all.
 * `px` is the drawn width of 96 authored pixels, so on a tile of W authored px
 * the ratio is `0.26 * 96 / W` and does not move with the scale: a 14px node
 * drew a 0.3566in disc on a 0.2000in tile, 178% of the service it was calling
 * out, and a 24px node drew one 104% of its tile. This is the same defect the
 * drawing exporter was corrected for four times over, and it was live in the
 * primary output format the whole time.
 *
 * Returns the legibility floor when a legible disc cannot also be a
 * proportionate one. That case is real - the smallest circle that holds a
 * readable digit is 0.18in, so a tile under about 0.33in drawn has no diameter
 * that is both - and the answer is not to drop the callout: the workflow slide
 * cites step numbers and every one of them has to be findable on the canvas,
 * which is a promise the deck keeps and a rule the gate already enforces. The
 * disc goes to the floor, and where even the floor is over the ceiling the
 * gate reports the empty intersection against the TILE, which is the only
 * object in that picture behaving badly.
 */
const BADGE_TILE_SHARE = 0.55;
/**
 * The smallest type a callout may be set in, in points.
 *
 * The deck's own legibility floor, the same constant the tile labels use. It
 * is here rather than inline because the disc's floor is derived from it and
 * the two must not drift.
 */
const BADGE_LEGIBLE_PT = LEGIBLE_TILE_PT;
/**
 * Diameter per inch of type, for a number of this many digits.
 *
 * PowerPoint lays a callout out as a text box centred in an ellipse, so the
 * thing that has to fit is a box of `digits * 0.62` by `1.3` ems - the same
 * model the gate's own bubble rule uses, and the line height is why this is
 * not the drawing exporter's chord formula. The box is inscribed rather than
 * merely contained, because a box that fits the bounding SQUARE still pushes
 * its corners through the circle, and a tenth of the disc is kept as a ring
 * because a white number with no dark disc behind it is not a callout.
 *
 * What this replaces is a FLAT 0.18in floor. A floor that does not move with
 * the type it holds is not a legibility floor, it is a magic number that
 * happens to be safe: for a 7pt digit the drawing exporter derived 0.1119in
 * and PowerPoint used 0.18in, 61% larger, and that one constant was the whole
 * of the 93%-versus-55% disagreement between the two formats on the same
 * nine-node diagram.
 */
function badgeDiameterPerIn(digits: number): number {
  return Math.hypot(Math.max(1, digits) * 0.62, 1.3) / 0.9;
}
function badgeFloorIn(stepNumber: number, fontPt: number): number {
  const digits = String(Math.max(1, Math.abs(Math.trunc(stepNumber)))).length;
  return (fontPt / 72) * badgeDiameterPerIn(digits);
}
/**
 * The largest type that fits inside a disc of this diameter.
 *
 * The disc is sized by the tile and the type is then sized by the disc, so the
 * two can never disagree. Before this the type was chosen from the chip beside
 * it and the disc from a constant, which is why the file needs a rule watching
 * for numbers that run outside their own bubble at all.
 */
function badgeFontPtFor(stepNumber: number, diameterIn: number): number {
  const digits = String(Math.max(1, Math.abs(Math.trunc(stepNumber)))).length;
  return (diameterIn / badgeDiameterPerIn(digits)) * 72;
}
/**
 * The index row a service gets when its tile is on the canvas and carries no
 * mark at all.
 *
 * A LOCATOR, not just a confession. `(drawn unlabelled)  =  Private DNS zone`
 * tells the reader that one of the boxes in front of them is this service and
 * nothing more, and it is the same string for every dark service in the deck,
 * so on a drawing with several of them it does not even narrow the field. That
 * is the whole of the reader's route on a 12px node sharing a drawing with
 * 150px ones: measured on `probe-glyph12`, the authored name "Private DNS zone"
 * occurs exactly ONCE in the entire export, in that row, while the service is
 * cited by numbered steps 1 and 2.
 *
 * Naming it on the canvas is not available. A mark needs `MARKABLE_TILE_W_IN`,
 * which for a 12px node is 0.0167 in/px, which draws its 150px neighbours 2.5in
 * wide and shreds the deck to 24 slides with 19 of them carrying a single tile
 * - the defect the planner's own bound exists to refuse. So the route is given
 * where the deck already reserves room for exactly this: the index. The part
 * label is the one the slide title already shows the reader, and the ordinal is
 * reading order among the tiles on that slide, so both halves are things the
 * reader can count off the page they are holding.
 */
const UNLABELLED_PREFIX = '(drawn unlabelled';
const UNLABELLED_ROW = `${UNLABELLED_PREFIX})`;
const unlabelledRow = (at: string): string => (at ? `${UNLABELLED_PREFIX}, ${at})` : UNLABELLED_ROW);

function stepBadgeDiameterIn(route: ExportRoute, transform: FitTransform, px: number): number {
  const step = route.stepNumber ?? 1;
  const floor = badgeFloorIn(step, BADGE_LEGIBLE_PT);
  const natural = clamp(0.26 * px, floor, 0.42);
  const ends = [route.sourceW, route.targetW]
    .filter((w) => typeof w === 'number' && w > 0)
    .map((w) => w * transform.scale);
  if (ends.length === 0) return natural;
  // The narrower end, not the average. A disc that is proportionate to one tile
  // and swamps the other is still swamping a tile.
  const ceiling = Math.min(...ends) * BADGE_TILE_SHARE;
  return Math.max(floor, Math.min(natural, ceiling));
}

/**
 * Whether this hop's tiles admit no callout at all.
 *
 * Under `floor / 0.55` the legible diameter and the proportionate one do not
 * intersect, so whatever is drawn is either unreadable or swamps the service
 * it points at. Every slide but the overview can answer that by splitting
 * further; the overview cannot, so it is the one place the callout is dropped.
 */
function stepBadgeConflicts(route: ExportRoute, transform: FitTransform): boolean {
  const ends = [route.sourceW, route.targetW]
    .filter((w) => typeof w === 'number' && w > 0)
    .map((w) => w * transform.scale);
  if (ends.length === 0) return false;
  return Math.min(...ends) * BADGE_TILE_SHARE
    < badgeFloorIn(route.stepNumber ?? 1, BADGE_LEGIBLE_PT) - 1e-6;
}

/**
 * Numbered callout on a connector, matching the workflow list.
 *
 * Reference architectures on the Azure Architecture Center number every arrow
 * and repeat those numbers in the prose, so a deck exported without them makes
 * the reader guess which sentence describes which hop. Drawn as a real
 * PowerPoint oval so it stays editable rather than being baked into an image.
 */
function stepBadgeBox(
  route: ExportRoute,
  transform: FitTransform,
  px: number,
  clampTo: DiagramFrame | undefined,
  chip: ReturnType<typeof connectorLabelBox>,
  obstacles: readonly Obstacle[] = [],
  ownGap?: (x: number, y: number) => number,
  foreignGap?: (x: number, y: number) => number,
  thumbnail = false,
): { x: number; y: number; d: number } | null {
  if (route.stepNumber === undefined) return null;
  // Not on the overview, when the tile cannot carry one.
  //
  // The overview is the whole drawing shown small on one slide, so its tiles
  // shrink with the size of the estate and no scale choice can rescue them: at
  // 120 services the disc drew 56% of its tile, at 240 it drew 78%, and the
  // trend has no end. Every other slide can be split until the tile is big
  // enough; this one cannot, by construction.
  //
  // Dropping it costs nothing the reader needs. The overview is a locator - it
  // exists so the reader can see where each part sits in the whole - and the
  // numbered story is told on the reading windows, which is also where the
  // rule requiring every cited step to be findable does its counting.
  //
  // Tried and reverted: dropping it on reading windows too, where the planner
  // has refused the split and the disc comes out wider than its tile. That
  // trades a swamped icon for a broken workflow band - the audit reported 20
  // issues of the form "workflow cites steps 2, 3 with no callout on the
  // canvas" - and a step the reader cannot locate at all is worse than one
  // drawn too large. The fix has to be in the plan, not in the disc.
  if (thumbnail && stepBadgeConflicts(route, transform)) return null;

  // No chip to hang off: either an unlabelled but numbered hop, or one whose
  // wording was muted because it had nowhere legible to stand. The anchor is
  // the natural home, but it is the middle of the arrow, which on a dense
  // drawing is routinely the middle of a tile — and a number printed over an
  // icon is the one thing on the slide the workflow list cannot survive
  // without. So walk outwards for a clear slot the way a chip does.
  const anchor = toInches(route.labelAnchor, transform);
  const d = stepBadgeDiameterIn(route, transform, px);
  const fit = (x: number, y: number): { x: number; y: number } => (clampTo
    ? {
      x: clamp(x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - d)),
      y: clamp(y, clampTo.y, Math.max(clampTo.y, clampTo.y + clampTo.h - d)),
    }
    : { x, y });
  // A number printed over another number cannot be read at all, while a number
  // over the corner of a tile can. So overlapping an existing annotation — a
  // chip block or an already-placed callout — is priced far above overlapping
  // a service. Without this the walk's own lattice, which is finer than the
  // disc it places, let a muted fan stack two callouts on each other: the
  // overlap was real but cost less than the clear slot further out.
  const ANNOTATION_OVERLAP_WEIGHT = 8;
  // A disc entirely inside one tile is not a callout any more — it reads as
  // that service's own badge, and the step list then describes a hop the
  // reader cannot find. A disc straddling the gutter between two tiles covers
  // almost the same area but is still unmistakably a callout, so plain area is
  // blind to the difference that matters. On a dense grid there is no clear
  // paper within reach at all, so without this the walk had no reason to move
  // and left the number in the middle of a service. Priced under the
  // misattribution penalty below: a number swallowed whole by a tile is *also*
  // misattributed — it reads as that service's own badge — so escaping is worth
  // slightly more than the price of landing near a stranger's arrow, where at
  // least the digit is still legible.
  const FULL_BURIAL_WEIGHT = 5;
  const deepest = (at: { x: number; y: number }, dd: number): number => {
    let worst = 0;
    for (const other of obstacles) {
      if (other.annotation) continue;
      const dx = Math.min(at.x + dd, other.x + other.w) - Math.max(at.x, other.x);
      const dy = Math.min(at.y + dd, other.y + other.h) - Math.max(at.y, other.y);
      if (dx > 0 && dy > 0) worst = Math.max(worst, (dx * dy) / (dd * dd));
    }
    return worst;
  };
  // A number hanging off its own chip is the best outcome there is, so the
  // chip's placement is taken whenever it is legible. But the disc is nailed
  // to the chip and can only slide along it, so on a grid whose gutters are
  // narrower than the disc the block comes to rest where the WORDING fits and
  // the number ends up buried in a tile — where it reads as that service's own
  // badge and the step list describes a hop the reader cannot find. A buried
  // disc is worth less than a free-standing one, so it falls through to the
  // walk below, which prices burial and misattribution against each other
  // instead of being unable to move at all.
  //
  // The bar is 0.7, not the 0.9 the walk itself uses. The two are asking
  // different questions. The walk's is "is this candidate slot ruined", and it
  // is choosing between real alternatives, so a near-miss there is genuinely
  // survivable. This one is "is the chip's slot so much better than anything
  // the walk could find that the walk need not even run", and 87% inside a
  // tile — the residue a grid one node wider than `tight-grid` leaves — is not.
  // Falling through does not commit the disc to moving: the walk keeps it
  // where it is unless it finds something better.
  if (chip?.badge && deepest(chip.badge, chip.badge.d) < 0.7) return chip.badge;
  const cover = (at: { x: number; y: number }): number => {
    let sum = 0;
    for (const other of obstacles) {
      const dx = Math.min(at.x + d, other.x + other.w) - Math.max(at.x, other.x);
      const dy = Math.min(at.y + d, other.y + other.h) - Math.max(at.y, other.y);
      if (dx > 0 && dy > 0) sum += dx * dy * (other.annotation ? ANNOTATION_OVERLAP_WEIGHT : 1);
    }
    sum += deepest(at, d) >= 0.9 ? FULL_BURIAL_WEIGHT * d * d : 0;
    return sum;
  };
  // A muted hop has nothing left but this number, so where it lands decides
  // which arrow the workflow list appears to be describing. The walk below is
  // free to travel more than an inch looking for clear paper, which on a
  // ladder is far enough to come to rest against a completely different hop -
  // measured at 0.32in from a stranger's arrow while its own ran 0.67in away.
  // So a slot that reads as somebody else's is priced like half a covered
  // badge rather than forbidden: the number is never dropped, it just prefers
  // to stay where it can still be attributed.
  const misread = (at: { x: number; y: number }): number => {
    if (!ownGap || !foreignGap) return 0;
    const cx = at.x + d / 2;
    const cy = at.y + d / 2;
    if (foreignGap(cx, cy) >= ownGap(cx, cy) - 0.1) return 0;
    // For a hop whose wording was muted this number is the whole label: the
    // step list describes an arrow, and if the number sits nearer a different
    // one the sentence is simply attached to the wrong hop. There is no chip
    // beside it to say otherwise, so for a bare callout misattribution is
    // priced above being covered rather than at half of it.
    return (chip ? 0.5 : 4) * d * d;
  };

  const cost = (at: { x: number; y: number }): number => cover(at) + misread(at);
  let spot = fit(anchor.x - d / 2, anchor.y - d / 2);
  let spotCover = cost(spot);
  const step = d * 0.9;
  for (let ring = 1; spotCover > 0 && ring <= 6; ring += 1) {
    for (let a = -ring; a <= ring && spotCover > 0; a += 1) {
      for (let b = -ring; b <= ring; b += 1) {
        if (Math.max(Math.abs(a), Math.abs(b)) !== ring) continue;
        const candidate = fit(anchor.x - d / 2 + a * step, anchor.y - d / 2 + b * step);
        const score = cost(candidate);
        if (score < spotCover - 0.0001) {
          spot = candidate;
          spotCover = score;
        }
        if (spotCover <= 0) break;
      }
    }
  }
  // Falling through is not the same as leaving. If the walk cannot better the
  // chip's own slot — which on a drawing with no clear paper at all is the
  // usual outcome — the disc stays where it belongs, beside its wording.
  if (chip?.badge && cost(chip.badge) <= spotCover) return chip.badge;
  return { x: spot.x, y: spot.y, d };
}
function addStepBadge(
  slide: Slide,
  route: ExportRoute,
  fontSize: number,
  box: { x: number; y: number; d: number } | null,
): void {
  if (!box) return;
  const { x, y, d } = box;
  // Sized by the disc, not by the chip beside it. The disc is sized by the
  // tile, so a callout on a small service gets a small disc, and type chosen
  // independently of it is type that runs to the rim or over it. Never larger
  // than asked for, so an ordinary deck is byte-identical.
  const fits = Math.min(fontSize, badgeFontPtFor(route.stepNumber ?? 1, d));

  slide.addText(String(route.stepNumber), {
    x,
    y,
    w: d,
    h: d,
    shape: 'ellipse',
    fill: { color: '1F2937' },
    line: { color: 'FFFFFF', width: 1.25 },
    color: 'FFFFFF',
    bold: true,
    fontSize: fits,
    fontFace: 'Yu Gothic UI',
    align: 'center',
    valign: 'middle',
    margin: 0,
    objectName: `connector-step-${route.id}`,
  });
}

/**
 * Where a box lands on the slide. `clampTo` pulls strays that sit outside the
 * fitted frame back onto the page so they stay visible (see
 * {@link chooseExportBounds}); without it an outlier is drawn off-slide and
 * simply disappears in PowerPoint.
 */
function placeBox(
  box: ExportBox,
  transform: FitTransform,
  clampTo?: DiagramFrame,
  clip = false,
): { x: number; y: number; w: number; h: number } {
  const topLeft = toInches({ x: box.x, y: box.y }, transform);
  const w = box.w * transform.scale;
  const h = box.h * transform.scale;
  if (!clampTo) return { x: topLeft.x, y: topLeft.y, w, h };
  // A zone is routinely wider than the band it is drawn on, and clamping only
  // the origin would slide the whole rectangle to the left margin at full
  // width, painting it across the band and enclosing tiles that are not in it.
  // PowerPoint's writer makes that worse: handed a width larger than a slide
  // it emits the raw inch count as EMU and the boundary disappears. So a zone
  // that meets the frame is cut at its edge.
  if (clip) {
    const left = Math.max(topLeft.x, clampTo.x);
    const top = Math.max(topLeft.y, clampTo.y);
    const right = Math.min(topLeft.x + w, clampTo.x + clampTo.w);
    const bottom = Math.min(topLeft.y + h, clampTo.y + clampTo.h);
    if (right > left && bottom > top) return { x: left, y: top, w: right - left, h: bottom - top };
    // A zone that misses the frame entirely is a trimmed outlier, and cutting
    // that leaves a hairline at an off-page coordinate -- the "outlier
    // silently disappears" defect again. Clamp it back on instead, shrinking
    // it if it is larger than the frame so no dimension can escape the page.
    const cw = Math.min(w, clampTo.w);
    const ch = Math.min(h, clampTo.h);
    return {
      x: clamp(topLeft.x, clampTo.x, clampTo.x + clampTo.w - cw),
      y: clamp(topLeft.y, clampTo.y, clampTo.y + clampTo.h - ch),
      w: cw,
      h: ch,
    };
  }
  return {
    x: clamp(topLeft.x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - w)),
    y: clamp(topLeft.y, clampTo.y, Math.max(clampTo.y, clampTo.y + clampTo.h - h)),
    w,
    h,
  };
}

/**
 * Records that `mark` is what some slide actually drew for the service authored
 * as `authored`.
 *
 * A service can be drawn more than once in a tiled deck - small on the overview
 * and larger on its reading slide - and the two tiles are different widths, so
 * the same name legitimately shortens to two different stubs. Keeping only the
 * longest of them left the other one drawn on a tile and defined nowhere, which
 * is the exact defect the index exists to prevent. Every distinct stub is
 * recorded, and the index row lists them all against the one name they mean.
 */
function recordMark(into: Map<string, Set<string>> | undefined, authored: string, mark: string): void {
  if (!into) return;
  const marks = into.get(authored);
  if (marks) marks.add(mark);
  else into.set(authored, new Set([mark]));
}

function addNodeShape(
  pptx: PptxGenJS,
  slide: Slide,
  box: ExportBox,
  transform: FitTransform,
  icon: RasterizedIcon | undefined,
  px: number,
  clampTo?: DiagramFrame,
  /**
   * The whole drawing shown small ahead of the readable slices of it. Names
   * are carried by those slices, so the thumbnail shows the shapes.
   */
  thumbnail = false,
  /**
   * The strings already drawn on THIS slide.
   *
   * A shortened name is a lookup key into the index slide, and the reader
   * cannot see the index and the drawing at the same time, so the drawing has
   * to be self-consistent on its own. Eight services sharing a long prefix on
   * narrow tiles all cut to the same stub, and four of them cut all the way to
   * a bare ellipsis, which tells the reader nothing and leaves four index rows
   * unmatchable. Passing what has already been drawn lets a colliding tile fall
   * back to a numeric key: unique, narrower than most letters, and resolved on
   * the index slide exactly as any other stub is.
   */
  drawnHere?: Map<string, string>,
  /** Stable, deck-global ordinals so a key means the same thing on every slide. */
  keyOrdinal?: Map<string, number>,
): {
  /** The box the service NAME is drawn in, not the room left over for it. */
  caption: { x: number; y: number; w: number; h: number } | null;
  /** The box the SKU / region / price sub-line is drawn in, when it is shown. */
  meta: { x: number; y: number; w: number; h: number } | null;
  /** The authored name, when the tile could not hold all of it. */
  clipped: string | null;
  /** What the tile actually drew, which is the key the index is looked up by. */
  drawn: string;
} {
  const topLeft = placeBox(box, transform, clampTo);
  const w = topLeft.w;
  const h = topLeft.h;
  const palette = styleForBox(box);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: topLeft.x,
    y: topLeft.y,
    w,
    h,
    rectRadius: Math.min(0.08, h / 4),
    fill: { color: palette.bg },
    line: { color: palette.border, width: 1.25 },
    shadow: {
      type: 'outer',
      color: '94A3B8',
      blur: 4,
      offset: 1,
      angle: 90,
      opacity: 0.35,
    },
    objectName: `service-${box.id}`,
  });

  const pad = Math.min(0.06, h * 0.09);
  // Every typographic dimension is proportional to the drawing scale, so the
  // number of wrapped lines is identical whatever size the diagram is drawn at.
  //
  // The floor differs by what the slide is for. A thumbnail may go below the
  // legibility floor because the `named` test below then takes the name off it
  // entirely and the slice that follows carries it in full. A window slide has
  // no later slice — it *is* the readable view — so a tile too short for
  // legible type must still get legible type. The planner's whole contract is
  // that the grid it picks clears `LEGIBLE_TILE_PT` for the shortest service on
  // the sheet, but when no grid can (one 20px node among eighty-one ordinary
  // ones is enough, because the shortest tile sets the target for all of them)
  // it returns the best grid it found and this clamp quietly drew that tile's
  // name at four points. Two floors for one contract, and only the planner's
  // was ever checked.
  const heightFontSize = clamp(h * 12, thumbnail ? 4 : LEGIBLE_TILE_PT, 13);
  // At 72 services the overview clamps this to 4pt, which is not small type —
  // it is grey ink the reader cannot resolve, and it makes the thumbnail
  // harder to read rather than more informative. The overview exists to show
  // the shape of the architecture before the reader pans through the readable
  // slices of it, and every one of those names appears in full on the slice
  // that follows, so below the resolvable floor the thumbnail draws the icon
  // and the tile and leaves the naming to them.
  // Width, as well as height, and on every slide rather than only the overview.
  // The height test asks whether the type would be resolvable; this asks
  // whether the tile has a column to set it in. A tile 0.08in wide draws "…",
  // which names nothing — the same empty claim a zone caption cut to "P…"
  // makes, and it is refused one function below for exactly that reason.
  //
  // Four characters is the bar the gate uses (`charsPerLine < 4`), so the
  // drawing and the thing that checks it cannot disagree about when a name has
  // stopped being a name. Below it the tile is icon-and-box, and `clipped`
  // sends the full name to the index slide, which is where a cut tile name has
  // always been recoverable.
  //
  // Measured against the SMALLEST font the tile is willing to draw, not the
  // one its height implies. `fontSize` comes from the height, so testing the
  // width against it made a taller tile demand a wider column: past
  // h = 1.0833in the bar saturates at 4 x 13/72 = 0.7222in and every tile
  // narrower than 0.7822in lost its name however tall it was. A 0.7813 x
  // 3.1250in tile — 2.44 square inches — missed by 0.0009in and drew no text
  // at all, while at 7pt the same column sets 7.9 capitals per line with 33
  // lines of room. That is the "this size doesn't fit" mistake rather than
  // "no legible size fits"; the name shrink loop below is what comes down to
  // meet it, exactly as the Visio exporter already does.
  const nameFloorPt = thumbnail ? OVERVIEW_LEGIBLE_PT : LEGIBLE_TILE_PT;
  const nameColumn = Math.max(0.05, w - 0.06);
  // Two of the widest glyph the name actually contains, which is the same bar
  // the stub below and the connector chips already use.
  //
  // "Four characters" charged a flat 1 em to every glyph, so it was strictly
  // harsher than any real string except solid CJK, and it disagreed with the
  // stub bar sitting a few lines away: at a 0.377in width an ICON-LESS tile
  // drew its name in full while an ICONED one of the same size drew nothing —
  // even though the iconed tile passes the two-glyph test with room for 3.8
  // characters a line. The cross-format rule caught it as a divergence from
  // Visio, which named both; two tiles of 0.57 square inches were losing their
  // names to an arithmetic accident rather than to a legibility limit.
  const namedWidth = drawableInColumn(box.label, nameFloorPt, nameColumn);
  // And once the tile is allowed to draw its name, the name is set at a size
  // the COLUMN can hold, not only one the height implies. A 0.78in-wide tile
  // is 13pt tall enough and 2.9 characters wide, so the height-derived size
  // set "Azure Firewall Premium" three characters to a line down a shape that
  // had room for eight at the floor. This is the same shrink Visio does before
  // it decides whether to draw at all.
  const fontSize = Math.max(
    nameFloorPt,
    Math.min(heightFontSize, Math.floor((nameColumn / 4) * 72 * 10) / 10),
  );
  const named = (!thumbnail || h * 12 >= OVERVIEW_LEGIBLE_PT) && namedWidth;
  // Giving up the name only works when the icon is left to carry the tile. A
  // service with no icon would otherwise be drawn as an empty grey box, which
  // says strictly less than type that is merely small. So the name comes back
  // at exactly the floor, cut to what the tile can hold: a short legible word
  // beats both an empty box and a paragraph of grey mush.
  const stub = !named && !icon;
  // Reassigned below when the name has to give type size back to the icon.
  // The floor is the slide's own, not the overview's: handing a reading slide
  // OVERVIEW_LEGIBLE_PT drew stub names at 6pt, under the 7pt floor the line
  // above enforces and under the gate's own `minFont < 7` rule.
  let drawnFont = named ? fontSize : nameFloorPt;
  const meta = metaSubline(box);
  const metaFontSize = clamp(fontSize - 2, 3.5, 9);
  // Sized from the sub-line's own font, not the name's. Deriving the band from
  // `fontSize` reserved 0.232in for a line needing 0.117in on every tile in the
  // corpus, and on a tight deck that phantom 0.05-0.09in was the whole reason
  // the icon did not fit and was dropped.
  let metaBand = named && showsMeta(h, px) && !!meta ? metaFontSize * 1.55 / 72 + 0.03 : 0;

  // Floored above one ellipsis at the 7pt type floor (0.0525in), not at an
  // arbitrary 0.05in: below that the fitter's own last resort does not fit the
  // column it is being fitted to, which is a contradiction the shrink loop
  // used to spin on.
  const innerW = Math.max(0.06, w - 0.06);
  // How much of the name the tile can actually hold, rather than a flat 40
  // cells. The flat cap clipped names a three-line tile had ample room for —
  // "Azure Database for PostgreSQL フレキシ…" on a tile that fits the whole
  // thing — and what it cut was not written down anywhere, so the reader had
  // no way to recover it. Cut to the tile, and only when the tile is really
  // too small.
  //
  // Fitted to the *lines* the tile has, not to `innerW * nameLines` of total
  // ink. Word wrap abandons the tail of a line whenever the next word will not
  // fit, so a name whose ink fits three lines routinely draws four or five:
  // "Azure Kubernetes Service Automatic cluster" was admitted whole at 13pt on
  // a 160x110 tile, drew 5 lines of a 3-line box, and painted 0.224in — all of
  // it — straight through the "P1v3 · eastus" sub-line below.
  const nameLines = Math.max(1, Math.floor((h - pad * 2 - metaBand) / ((fontSize * 1.35) / 72)));
  const linesIn = (text: string, columnIn: number, sizeIn: number): number => wrappedLineCount(text, columnIn, sizeIn * 72);
  const full = fitLabelToLines(box.label, innerW, fontSize / 72, nameLines, linesIn);
  // A stub gets as much of the name as fits the tile at the floor size, on as
  // many lines as the tile is tall enough for, and an ellipsis for the rest.
  //
  // Measured at the size it is PAINTED at. `drawnFont` moved to the reading
  // slide's 7pt floor last round and these two measurements did not, so the
  // string was fitted at 6pt and drawn at 7 — 16.7% under-measured in width,
  // in line height and in the line count. A 0.177 x 0.965in tile kept 14
  // characters where 6pt fitting said 8 lines would fit and only 7 do, and
  // painted 0.104in — one whole line — below its own box and 0.044in past the
  // bottom of the tile.
  const stubLines = stub
    ? Math.max(1, Math.floor((h - pad * 2) / ((drawnFont * 1.35) / 72)))
    : 0;
  const labelLinesFor = (text: string): number => (named ? wrappedLineCount(text, innerW, fontSize) : 0);
  let label = stub ? fitLabelToLines(full, innerW, drawnFont / 72, stubLines, linesIn) : full;
  // And a stub whose column cannot hold two of its own widest glyph draws
  // nothing at all. `namedWidth` has just decided this column is too narrow
  // for a name; drawing one anyway with no column test at all was the same
  // "a chip narrower than one letter is not a small chip, it is a broken one"
  // artefact, one function away. At 0.080in and 0.060in the widest glyph in
  // the drawn string is the ellipsis — 1.0 em, WIDER than the whole column —
  // so nothing can set on one line and PowerPoint centres each glyph and
  // overflows both sides. The premise that an icon-less tile says less than
  // small type fails once no glyph fits at all, and the whole name is already
  // on the index slide by the route `clipped` takes.
  if (stub && innerW < 2 * widestGlyphIn(label, drawnFont)) label = '';
  // Counted by wrapping, not by dividing total ink by the column. The ratio is
  // the break-anywhere assumption: it says how many lines the ink would need if
  // words could be split at any character, which is a lower bound and never the
  // answer. `labelBlockH` feeds the icon size and the top-aligned text box, so
  // under-counting here is what let the surplus lines out of the box.
  let labelLines = labelLinesFor(label);
  let labelBlockH = (labelLines * fontSize * 1.35) / 72;

  const iconFloor = 0.08 * px;
  // And if the name is what is crowding the icon out, the name yields — not the
  // icon. The same reasoning as the sub-line above, one step further: on the
  // Architecture Center the icon is what says which service a tile is, and it
  // is the one thing on the tile that cannot be recovered anywhere else. A tile
  // that loses its icon is a grey box of type; a name set two points smaller is
  // still the whole name.
  //
  // Type size first, and characters only if that is not enough. Dropping a line
  // outright looks equivalent — both free the same room — but it is not: on an
  // inventory of names that differ only in their last token it cut six
  // distinct services down to one drawn string, which is a worse failure than
  // either a small name or a missing icon. Shrinking keeps every character.
  // Counting the lines honestly is what made this reachable at all: the old
  // ratio under-counted the block, so the icon kept a share of the tile the
  // words were already using.
  const squeeze = (band: number): {
    font: number; label: string; lines: number; blockH: number; band: number; room: number;
  } => {
    let font = fontSize;
    const asks = Math.max(1, Math.floor((h - pad * 2 - band) / ((fontSize * 1.35) / 72)));
    let text = fitLabelToLines(box.label, innerW, font / 72, asks, linesIn);
    let count = Math.max(1, wrappedLineCount(text, innerW, font));
    let blockH = (count * font * 1.35) / 72;
    const room = (): number =>
      Math.min(h * 0.42, w * 0.34, Math.max(0, h - pad * 2 - band - blockH - 0.02));
    while (font > LEGIBLE_TILE_PT && room() < iconFloor) {
      font = Math.max(LEGIBLE_TILE_PT, font - 0.5);
      text = fitLabelToLines(box.label, innerW, font / 72, asks, linesIn);
      count = Math.max(1, wrappedLineCount(text, innerW, font));
      blockH = (count * font * 1.35) / 72;
    }
    // `count` is both the loop's control variable AND re-derived from a
    // measurement inside the body, so the body can put it back UP — and a
    // `while` on a value the body re-measures cannot be proved to terminate.
    // It did not: at the type floor `fitLabelToLines` returns "…", which is
    // 0.0733in at 7pt and does not fit `innerW`'s own 0.06in floor, so a
    // request for ONE line measures as two, and 2 → 1 → 2 → 1 forever. The
    // exit condition cannot save it either, because the room is capped by
    // `w * 0.34`, which does not depend on the font.
    //
    // This hangs the tab synchronously for any tile narrower than 0.133in at
    // 7pt: no error, no watchdog, no way to close the page. File → Load
    // reaches it, because the restore validator never checks `width`.
    //
    // So drive the loop by the REQUEST, which only ever falls, and stop the
    // moment asking for fewer lines stops producing fewer.
    let asked = count;
    while (asked > 1 && room() < iconFloor) {
      asked -= 1;
      const shorter = fitLabelToLines(box.label, innerW, font / 72, asked, linesIn);
      const measured = Math.max(1, wrappedLineCount(shorter, innerW, font));
      if (measured >= count) break;
      text = shorter;
      count = measured;
      blockH = (count * font * 1.35) / 72;
    }
    return { font, label: text, lines: count, blockH, band, room: room() };
  };
  let nameFont = fontSize;
  if (icon && named && !stub) {
    // Which of the three things a tile carries yields when it cannot hold all
    // three. The sub-line used to be dropped FIRST, at full-size type, on the
    // reasoning that the icon outranks it — which is true, and was the wrong
    // conclusion, because it never asked whether the icon needed it. Charging
    // a space the width it draws made names one line taller across the corpus
    // and the rule then deleted the SKU and the region from every tile in a
    // scenario that exists to carry them, while the name sat at full size.
    //
    // A name set half a point smaller is still the whole name; a deleted
    // sub-line is information the tile cannot get back and that nothing else
    // in the deck carries. So the free move is tried first, and the sub-line
    // is dropped only when shrinking the name does not save the icon AND
    // dropping it does.
    const withMeta = squeeze(metaBand);
    let chosen = withMeta;
    if (metaBand > 0 && withMeta.room < iconFloor) {
      const withoutMeta = squeeze(0);
      if (withoutMeta.room >= iconFloor) chosen = withoutMeta;
    }
    metaBand = chosen.band;
    nameFont = chosen.font;
    label = chosen.label;
    labelLines = chosen.lines;
    labelBlockH = chosen.blockH;
    drawnFont = nameFont;
  }

  // A KEY THAT REPEATS IS NOT A KEY, and this is the LAST word on what the tile
  // draws - the squeeze above re-picks the label to save the icon, so checking
  // any earlier reads a string that is then thrown away. Everything above
  // decides what THIS tile can hold; none of it can see what the tile beside it
  // drew. Once a name is shortened it stops being a name and becomes a lookup
  // key into the index slide, and two tiles holding the same key - or holding
  // none at all - are indistinguishable to a reader who cannot see the index
  // and the drawing at the same time. Eight services sharing a long prefix cut
  // to a bare ellipsis and four of them landed on one slide. Lengthening is not
  // always available: a column with room for 1.79 characters has nothing to
  // lengthen into. A number is, and it is narrower than most letters.
  // Compared against the AUTHORED name, not against `full` - `full` is itself
  // already the fitted string, so the two are equal on exactly the tiles that
  // cut the hardest and the test excluded the only case it was written for.
  const authoredLabel = String(box.label ?? '').trim();
  // The legibility test has to be applied to the string that is ACTUALLY drawn.
  // `namedWidth` asks whether the tile can set `box.label`, whose widest glyph
  // is an ordinary letter; what lands on a hard-cut tile is "…", whose glyph is
  // 1.0 em - 0.071in against 0.052in - so a column judged wide enough for the
  // name was 0.045in too narrow for the mark it ended up carrying, and four
  // tiles drew an ellipsis the gate then reported as illegible. The stub branch
  // has always been guarded this way; this is the same guard applied to the
  // final string rather than to one of the two paths that produce it, and it is
  // placed after the icon-saving squeeze because that squeeze re-picks the
  // label - anything earlier tests a string that is then thrown away.
  if (label && innerW < 2 * widestGlyphIn(label, drawnFont)) label = '';
  // A bare ellipsis is not a short name, it is an absent one: it carries no
  // character of the service it stands for, so it is exactly as useful as a
  // blank tile and no more distinguishable. Anything with no letter or digit in
  // it fails as a lookup key for the same reason a repeated one does.
  const informative = /[\p{L}\p{N}]/u.test(label);
  const claimed = drawnHere?.get(label);
  // Whether the KEY is the only thing this tile will carry. An iconed tile
  // below the naming floor draws no text at all - the icon and the box are
  // meant to carry it - which is sound while the icon says something the tile
  // beside it does not. Eight instances of one service defeat that: eight
  // identical icons, eight names too long for the column, and nothing on the
  // drawing to tell any of them from any other. A single digit is not a name
  // and is not asked to be one; it is the smallest mark that identifies, and
  // the index defines it.
  let keyed = false;
  if (drawnHere && label !== authoredLabel
    && (!informative || (claimed !== undefined && claimed !== authoredLabel))) {
    const key = `${keyOrdinal?.get(String(box.id)) ?? drawnHere.size + 1}`;
    // TWO of the key's widest glyph, the same bar the gate's legibility rule
    // uses. Relaxing it to one - on the argument that a single character cannot
    // stack down the side of a tile - drew digits the gate then reported as
    // illegible in a 0.065in box, so the renderer and the thing that checks it
    // disagreed about the same quantity. The narrow tile is fixed where it is
    // caused instead: `MARKABLE_TILE_W_IN` lifts the transform cap so the tile
    // arrives wide enough to carry the key in the first place.
    const room = 2 * widestGlyphIn(key, drawnFont);
    // And the room to SET it: one line at the floor. Without this the key is
    // drawn in a box taller than the tile and overhangs its neighbours.
    const lineH = (drawnFont * 1.35) / 72;
    if (innerW >= room && h >= lineH) {
      label = key;
      keyed = !named && !stub;
      labelLines = Math.max(1, labelLinesFor(label));
      labelBlockH = (labelLines * drawnFont * 1.35) / 72;
    }
  }
  if (drawnHere && label) drawnHere.set(label, authoredLabel);

  // Fit the icon into whatever vertical room the label does not need, instead
  // of forcing a minimum that pushes the text out of the tile.
  const available = h - pad * 2 - metaBand;
  let iconSize = 0;
  if (icon) {
    iconSize = clamp(Math.min(h * 0.42, w * 0.34, Math.max(0, available - labelBlockH - 0.02)), 0, 0.6);
    if (iconSize < 0.08 * px) iconSize = 0; // too small to read — drop it and keep the words
  }
  // With no name to make room for, the icon is the only thing carrying meaning
  // in the tile, so it takes the whole of it.
  if (!named && icon) {
    iconSize = clamp(Math.min(h * 0.78, w * 0.78, available), 0, 0.6);
  }

  if (iconSize > 0 && icon) {
    slide.addImage({
      data: icon.dataUrl,
      x: topLeft.x + (w - iconSize) / 2,
      y: named ? topLeft.y + pad : topLeft.y + (h - iconSize) / 2,
      w: iconSize,
      h: iconSize,
      objectName: `icon-${box.id}`,
    });
  }

  const textTop = iconSize > 0 ? topLeft.y + pad + iconSize + 0.02 : topLeft.y + pad;
  const textHeight = Math.max(0.08, topLeft.y + h - pad - metaBand - textTop);

  let captionBand: { x: number; y: number; w: number; h: number } | null = null;
  if ((named || stub || keyed) && label !== '') {
    const boxY = stub ? topLeft.y + pad : textTop;
    const boxH = stub ? Math.max(0.08, h - pad * 2) : textHeight;
    // The box the words are drawn in, not the room left over for them. With no
    // icon the leftover room is nearly the whole tile, and a caption box that
    // claims the whole tile tells every later pass nothing: a chip weighed
    // against it is weighed against the tile it already knew about, and a rule
    // measuring "how much of the name is covered" is really measuring the tile.
    // Vertically centred text inside a shrunk, centred box draws in exactly the
    // same place, so this describes the caption without moving it.
    //
    // The band is what the type needs, not what the tile has. Clamping it to
    // the tile was right while type was derived from the tile, because then it
    // always fit; now that a window tile's type is floored at the legibility
    // limit however short the tile is, a collapsed node's 0.08in box carries a
    // 7pt line needing 0.12in, and clamping described a line that reaches past
    // the band it was measured in. On every ordinary tile the type still fits
    // and this is exactly the old value.
    const needH = Math.max(0.08, (Math.max(1, labelLines) * drawnFont * 1.35) / 72);
    const drawnH = needH;
    // Growing the band grows the box the words are actually drawn in, kept
    // centred on the tile so a sliver's name overhangs evenly instead of
    // hanging off one edge. Identical to `boxY`/`boxH` whenever the type fits.
    const textBoxH = Math.max(boxH, needH);
    const textBoxY = boxY - (textBoxH - boxH) / 2;
    const topAligned = !stub && iconSize > 0;
    captionBand = {
      x: topLeft.x + 0.03,
      y: topAligned ? boxY : textBoxY + (textBoxH - drawnH) / 2,
      w: innerW,
      h: drawnH,
    };
    slide.addText(label, {
      x: topLeft.x + 0.03,
      y: topAligned ? boxY : textBoxY,
      w: innerW,
      h: topAligned ? boxH : textBoxH,
      fontSize: drawnFont,
      color: '1F2937',
      fontFace: 'Yu Gothic UI',
      align: 'center',
      valign: !stub && iconSize > 0 ? 'top' : 'middle',
      margin: 0,
      lineSpacingMultiple: 0.9,
      wrap: true,
      objectName: `service-label-${box.id}`,
    });
  }
  let metaBandRect: { x: number; y: number; w: number; h: number } | null = null;
  if (metaBand > 0 && meta) {
    // Fitted to the tile, not to a flat 44 characters. This line is
    // `wrap: false`, so a string the tile cannot hold does not wrap or clip —
    // it is drawn centred at its full natural width and spills out of both
    // sides of the tile, over whatever the neighbours put there. A count of
    // characters cannot know that: "P1v3 · japaneast · $9.60/mo" and
    // "Standard_D4s_v5 · japaneast · $128.40/mo" are both under 44 and only
    // one of them fits.
    //
    // Shrink before cutting. Every character of a SKU, a region and a price is
    // load-bearing and none of it is recoverable from an ellipsis, so a point
    // of type is a far better trade than the end of the string.
    let metaPt = metaFontSize;
    while (metaPt > META_LEGIBLE_PT && estimateTextWidthIn(meta, metaPt) > innerW) {
      metaPt = Math.max(META_LEGIBLE_PT, metaPt - 0.5);
    }
    // Still too wide at the smallest legible size. Drop whole facts from the
    // least essential end rather than cutting mid-token: "Standard_D4s_v…" is
    // a SKU nobody can look up, while "Standard_D4s_v5 · japaneast" is two
    // true statements and the price it dropped is on the cost slides. A tile
    // too small for even one whole fact carries none — an ellipsis there would
    // only claim to say something it does not.
    let shown = meta;
    if (estimateTextWidthIn(shown, metaPt) > innerW) {
      const facts = meta.split(' · ');
      shown = '';
      while (facts.length > 1) {
        facts.pop();
        const candidate = facts.join(' · ');
        if (estimateTextWidthIn(candidate, metaPt) <= innerW) { shown = candidate; break; }
      }
      if (shown === '' && facts.length === 1 && estimateTextWidthIn(facts[0], metaPt) <= innerW) shown = facts[0];
    }
    const drawnW = Math.min(innerW, estimateTextWidthIn(shown, metaPt));
    // The height of the glyphs, not of the band. The line is bottom-aligned in
    // a band sized for the tile, so the room above the words holds nothing —
    // and an obstacle that claims it makes every "how deep is this bite" test
    // read shallower than what is actually drawn.
    const drawnH = Math.min(metaBand, (metaPt * 1.35) / 72);
    metaBandRect = {
      x: topLeft.x + 0.03 + (innerW - drawnW) / 2,
      y: topLeft.y + h - pad - drawnH,
      w: drawnW,
      h: drawnH,
    };
    if (shown === '') metaBandRect = null;
    else slide.addText(shown, {
      x: topLeft.x + 0.03,
      y: topLeft.y + h - pad - metaBand,
      w: innerW,
      h: metaBand,
      fontSize: metaPt,
      // The tile fill is category-dependent, so a fixed grey reads at 4.26:1 on
      // the lighter categories. Derive it from the panel it is printed on.
      color: stripHash(readableTextOn('#64748B', `#${stripHash(palette.bg)}`)),
      fontFace: 'Yu Gothic UI',
      align: 'center',
      valign: 'bottom',
      margin: 0,
      wrap: false,
      objectName: `service-meta-${box.id}`,
    });
  }
  return {
    caption: captionBand,
    meta: metaBandRect,
    // A name the tile refused to set at all is as lost as one it cut, and it is
    // the index slide that gets it back either way.
    clipped: !named && box.label ? box.label : (label === box.label ? null : box.label),
    // What the tile ACTUALLY drew. Reporting the label the sizing arrived at
    // made the index define marks that appear on no shape in the file: an
    // iconed tile below the naming floor emits no text at all, and the index
    // still printed the string that tile would have drawn if it had. An index
    // row promising a mark the reader cannot find is worse than one admitting
    // the name was never drawn, which is what an empty mark prints.
    drawn: (named || stub || keyed) ? label : '',
  };
}

/** The SKU · region · cost sub-line only earns its space on a legible tile. */
function showsMeta(heightIn: number, px: number): boolean {
  return heightIn > 0.5 * px;
}


function addGroupShape(
  pptx: PptxGenJS,
  slide: Slide,
  box: ExportBox,
  index: number,
  transform: FitTransform,
  clampTo?: DiagramFrame,
  /** Members of this zone on this slide, and in the drawing as a whole. */
  held?: { here: number; all: number },
  /** Where the tiles already landed, so a title is not written on top of one. */
  occupied?: readonly { x: number; y: number; w: number; h: number }[],
  /** The other zones, which a title may never be written inside. */
  foreign?: readonly { x: number; y: number; w: number; h: number }[],
  /**
   * Where the arrow labels want to sit. Not forbidden — merely expensive.
   *
   * A caption and a chip are placed by two searches that ran without knowing
   * about each other, and both prefer the same clear paper: the corridor
   * between two rows of tiles. The caption is drawn first, so the chip loses,
   * and the chip has nowhere to fall back to when its route carries no step
   * number — the wording would simply be gone. Charging the caption for the
   * corridor lets it step aside while it still has other rows to take, and
   * lets it stand its ground when it does not.
   */
  /** Cut names, spelled out in full on the index slide. */
  truncatedNames?: Map<string, Set<string>>,
): { caption: { x: number; y: number; w: number; h: number } } {
  const topLeft = placeBox(box, transform, clampTo, true);
  const w = topLeft.w;
  const h = topLeft.h;
  // Whether the window cut this zone, which decides whether the drawn
  // rectangle is the zone or only the part of it that survived the cut.
  const uncut = placeBox(box, transform);
  const clipped = Math.abs(uncut.w - w) > 1e-6 || Math.abs(uncut.h - h) > 1e-6;
  const palette = zoneStyleFor(box, index);
  const bg = stripHash(palette.bg);
  const border = stripHash(palette.border);
  // The border colour is tuned to be seen as a 1pt line, not read as words: on
  // the light slide it lands at 2.6-4.4:1, below the WCAG AA bar. The palette
  // already carries a text colour chosen for reading; use it.
  const labelColor = stripHash(palette.text);

  // Let a long zone title wrap to two lines instead of clipping at a fixed band.
  const titleH = clamp(h * 0.16, 0.24, 0.5);
  // A closed box says "these are all of them". When a drawing is split across
  // slides the same zone is redrawn on each one, so a zone of 28 services can
  // appear as a closed box around 3 with nothing to tell the reader it is a
  // fragment. Say so in the title, the way a split reference architecture does.
  const fragment = held && held.all > held.here ? ` (${held.here} / ${held.all})` : '';
  // A zone's title band is empty because the author left it empty — and a
  // window that cuts the zone's top away takes that room with it, so the band
  // lands on whatever tiles are nearest the cut and the fragment loses its
  // name. Every candidate here keeps the title attached to the fragment; the
  // one that covers the fewest tiles wins, and the top-left band still wins
  // outright whenever it is clear, which is every drawing that is not split.
  // Never wider than the zone it names. `max(0.4, w - 0.12)` gave a 0.30in
  // zone a 0.40in band: wider than the box it belongs to, hanging over both
  // neighbours, and no placement could fix it because the geometry made
  // overlap compulsory — one scaled row of small zones produced 2385 pairs of
  // captions written across each other. A floor is worth having, but only up
  // to the zone's own width; past that the band stops naming this zone and
  // starts naming the one beside it.
  const titleW = Math.max(Math.min(0.4, w), w - 0.12);
  const titleX = topLeft.x + (w - titleW) / 2;
  // What the name will actually take, at the pitch the gate measures painted
  // ink with, so the exporter's budget and the gate's check are one number
  // rather than two guesses that happen to agree today.
  const captionPt = clamp(Math.round(h * 5), 8, 12);
  const captionText = `${box.label}${fragment}`;
  // PowerPoint applies a 0.1in inset on each side of a text box that declares
  // no margin, so `width - 0.2` is the column the words actually get. The
  // `max(0.2, …)` floor that used to sit here was a lie the fitter believed:
  // on a band narrower than 0.4in it promised 0.2in of line that does not
  // exist, so the name was fitted onto one line and PowerPoint stacked it one
  // letter per line, running out of the bottom of the band. Zero is the honest
  // answer for a band that has no column at all; the guard below stops before
  // anything is drawn into it.
  const captionColumn = (width: number): number => Math.max(0, width - 0.2);
  const captionRoom = (band: number, pt: number): number => Math.max(1, Math.floor(band / ((pt * 1.35) / 72)));  // A zone cut to a sliver at the frame edge has less width than its own title
  // needs, so the band has to be pulled back onto the page — placed raw it ran
  // off the slide and PowerPoint dropped it, taking the zone's name with it.
  const fit = <T extends { x: number; y: number; w: number }>(c: T): T => (clampTo
    ? {
      ...c,
      x: clamp(c.x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - c.w)),
      y: clamp(c.y, clampTo.y, Math.max(clampTo.y, clampTo.y + clampTo.h - titleH)),
    }
    : c);
  const top = topLeft.y + 0.04;
  const foot = topLeft.y + h - titleH - 0.04;
  const part = (share: number): number => Math.min(titleW, Math.max(0.4, w * share - 0.12));
  // The bands the tiles actually left free, rather than the fractions of the
  // width somebody guessed at. Fixed shares only work while the row has a
  // quarter of itself spare: fill a subnet — which is what a subnet drawn to
  // scale looks like — and every band on offer, full width, half or third,
  // lands on the same tiles, so a rule that fails a title at 25% coverage had
  // no legal placement left to choose. Reading the gaps finds the one the
  // author left, wherever it happens to be.
  const runs = (y: number): Array<{ x: number; y: number; w: number }> => {
    const lo = Math.min(titleX, topLeft.x + 0.06);
    const hi = topLeft.x + w - 0.06;
    const blockers = (occupied ?? [])
      .filter((t) => t.y < y + titleH && t.y + t.h > y && t.x + t.w > lo && t.x < hi)
      .map((t) => [Math.max(lo, t.x), Math.min(hi, t.x + t.w)] as [number, number])
      .sort((a, b) => a[0] - b[0]);
    const out: Array<{ x: number; y: number; w: number }> = [];
    let cursor = lo;
    for (const [from, to] of blockers) {
      if (from - cursor >= 0.35) out.push({ x: cursor, y, w: Math.min(titleW, from - cursor) });
      cursor = Math.max(cursor, to);
    }
    if (hi - cursor >= 0.35) out.push({ x: cursor, y, w: Math.min(titleW, hi - cursor) });
    return out.sort((a, b) => b.w - a.w);
  };
  // Inside the box it names, always — a name is a claim about what the box
  // contains, and a name printed anywhere else is a different claim. When the
  // full-width band across the top is standing on the zone's own tiles, the
  // answer is a narrower band in the part of the zone the tiles left free, not
  // a clear band belonging to somebody else: a subnet stack drawn tight has no
  // room above any box except the box above it, and "Data subnet" printed
  // there says the data tier is part of the application tier.
  //
  // Rows, plural. Every candidate used to sit at either `top` or `foot`, so a
  // zone whose tiles are stacked down its length was offered nothing but the
  // two rows its tiles are thickest in — and a tall narrow zone with clear
  // paper between every pair of tiles took a 0.40in band standing on a tile
  // while 0.86in of clear width sat one row below. The gaps between a zone's
  // own tiles are exactly the rows it has to give, so offer them.
  //
  // They are offered free. The corridor between two rows of tiles is also
  // where the arrows run and where connector chips are seated, so it is
  // tempting to charge for it — but a surcharge was measured and it costs more
  // than it saves: it moved no chip off anything in the corpus, and it took a
  // flush-to-the-top zone's caption from a 9.26in band down to a 1.50in one.
  // A caption that gives up characters so a chip can keep its first choice has
  // traded the zone's name for one hop's verb, which is the wrong way round.
  // Chips are kept off captions on the chip's side instead, by putting the
  // chosen band into `chipObstacles`.
  const gapRows: number[] = [];
  {
    const mine = (occupied ?? [])
      .filter((t) => t.x < topLeft.x + w && topLeft.x < t.x + t.w
        && t.y < topLeft.y + h && topLeft.y < t.y + t.h)
      .sort((a, b) => a.y - b.y);
    let cursor = top;
    for (const tile of mine) {
      if (tile.y - cursor >= titleH && gapRows.length < 4) gapRows.push(cursor);
      cursor = Math.max(cursor, tile.y + tile.h + 0.02);
    }
    if (cursor + titleH <= foot && gapRows.length < 4) gapRows.push(cursor);
  }
  const inside = [
    { x: titleX, y: top, w: titleW },
    { x: titleX, y: foot, w: titleW },
    ...gapRows.map((y) => ({ x: titleX, y, w: titleW })),
    ...runs(top),
    ...runs(foot),
    ...gapRows.flatMap((y) => runs(y)),
    { x: topLeft.x + Math.max(0, w - part(0.5) - 0.06), y: top, w: part(0.5) },
    { x: titleX, y: top, w: part(0.5) },
    { x: topLeft.x + Math.max(0, w - part(0.34) - 0.06), y: top, w: part(0.34) },
    { x: titleX, y: top, w: part(0.34) },
    { x: topLeft.x + Math.max(0, w - part(0.34) - 0.06), y: foot, w: part(0.34) },
    { x: titleX, y: foot, w: part(0.34) },
  ];
  // A fragment is the one exception. Its drawn rectangle is not the zone — it
  // is whatever survived the window cut — so there may be no room inside it at
  // all, and the band just outside the cut is still inside the zone the reader
  // is being shown.
  const outside = clipped
    ? [
      { x: titleX, y: topLeft.y - titleH - 0.02, w: titleW },
      { x: titleX, y: topLeft.y + h + 0.02, w: titleW },
    ].filter((c) => !clampTo || (c.y >= clampTo.y && c.y + titleH <= clampTo.y + clampTo.h))
    : [];
  const candidates = [...inside, ...outside].map(fit);
  const cover = (c: { x: number; y: number; w: number }): number => (occupied ?? []).reduce((sum, tile) => {
    const ox = Math.max(0, Math.min(c.x + c.w, tile.x + tile.w) - Math.max(c.x, tile.x));
    const oy = Math.max(0, Math.min(c.y + titleH, tile.y + tile.h) - Math.max(c.y, tile.y));
    return sum + ox * oy;
  }, 0);
  // Whatever the band gains in clear space it must not buy from a neighbour.
  const trespass = (c: { x: number; y: number; w: number }): number => (foreign ?? []).reduce((sum, zone) => {
    const ox = Math.max(0, Math.min(c.x + c.w, zone.x + zone.w) - Math.max(c.x, zone.x));
    const oy = Math.max(0, Math.min(c.y + titleH, zone.y + zone.h) - Math.max(c.y, zone.y));
    return Math.max(sum, ox * oy);
  }, 0);
  // Scored as a fraction of the band, not as absolute area: a narrower band
  // covers less simply by being narrower, so absolute area would always prefer
  // the smallest one on offer even when the widest is completely clear.
  //
  // Trespass is weighted above coverage because the two failures are not
  // comparable. A name lying over its own icon is crowded; a name lying inside
  // a different zone's box asserts a containment the architecture does not
  // have. Both are scored rather than filtered so that when every band
  // trespasses — which is what two overlapping zones give you — the least bad
  // one still wins instead of the first one tried.
  // A fit term was tried here and removed. The premise as first stated — that a
  // clear 0.4in band can beat a 3.7in one on score — does not survive reading
  // the candidate order: the full-width band is `inside[0]`, it is scored
  // first, and the loop only replaces on a strictly better score, so a clear
  // full-width band always wins on score. Scoring fit changed the chosen band
  // in no fixture in the corpus, and at any weight large enough to change one
  // it bought a 48% covered band in `flush-subnets`.
  //
  // The defect was in the *tie-break*, not the score. Two candidates that are
  // both completely clear both score 0, and `if (best <= 0.01) break` stopped
  // at whichever came first in the array — so a zone with its tiles flush to
  // the top handed its caption to the 0.43in gap above them and never looked at
  // the 9.35in of clear paper below, drawing "Producti…dant" for want of a
  // comparison it had already decided not to make. Score all of them, and when
  // the score cannot separate two bands, take the wider one: it holds more of
  // the name for exactly the same cost.
  const score = (c: { x: number; y: number; w: number }): number => {
    const area = Math.max(1e-6, c.w * titleH);
    return cover(c) / area + 2 * (trespass(c) / area);
  };
  let title = candidates[0];
  let best = score(title);
  for (const candidate of candidates.slice(1)) {
    const next = score(candidate);
    if (next < best - 1e-6 || (Math.abs(next - best) <= 1e-6 && candidate.w > title.w + 1e-6)) {
      best = next;
      title = candidate;
    }
  }
  // A full box has no gap to give, and every band inside it is standing on a
  // tile. The answer a container diagram has always used is a header strip:
  // the name lives in a band that is part of the box, above the things the box
  // holds. The author drew the rectangle around their tiles, so the strip is
  // taken from just outside it and the drawn rectangle grows to include it —
  // the title is then inside the boundary it names, covering nothing, which is
  // what both rules ask for and what neither could otherwise be given.
  let rectY = topLeft.y;
  let rectH = h;
  let bandH = titleH;
  if (best > 0.2) {
    const lo = Math.min(titleX, topLeft.x + 0.06);
    const hi = topLeft.x + w - 0.06;
    // How far the box can grow before it meets something, rather than whether
    // a fixed strip happens to be clear. A subnet stack is drawn with its tiers
    // close together — the gap between two of them is the gap, not a
    // negotiation — and asking for a comfortable strip meant the middle tier of
    // three got none at all while its neighbours, which had the page edge to
    // grow into, got theirs. The band only has to hold one line.
    const blockers = [...(occupied ?? []), ...(foreign ?? [])].filter((r) => r.x < hi && r.x + r.w > lo);
    const ceiling = clampTo ? clampTo.y : topLeft.y - titleH - 0.1;
    const floorY = clampTo ? clampTo.y + clampTo.h : topLeft.y + h + titleH + 0.1;
    const roomAbove = topLeft.y - Math.max(
      ceiling,
      ...blockers.filter((r) => r.y + r.h <= topLeft.y + 1e-6).map((r) => r.y + r.h),
    );
    const roomBelow = Math.min(
      floorY,
      ...blockers.filter((r) => r.y >= topLeft.y + h - 1e-6).map((r) => r.y),
    ) - (topLeft.y + h);
    const fits = (available: number): number => Math.min(titleH, available - 0.015);
    const above = fits(roomAbove);
    const below = fits(roomBelow);
    // One line of the smallest type this title is ever set in, plus its margin.
    const MIN_STRIP = 0.16;
    if (above >= MIN_STRIP && above >= below) {
      bandH = above;
      rectY = topLeft.y - above;
      rectH = h + above;
      title = { x: lo, y: rectY + 0.005, w: titleW };
    } else if (below >= MIN_STRIP) {
      bandH = below;
      rectH = h + below;
      title = { x: lo, y: topLeft.y + h + 0.005, w: titleW };
    }
  }

  slide.addShape(pptx.ShapeType.roundRect, {
    x: topLeft.x,
    y: rectY,
    w,
    h: rectH,
    rectRadius: 0.06,
    fill: { color: bg, transparency: 15 },
    line: { color: border, width: 1, dashType: 'dash' },
    objectName: `zone-${box.id}`,
  });
  // Shrink first, cut last — and cut by measurement, not by counting cells. A
  // 60-character cap says nothing about how wide those characters are: the
  // band that failed here was 0.4in holding a name the cap had already
  // "shortened" to 60 characters, which is 22 wrapped lines of it.
  let captionSize = captionPt;
  let caption = captionText;
  const column = captionColumn(title.w);
  for (let pt = captionPt; pt >= LEGIBLE_TILE_PT; pt -= 1) {
    captionSize = pt;
    if (wrappedLineCount(captionText, column, pt) <= captionRoom(bandH, pt)) break;
    if (pt - 1 < LEGIBLE_TILE_PT) {
      const room = captionRoom(bandH, pt);
      // The fragment marker is never cut. It is the only thing telling the
      // reader this closed box is a slice of a larger zone, and a box that
      // silently drops "(3 / 28)" is not abbreviated — it is claiming to hold
      // everything it names. So the name yields to the marker, down to nothing
      // if that is what the band affords.
      caption = fitLabelToLines(
        box.label,
        column,
        pt / 72,
        room,
        (text, col, sizeIn) => wrappedLineCount(`${text}${fragment}`, col, sizeIn * 72),
      ) + fragment;
      // A zone cut to a sliver by the window gets a caption cut with it, and a
      // box labelled "P…" names nothing. A cut service name has always been
      // spelled out on the index slide; a cut zone name is no different, and
      // is the only place the reader can now recover it. The CUT text is the
      // mark the index row is keyed by, exactly as a tile's stub is.
      recordMark(truncatedNames, box.label, caption);
    }
  }

  // A column narrower than a single character cannot paint a name inside the
  // band. PowerPoint does not clip a text box, so what it does instead is stack
  // the word one letter per line and print the tail below the band and across
  // the zones on either side — on a row of subnets scaled small, every caption
  // was written over its neighbours' and the reader could not tell which box
  // any of them named. Silence is the honest answer here, and it is not a loss:
  // an undrawable zone name goes to the index slide by exactly the route a
  // window-clipped one already takes.
  //
  // Measured against the COLUMN, not the box. Against the box this fired only
  // below about 0.043in — the one geometry the corpus happened to sample — and
  // was silent across the whole ordinary range of small zones between.
  const widest = widestGlyphIn(caption, captionSize);
  // Two of them for a caption of more than one character, for the reason the
  // chip guard has: a column that holds one glyph draws the zone's name down
  // the side of it, one letter per line, which no reader follows back to the
  // box it names.
  const needs = caption.length > 1 ? widest * 2 : widest;
  if (caption && needs > captionColumn(title.w) + 0.01) {
    // A zone whose caption cannot be drawn at all has NO mark to look up by,
    // so its row keys on the empty string and the index prints "(not drawn)".
    recordMark(truncatedNames, box.label, '');
    return { caption: { x: title.x, y: title.y, w: 0, h: 0 } };
  }
  slide.addText(caption, {
    x: title.x,
    y: title.y,
    w: title.w,
    h: bandH,
    fontSize: captionSize,
    bold: true,
    color: labelColor,
    fontFace: 'Yu Gothic UI',
    align: 'left',
    valign: 'top',
    wrap: true,
    lineSpacingMultiple: 0.9,
    objectName: `zone-label-${box.id}`,
  });
  // Hand the band back so the chips can keep off it. A zone caption is drawn
  // first and a chip is drawn last at 92% opacity, so whatever the chip lands
  // on is simply gone — and until the gaps between a zone's own tiles became
  // caption rows, the caption always sat in the zone's margins, where chips do
  // not go, so the collision was structurally rare. It is now the default
  // shape of the commonest drawing in the corpus: two tiers of tiles with
  // edges running between them, in the very corridor both want.
  return { caption: { x: title.x, y: title.y, w: title.w, h: bandH } };
}

/** Where the colour key lands, so the drawing can keep out from under it. */
function connectionLegendRect(
  edges: Edge[],
  frame: DiagramFrame,
): { x: number; y: number; w: number; h: number } | null {
  const entries = usedConnectionLegend(edges);
  if (entries.length === 0) return null;
  // One row of swatches along the bottom rather than a stacked card in a
  // corner. The card was 92% opaque and drawn last, so on a full grid it simply
  // deleted whatever tile it landed on, and no corner of a full grid is free;
  // a strip costs a third of the height to reserve and collides with nothing.
  const w = Math.min(frame.w - 0.1, entries.length * 1.55);
  const h = 0.24;
  return { x: frame.x + 0.05, y: frame.y + frame.h + 0.03, w, h };
}

/** Small colour key so the deck agrees with the PNG's connection legend. */
function addConnectionLegend(
  pptx: PptxGenJS,
  slide: Slide,
  edges: Edge[],
  frame: DiagramFrame,
): void {
  const entries = usedConnectionLegend(edges);
  const seat = connectionLegendRect(edges, frame);
  if (entries.length === 0 || !seat) return;

  const swatchW = 0.3;
  const cellW = seat.w / entries.length;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: seat.x, y: seat.y, w: seat.w, h: seat.h,
    rectRadius: 0.04,
    fill: { color: 'FFFFFF', transparency: 8 },
    line: { color: 'CBD5E1', width: 0.5 },
    objectName: 'connection-legend',
  });
  entries.forEach((entry, i) => {
    const cx = seat.x + i * cellW + 0.08;
    slide.addShape(pptx.ShapeType.line, {
      x: cx, y: seat.y + seat.h / 2, w: swatchW, h: 0,
      line: {
        color: stripHash(entry.color),
        width: 1.5,
        dashType: entry.dashed ? (entry.type === 'telemetry' ? 'dashDot' : entry.type === 'async' ? 'dash' : 'sysDot') : 'solid',
      },
    });
    slide.addText(entry.label, {
      x: cx + swatchW + 0.06, y: seat.y, w: Math.max(cellW - swatchW - 0.2, 0.3), h: seat.h,
      fontSize: 8,
      color: '475569',
      fontFace: 'Yu Gothic UI',
      valign: 'middle',
    });
  });
}

/**
 * Render the diagram onto `slide` using native PowerPoint shapes.
 * Returns false when there is nothing to draw so callers can fall back to the
 * captured PNG.
 */
async function addEditableDiagram(
  pptx: PptxGenJS,
  slide: Slide,
  diagram: DiagramShapeSource,
  fullFrame: DiagramFrame,
  _isDarkMode: boolean,
  window?: DiagramWindow,
  /**
   * Wording that a muted chip handed over, by step number. The caller writes
   * the workflow slide, and a muted label survives only if that slide says it.
   */
  mutedWording: Map<number, string> = new Map(),
  /**
   * Names the tiles had to cut, filled in as they are drawn. The caller lists
   * them on an index slide, because a name clipped on the drawing and written
   * down nowhere else has been thrown away.
   */
  truncatedNames: Map<string, Set<string>> = new Map(),
  /**
   * This is the whole drawing shown small ahead of the readable slices of it,
   * so anything that would land under the resolvable floor is left to them.
   */
  thumbnail = false,
  presetIcons?: Map<string, RasterizedIcon>,
  /**
   * Wording promoted out of a chip that had nowhere legible to stand, by the
   * step number handed to it. Distinct from `mutedWording`: that trades a chip
   * for a row the author already wrote, while this one has no row at all until
   * the caller adds it, because the edge carried a label and no step.
   */
  promotedSteps: Map<number, string> = new Map(),
  /** The mark each service draws, shared by every slide in the deck. */
  drawnHere: Map<string, string> = new Map(),
  /** Stable, deck-global ordinals for the numeric key fallback. */
  keyOrdinal: Map<string, number> = new Map(),
  /**
   * How the slide titles itself, e.g. `2 / 3`. Printed in the index row of any
   * service this slide draws without a mark, so the reader knows which sheet to
   * turn to. Empty on a one-slide deck, where there is nothing to disambiguate.
   */
  slideLabel = '',
): Promise<boolean> {
  const frame = fullFrame;

  // Size and draw from the SAME bounds. Sizing the page for the dense cluster
  // while drawing every box is what silently pushed far-placed services off
  // the slide, so when outliers are excluded from the fit they are clamped
  // back onto the page instead of being drawn into the void.
  //
  // Parked once, by the shared helper, so the whole slide pipeline — page
  // sizing, window planning, routing, drawing — agrees on where a stray ended
  // up. Doing it inside `placeBox` on each slide meant the tile, the arrow
  // aimed at it and the window that claimed it could each pick a different
  // answer.
  const { boxes, bounds, clamped } = parkedLayout(diagram.nodes ?? []);
  if (boxes.size === 0) return false;
  const { groups, services } = partitionBoxes(boxes);
  if (services.length === 0) return false;
  // A banded slide is sized from its own tile, which is what buys back the
  // legible scale; the tile is then clamped so a shape straddling the seam is
  // cut at the page edge instead of spilling into the void.
  const banded = !!window && (
    window.fit.minX > bounds.minX + 0.5 || window.fit.maxX < bounds.maxX - 0.5
    || window.fit.minY > bounds.minY + 0.5 || window.fit.maxY < bounds.maxY - 0.5
  );
  const fitBounds = window?.fit ?? bounds;
  const ownBounds = window?.own ?? bounds;
  // Ownership is decided by the exact window; what the slide draws is the
  // window plus a bleed, so a chip anchored near a seam has room to be drawn
  // where it belongs instead of being clamped on top of a tile. The bleed is
  // deliberately *not* trimmed at the drawing's edge: every window is the same
  // size, so every part of the deck then renders at exactly the same scale.
  const view = banded
    ? {
      minX: fitBounds.minX - WINDOW_BLEED_PX,
      maxX: fitBounds.maxX + WINDOW_BLEED_PX,
      minY: fitBounds.minY - WINDOW_BLEED_PX,
      maxY: fitBounds.maxY + WINDOW_BLEED_PX,
    }
    : fitBounds;
  // Routes are planned from where the tiles are, and on a clamped drawing that
  // is where `clampedBoxes` above put them — miles from the node's declared
  // position. Planning from the declared position aimed every hop touching a
  // stray off the sheet, where the clip then threw it away: the arrow vanished
  // while its chip, its numbered callout and its line in the step list all
  // stayed behind.
  const routes = buildExportRoutes(diagram.edges ?? [], boxes);
  // An arrow is drawn as surely as a tile is, and a detour lane placed just
  // past the last obstacle is frequently just past the last tile as well. The
  // page was sized from the boxes alone, so those few pixels fell outside the
  // frame and the clip cut the hop down to whichever fragment survived — one
  // that stood nearly two inches from the service it names, with no arrowhead.
  // So the view is the union of everything drawn, not of the boxes alone.
  // Routing is untouched by this, and wherever the drawing is smaller than the
  // frame the scale cap absorbs the extra room without changing anything.
  //
  // Slices keep exactly their own window: every window is the same size so that
  // every slide renders at one scale, and the overview covers the whole drawing
  // anyway.
  const drawnView = banded ? view : routes.reduce(
    (acc, route) => route.points.reduce((box, point) => ({
      minX: Math.min(box.minX, point.x),
      maxX: Math.max(box.maxX, point.x),
      minY: Math.min(box.minY, point.y),
      maxY: Math.max(box.maxY, point.y),
    }), acc),
    view,
  );
  // The cap is a DENSITY, and a density is the wrong thing to cap when the
  // drawing itself was authored small. At 96 px per inch a 12px node is drawn
  // 0.125in wide, whose text column is 0.065in - narrower than two of the
  // digit "1" - so the tile can carry no name, no stub and not even a key, and
  // eight instances of one service came out as eight anonymous dots with an
  // index that could only say "(not drawn)" eight times. The cap exists to stop
  // a two-node diagram being blown up to absurd tiles, and it still does: this
  // only RAISES it, only for a drawing whose smallest tile cannot carry a mark,
  // and only far enough to reach that bar. Every drawing already above it keeps
  // the identical transform, so the geometry of the rest of the corpus is
  // untouched by construction. The frame terms of `computeFitTransform` still
  // bind, so nothing is magnified past the page. The expression itself lives on
  // `rendererMaxScale`, which the planner reads too - the two had diverged.
  const minBoxW = narrowestBoxW(services);
  // The same bar the planner used. A renderer that stops magnifying below the
  // scale the planner split to leaves tiles under the bar with the slides
  // already spent, which is the divergence this expression was pulled into
  // `rendererMaxScale` to prevent - and a numbered drawing moves the bar.
  const maxScale = rendererMaxScale(minBoxW, markableTileWIn(routes.map((r) => r.stepNumber)));
  const transform = computeFitTransform(drawnView, frame, { maxScale });
  const clampTo = clamped || banded ? frame : undefined;
  // A tile is drawn where the drawing says; a chip is drawn *around* its arrow
  // and is therefore the one shape that can be pushed off the sheet by its own
  // size. Parallel ordinals on a short hop used to land in the header strip and
  // below the footer, and their step numbers vanished with them. Chips and
  // badges are always held inside the frame, on every slide, banded or not.
  const labelFrame = clampTo ?? frame;
  const px = transform.scale * PX_PER_IN;
  const first = { x: fitBounds.minX <= bounds.minX + 0.5, y: fitBounds.minY <= bounds.minY + 0.5 };
  const last = { x: fitBounds.maxX >= bounds.maxX - 0.5, y: fitBounds.maxY >= bounds.maxY - 0.5 };

  // A window owns whatever falls inside its `own` rectangle, so a shape
  // straddling a seam is drawn once instead of twice, and `clampTo` pulls
  // strays back onto the page exactly as on an unbanded slide. The `own`
  // rectangles cover the drawing exactly once between them even where an empty
  // cell was not worth a slide of its own, so nothing anchored anywhere in the
  // drawing can end up belonging to no part.
  const owns = (x: number, y: number): boolean => !banded || windowOwnsPoint(ownBounds, bounds, x, y);
  const visibleBox = (box: ExportBox): boolean => owns(box.x + box.w / 2, box.y + box.h / 2);
  // A zone is routinely larger than a whole window, so centre-ownership would
  // print the boundary and its name on one slide and leave the services on the
  // other slides floating with no container. Unlike a service tile, a zone is
  // continued on every window it overlaps — the palette index below is already
  // stable across slices, and a partial rectangle reads as a boundary that
  // carries on, which is exactly what it does.
  const overlapsAxis = (lo: number, hi: number, wLo: number, wHi: number, isFirst: boolean, isLast: boolean): boolean =>
    (hi >= wLo || isFirst) && (lo <= wHi || isLast);
  const visibleGroup = (box: ExportBox): boolean => !banded || (
    overlapsAxis(box.x, box.x + box.w, view.minX, view.maxX, first.x, last.x)
    && overlapsAxis(box.y, box.y + box.h, view.minY, view.maxY, first.y, last.y)
  );
  // A connector is continued on every window it crosses so the reader can
  // follow where it goes; only the window holding its anchor draws the chip and
  // number.
  const visibleRoute = (route: ExportRoute): boolean => {
    if (!banded) return true;
    if (route.points.length === 0) return owns(route.labelAnchor.x, route.labelAnchor.y);
    const xs = route.points.map((point) => point.x);
    const ys = route.points.map((point) => point.y);
    if (!(overlapsAxis(Math.min(...xs), Math.max(...xs), view.minX, view.maxX, first.x, last.x)
      && overlapsAxis(Math.min(...ys), Math.max(...ys), view.minY, view.maxY, first.y, last.y))) return false;
    // Bounding-box overlap alone lets a hop that merely clips a corner of this
    // window be drawn as a stub at the seam. Several wrap-arounds leaving the
    // same edge reduce to the same stub, so the reader sees one short line
    // standing for three hops that go somewhere else entirely. A quarter of the
    // hop has to be on this paper for the fragment to be worth drawing.
    //
    // Safe to drop because a banded deck always opens with an overview slide
    // covering the whole drawing, so every route that meets the fitted bounds
    // is drawn there whatever the slices decide — and the audit fails any deck
    // with an edge drawn on no slide at all.
    let total = 0;
    let inView = 0;
    for (let i = 1; i < route.points.length; i += 1) {
      const a = route.points[i - 1];
      const b = route.points[i];
      const len = Math.hypot(b.x - a.x, b.y - a.y);
      total += len;
      const steps = Math.max(1, Math.ceil(len / 8));
      for (let s = 0; s < steps; s += 1) {
        const t = (s + 0.5) / steps;
        const x = a.x + (b.x - a.x) * t;
        const y = a.y + (b.y - a.y) * t;
        if (x >= view.minX && x <= view.maxX && y >= view.minY && y <= view.maxY) inView += len / steps;
      }
    }
    // A share alone is backwards for exactly the hops that matter most. The
    // longer a hop is, the smaller its share of any one window — so a pure
    // fraction suppresses hardest the wrap-around that explains how one row
    // reaches the next. On a 24-wide chain the row-turn hop had 17% of itself
    // on every window and was dropped from all seven, although the visible
    // piece was longer than the window is wide. So a fragment also earns its
    // place by absolute length: more than half a window of arrow is never a
    // meaningless stub, whatever fraction of the whole it happens to be.
    const span = Math.max(view.maxX - view.minX, view.maxY - view.minY);
    return total <= 0 || inView >= 0.25 * total || inView >= 0.6 * span;
  };
  const shownGroups = groups.filter(visibleGroup);
  const shownServices = services.filter(visibleBox);
  const shownRoutes = routes.filter(visibleRoute);
  // Ownership of a label and visibility of its arrow were decided by two
  // different tests, so a long diagonal hop could have its midpoint owned by a
  // window that carried too little of the arrow to draw it, while the windows
  // that did draw it owned no part of the label. The wording was then written
  // on no slide at all. The window holding the anchor draws the hop, whatever
  // fraction of it lands there: exactly one window owns any point, so this adds
  // an arrow, never a duplicate.
  const annotatedRoutes = routes.filter((route) => owns(route.labelAnchor.x, route.labelAnchor.y));
  for (const route of annotatedRoutes) if (!shownRoutes.includes(route)) shownRoutes.push(route);
  const icons = presetIcons ?? await rasterizeIcons(shownServices.map((service) => service.iconPath), 128);

  // Index by the full group list so a zone keeps its palette colour on every
  // slice it appears on.
  const zoneMembers = (group: ExportBox, list: readonly ExportBox[]): number => list.filter((service) => {
    const cx = service.x + service.w / 2;
    const cy = service.y + service.h / 2;
    return cx >= group.x && cx <= group.x + group.w && cy >= group.y && cy <= group.y + group.h;
  }).length;
  const placedTiles = shownServices.map((service) => placeBox(service, transform, clampTo));
  // A zone whose members all landed on other slides is not drawn here. The
  // window can cut a zone to a 0.4in sliver holding none of its services, and
  // a closed box around nothing, captioned "… (0 / 1)" because 0.2in of column
  // holds no name, tells the reader nothing and paints its marker outside its
  // own band doing it. A boundary is a claim about contents; with no contents
  // on the slide there is no claim to make. Zones the author drew empty are
  // untouched — they have no members anywhere, so nothing is being hidden.
  const drawnGroups = shownGroups.filter(
    (group) => zoneMembers(group, shownServices) > 0 || zoneMembers(group, services) === 0,
  );
  const placedZones = new Map(drawnGroups.map((group) => [group.id, placeBox(group, transform, clampTo, true)]));
  // Where each arrow's label wants to sit, sized as the chip that will be put
  // there. Approximate on purpose: the chips have not been placed yet and
  // cannot be, since their own search reads the caption bands this call is
  // about to produce. The anchor is where `connectorLabelBox` starts its walk,
  // so it is the best available statement of which paper is spoken for.

  const captionBands: Obstacle[] = [];
  drawnGroups.forEach((group) => {
    const bands = addGroupShape(
      pptx, slide, group, groups.indexOf(group), transform, clampTo,
      { here: zoneMembers(group, shownServices), all: zoneMembers(group, services) },
      // Captions already chosen are paper too. Nested zones are the case:
      // `trespass` charges for writing inside a foreign zone, but a zone drawn
      // inside another one is not trespassing, so two captions on the same
      // rows scored identically and were written 0.09in apart on top of each
      // other. The list is live and the loop is sequential, so each zone sees
      // exactly the bands settled before it.
      [...placedTiles, ...captionBands.map((band) => ({ x: band.x, y: band.y, w: band.w, h: band.h }))],
      drawnGroups.filter((other) => other !== group).map((other) => placedZones.get(other.id)!),
      thumbnail ? undefined : truncatedNames,
    );
    // A zone caption is worth exactly what a tile caption is worth: it is the
    // only thing that says what the box contains, and a chip over it is a
    // wrong claim rather than a blemish.
    captionBands.push({ ...bands.caption, weight: 60, caption: true });
  });
  // Where a dark tile sits, in terms the reader can count off the page.
  //
  // Reading order among the tiles THIS slide draws, banded by row: sorting on
  // `y` alone makes two tiles whose tops differ by a pixel into separate rows,
  // and sorting on `x` alone interleaves rows. The band is the median tile
  // height, which is the same quantity the eye uses to decide what is a row.
  const rowBand = Math.max(
    1,
    [...shownServices].map((s) => s.h).sort((a, b) => a - b)[Math.floor(shownServices.length / 2)] ?? 1,
  );
  const readingRank = new Map<string, number>(
    [...shownServices]
      .sort((a, b) => (Math.round(a.y / rowBand) - Math.round(b.y / rowBand)) || (a.x - b.x))
      .map((service, i) => [String(service.id), i + 1]),
  );
  const locate = (service: ExportBox): string => {
    const rank = readingRank.get(String(service.id));
    if (rank === undefined) return slideLabel;
    const where = `box ${rank} of ${shownServices.length} in reading order`;
    return slideLabel ? `${slideLabel}, ${where}` : where;
  };
  // Per SLIDE, not per deck: the reader looks at one slide at a time, so the
  // keys only have to be distinguishable among the tiles they are drawn beside.
  for (const service of shownServices) {
    const bands = addNodeShape(
      pptx, slide, service, transform,
      service.iconPath ? icons.get(service.iconPath) : undefined,
      px, clampTo, thumbnail, drawnHere, keyOrdinal,
    );
    // The overview is allowed to clip: every name it clips is drawn in full on
    // the slice that follows. Only a window slide's clipping is a real loss.
    //
    // The MARK is not allowed to go with it. An index that prints only the full
    // name leaves the reader holding a tile marked "3" and a list of sentences,
    // none of which contains a 3; the row has to define the mark it is looked
    // up by. EVERY distinct mark is kept: two slides draw the same service at
    // two widths, and quoting only the longer one left the shorter one drawn on
    // a tile and defined nowhere - the very thing the index exists to prevent.
    //
    // AND EVERY SLIDE THAT DRAWS ONE, INCLUDING THE OVERVIEW. Gating the
    // recording on `!thumbnail` alongside the clipping made the gap structural:
    // the overview is where tiles are smallest, so it is where keys are most
    // often needed, and it was the one slide guaranteed to define none of them.
    // `wide-chain` shipped an overview covered in 38 bare integers in a deck
    // with no index slide at all, in a drawing whose workflow numbers its own
    // steps with bare integers - and "3" sat on the tile for Service 10.
    // Whether the overview's clipping counts as a LOSS is a separate question
    // from whether the mark it draws needs defining; it does.
    if (bands.clipped && (!thumbnail || bands.drawn)) {
      recordMark(
        truncatedNames,
        bands.clipped,
        bands.drawn || unlabelledRow(locate(service)),
      );
    }
    // A tile can be leaned on: the reader still sees which service it is. Its
    // name cannot, because the name is the only thing that says so, and a chip
    // is drawn over it at 92% opacity. Weighted far above a tile so that even
    // a chip allowed to touch its own endpoint is pushed off the words.
    if (bands.caption) captionBands.push({ ...bands.caption, weight: 60, caption: true });
    // The sub-line deliberately gets no obstacle of its own. It is drawn INSIDE
    // the tile, and the tile is already an obstacle, so the only way a chip
    // could ever reach it was by the sub-line escaping the tile — which is what
    // a `wrap="none"` line wider than its box did. Fitting it to the tile
    // closes that, and an extra band here proved indistinguishable from nothing
    // across every fixture and every row spacing tried. The audit measures it
    // regardless, which is what keeps this honest.
  }

  for (const route of shownRoutes) addConnector(pptx, slide, route, transform, clampTo);
  // Labels are drawn after every connector so a chip is never hidden by a line
  // that is rendered later.
  //
  // The floor is the tiles' floor, deliberately. A chip was allowed down to 4pt
  // while a tile name stopped at 7, so a window slide wrote its arrow labels at
  // 6.74pt beside tile names held at 7.04 — grey mush on a projector, and the
  // one piece of text on the slide that says *why* two services are connected.
  // A chip that cannot be written legibly has somewhere better to go:
  // `connectorLabelBox` drops it and the workflow list on the slide still
  // carries the sentence against the same step number. Smaller-but-drawn is the
  // worse of the two outcomes.
  //
  // The overview keeps its own, lower floor. It is a map rather than a reading
  // surface, its names are carried by the slides that follow, and forcing its
  // chips up to reading size would only crowd the picture it exists to give.
  const labelFontSize = clamp(9 * px, thumbnail ? OVERVIEW_LEGIBLE_PT : LEGIBLE_TILE_PT, 10);
  // Chips and numbers dodge the tiles that are actually on this slide, so a
  // label on a short hop is pushed clear instead of covering a service.
  const tileRects = shownServices.map((service) => ({ ...placeBox(service, transform, clampTo), node: service.id }));
  // Place every chip before drawing any of them, adding each to the obstacle
  // list as it is settled. Parallel edges between the same pair are staggered,
  // but the tile-avoidance walk could drag two of them back onto the same spot
  // because a chip could not see the chips already placed.
  const chipObstacles: Obstacle[] = [...tileRects, ...captionBands];
  // The colour key is drawn last and is all but opaque, so anything it lands on
  // is simply gone: a numbered callout under it leaves the workflow band citing
  // a step the reader cannot find. Reserve whichever corner it will take.
  const legendRect = connectionLegendRect(diagram.edges ?? [], frame);
  if (legendRect) chipObstacles.push({ ...legendRect, weight: 4 });
  const chips = new Map<string, ReturnType<typeof connectorLabelBox>>();
  const badges = new Map<string, ReturnType<typeof stepBadgeBox>>();
  const parallel = new Map<string, number>();
  const bundleKey = (route: ExportRoute): string => (route.sourceId < route.targetId
    ? `${route.sourceId}|${route.targetId}`
    : `${route.targetId}|${route.sourceId}`);
  for (const route of annotatedRoutes) {
    const key = bundleKey(route);
    parallel.set(key, (parallel.get(key) ?? 0) + 1);
  }
  // Every arrow on the slide as a thin ribbon, so a chip pushed off a tile does
  // not come to rest on somebody else's line. A chip's own bundle is skipped —
  // it is meant to sit on its own hop — and the ribbons go in with the tiles so
  // the repair pass treats them as fixed scenery.
  for (const route of shownRoutes) {
    const key = bundleKey(route);
    const points = route.points.map((point) => toInches(point, transform));
    for (let i = 1; i < points.length; i += 1) {
      const a = points[i - 1];
      const b = points[i];
      const x = Math.min(a.x, b.x);
      const y = Math.min(a.y, b.y);
      const w = Math.abs(b.x - a.x);
      const h = Math.abs(b.y - a.y);
      if (w + h < 0.05) continue;
      chipObstacles.push({ x: x - 0.03, y: y - 0.03, w: w + 0.06, h: h + 0.06, owner: key, weight: RIBBON_WEIGHT });
    }
  }
  // Measure each bundle before placing any of it: the rungs of a ladder have to
  // step by the tallest block in that ladder, and hang off the bundle's own
  // centre, neither of which is known until every label in it has been sized.
  const bundles = new Map<string, {
    count: number; rung: number; x: number; y: number;
    shift?: number; across?: number; font?: number; perCol?: number;
    rank?: ReadonlyMap<string, number>;
    span?: number;
    maxWidth?: number;
    badgesOnly?: boolean;
    dirty?: number;
    stray?: number;
  }>();
  const shapes = new Map<string, { w: number; h: number; alongX: boolean }>();
  for (const route of annotatedRoutes) {
    const key = bundleKey(route);
    const count = parallel.get(key) ?? 1;
    if (count < 2) continue;
    const at = toInches(route.labelAnchor, transform);
    const seed = bundles.get(key) ?? { count, rung: 0, x: 0, y: 0 };
    const probe = connectorLabelBox(route, transform, labelFontSize, px, labelFrame, [], { count, rung: 0, x: at.x, y: at.y });
    const shape = shapes.get(key);
    if (probe) {
      shapes.set(key, {
        w: Math.max(shape?.w ?? 0, probe.w),
        h: Math.max(shape?.h ?? 0, probe.block.h),
        alongX: probe.alongX,
      });
    }
    const ends = [route.points[0] ?? route.labelAnchor, route.points[route.points.length - 1] ?? route.labelAnchor]
      .map((point) => toInches(point, transform));
    bundles.set(key, {
      count,
      rung: Math.max(seed.rung, probe ? (probe.alongX ? probe.block.h + 0.05 : probe.w + 0.12) : 0),
      x: seed.x + at.x / count,
      y: seed.y + at.y / count,
      span: Math.max(seed.span ?? 0, Math.abs(ends[1].x - ends[0].x), Math.abs(ends[1].y - ends[0].y)),
    });
  }
  // Rung order follows the arrows, not the edge order. `parallelOffset` fans
  // the routes alternately about the centre, so a ladder numbered by ordinal
  // runs in a different order from the arrows it labels and the callouts end up
  // beside the wrong ones.
  for (const [key, bundle] of bundles) {
    const members = annotatedRoutes
      .filter((route) => bundleKey(route) === key)
      .sort((a, b) => (a.fanOffset - b.fanOffset) || (a.ordinal - b.ordinal));
    bundle.rank = new Map(members.map((route, index) => [route.id, index]));
  }
  // Move the whole ladder off the tiles rather than letting each rung walk on
  // its own: a rung that leaves the lattice lands part-way inside the one above
  // it, so the group takes a single offset that every rung shares. Ordinary
  // single connectors are settled first, so a ladder — which is far the larger
  // block — is the thing that has to dodge, rather than displacing a chip that
  // has nowhere to go but onto its neighbour.
  const ordered = [
    ...annotatedRoutes.filter((route) => !bundles.has(bundleKey(route))),
    ...annotatedRoutes.filter((route) => bundles.has(bundleKey(route))),
  ];
  // A chip landing on a service tile is untidy; a chip landing on another
  // annotation makes one of the two unreadable. Keep the two apart so the
  // ladder search always trades tile coverage for annotation clearance.
  const tileObstacleCount = chipObstacles.length;
  // Every arrow as line segments in page inches, grouped by the bundle it
  // belongs to, so a ladder can be asked the one question a collision check
  // cannot answer: is this rung still nearer its OWN hop than anybody else's?
  type Seg = { ax: number; ay: number; bx: number; by: number };
  const segsByBundle = new Map<string, Seg[]>();
  const segsByRoute = new Map<string, Seg[]>();
  for (const route of shownRoutes) {
    const key = bundleKey(route);
    const pts = route.points.map((point) => toInches(point, transform));
    const list = segsByBundle.get(key) ?? [];
    const own: Seg[] = [];
    for (let i = 1; i < pts.length; i += 1) {
      own.push({ ax: pts[i - 1].x, ay: pts[i - 1].y, bx: pts[i].x, by: pts[i].y });
    }
    list.push(...own);
    if (own.length) segsByRoute.set(route.id, own);
    if (list.length) segsByBundle.set(key, list);
  }
  const gapToSeg = (x: number, y: number, s: Seg): number => {
    const vx = s.bx - s.ax;
    const vy = s.by - s.ay;
    const len = vx * vx + vy * vy;
    const t = len <= 0 ? 0 : Math.max(0, Math.min(1, ((x - s.ax) * vx + (y - s.ay) * vy) / len));
    return Math.hypot(x - (s.ax + t * vx), y - (s.ay + t * vy));
  };
  const gapToSegs = (x: number, y: number, segs: readonly Seg[]): number => {
    let best = Number.POSITIVE_INFINITY;
    for (const s of segs) {
      const d = gapToSeg(x, y, s);
      if (d < best) best = d;
    }
    return best;
  };
  // How much nearer a foreign arrow has to be before a reader would credit the
  // label to it. The audit uses the same slack, because this is the rule it
  // enforces on the finished deck.
  const CONFUSION_SLACK = 0.25;
  // What one misread rung is worth, in the square inches the rest of the score
  // is measured in. A service tile is about 1.2 sq in, so a rung parked beside
  // the wrong arrow costs roughly half a covered tile: enough that a ladder
  // will sit on a corner rather than walk to a clean slot beside a stranger,
  // and not so much that it refuses to move when there is nowhere else at all.
  const DIRTY_RUNG_COST = 0.5;
  // What one MISREAD rung is worth. Dearer than a clipped one, and dearer than
  // a whole covered tile: a rung clipping a corner is untidy and the reader
  // still knows what it says about which arrow, but a rung parked beside a
  // stranger's hop is read, believed, and wrong. Kept apart from the clipping
  // count because the two have different cures — a clip usually moves away,
  // and a misread often cannot move anywhere at all.
  const STRAY_RUNG_COST = 1.5;
  // The arrows a given hop's own label could be mistaken for: every bundle but
  // its own, near enough to compete. Cached, because the walk asks per
  // candidate position and a wide estate has hundreds of segments.
  const rivalCache = new Map<string, Seg[]>();
  const foreignGapFor = (route: ExportRoute): ((x: number, y: number) => number) | undefined => {
    const key = bundleKey(route);
    let rivals = rivalCache.get(key);
    if (!rivals) {
      rivals = [];
      for (const [other, segs] of segsByBundle) {
        if (other === key) continue;
        rivals.push(...segs);
      }
      rivalCache.set(key, rivals);
    }
    return rivals.length ? (x: number, y: number) => gapToSegs(x, y, rivals) : undefined;
  };
  // Which step numbers the workflow slide will actually narrate. Dropping a
  // chip is only ever a trade against that list; with no row to read it is a
  // deletion.
  const narratedRows = new Map(workflowListFromEdges(diagram.edges ?? []).map((entry) => [entry.step, entry.description]));
  const narratedSteps = new Set(narratedRows.keys());
  // Numbers a labelled-but-unnarrated edge may be promoted into if its chip
  // turns out to have nowhere legible to stand. Allocated here, from the whole
  // edge list rather than the routes this slide happens to show, so an edge
  // gets the same number on the overview as it does on the slice that promotes
  // it — the deck would otherwise print two different badges for one hop.
  const promotable = new Map<string, number>();
  {
    let next = 0;
    for (const edge of diagram.edges ?? []) {
      // `readStepValue` is the one predicate every other reader of this field
      // uses. A model emits `"1"` about as often as `1`, and the load path
      // never coerces it, so reading the raw field here disagreed with the
      // workflow list in both directions: it skipped an authored `"2"` (then
      // handed 2 out again to a promoted edge, whose row the list dropped as a
      // duplicate — losing exactly the label promotion exists to carry) and it
      // accepted a `2.5` that no list row can match.
      const step = readStepValue((edge.data as { stepNumber?: unknown } | undefined)?.stepNumber);
      if (step !== undefined) next = Math.max(next, step);
    }
    const waiting = (diagram.edges ?? [])
      // Only an edge with a label has anything to promote. Counting the
      // unlabelled ones spent numbers on rows that will never exist, so a deck
      // with three authored steps and twelve plain connectors numbered its one
      // promoted hop 16 and read 1, 2, 3, 16.
      .filter((edge) => readStepValue((edge.data as { stepNumber?: unknown } | undefined)?.stepNumber) === undefined
        && readEdgeLabel(edge) !== '')
      .map((edge) => edge.id)
      .sort();
    waiting.forEach((id, at) => promotable.set(id, next + 1 + at));
  }

  const ownGapFor = (route: ExportRoute): ((x: number, y: number) => number) | undefined => {
    // Its OWN arrow, not its bundle's. A fan is fanned — `parallelOffset`
    // spreads the members apart — so "nearest arrow in my bundle" is a much
    // shorter distance than "nearest point of the arrow I am labelling", and
    // measuring the first lets a callout drift onto a stranger's hop while
    // still passing the attribution test.
    const segs = segsByRoute.get(route.id);
    return segs && segs.length ? (x: number, y: number) => gapToSegs(x, y, segs) : undefined;
  };
  const settle = (routes: readonly ExportRoute[]): void => {
    for (const route of routes) {
      const bundle = bundles.get(bundleKey(route));
      // A muted fan carries no wording at all, so the rigid lattice that keeps
      // text from colliding has nothing left to protect - and a lattice cannot
      // bend, so on a crowded slide it comes down as a block beside somebody
      // else's hop. A bare number is small enough to sit on the arrow it
      // belongs to, which is what the Architecture Center draws, and it is then
      // placed one at a time against the same attribution test as everything
      // else.
      let box = bundle?.badgesOnly ? null : connectorLabelBox(
        route, transform, labelFontSize, px, labelFrame, chipObstacles, bundle,
        undefined, foreignGapFor(route),
      );
      // The sizer returns null for a labelled hop whose corridor cannot host
      // one letter, and it is right to: a smear is worse than nothing. But it
      // decides that against `route.stepNumber`, which is read BEFORE this walk
      // can promote the route — so an un-numbered labelled edge took the drop
      // and never reached the promotion branch below, built for exactly this
      // edge. The wording then appeared on no slide at all: no chip, no badge,
      // and `workflowListFromEdges` needs a number the edge does not have.
      //
      // Promote first, then ask again. The route now has a number, so the sizer
      // keeps its callout and the sentence goes to the workflow slide.
      // Membership of `promotable` already means "labelled and un-numbered", so
      // it is the whole test — the route's own `label` is the resolved one.
      //
      // Reachability, measured: every assignment to the sizer's `roomW` is
      // clamped with lower bound `0.34 * px`, and the chip font is
      // `clamp(9 * px, 7, 10)`. Since the bar became TWO of the widest glyph
      // the drop needs `0.34*px + 0.002 < 2*widestGlyph + 0.12`, which fires
      // well inside the corpus range (min px 0.7473 over 3,124 non-thumbnail
      // samples) — before the two-glyph bar it needed px < 0.633 and was dead
      // code. Kept, and now live.
      if (box === null && !bundle?.badgesOnly && !thumbnail
        && route.stepNumber === undefined) {
        const step = promotable.get(route.id);
        if (step !== undefined) {
          route.stepNumber = step;
          promotedSteps.set(step, route.label);
          box = connectorLabelBox(
            route, transform, labelFontSize, px, labelFrame, chipObstacles, bundle,
            undefined, foreignGapFor(route),
          );
        }
      }
      // A chip that still lands on a service name or on another callout after
      // the walk has done its best is worse than no chip: it is drawn at 92%
      // opacity over the one thing that says which service this is, or over a
      // step number. On a slide this crowded the Architecture Center leaves a
      // numbered callout on the arrow and puts the sentence in the step list,
      // which is exactly the trade `mutedWording` already implements — so make
      // it per route, not only per fan.
      if (box && route.stepNumber !== undefined && narratedSteps.has(route.stepNumber)) {
        if (chipSpoils(box.block, { sourceId: route.sourceId, targetId: route.targetId, bundleKey: bundleKey(route) }, chipObstacles)) {
          if (!thumbnail && route.label
            && !carriesWording(narratedRows.get(route.stepNumber) ?? '', route.label)) {
            mutedWording.set(route.stepNumber, route.label);
          }
          box = null;
        }
      } else if (box && chipSpoils(box.block, { sourceId: route.sourceId, targetId: route.targetId, bundleKey: bundleKey(route) }, chipObstacles)) {
        // A route with no step number has no step list to hand its wording to,
        // so dropping the chip would simply lose it — which is why this check
        // was written to run only on narrated routes, and why an un-narrated
        // chip was free to sit on a name. Free is the wrong answer: an unnamed
        // tile no longer says which service it is, and that is a worse loss
        // than a chip that has become a numbered badge.
        //
        // The walk cannot be sent further to fix it — `inReach` holds a chip
        // beside the arrow it labels, and it must, because a chip an inch from
        // its hop is read, believed and wrong. So build the row that was
        // missing instead: give the edge the next free number, hand its label
        // to the workflow slide, and let the badge stand where the chip could
        // not. That is exactly what `narrateEdgeCallouts` already does for the
        // members of a deep fan, applied to the other case that has nowhere to
        // put its words.
        const step = promotable.get(route.id);
        if (step !== undefined && !thumbnail) {
          route.stepNumber = step;
          promotedSteps.set(step, route.label);
          box = null;
        }
      }
      chips.set(route.id, box);
      if (box) chipObstacles.push(box.block);
      const badge = stepBadgeBox(
        route, transform, px, labelFrame, box, chipObstacles,
        ownGapFor(route), foreignGapFor(route), thumbnail,
      );
      badges.set(route.id, badge);
      // The chip reserves room for its own badge inside its block, so a badge
      // still sitting there needs no obstacle of its own. One that left — the
      // walk moves it when the chip's slot is buried — is outside that block
      // and invisible to every annotation placed afterwards, which is how a
      // callout came to be covered by a chip settled six hops later.
      if (badge && badge !== box?.badge) {
        chipObstacles.push({ x: badge.x, y: badge.y, w: badge.d, h: badge.d, annotation: true });
      }
    }
  };
  settle(ordered.filter((route) => !bundles.has(bundleKey(route))));

  type Block = Obstacle;
  type Bundle = NonNullable<ReturnType<typeof bundles.get>>;
  const area = (a: Block, b: Block): number => {
    const dx = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
    const dy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return dx > 0 && dy > 0 ? dx * dy : 0;
  };
  // Score the bundle at its REAL geometry: probe the same routine that will
  // place it, with no obstacles so no walk runs. Modelling the lattice by hand
  // kept choosing offsets that the in-frame clamps then undid, so the search
  // believed it had found a clear slot that never materialised.
  //
  // `ignore` lifts a bundle's own already-placed blocks out of the obstacle
  // list so the same search can be re-run later. Bundles are otherwise settled
  // one after another and frozen, which leaves the first one on a slide unable
  // to move even when a later one has come down on top of it.
  const searchBundle = (key: string, bundle: Bundle, ignore: ReadonlySet<Block> = new Set()): number => {
    const shape = shapes.get(key);
    if (!shape || bundle.rung <= 0) return 0;
    const members = annotatedRoutes.filter((route) => bundleKey(route) === key);
    // The arrows this ladder could be mistaken for. Only the ones near enough
    // to compete: a rung is never credited to a hop on the far side of the
    // slide, and scoring against every segment on a busy drawing would put the
    // whole export back into the quadratic regime the stride pass escaped.
    const ownSegs = segsByBundle.get(key) ?? [];
    const rivals: Seg[] = [];
    if (ownSegs.length) {
      let lo = { x: Infinity, y: Infinity };
      let hi = { x: -Infinity, y: -Infinity };
      for (const s of ownSegs) {
        lo = { x: Math.min(lo.x, s.ax, s.bx), y: Math.min(lo.y, s.ay, s.by) };
        hi = { x: Math.max(hi.x, s.ax, s.bx), y: Math.max(hi.y, s.ay, s.by) };
      }
      const REACH = 4;
      for (const [other, segs] of segsByBundle) {
        if (other === key) continue;
        for (const s of segs) {
          if (Math.min(s.ax, s.bx) > hi.x + REACH || Math.max(s.ax, s.bx) < lo.x - REACH) continue;
          if (Math.min(s.ay, s.by) > hi.y + REACH || Math.max(s.ay, s.by) < lo.y - REACH) continue;
          rivals.push(s);
        }
      }
    }
    // A ladder's rungs are a fixed shape that the search only ever TRANSLATES,
    // so measure them once. Re-wrapping every label for every candidate offset
    // made a 30-node grid of three-way fans take 42 seconds to export; the
    // translation below is arithmetic, and the exact routine is still used for
    // any candidate near enough to the frame edge for a clamp to fire.
    const measured = members.map((member) => connectorLabelBox(
      member, transform, labelFontSize, px, labelFrame, [], { ...bundle, shift: 0, across: 0 },
    )).filter((box): box is NonNullable<typeof box> => box !== null);
    const inset = 0.03;
    // The ladder clamp is driven by each rung's own size measured from the
    // bundle centre, so a short rung can be clamped while its own rect is still
    // inside the frame. Keep the fast path away from the edge by the spread of
    // rung sizes and that case cannot arise.
    const ladderSize = (box: (typeof measured)[number]): number => (box.alongX ? box.block.h : box.block.w);
    const sizes = measured.map(ladderSize);
    const slack = sizes.length > 0 ? (Math.max(...sizes) - Math.min(...sizes)) / 2 : 0;
    const free = (block: Block): boolean => block.x >= labelFrame.x + inset + slack
      && block.y >= labelFrame.y + inset + slack
      && block.x + block.w <= labelFrame.x + labelFrame.w - inset - slack
      && block.y + block.h <= labelFrame.y + labelFrame.h - inset - slack;
    // Only safe to translate when nothing was clamped at the seed either, or
    // the "unclamped" rects would be translated copies of a clamped one.
    const seedFree = measured.length === members.length && measured.every((box) => free(box.block));
    // Prove it rather than assume it: the ladder clamp is driven by each rung's
    // own height, so a fan can be clamped while every individual rung still
    // sits inside the frame. One exact probe per bundle settles it.
    const probe = (dAlong: number, dAcross: number): boolean => {
      for (let i = 0; i < members.length; i += 1) {
        const box = connectorLabelBox(
          members[i], transform, labelFontSize, px, labelFrame, [],
          { ...bundle, shift: dAlong, across: dAcross },
        );
        const want = measured[i];
        if (!box || !want) return false;
        const moved = {
          x: want.block.x + (want.alongX ? dAcross : dAlong),
          y: want.block.y + (want.alongX ? dAlong : dAcross),
        };
        if (Math.abs(box.block.x - moved.x) > 1e-6 || Math.abs(box.block.y - moved.y) > 1e-6) return false;
        if (Math.abs(box.block.w - want.block.w) > 1e-6 || Math.abs(box.block.h - want.block.h) > 1e-6) return false;
      }
      return true;
    };
    const translatable = seedFree && probe(0.11, 0.07) && probe(-0.09, -0.13);
    const taken = chipObstacles.filter((block) => !ignore.has(block) && block.owner !== key);
    // A ladder is scored against every obstacle on the slide for every
    // candidate offset, which on a dense grid is tens of millions of rectangle
    // tests. Bucket them once so each rung only meets the handful it could
    // possibly touch. `stamp` keeps an obstacle spanning several cells from
    // being counted twice, so the score is identical to the linear scan.
    const CELL = 0.5;
    const buckets = new Map<string, number[]>();
    const cellsOf = (r: { x: number; y: number; w: number; h: number }): string[] => {
      const keys: string[] = [];
      for (let cx = Math.floor(r.x / CELL); cx <= Math.floor((r.x + r.w) / CELL); cx += 1) {
        for (let cy = Math.floor(r.y / CELL); cy <= Math.floor((r.y + r.h) / CELL); cy += 1) keys.push(`${cx},${cy}`);
      }
      return keys;
    };
    taken.forEach((block, index) => {
      for (const cell of cellsOf(block)) {
        const list = buckets.get(cell);
        if (list) list.push(index); else buckets.set(cell, [index]);
      }
    });
    const stamp = new Int32Array(taken.length);
    let visit = 0;
    const cost = (shift: number, across: number): { total: number; onLabel: number; overlap: number; dirty: number; stray: number } => {
      // Each rung paired with the segments of the ONE arrow it labels. A fan is
      // fanned, so the bundle's nearest segment to a given rung is usually a
      // sibling's arrow rather than its own, and scoring attribution against
      // the bundle lets a rung drift onto a stranger's hop while still reading
      // as correctly attributed. Carried alongside the block because a rung
      // that fails to measure is skipped, and a bare index would then pair
      // every later rung with the wrong arrow.
      const placed: { block: Block; own: Seg[] }[] = [];
      if (translatable) {
        for (let i = 0; i < measured.length; i += 1) {
          const box = measured[i];
          const moved = {
            ...box.block,
            x: box.block.x + (box.alongX ? across : shift),
            y: box.block.y + (box.alongX ? shift : across),
          };
          if (!free(moved)) { placed.length = 0; break; }
          placed.push({ block: moved, own: segsByRoute.get(members[i].id) ?? ownSegs });
        }
      }
      if (placed.length !== members.length) {
        placed.length = 0;
        for (const member of members) {
          const box = connectorLabelBox(
            member, transform, labelFontSize, px, labelFrame, [], { ...bundle, shift, across },
          );
          if (box) placed.push({ block: box.block, own: segsByRoute.get(member.id) ?? ownSegs });
        }
      }
      // Anything the ladder lands on, plus any rung the frame clamp has stacked
      // back on top of a sibling. Landing on another annotation is weighted far
      // above landing on a tile: a covered tile is untidy, a covered label is
      // lost.
      let tiles = 0;
      let labels = 0;
      // How many rungs a reader would see as wrong, rather than how much area
      // is covered: the fallback is a judgement about legibility, and one chip
      // buried under a tile with eight clean ones is a different picture from
      // nine chips each clipping a corner.
      let dirty = 0;
      // Rungs a reader would credit to the wrong arrow. Counted apart from the
      // clipped ones because it is the only kind of dirt a ladder cannot always
      // walk away from, and that is what decides whether it should stop being a
      // ladder at all.
      let stray = 0;
      for (let i = 0; i < placed.length; i += 1) {
        const rect = placed[i].block;
        const own = Math.max(0.0001, rect.w * rect.h);
        let mine = 0;
        visit += 1;
        for (const cell of cellsOf(rect)) {
          for (const index of buckets.get(cell) ?? []) {
            if (stamp[index] === visit) continue;
            stamp[index] = visit;
            const other = taken[index];
            const hit = area(rect, other);
            if (hit <= 0) continue;
            if (other.annotation) labels += hit; else tiles += hit * (other.weight ?? 1);
            mine += other.annotation ? hit * 4 : hit * (other.weight ?? 1);
          }
        }
        for (let j = 0; j < i; j += 1) {
          const hit = area(rect, placed[j].block);
          labels += hit;
          mine += hit * 4;
        }
        if (mine / own >= 0.02) dirty += 1;
        // A clean slot beside somebody else's arrow is not a solution. A rung a
        // reader would credit to the wrong hop is as wrong as one buried under
        // a tile, and counts the same, so a ladder that can only find clear air
        // by emigrating is recognised as stuck and falls back to bare numbered
        // callouts — which is what the Architecture Center draws for a bundle
        // of parallel flows anyway.
        //
        // Asked of every rung, not only the clean ones. As an `else if` a rung
        // that was BOTH buried and misattributed counted as dirt alone, and
        // dirt only mutes a fan of five or more — so a fan of three whose rungs
        // clipped a tile *and* read as a stranger's label scored `stray = 0`
        // and was drawn exactly where the reader misreads it.
        if (rivals.length > 0) {
          const mineSegs = placed[i].own;
          const cx = rect.x + rect.w / 2;
          // The centre alone is not the rung. A numbered callout hangs off the
          // bottom of the block, so a rung whose centre reads correctly can
          // still have its badge sitting on somebody else's arrow — and after a
          // fan mutes, the badge is the ONLY thing tying a sentence to a hop.
          const ys = [
            rect.y + rect.h / 2,
            rect.y + Math.min(0.08, rect.h / 2),
            rect.y + rect.h - Math.min(0.08, rect.h / 2),
          ];
          if (ys.some((cy) => gapToSegs(cx, cy, rivals) < gapToSegs(cx, cy, mineSegs) - CONFUSION_SLACK)) {
            stray += 1;
          }
        }
      }
      // The same price a single chip pays. A ladder moves as one object, so the
      // trip is charged once, not once per rung: priced per rung a deep fan
      // could not afford to step off a tile it was completely covering.
      const drift = DRIFT_COST_PER_IN * Math.hypot(shift, across);
      return {
        total: tiles + labels * ANNOTATION_WEIGHT + dirty * DIRTY_RUNG_COST + stray * STRAY_RUNG_COST + drift,
        onLabel: labels,
        overlap: tiles + labels * ANNOTATION_WEIGHT,
        dirty,
        stray,
      };
    };
    // Search along the ladder AND across it. A ladder is far taller than one
    // chip, so on a busy slide there is often no clear band anywhere along the
    // arrow and the only way out is to step the whole thing off the corridor.
    // Seeded at wherever the bundle already sits, so a re-run can only improve
    // on it and never trades an equal position for a different one.
    const alongStep = bundle.rung / 2;
    const acrossStep = (shape.alongX ? shape.w + 0.12 : shape.h + 0.05) / 2;
    // Far enough to cross the whole drawing. A fixed ring budget is a budget in
    // units of the object's own size, so a ladder of bare callouts — the very
    // case that has already failed to place at full size — could only ever
    // travel an inch and never reached the clear band two rows away.
    const rings = Math.max(
      6,
      Math.min(48, Math.ceil(Math.max(labelFrame.w / Math.max(acrossStep, 0.05), labelFrame.h / Math.max(alongStep, 0.05)))),
    );
    const sweep = (acrossLimit: number, shiftLimit: number): { shift: number; across: number; cost: number; onLabel: number; overlap: number; dirty: number; stray: number } => {
      let best = { shift: bundle.shift ?? 0, across: bundle.across ?? 0 };
      let at = cost(best.shift, best.across);
      // Coarse to fine. Scanning every lattice point out to the far side of the
      // page costs thousands of placements per bundle and is quadratic in the
      // reach; a stride-4 pass finds the clear REGION for a sixteenth of that,
      // and a full-resolution pass around it finds the point.
      const pass = (stride: number, centreShift: number, centreAcross: number, span: number): void => {
        const limit = Math.max(1, Math.ceil(span / stride));
        for (let ring = 1; at.total > 0 && ring <= limit; ring += 1) {
          for (let dx = -ring; dx <= ring && at.total > 0; dx += 1) {
            for (let dy = -ring; dy <= ring; dy += 1) {
              if (Math.max(Math.abs(dx), Math.abs(dy)) !== ring) continue;
              const shift = centreShift + dx * stride * alongStep;
              const across = centreAcross + dy * stride * acrossStep;
              if (Math.abs(across) > acrossLimit + 0.001) continue;
              if (Math.abs(shift) > shiftLimit + 0.001) continue;
              const c = cost(shift, across);
              if (c.total < at.total) {
                best = { shift, across };
                at = c;
              }
              if (at.total <= 0) break;
            }
          }
        }
      };
      pass(4, 0, 0, rings);
      if (at.total > 0) pass(1, best.shift, best.across, 4);
      return { ...best, cost: at.total, onLabel: at.onLabel, overlap: at.overlap, dirty: at.dirty, stray: at.stray };
    };
    // The ladder steps across its arrows freely — that is how it dodges the
    // tiles — but sliding it along them past the hop's own ends parks the whole
    // fan beside the two services instead of between them. The one thing worth
    // that is another label: two ladders on neighbouring rows have nowhere else
    // to go, and a buried label is not a label at all.
    let picked = sweep(Math.max((bundle.span ?? 0) / 2 + 2 * shape.w, 3 * acrossStep), Number.POSITIVE_INFINITY);
    if (picked.onLabel > 0) {
      const wider = sweep(Number.POSITIVE_INFINITY, Number.POSITIVE_INFINITY);
      if (wider.onLabel < picked.onLabel) picked = wider;
    }
    // The ring search moves in half-chip steps, so a ladder that comes to rest
    // a few hundredths of an inch over a tile has no lattice point left to shed
    // that last sliver. Refine on a quarter lattice around whatever it chose.
    for (const frac of [0.5, 0.25]) {
      if (picked.cost <= 0) break;
      for (let dx = -2; dx <= 2; dx += 1) {
        for (let dy = -2; dy <= 2; dy += 1) {
          if (dx === 0 && dy === 0) continue;
          const shift = picked.shift + dx * frac * alongStep;
          const across = picked.across + dy * frac * acrossStep;
          const c = cost(shift, across);
          if (c.total < picked.cost) {
            picked = {
              shift, across, cost: c.total, onLabel: c.onLabel, overlap: c.overlap, dirty: c.dirty, stray: c.stray,
            };
          }
        }
      }
    }
    bundle.shift = picked.shift;
    bundle.across = picked.across;
    bundle.dirty = picked.dirty;
    bundle.stray = picked.stray;
    // Covered area PLUS the rungs a reader would credit to the wrong arrow, and
    // never the drift tie-breaker (which is never zero, so a total would report
    // every bundle as unplaceable). Returning area alone made the confusion
    // machinery unreachable: a ladder that emigrated into perfectly EMPTY air
    // beside a foreign hop scored 0, so the caller skipped the retry and the
    // mute, and drew every rung beside the wrong arrow. That is the one case
    // the rule exists to catch.
    return picked.overlap + picked.dirty * DIRTY_RUNG_COST + picked.stray * STRAY_RUNG_COST;
  };

  // Which step numbers the workflow slide will actually narrate. Dropping a
  // chip is only ever a trade against that list; with no row to read it is a
  // deletion.

  for (const [key, bundle] of bundles) {
    const shape = shapes.get(key);
    if (!shape || bundle.rung <= 0) continue;
    const members = annotatedRoutes.filter((route) => bundleKey(route) === key);

    // A fan deeper than the frame cannot be laid out at the ordinary label size:
    // the rungs get clamped onto the page edge and restack. Shrink the whole
    // bundle just enough to fit instead, so every label survives intact, and if
    // even the smallest readable size will not do it, wrap into columns.
    const corridor = shape.alongX ? labelFrame.h : labelFrame.w;
    // Re-measure every member and take the max again, exactly as the seeding
    // loop does. Re-probing only the first route drops the rung below a taller
    // sibling's own step, and the bundle stops being a lattice.
    const remeasure = (): void => {
      let rung = 0;
      let wide = 0;
      let tall = 0;
      for (const member of annotatedRoutes) {
        if (bundleKey(member) !== key) continue;
        const again = connectorLabelBox(member, transform, labelFontSize, px, labelFrame, [], { ...bundle, rung: 0 });
        if (!again) continue;
        rung = Math.max(rung, again.alongX ? again.block.h + 0.05 : again.w + 0.12);
        wide = Math.max(wide, again.w);
        tall = Math.max(tall, again.block.h);
      }
      if (rung <= 0) return;
      bundle.rung = rung;
      shape.w = wide;
      shape.h = tall;
    };
    // How deep one rung is along the ladder's own axis: a horizontal arrow
    // stacks its chips vertically and needs their height, a vertical one stacks
    // them side by side and needs their width.
    const depth = (): number => (shape.alongX ? shape.h : shape.w);
    const fitCorridor = (): void => {
      for (let pass = 0; pass < 4; pass += 1) {
        const needed = (bundle.count - 1) * bundle.rung + depth();
        if (needed <= corridor) break;
        const font = Math.max(LEGIBLE_TILE_PT, (bundle.font ?? labelFontSize) * (corridor / needed));
        if (font >= (bundle.font ?? labelFontSize) - 0.01) break;
        bundle.font = font;
        remeasure();
      }
      if ((bundle.count - 1) * bundle.rung + depth() > corridor) {
        bundle.perCol = Math.max(1, Math.floor((corridor - depth()) / bundle.rung) + 1);
        remeasure();
      }
    };
    fitCorridor();

    // Score the bundle at its REAL geometry: probe the same routine that will
    // place it, with no obstacles so no walk runs. Modelling the lattice by
    // hand kept choosing offsets that the in-frame clamps then undid, so the
    // search believed it had found a clear slot that never materialised.
    const cost = searchBundle(key, bundle);
    if (cost <= 0) continue;

    // Nowhere clear at its natural width. Narrow the whole ladder to the lane
    // between the two services and look again: a fan wrapped onto more lines
    // fits the gap it belongs in, where at full width it can only stand on the
    // tiles either side of that gap.
    const wide = { rung: bundle.rung, font: bundle.font, perCol: bundle.perCol, shift: bundle.shift, across: bundle.across, w: shape.w, h: shape.h, dirty: bundle.dirty, stray: bundle.stray };
    bundle.maxWidth = Math.max(0.34 * px, ((bundle.span ?? 0) - 0.16) || 0.34 * px);
    bundle.shift = undefined;
    bundle.across = undefined;
    remeasure();
    fitCorridor();
    const narrowed = searchBundle(key, bundle);
    if (narrowed >= cost) {
      bundle.maxWidth = undefined;
      bundle.rung = wide.rung;
      bundle.font = wide.font;
      bundle.perCol = wide.perCol;
      bundle.shift = wide.shift;
      bundle.across = wide.across;
      shape.w = wide.w;
      shape.h = wide.h;
      bundle.dirty = wide.dirty;
      bundle.stray = wide.stray;
    }
    // Still nothing. A deep fan across a crowded drawing has no honest inline
    // position left: every slot it can reach is on top of a service or another
    // label. Fall back to what the Architecture Center itself does with a
    // bundle of parallel flows — number the arrows and let the workflow slide
    // carry the wording. Only for a fan that is both deep and badly stuck: a
    // rung clipping a tile corner is not worth losing every label for.
    //
    // And only when the wording genuinely survives. A route with no step
    // number gets no callout at all, and a step with no description gets no
    // workflow row, so dropping the chip in either case deletes the author's
    // text outright — which is exactly the fan a model produces when it emits
    // four duplicate connections between the same pair.
    const carried = members.every(
      (route) => route.stepNumber !== undefined && narratedSteps.has(route.stepNumber),
    );
    // Judged by how many rungs a reader would see as wrong, not by how much
    // area is covered: summed area has no scale a human recognises, and
    // measured against one chip a whole fan was erased for clipping a sixth of
    // one tile.
    const soiled = Math.max(2, Math.ceil(0.35 * bundle.count));
    // Two different reasons to stop being a ladder.
    //
    // Deep and clipped: a big fan that lands on tiles wherever it stands. The
    // depth gate is what keeps a shallow fan from losing every label because
    // one rung clips a corner.
    const clipped = bundle.count >= 5 && (bundle.dirty ?? 0) + (bundle.stray ?? 0) >= soiled;
    // Or misread anywhere it can reach. The search has already swept the whole
    // frame and this is the best it found, so a rung still credited to a
    // stranger's arrow is not a placement the ladder can improve on — it is
    // proof that no honest position exists for an object this shape. Depth is
    // beside the point: a fan of three that reads as belonging to the wrong hop
    // is wrong at every depth, and unlike a clip it cannot be walked away from.
    // Muting is not a loss here, because `carried` has already established that
    // every rung's sentence reaches the workflow slide.
    const misread = bundle.count >= 3 && (bundle.stray ?? 0) > 0;
    if (carried && (clipped || misread)) {
      bundle.badgesOnly = true;
      // A row that exists is still not a row that says anything. An author who
      // writes both a terse description ("Step 13") and a real label loses the
      // label the moment its chip is muted: the deck then reads "13. Step 13"
      // and the sentence the arrow carried is gone. Muting is the right trade
      // on a stuck fan, so it stands — but the wording it trades away is
      // handed to the row that is supposed to be carrying it.
      for (const route of members) {
        if (route.stepNumber === undefined || !route.label) continue;
        if (carriesWording(narratedRows.get(route.stepNumber) ?? '', route.label)) continue;
        if (!thumbnail) mutedWording.set(route.stepNumber, route.label);
      }
      // A ladder of bare callouts is a different object from the one the pitch
      // and the offsets were chosen for: at the wrapped-label pitch the badges
      // stay strung across the same rows the labels could not fit in. Measure
      // it at its real size and look for a slot again from scratch.
      bundle.perCol = undefined;
      bundle.shift = undefined;
      bundle.across = undefined;
      remeasure();
      // Wrap the callouts into a block roughly as long as the hop they belong
      // to. Left as one column a fan of ten runs the full height of the frame,
      // which on a grid means crossing every row above and below its own — the
      // ladder is compact but it is still strung across other people's tiles.
      //
      // Which shape fits is not something the hop's length alone can decide:
      // one column is too tall, and enough columns to match the hop are wider
      // than the gap between the two services. Try the handful of shapes a
      // block of this many callouts can take and keep whichever one the search
      // scores best — the same measurement that decides everything else.
      const lane = Math.max(bundle.span ?? 0, 2 * bundle.rung);
      const shapesToTry = new Set<number>([
        Math.max(1, Math.min(bundle.count, Math.floor(lane / Math.max(bundle.rung, 0.01)))),
      ]);
      for (let cols = 1; cols <= 6; cols += 1) shapesToTry.add(Math.max(1, Math.ceil(bundle.count / cols)));
      let bestScore = Number.POSITIVE_INFINITY;
      let bestShape: { perCol: number; shift?: number; across?: number; w: number; h: number; dirty?: number; stray?: number } | null = null;
      for (const perCol of shapesToTry) {
        bundle.perCol = perCol;
        bundle.shift = undefined;
        bundle.across = undefined;
        remeasure();
        fitCorridor();
        const score = searchBundle(key, bundle);
        if (score < bestScore) {
          bestScore = score;
          bestShape = {
            perCol, shift: bundle.shift, across: bundle.across, w: shape.w, h: shape.h, dirty: bundle.dirty, stray: bundle.stray,
          };
        }
        if (bestScore <= 0) break;
      }
      if (bestShape) {
        bundle.perCol = bestShape.perCol;
        remeasure();
        fitCorridor();
        bundle.shift = bestShape.shift;
        bundle.across = bestShape.across;
        bundle.dirty = bestShape.dirty;
        bundle.stray = bestShape.stray;
        shape.w = bestShape.w;
        shape.h = bestShape.h;
      }
    }
  }
  settle(ordered.filter((route) => bundles.has(bundleKey(route))));
  // Repair pass. A ladder is rigid — it has one offset for every rung — so where
  // it still had to come down on an ordinary chip, move that chip instead: a
  // single connector label can walk in any direction and usually has a free slot
  // within a step or two. Repeat, because moving one chip can hand its old
  // problem to the next.
  const overlaps = (
    a: { x: number; y: number; w: number; h: number },
    b: { x: number; y: number; w: number; h: number },
  ): boolean => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) > 0.01
    && Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > 0.01;
  for (let round = 0; round < 4; round += 1) {
    let repaired = 0;
    for (const route of ordered) {
      if (bundles.has(bundleKey(route))) continue;
      const box = chips.get(route.id);
      if (!box) continue;
      const others = chipObstacles.slice(tileObstacleCount).filter((taken) => taken !== box.block);
      if (!others.some((taken) => overlaps(box.block, taken))) continue;
      const moved = connectorLabelBox(
        route, transform, labelFontSize, px, labelFrame,
        chipObstacles.filter((taken) => taken !== box.block),
        undefined, undefined, foreignGapFor(route),
      );
      if (!moved || (moved.block.x === box.block.x && moved.block.y === box.block.y)) continue;
      // A chip only reads as belonging to the arrow under it. Refuse a repair
      // that carries it further than its own size from where it started: the
      // move is not worth an unreadable diagram, and a numbered callout landed
      // beside a different arrow entirely.
      const reach = 1.5 * Math.max(box.w, box.block.h);
      if (Math.hypot(moved.block.x - box.block.x, moved.block.y - box.block.y) > reach) continue;
      // The repair has to answer the same question the first seat did. It used
      // to ask only whether the new slot hit another chip, so a chip refused
      // for standing on a caption could be moved onto a different caption and
      // kept — the refusal was undone by the fix for a different defect.
      //
      // And it asked it only of narrated routes, so an un-numbered chip could
      // be moved onto a service tile and kept while the identical numbered one
      // beside it was refused. `settle` has no such exemption; neither should
      // its repair. The seat this move came from was already cleared by
      // `settle`, so refusing the move is always the safe answer.
      if (chipSpoils(moved.block, { sourceId: route.sourceId, targetId: route.targetId, bundleKey: bundleKey(route) },
        chipObstacles.filter((taken) => taken !== box.block))) {
        continue;
      }
      const slot = chipObstacles.indexOf(box.block);
      if (slot >= 0) chipObstacles[slot] = moved.block;
      chips.set(route.id, moved);
      badges.set(route.id, stepBadgeBox(
        route, transform, px, labelFrame, moved, chipObstacles,
        ownGapFor(route), foreignGapFor(route), thumbnail,
      ));
      repaired += 1;
    }
    // A second ladder on the same slide is placed against the first, but the
    // first was frozen before the second existed. Re-run its search with its own
    // rungs lifted out of the way; seeded at where it already sits, so it only
    // ever moves to a strictly better slot.
    for (const [key, bundle] of bundles) {
      const members = annotatedRoutes.filter((route) => bundleKey(route) === key);
      const own = new Set<Block>();
      for (const route of members) {
        const block = chips.get(route.id)?.block;
        if (block) own.add(block);
      }
      if (own.size === 0) continue;
      const others = chipObstacles.slice(tileObstacleCount).filter((taken) => !own.has(taken));
      if (![...own].some((block) => others.some((taken) => overlaps(block, taken)))) continue;
      const before = { shift: bundle.shift ?? 0, across: bundle.across ?? 0 };
      searchBundle(key, bundle, own);
      if (bundle.shift === before.shift && bundle.across === before.across) continue;
      const pool = chipObstacles.filter((taken) => !own.has(taken));
      for (const route of members) {
        const old = chips.get(route.id);
        let moved = connectorLabelBox(
          route, transform, labelFontSize, px, labelFrame, pool, bundle,
          undefined, foreignGapFor(route),
        );
        // The ladder repair re-seats every rung in the bundle, so each new seat
        // is a fresh placement and owes the same answer as a first one. A rung
        // that now stands on a name is handed to the step list rather than
        // drawn, which is the trade the whole chip mechanism is built on.
        //
        // Every rung, not just the narrated ones: a ladder is re-seated as a
        // unit, so exempting its un-numbered members left them standing on the
        // tiles the numbered ones had just been moved off. An un-numbered rung
        // takes the same promotion `settle` gives it — a number, a callout, and
        // its sentence on the workflow slide.
        if (moved && chipSpoils(
          { ...moved.block }, { sourceId: route.sourceId, targetId: route.targetId, bundleKey: key }, pool,
        )) {
          if (route.stepNumber !== undefined && narratedSteps.has(route.stepNumber)) {
            if (!thumbnail && route.label
              && !carriesWording(narratedRows.get(route.stepNumber) ?? '', route.label)) {
              mutedWording.set(route.stepNumber, route.label);
            }
            moved = null;
          } else if (route.stepNumber === undefined && !thumbnail) {
            const step = promotable.get(route.id);
            if (step !== undefined) {
              route.stepNumber = step;
              promotedSteps.set(step, route.label);
              moved = null;
            }
          } else {
            moved = null;
          }
        }
        chips.set(route.id, moved);
        badges.set(route.id, stepBadgeBox(
          route, transform, px, labelFrame, moved, pool,
          ownGapFor(route), foreignGapFor(route), thumbnail,
        ));
        const slot = old?.block ? chipObstacles.indexOf(old.block) : -1;
        if (moved) {
          pool.push(moved.block);
          if (slot >= 0) chipObstacles[slot] = moved.block;
          else chipObstacles.push(moved.block);
        } else if (slot >= 0) {
          chipObstacles.splice(slot, 1);
        }
      }
      repaired += 1;
    }
    if (repaired === 0) break;
  }
  // Draw at the size the box was measured at. A shrunk bundle written back out
  // at the ordinary size spills its text past its own chip and over its own
  // numbered callout, because the box is smaller than the text inside it.
  //
  // On the thumbnail, anything that has been squeezed below the resolvable
  // floor is dropped rather than drawn as grey noise: the same annotation is
  // legible on the detail slide that follows, and on the workflow list.
  const drawable = (size: number): boolean => !thumbnail || size >= OVERVIEW_LEGIBLE_PT;
  for (const route of annotatedRoutes) {
    const box = chips.get(route.id) ?? null;
    const size = box?.fontSize ?? labelFontSize;
    if (!drawable(size)) continue;
    addConnectorLabel(slide, route, size, box);
  }
  for (const route of annotatedRoutes) {
    const size = chips.get(route.id)?.fontSize ?? labelFontSize;
    if (!drawable(size)) continue;
    addStepBadge(slide, route, size, badges.get(route.id) ?? null);
  }

  // Colour key so the deck's connectors agree with the PNG legend. Drawn in the
  // strip reserved for it below the diagram, not over the drawing.
  addConnectionLegend(pptx, slide, diagram.edges ?? [], fullFrame);

  return true;
}

/**
 * Build the single-slide presentation in memory.
 *
 * Exposed separately from the download helper so tests (and future callers such
 * as server-side rendering) can inspect the generated package.
 */
export async function buildDiagramSlidePptx(
  imageDataUrl: string,
  options: PptxExportOptions,
): Promise<PptxGenJS> {
  const { diagramName: rawDiagramName, author: rawAuthor, date: rawDate, isDarkMode } = options;
  // The header triple is drawn on every slide and is free text the user
  // typed, so it goes through the same single-line composition as every other
  // drawn string. Without it an NFD name measures wider than the NFC one that
  // means the same thing, and the header fitter shrinks the two differently.
  const diagramName = singleLineName(rawDiagramName);
  const author = singleLineName(rawAuthor);
  const date = singleLineName(rawDate);
  const t = isDarkMode ? DARK_THEME : LIGHT_THEME;
  // Number the callouts before anything measures them, so the drawing, the
  // badges and the workflow list are all built from the same edges.
  const diagram = options.diagram
    ? { ...options.diagram, edges: narrateEdgeCallouts(options.diagram.edges ?? []) }
    : options.diagram;

  const pptx = newDeck();
  const geom = planSlideGeometry(diagram);
  if (geom.w > BASE_W + 0.001 || geom.h > BASE_H + 0.001) {
    // A custom page keeps every shape at its true size instead of squeezing a
    // wide architecture until the labels break apart.
    pptx.defineLayout({ name: 'AZD_FIT', width: geom.w, height: geom.h });
    pptx.layout = 'AZD_FIT';
  } else {
    pptx.layout = 'LAYOUT_WIDE';
  }
  const W = geom.w;
  const FOOTER_Y = geom.footerY;
  pptx.author = author;
  pptx.title = diagramName;
  pptx.subject = 'Microsoft Product Architecture Diagram';
  pptx.company = 'Swarm Data SE, Jiayi Yang';

  // A tiled deck opens with the whole drawing, deliberately below the
  // legibility floor, so the reader sees the architecture before panning
  // through the readable parts of it. This is how the Azure Architecture
  // Center presents one: an overview, then the numbered workflow.
  const parts = geom.windows;
  const windows: (DiagramWindow | undefined)[] = parts.length > 0 ? [undefined, ...parts] : [undefined];
  let renderedNatively = false;
  // Wording that a muted chip handed to the workflow slide, filled in by the
  // diagram pass and read when that slide is written.
  const mutedWording = new Map<number, string>();
  // Wording promoted into a numbered step because its chip had nowhere to
  // stand and its edge carried no step of its own. Unlike a muted chip, there
  // is no authored row waiting for it, so the workflow list has to grow one.
  const promotedSteps = new Map<number, string>();
  // Names the tiles had to cut, for the index slide at the end of the deck.
  //
  // A MAP, because the index has to print the pair. A shortened name stops
  // being a name and becomes a lookup key, and a key the index never defines is
  // not a key at all - the reader sees a tile marked "3" and an index of
  // sentences, none of which contains a 3. Keyed by the AUTHORED name so a
  // service drawn on the overview and again on its own window slide occupies
  // one row, not two.
  const truncatedNames = new Map<string, Set<string>>();
  // The mark each service draws, shared by every slide in the deck.
  //
  // Per-slide, this counter restarted: two window slides issued the same four
  // keys to eight different services, and a service was "5" on the overview and
  // "1" on its own slide. A mark has to mean one thing in one file, so both the
  // collision test and the key it falls back to are deck-global.
  const drawnHere = new Map<string, string>();
  // Assigned once, from a stable sort of the node ids, so adding a node does
  // not renumber the ones already there. Digits are the narrowest glyphs in the
  // model - "123" is 0.1572in against 0.2379in for a two-letter stub - so a
  // stable key still fits every column a positional one did.
  const keyOrdinal = new Map<string, number>();
  [...(diagram?.nodes ?? [])]
    .map((node) => String(node.id))
    .sort((a, b) => a.localeCompare(b))
    .forEach((nodeId, i) => keyOrdinal.set(nodeId, i + 1));

  for (const [index, window] of windows.entries()) {
    const slide = pptx.addSlide();
    slide.background = { color: t.bg };
    const partOf = parts.length === 0
      ? ''
      : index === 0
        ? '  (Overview)'
        : `  (${index} / ${parts.length})`;
    // Quoted from the heading the reader is looking at, not re-derived, so the
    // index cannot name a sheet the deck does not print.
    const slideLabel = partOf ? `slide "${partOf.trim().replace(/^\(|\)$/g, '')}"` : '';

    // ── Top accent bar (Azure blue) ───────────────────────────────────────────
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: 0, w: W, h: ACCENT_H,
      fill: { color: t.accent },
      line: { color: t.accent, width: 0 },
    });

    // ── Header strip ──────────────────────────────────────────────────────────
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: ACCENT_H, w: W, h: HEADER_H,
      fill: { color: t.headerBg },
      line: { color: t.headerBg, width: 0 },
    });

    // ── Diagram title ─────────────────────────────────────────────────────────
    const titleW = Math.max(3, W - 3.85);
    const head = fitHeadingToBox(`${diagramName}${partOf}`, titleW, HEADER_H - 0.1, 24);
    slide.addText(head.body, {
      x: 0.35, y: ACCENT_H + 0.05, w: titleW, h: HEADER_H - 0.1,
      fontSize: head.fontSize,
      bold: true,
      color: t.titleText,
      fontFace: 'Yu Gothic UI',
      valign: 'middle',
      wrap: true,
    });

    // ── Author + date (right side of header) ──────────────────────────────────
    slide.addText(`${author}  ·  ${date}`, {
      x: W - 3.43, y: ACCENT_H + 0.05, w: 3.08, h: HEADER_H - 0.1,
      fontSize: 10,
      color: t.metaText,
      fontFace: 'Yu Gothic UI',
      align: 'right',
      valign: 'middle',
    });

    // ── Thin separator between header and image ───────────────────────────────
    slide.addShape(pptx.ShapeType.rect, {
      x: 0, y: HEADER_END, w: W, h: SEP_H,
      fill: { color: t.accent },
      line: { color: t.accent, width: 0 },
    });

    // ── Diagram body — native shapes when available, captured PNG otherwise ───
    renderedNatively = diagram
      ? await addEditableDiagram(pptx, slide, diagram, geom.frame, isDarkMode, window, mutedWording, truncatedNames, window === undefined && parts.length > 0, options.presetIcons, promotedSteps, drawnHere, keyOrdinal, slideLabel)
      : false;

    if (!renderedNatively) {
      slide.addImage({
        data: imageDataUrl,
        x: geom.frame.x,
        y: geom.frame.y,
        w: geom.frame.w,
        h: geom.frame.h,
        sizing: { type: 'contain', w: geom.frame.w, h: geom.frame.h },
      });
    }

    // ── Footer text ───────────────────────────────────────────────────────────
    const note = parts.length > 0
      ? index === 0
        ? `The whole architecture, shown small enough to fit one slide. The next ${parts.length} slides repeat it at a readable size, in reading order.`
        : `This architecture needs more than one readable slide, so it continues across ${parts.length} of them — this is part ${index}. Export to Visio (.vsdx) for the whole drawing on a single sheet.`
      : geom.overflow
        ? 'This architecture is wider than PowerPoint\'s 56" page limit, so it was scaled down to fit. Export to Visio (.vsdx) for a full-size, fully legible drawing.'
        : geom.outliersClamped
          ? 'One or more services sat far outside the main layout. They were moved to the page edge so they remain visible — reposition them on the canvas for an exact layout.'
          : '';
    if (note && renderedNatively) {
      slide.addText(note, {
        objectName: 'overflow-note',
        x: 0.35, y: FOOTER_Y - 0.26, w: W - 0.7, h: 0.24,
        fontSize: 9,
        bold: true,
        color: stripHash(readableTextOn('#B45309', `#${stripHash(t.bg)}`)),
        fontFace: 'Yu Gothic UI',
        valign: 'middle',
      });
    }
    slide.addText('Generated by Microsoft Product Architecture Diagram Builder  ·  Swarm Data SE, Jiayi Yang', {
      x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H,
      fontSize: 8,
      color: t.footerText,
      fontFace: 'Yu Gothic UI',
      valign: 'middle',
    });
  }

  // A numbered callout means nothing without the sentence it points at, so the
  // Azure Architecture Center always pairs the badges with a numbered list.
  const workflow = workflowListFromEdges(diagram?.edges ?? []).map((row) => {
    // A chip that was muted traded its wording for this row, so the row has to
    // say it. The label goes in the parenthesis the Architecture Center uses
    // for the mechanism on a numbered flow.
    const handed = mutedWording.get(row.step);
    return handed ? { ...row, description: `${row.description}（${handed}）` } : row;
  });
  // A promoted chip has a badge on the drawing and, until now, nothing to look
  // it up in: `workflowListFromEdges` builds rows only from edges the author
  // numbered, and a promoted edge is by definition one they did not. Append
  // the rows it is owed and keep the list in step order, so the badges read
  // down the page the way the numbers do.
  for (const [step, text] of promotedSteps) {
    if (workflow.some((row) => row.step === step)) continue;
    workflow.push({ step, description: text });
  }
  workflow.sort((a, b) => a.step - b.step);
  if (workflow.length > 0) {
    // Rows stop shrinking at a legible minimum, so a long workflow continues on
    // another slide. Dropping the tail would leave badges on the drawing whose
    // sentence appears nowhere in the deck.
    const listTop = IMAGE_Y + 0.1;
    const available = Math.max(MIN_WORKFLOW_ROW_IN, geom.footerY - 0.1 - listTop);
    // The sentence column, measured against the widest the badge is ever
    // allowed to be so the estimate is never optimistic.
    const rowTextW = Math.max(1, W - (0.42 + 0.34 + 0.16) - 0.42 - 0.2);
    const rowHeightIn = (text: string, pt: number): number => {
      // 1.35, matching every other line-height in this file: `Yu Gothic UI`'s
      // own hhea and OS/2 win metrics give 1.3301, and the same face carries
      // the Latin and the CJK. 1.25 was 6% optimistic, which the row slack
      // absorbed only up to about six lines. And measure the wrap the way a
      // renderer wraps it — the ratio is a lower bound for text that breaks
      // between words. The 0.2in off the column and 0.1in on the row are the
      // text box's own insets, which come out of the room the words get.
      const lines = wrappedLineCount(text, rowTextW, pt);
      return lines * pt * 1.35 / 72 + 0.1;
    };
    // Paginating on a flat 0.34in assumed every step was one line. Real
    // Architecture-Center prose wraps to two or three, and PowerPoint does not
    // clip a `valign: middle` box — it spills symmetrically — so a wrapped row
    // ran into the rows above and below it. Give each slide as many rows as
    // actually fit once the longest sentence is allowed to wrap at the smallest
    // size still worth reading.
    const neededRow = Math.max(
      MIN_WORKFLOW_ROW_IN,
      ...workflow.map(entry => rowHeightIn(entry.description, WORKFLOW_MIN_PT) + 0.06),
    );
    const perSlide = Math.max(1, Math.floor(available / neededRow));
    const parts = Math.ceil(workflow.length / perSlide);

    for (let part = 0; part < parts; part += 1) {
      const rows = workflow.slice(part * perSlide, (part + 1) * perSlide);
      const slide = pptx.addSlide();
      slide.background = { color: t.bg };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: W, h: ACCENT_H,
        fill: { color: t.accent }, line: { color: t.accent, width: 0 },
      });
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: ACCENT_H, w: W, h: HEADER_H,
        fill: { color: t.headerBg }, line: { color: t.headerBg, width: 0 },
      });
      slide.addText(parts > 1 ? `Workflow (${part + 1} / ${parts})` : 'Workflow', {
        objectName: 'workflow-heading',
        x: 0.35, y: ACCENT_H + 0.05, w: Math.max(3, W - 3.85), h: HEADER_H - 0.1,
        fontSize: 24, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'middle',
      });
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: HEADER_END, w: W, h: SEP_H,
        fill: { color: t.accent }, line: { color: t.accent, width: 0 },
      });

      // The badge colour and shape repeat here so a reader can match a number on
      // the drawing to its row without hunting.
      // The 0.62in cap keeps a short list from turning into widely-spaced
      // bullets — but pagination has just reserved `neededRow` per row for the
      // longest sentence at the 9pt floor, and capping below that threw the
      // reservation away and printed the text outside its own box. Never
      // shrink a row below what the pagination promised it.
      const rowGap = Math.max(
        neededRow,
        Math.min(0.62, Math.max(MIN_WORKFLOW_ROW_IN, available / rows.length)),
      );
      const badge = Math.min(0.34, rowGap - 0.06);
      // Pagination already reserved room for the longest sentence at the floor
      // size; this hands every shorter row the largest size that still fits its
      // own box, so only the sentences that need to shrink do.
      const rowFontPt = (text: string, boxH: number): number => {
        for (let pt = WORKFLOW_ROW_PT; pt > WORKFLOW_MIN_PT; pt -= 0.5) {
          if (rowHeightIn(text, pt) <= boxH) return pt;
        }
        return WORKFLOW_MIN_PT;
      };
      rows.forEach((entry, index) => {
        const y = listTop + index * rowGap;
        slide.addText(String(entry.step), {
          objectName: `workflow-step-${entry.step}`,
          shape: pptx.ShapeType.ellipse,
          x: 0.42, y, w: badge, h: badge,
          fill: { color: t.accent },
          line: { color: 'FFFFFF', width: 1 },
          fontSize: Math.max(8, badge * 26),
          bold: true, color: 'FFFFFF', fontFace: 'Yu Gothic UI',
          align: 'center', valign: 'middle', margin: 0,
        });
        slide.addText(entry.description, {
          objectName: `workflow-text-${entry.step}`,
          x: 0.42 + badge + 0.16, y, w: W - (0.42 + badge + 0.16) - 0.42, h: rowGap - 0.04,
          fontSize: rowFontPt(entry.description, rowGap - 0.04),
          color: t.titleText, fontFace: 'Yu Gothic UI',
          valign: 'middle', wrap: true,
        });
      });
      slide.addText('Generated by Microsoft Product Architecture Diagram Builder  ·  Swarm Data SE, Jiayi Yang', {
        x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H,
        fontSize: 8, color: t.footerText, fontFace: 'Yu Gothic UI', valign: 'middle',
      });
    }
  }

  // Every name the drawing had to cut, spelled out.
  //
  // A tile too small for its name is a fact of any large estate — the drawing
  // is a map, and a map abbreviates. Throwing the name away is not. Without
  // this slide the reader of a 400-service deck meets rows of
  // "Azure Kubernetes Ser…" and has nowhere to go, which is precisely the
  // "export it and then retype it by hand" outcome this exporter exists to
  // avoid. Columns are sized from the longest name so nothing is clipped twice.
  if (truncatedNames.size > 0) {
    // THE PAIR, in the same format the Visio index uses: "<mark>  =  <name>".
    // A tile drew a mark and the reader has nothing but that mark to look the
    // service up by, so a row printing the name alone defines nothing. An empty
    // mark - a name the slide could not draw at all - reads "(not drawn)".
    // ALL of a service's marks are listed, because the overview and the reading
    // slide shorten one name to two different widths; a row quoting one of them
    // leaves the other drawn on a tile and defined nowhere. Separated by
    // "  |  " rather than by a comma: a cut name routinely contains a comma
    // ("... (Production, Zone Redundant)"), and splitting a row on one shredded
    // the mark into fragments that matched no tile. Two spaces on each side,
    // which no drawn string can contain - singleLineName collapses every run of
    // spaces before a label is ever measured - so this separator is unambiguous
    // for exactly the reason "  =  " is.
    const names = [...truncatedNames]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([authored, marks]) => {
        const printed = [...marks]
          // "(drawn unlabelled)", not "(not drawn)", because the tile IS on the
          // canvas. Measured across the corpus, 16 services on 7 drawings emit
          // a service shape and no caption, and the index told the reader each
          // one was absent while they were looking straight at it - on
          // probe-refused-raise, 4 of 6 services, every one of them cited by a
          // numbered step. A row that denies a drawn shape is worse than no row.
          .map((m) => (m === '' ? UNLABELLED_ROW : m))
          .sort((a, b) => a.localeCompare(b))
          .join('  |  ');
        return `${printed}  =  ${authored}`;
      });
    const listTop = IMAGE_Y + 0.1;
    // The note that heads the list is drawn 0.24in below `listTop` and the
    // first row starts under it, so the rows have 0.24in less room than the
    // page does. Fitting the grid to the page and then placing it 0.24in lower
    // pushed the last row of every column into the footer band: 0.0403in in
    // seven shipped scenarios, and up to 0.1309in - two thirds of the row -
    // once the fit loop made the pitch a variable.
    const INDEX_NOTE_H = 0.24;
    const available = Math.max(0.3, geom.footerY - 0.1 - listTop - INDEX_NOTE_H);
    // THE INDEX IS THE ONE PAGE A READER OPENS WHEN A MARK MEANS NOTHING TO
    // THEM, so a row that runs off the sheet defeats the entire page. At a
    // fixed 10pt with `wrap: false`, PowerPoint neither clips the row nor
    // complains: it simply paints past the slide edge and the characters are
    // gone. A 130-character pair drew 0.140in past the edge and a 175-character
    // pair drew 3.127in past it, losing about 45 characters, and the corpus was
    // already inside 0.040in of the margin - so this was one long service name
    // away from shipping. The type shrinks to fit its column first, and what
    // still will not fit at the floor WRAPS and takes the height it needs,
    // which is the same remedy the Visio index was given.
    const INDEX_MAX_PT = 10;
    const INDEX_MIN_PT = 7;
    let indexPt = INDEX_MAX_PT;
    let cols = 1;
    let colW = W - 0.7;
    // A SHRINK IS ONLY WORTH WHAT IT BUYS. The loop stepped the type down half
    // a point at a time until the widest name fitted its column on one line,
    // but it also stopped at the 7pt floor - and a name too long for a
    // full-width column at 10pt is still too long at 7pt. `probe-overlong-index`
    // walked all the way down and wrapped anyway: the reader was handed 7pt
    // type AND the wrapping the shrink existed to prevent. Three points of
    // legibility spent on nothing, on the one page a reader opens precisely
    // because a mark on the drawing meant nothing to them.
    //
    // So the size is chosen rather than walked into. Take the largest size that
    // actually fits; if none does, keep the largest size and let the rows wrap.
    // Where the shrink pays - `probe-shrinkable-index` - the answer is
    // unchanged, because the first size that fits is the one the walk reached.
    const indexFitAt = (pt: number) => {
      const widest = Math.max(...names.map((n) => estimateTextWidthIn(n, pt)));
      const rowsPer = Math.max(1, Math.floor(available / (pt * 1.45 / 72)));
      const at = Math.max(1, Math.min(
        Math.floor((W - 0.7) / (widest + 0.25)),
        Math.ceil(names.length / rowsPer),
      ));
      const width = (W - 0.7) / at;
      return { cols: at, colW: width, fits: widest <= width - 0.15 };
    };
    let indexFit = indexFitAt(INDEX_MAX_PT);
    for (let pt = INDEX_MAX_PT; pt >= INDEX_MIN_PT - 1e-9; pt -= 0.5) {
      const at = indexFitAt(pt);
      if (!at.fits) continue;
      indexPt = pt;
      indexFit = at;
      break;
    }
    cols = indexFit.cols;
    colW = indexFit.colW;
    const lineH = indexPt * 1.45 / 72;
    const indexTextW = Math.max(0.05, colW - 0.15);
    // EACH ROW TAKES ITS OWN HEIGHT, and the rows are packed by accumulated
    // height rather than laid on a fixed grid. Giving every row the tallest
    // row's pitch made one long name re-pitch the whole page: 44 rows of
    // 0.1410in ink were each given 0.2819in, which stranded half the sheet and
    // cost a whole extra index page. It also put the column count and the row
    // count permanently out of step - the columns were chosen assuming
    // single-line rows and the rows counted at the grown pitch, so a page held
    // a fraction of what it was sized for.
    const indexRows = names.map((text) => {
      const lines = wrappedLineCount(text, indexTextW, indexPt);
      return { text, lines, h: lineH * lines };
    });
    type PlacedRow = { text: string; lines: number; h: number; col: number; top: number };
    const indexPages: PlacedRow[][] = [];
    {
      let current: PlacedRow[] = [];
      let col = 0;
      let cursor = 0;
      for (const row of indexRows) {
        if (cursor > 0 && cursor + row.h > available + 1e-6) {
          col += 1;
          cursor = 0;
        }
        if (col >= cols) {
          indexPages.push(current);
          current = [];
          col = 0;
          cursor = 0;
        }
        current.push({ ...row, col, top: cursor });
        cursor += row.h;
      }
      if (current.length > 0) indexPages.push(current);
    }
    const pages = indexPages.length;

    let indexOrdinal = 0;
    for (let page = 0; page < pages; page += 1) {
      const slice = indexPages[page];
      const slide = pptx.addSlide();
      slide.background = { color: t.bg };
      addChrome(
        pptx, slide, t,
        pages > 1 ? `Service names (${page + 1} / ${pages})` : 'Service names',
        `${author}  ·  ${date}`,
      );
      slide.addText(
        'Names shortened on the drawing, in full.',
        {
          objectName: 'index-note',
          x: 0.35, y: HEADER_END + SEP_H + 0.04, w: W - 0.7, h: 0.22,
          fontSize: 9, color: t.metaText, fontFace: 'Yu Gothic UI', valign: 'middle',
        },
      );
      slice.forEach((row) => {
        slide.addText(row.text, {
          objectName: `index-name-${indexOrdinal}`,
          x: 0.35 + row.col * colW,
          y: listTop + INDEX_NOTE_H + row.top,
          w: indexTextW,
          h: row.h,
          fontSize: indexPt,
          color: t.titleText,
          fontFace: 'Yu Gothic UI',
          valign: 'middle',
          wrap: row.lines > 1,
        });
        indexOrdinal += 1;
      });
      slide.addText('Generated by Microsoft Product Architecture Diagram Builder  ·  Swarm Data SE, Jiayi Yang', {
        x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H,
        fontSize: 8, color: t.footerText, fontFace: 'Yu Gothic UI', valign: 'middle',
      });
    }
  }

  return pptx;
}

/**
 * Write the deck out after repairing it into shapes PowerPoint treats as its
 * own — real connectors glued to the services they join, service names living
 * inside their tiles, and each tile grouped with its icon.
 *
 * pptxgenjs cannot emit any of that, so it is done on the finished package.
 * If anything goes wrong the untouched deck is still downloaded: a deck that
 * is harder to edit is very much better than no deck at all.
 */
async function downloadNativePptx(pptx: PptxGenJS, fileName: string): Promise<void> {
  try {
    const blob = (await pptx.write({ outputType: 'blob' })) as Blob;
    const zip = await nativizePackage(await JSZip.loadAsync(blob));
    const repaired = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      // pptxgenjs writes the package uncompressed. Once this path owns the
      // write it may as well deflate it: measured ~9x smaller (425KB -> 46KB),
      // which is the difference between a deck that mails and one that bounces.
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
    const url = URL.createObjectURL(repaired);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.warn('PowerPoint shape conversion failed; exporting the unconverted deck.', error);
    await pptx.writeFile({ fileName });
  }
}

/**
 * Build and download a single-slide PPTX for the diagram.
 *
 * The diagram is drawn with native PowerPoint shapes whenever canvas contents
 * are supplied, so the recipient can move, restyle, and relabel every service
 * directly in PowerPoint. `imageDataUrl` is the fallback for empty canvases.
 * Returns the generated filename.
 */
export async function exportDiagramAsPptx(
  imageDataUrl: string,
  options: PptxExportOptions,
): Promise<string> {
  const pptx = await buildDiagramSlidePptx(imageDataUrl, options);
  const fileName = generateModelFilename('azure-diagram-slide', 'pptx');
  await downloadNativePptx(pptx, fileName);
  return fileName;
}

// ═══════════════════════════════════════════════════════════════════════════
//  Customer-ready deck (multi-slide)
// ═══════════════════════════════════════════════════════════════════════════

// PptxGenJS's runtime shapes are richer than its exported types in a few spots
// (e.g. table cell options); a narrow local alias keeps the call-sites clean.
type Slide = ReturnType<PptxGenJS['addSlide']>;

export interface DeckService {
  name: string;
  category?: string;
  group?: string;
}

export interface DeckPillar {
  pillar: string;
  score: number;
  /** Maturity label, e.g. "Adequate, with gaps" (falls back to a score band). */
  maturity?: string;
}

export interface DeckFinding {
  severity: 'critical' | 'high' | 'medium' | 'low' | string;
  category: string;
  issue: string;
  recommendation?: string;
}

export interface DeckValidation {
  overallScore: number;
  /** Overall maturity label, e.g. "Adequate, with gaps". */
  overallLabel?: string;
  summary?: string;
  pillars: DeckPillar[];
  findings: DeckFinding[];
  modelUsed?: string;
}

export interface DeckRegionCost {
  name: string;
  flag?: string;
  monthly: number;
  annual: number;
  isCurrent?: boolean;
  isCheapest?: boolean;
}

export interface DeckCost {
  totalMonthly: number;
  annual?: number;
  currency: string;
  term?: string;
  region?: string;
  pricesAsOf?: string;
  /**
   * The oldest still-unchanged price behind this estimate, when it has held for
   * over a year. Not a staleness warning — these are current Azure prices — but
   * a customer asking "how firm is this number?" is asking a fair question and
   * the deck should answer it.
   */
  oldestMeterAsOf?: string;
  /** Fixed (predictable) vs usage-based split, when derivable. */
  fixedCost?: number;
  usageCost?: number;
  byCategory: { category: string; cost: number; percentage: number }[];
  topServices: { serviceName: string; cost: number; tier?: string; percentage?: number }[];
  /** Multi-region comparison (sorted cheapest-first), when computed. */
  regions?: DeckRegionCost[];
  /** True when one or more regions could not preserve every selected SKU. */
  regionComparisonIncomplete?: boolean;
  /** Regions omitted from the like-for-like comparison, with the reason. */
  unavailableRegions?: string[];
}

export interface DeckWorkflowStep {
  /** 1-based step number, matching the numbered callout drawn on the arrow. */
  step: number;
  description: string;
  /** Services the step touches, in flow order. */
  services?: string[];
}

export interface ArchitectureDeckOptions extends PptxExportOptions {
  /** The original natural-language prompt ("the napkin"). */
  prompt?: string;
  /** Model that generated the architecture. */
  model?: string;
  /** Flat service inventory derived from the diagram nodes. */
  services: DeckService[];
  /**
   * Numbered dataflow. The Azure Architecture Center pairs every numbered
   * arrow with a matching numbered paragraph; without this the callouts on the
   * diagram slide refer to nothing. A slide is added only when present.
   */
  workflow?: DeckWorkflowStep[] | null;
  /** Optional WAF validation summary — a slide is added only when present. */
  validation?: DeckValidation | null;
  /** Optional cost estimate — a slide is added only when present. */
  cost?: DeckCost | null;
}

const BODY_TOP = HEADER_END + SEP_H + 0.18; // ~1.12"
const BODY_H = FOOTER_Y - BODY_TOP - 0.1;

function scoreColor(s: number): string {
  return s >= 80 ? '16a34a' : s >= 60 ? 'f59e0b' : 'dc2626';
}
function scoreBand(s: number): string {
  return s >= 80 ? 'Well-Architected' : s >= 60 ? 'Needs attention' : 'At risk';
}
function severityColor(sev: string): string {
  const s = sev.toLowerCase();
  return s === 'critical' || s === 'high' ? 'dc2626' : s === 'medium' ? 'f59e0b' : '0078d4';
}
function money(n: number, currency: string): string {
  const sym = currency === 'USD' ? '$' : '';
  return `${sym}${n.toLocaleString(undefined, { maximumFractionDigits: 0 })}`;
}

/** Shared header bar + separator + footer chrome for a content slide. */
function addChrome(pptx: PptxGenJS, slide: Slide, t: SlideTheme, title: string, meta?: string): void {
  slide.background = { color: t.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: ACCENT_H, fill: { color: t.accent }, line: { color: t.accent, width: 0 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: ACCENT_H, w: W, h: HEADER_H, fill: { color: t.headerBg }, line: { color: t.headerBg, width: 0 } });
  const head = fitHeadingToBox(title, 9.5, HEADER_H - 0.1, 22);
  slide.addText(head.body, { x: 0.35, y: ACCENT_H + 0.05, w: 9.5, h: HEADER_H - 0.1, fontSize: head.fontSize, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true });
  if (meta) {
    slide.addText(meta, { x: 9.9, y: ACCENT_H + 0.05, w: 3.08, h: HEADER_H - 0.1, fontSize: 10, color: t.metaText, fontFace: 'Yu Gothic UI', align: 'right', valign: 'middle' });
  }
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: HEADER_END, w: W, h: SEP_H, fill: { color: t.accent }, line: { color: t.accent, width: 0 } });
  slide.addText('Generated by Microsoft Product Architecture Diagram Builder  ·  Swarm Data SE, Jiayi Yang', { x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H, fontSize: 8, color: t.footerText, fontFace: 'Yu Gothic UI', valign: 'middle' });
}

/**
 * Prose fitted to a fixed box: shrink first, and trim by measurement only when
 * shrinking has run out.
 *
 * A character cap stopped bounding height the moment a hard break became a real
 * paragraph — a twenty-line bulleted brief of 213 characters, half of the 420
 * the cap allowed, drew 4.875in in a 1.700in box: 1.725in off the bottom of the
 * sheet and straight through the model credit. Characters never bounded width
 * either; 420 of them is 0.98in of ASCII and 1.71in of Japanese.
 */
function fitProseToBox(
  prefix: string, body: string, boxW: number, boxH: number, startPt: number,
  /**
   * The smallest size this text may shrink to before it is trimmed instead.
   *
   * Prose keeps the 7pt legibility floor: a brief that has to be read closely
   * is still worth more small than absent. A heading does not - a 40pt cover
   * title that shrank to 7pt has stopped being a title, and the reader is
   * better served by a large name with an ellipsis than by a paragraph where
   * the name should be. So headings pass a floor of their own and reach the
   * trimming path far sooner.
   */
  minPt: number = LEGIBLE_TILE_PT,
): { body: string; fontSize: number } {
  const usable = Math.max(0.4, boxW - 0.2);
  const floorPt = Math.max(LEGIBLE_TILE_PT, Math.min(minPt, startPt));
  const fits = (s: string, pt: number): boolean => (wrappedLineCount(prefix + s, usable, pt) * pt * 1.35) / 72 <= boxH;
  for (let pt = startPt; pt >= floorPt; pt -= 0.5) {
    if (fits(body, pt)) return { body, fontSize: pt };
  }
  const chars = [...body];
  let lo = 0;
  let hi = chars.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (fits(`${chars.slice(0, mid).join('')}…`, floorPt)) lo = mid;
    else hi = mid - 1;
  }
  return { body: `${chars.slice(0, lo).join('').trimEnd()}…`, fontSize: floorPt };
}

/**
 * A heading fitted to the band that was reserved for it.
 *
 * The cover title, the section headers and the diagram slide header were all
 * drawn at a fixed point size with `wrap: true` and no autofit, so PowerPoint
 * wrapped them to as many lines as they needed and grew the block out of the
 * box - upward too, at `anchor="ctr"`, which put a three-line header above
 * the top edge of the slide. Nothing in the deck bounded the name, because
 * the name is free text the user types.
 *
 * Same two-stage answer as the index panel: take the largest size that fits,
 * and only when the floor is reached trim by measurement. `TITLE_FLOOR_PT`
 * ratios rather than absolute sizes so each of the three keeps its place in
 * the hierarchy - a shrunk cover title is still bigger than a section header.
 */
const TITLE_FLOOR_RATIO = 0.5;

function fitHeadingToBox(
  text: string, boxW: number, boxH: number, startPt: number,
): { body: string; fontSize: number } {
  return fitProseToBox('', text, boxW, boxH, startPt, startPt * TITLE_FLOOR_RATIO);
}

/** Slide 1 — title / cover. */
function addTitleSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const slide = pptx.addSlide();
  slide.background = { color: t.headerBg };
  // Left accent band
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.28, h: 7.5, fill: { color: t.accent }, line: { color: t.accent, width: 0 } });
  slide.addText('AZURE ARCHITECTURE', { x: 0.9, y: 1.5, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: t.accent, fontFace: 'Yu Gothic UI', charSpacing: 3 });
  const cover = fitHeadingToBox(o.diagramName, 11.5, 1.6, 40);
  slide.addText(cover.body, { x: 0.9, y: 2.0, w: 11.5, h: 1.6, fontSize: cover.fontSize, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
  slide.addText(`${o.author}   ·   ${o.date}`, { x: 0.9, y: 3.7, w: 11.5, h: 0.4, fontSize: 14, color: t.metaText, fontFace: 'Yu Gothic UI' });
  if (o.prompt) {
    const brief = fitProseToBox('Brief:  ', o.prompt, 11.5, 1.7, 13);
    slide.addText([
      { text: 'Brief:  ', options: { bold: true, color: t.metaText } },
      { text: brief.body, options: { color: t.metaText } },
    ], { x: 0.9, y: 4.35, w: 11.5, h: 1.7, fontSize: brief.fontSize, fontFace: 'Yu Gothic UI', valign: 'top', italic: true });
  }
  if (o.model) {
    slide.addText(`Generated with ${o.model}`, { x: 0.9, y: 6.6, w: 11.5, h: 0.35, fontSize: 10, color: t.footerText, fontFace: 'Yu Gothic UI' });
  }
}

/**
 * Slide 2 — the diagram, drawn with native (editable) PowerPoint shapes.
 *
 * This deck carries title, workflow, services, review and cost slides that are
 * all designed for a standard 16:9 page, so unlike {@link buildDiagramSlidePptx}
 * it cannot grow the page for a large drawing — PowerPoint gives a deck exactly
 * one page size. It tiles instead: an overview, then one slide per readable
 * window, exactly as the diagram-only deck does when the page has stopped
 * growing. Squeezing the whole drawing onto the single fixed slide is what
 * produced 0.05in tiles and 4pt type on the deck the export button actually
 * ships, while the audited deck showed 0.44in tiles for the same architecture.
 */
async function addDiagramSlide(pptx: PptxGenJS, t: SlideTheme, imageDataUrl: string, o: ArchitectureDeckOptions): Promise<void> {
  // The colour key is seated just below the drawing frame, so the frame has to
  // give up that strip first. `planSlideGeometry` does exactly this for the
  // single-slide export; the deck used the raw constant and so seated the key
  // 0.27in lower — 7.07in to 7.31in on a 7.5in page, straight through the top
  // of the footer band at 7.14in. It overflowed nothing and left nothing off
  // the sheet, so it was a pure position defect on every drawing slide of
  // every scenario in the corpus.
  const legendH = usedConnectionLegend(o.diagram?.edges ?? []).length > 0 ? 0.24 + 0.03 : 0;
  const frame = { x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H - legendH };
  const parts = o.diagram ? planFixedPageWindows(o.diagram, frame) : [];
  const windows: (DiagramWindow | undefined)[] = parts.length > 0 ? [undefined, ...parts] : [undefined];
  for (const [index, window] of windows.entries()) {
    const slide = pptx.addSlide();
    const partOf = parts.length === 0
      ? ''
      : index === 0
        ? '  (Overview)'
        : `  (${index} / ${parts.length})`;
    addChrome(pptx, slide, t, `${o.diagramName}${partOf}`, `${o.author}  ·  ${o.date}`);
    const renderedNatively = o.diagram
      ? await addEditableDiagram(
        pptx,
        slide,
        o.diagram,
        frame,
        o.isDarkMode,
        window,
        undefined,
        undefined,
        window === undefined && parts.length > 0,
        o.presetIcons,
      )
      : false;
    if (!renderedNatively) {
      slide.addImage({ data: imageDataUrl, x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H, sizing: { type: 'contain', w: IMAGE_W, h: IMAGE_H } });
      return;
    }
  }
}

/**
 * Slide 3 — numbered dataflow.
 *
 * Mirrors the "Workflow" section of an Azure Architecture Center reference
 * architecture: each numbered paragraph corresponds to the callout drawn on the
 * matching arrow of the diagram slide. Emitted only when the architecture
 * actually has a workflow, so a topology-only diagram gains no empty slide.
 *
 * Continues onto further slides rather than stopping at a fixed count. A flat
 * cap of twelve turned a sixteen-step architecture into a drawing carrying
 * callouts 13 to 16 that the deck then never explained, and the reader had no
 * way to find out what they meant — the wording was not truncated, it was
 * discarded. The diagram deck has always paginated its workflow; this is the
 * same rule applied to the deck the customer is actually sent.
 */
function addWorkflowSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const authored = (o.workflow ?? [])
    .filter((entry) => entry && readStepValue(entry.step) !== undefined && !!entry.description)
    .sort((a, b) => a.step - b.step);
  if (authored.length === 0) return;

  // The smallest row still worth reading, which used to be all that set how
  // many fit — a row *count*, with nothing measuring the words in it. The box
  // is `wrap="square"` with no autofit, so a description longer than its row
  // renders at full size and spills symmetrically out of both ends: at nineteen
  // steps of ordinary prose the rows overlap by a fifth of their pitch, at
  // CJK lengths by four fifths, and the top row prints over the header bar.
  // This is the same defect the diagram deck's workflow band was rewritten to
  // fix, and it takes the same answer — measure the wrap, then paginate on it.
  const MIN_ROW_IN = 0.3;
  const FLOOR_PT = 10;
  const ROW_SLACK_IN = 0.06;
  // A text box has insets of its own — pptxgenjs emits 0.1in left and right,
  // 0.05in top and bottom — and they come out of the room the words get. The
  // column is therefore 0.2in narrower than the box, and the row 0.1in taller
  // than its type. Measuring the box instead of the column under-counts the
  // wrap by a line exactly where the line is expensive.
  const TEXT_INSET_H_IN = 0.2;
  const TEXT_INSET_V_IN = 0.1;
  const descW = W - 1.1 - 0.34 - TEXT_INSET_H_IN;
  const needsFor = (text: string, hasServices: boolean, pt: number): number =>
    wrappedLineCount(text, descW, pt) * pt * 1.35 / 72 + TEXT_INSET_V_IN
    + (hasServices ? 0.2 : 0) + ROW_SLACK_IN;
  // A flat 240-character cap cut roughly a sentence off every generated step
  // and put it nowhere else in the deck — the same "truncated with no
  // discharge" defect the callout convention exists to prevent. Now that the
  // slide paginates on measured height there is no reason for a cap short of
  // the physical one: a step is shortened only when it could not fit a whole
  // page at the legibility floor, which generated prose never reaches.
  const steps = authored.map((entry) => {
    const hasServices = (entry.services ?? []).filter(Boolean).length > 0;
    let text = entry.description;
    while (text.length > 40 && needsFor(text, hasServices, FLOOR_PT) > BODY_H) {
      text = text.slice(0, Math.max(40, Math.floor(text.length * 0.9)));
    }
    return {
      entry,
      hasServices,
      description: text === entry.description ? text : `${text.trimEnd()}…`,
    };
  });
  const rowNeeds = (s: typeof steps[number], pt: number): number =>
    needsFor(s.description, s.hasServices, pt);
  const neededRow = Math.max(MIN_ROW_IN, ...steps.map((s) => rowNeeds(s, FLOOR_PT)));
  const perSlide = Math.max(1, Math.floor(BODY_H / neededRow));
  const pages = Math.ceil(steps.length / perSlide);

  for (let page = 0; page < pages; page += 1) {
    const shown = steps.slice(page * perSlide, (page + 1) * perSlide);
    const slide = pptx.addSlide();
    addChrome(
      pptx, slide, t,
      pages > 1 ? `Workflow (${page + 1} / ${pages})` : 'Workflow',
      `${steps.length} step${steps.length === 1 ? '' : 's'}`,
    );

    const rowH = Math.min(0.62, BODY_H / shown.length);
    // The type a row can carry without spilling out of it, not the type its
    // pitch suggests. A ladder of `rowH >= 0.5 ? 13 : …` reads the row's height
    // and ignores its contents, so a tall row full of prose still overflowed.
    const pitch = Math.max(rowH, ...shown.map((entry) => rowNeeds(entry, FLOOR_PT)));
    let fontSize = rowH >= 0.5 ? 13 : rowH >= 0.38 ? 11 : 10;
    while (fontSize > FLOOR_PT
      && shown.some((entry) => rowNeeds(entry, fontSize) > pitch)) {
      fontSize -= 0.5;
    }
    const badgeD = Math.min(0.34, pitch - 0.06);

    shown.forEach((s, index) => {
      const entry = s.entry;
      const y = BODY_TOP + index * pitch;
      slide.addShape(pptx.ShapeType.ellipse, {
        x: 0.4, y: y + (pitch - badgeD) / 2, w: badgeD, h: badgeD,
        fill: { color: t.accent }, line: { color: t.accent, width: 0 },
      });
      slide.addText(String(entry.step), {
        x: 0.4, y: y + (pitch - badgeD) / 2, w: badgeD, h: badgeD,
        fontSize: Math.max(8, Math.round(fontSize * 0.8)), bold: true, color: 'ffffff',
        fontFace: 'Yu Gothic UI', align: 'center', valign: 'middle',
      });
      // A wrapped two-line description used to run straight through the services
      // strip, so the strip is reserved out of the description box's height.
      const services = (entry.services ?? []).filter(Boolean);
      const showsServices = services.length > 0 && pitch >= 0.5;
      slide.addText(s.description, {
        x: 0.4 + badgeD + 0.16, y, w: W - 1.1 - badgeD, h: showsServices ? pitch - 0.2 : pitch,
        fontSize, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true,
      });
      if (showsServices) {
        // One line, fitted. The strip is a breadcrumb under a sentence that
        // already names the same services, and it was given a flat 0.18in with
        // nothing measuring what went in it — two long resource names joined by
        // an arrow wrap to two lines and the second is painted over the row
        // below. Every name here is spelled out in full on the Services slide,
        // so shortening this one is a shortening, not a loss.
        const stripW = W - 1.1 - badgeD;
        slide.addText(fitLabelToBox(services.join('  →  '), stripW - 0.2, 9), {
          x: 0.4 + badgeD + 0.16, y: y + pitch - 0.2, w: stripW, h: 0.18,
          fontSize: 9, color: t.metaText, fontFace: 'Yu Gothic UI', valign: 'middle',
        });
      }
    });
  }
}

/**
 * The height a table row occupies once PowerPoint has wrapped it.
 *
 * `rowH` is a *minimum*, not a height: none of these tables declares autofit,
 * so PowerPoint grows any row whose text wraps and the table quietly gets
 * taller than the slide. Anything sizing a table has to measure the words —
 * including the vertical cell insets, which pptxgenjs emits (`marT`/`marB`,
 * 0.05in each) and which are charged to the row on top of the type.
 */
const TABLE_CELL_MARGIN_IN = 0.2; // marL + marR
const TABLE_CELL_INSET_V_IN = 0.1; // marT + marB
const TABLE_MIN_ROW_H = 0.325;
export function tableRowHeightIn(cells: string[], colW: number[], pt: number): number {
  const lines = Math.max(...cells.map((text, i) => wrappedLineCount(
    text, Math.max(0.5, (colW[i] ?? colW[0]) - TABLE_CELL_MARGIN_IN), pt,
  )));
  return Math.max(TABLE_MIN_ROW_H, lines * pt * 1.35 / 72 + TABLE_CELL_INSET_V_IN);
}

/**
 * As many rows as fit in `availableIn`, and the type they fit at.
 *
 * Used where the table cannot paginate — a summary slide shows the top N and
 * the tail is genuinely optional — so the answer to "too tall" is smaller type
 * first, down to the deck's legibility floor, and only then fewer rows. A
 * caller that drops rows has to say so; silently listing eight of ten cost
 * drivers under a heading that says ten is the same defect as a truncated
 * inventory.
 */
export function fitTableRows(
  rows: string[][], header: string[], colW: number[], availableIn: number, startPt: number,
): { rows: number; pt: number } {
  let pt = startPt;
  const heightAt = (p: number, count: number): number => tableRowHeightIn(header, colW, p)
    + rows.slice(0, count).reduce((sum, r) => sum + tableRowHeightIn(r, colW, p), 0);
  while (pt > LEGIBLE_TILE_PT && heightAt(pt, rows.length) > availableIn) pt -= 0.5;
  pt = Math.max(LEGIBLE_TILE_PT, pt);
  let count = rows.length;
  while (count > 1 && heightAt(pt, count) > availableIn) count -= 1;
  return { rows: count, pt };
}

/**
 * A warning banner whose middle is a list of unknown length, fitted to its box.
 *
 * These two notices were the last character caps in the file: 120 and 150
 * characters into boxes that hold neither. A character is not a width — a
 * Japanese region name is an em wide and a Latin one 0.54 — so the 120-cap
 * needed 1.125in of a 0.750in box in CJK and the 150-cap bit into the table
 * beneath it. Shrink first, and only then shorten the list, saying how many
 * were left out rather than ending on an ellipsis that names nothing.
 */
function fitNotice(
  head: string, list: string, tail: string,
  boxW: number, boxH: number, startPt: number,
): { text: string; fontSize: number } {
  const usable = Math.max(0.4, boxW - 0.2);
  const fits = (s: string, pt: number): boolean => (wrappedLineCount(s, usable, pt) * pt * 1.35) / 72 <= boxH;
  for (let pt = startPt; pt >= LEGIBLE_TILE_PT; pt -= 0.5) {
    if (fits(head + list + tail, pt)) return { text: head + list + tail, fontSize: pt };
  }
  const items = list.split(', ').filter(Boolean);
  for (let keep = items.length - 1; keep >= 1; keep -= 1) {
    const shortened = `${items.slice(0, keep).join(', ')} +${items.length - keep} more`;
    if (fits(head + shortened + tail, LEGIBLE_TILE_PT)) {
      return { text: head + shortened + tail, fontSize: LEGIBLE_TILE_PT };
    }
  }
  return { text: `${head}${items.length} regions${tail}`, fontSize: LEGIBLE_TILE_PT };
}

/** Slide 4 — service inventory. */
function addServicesSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  if (!o.services.length) return;

  // An inventory that omits its own contents is not an inventory. The heading
  // counted every service while the table stopped at twenty, so a deck for a
  // sixty-service estate announced sixty components and listed a third of them.
  //
  // Paginating on a flat row height then reintroduced the same defect one level
  // down. This table declares no autofit, so PowerPoint treats `<a:tr h>` as a
  // minimum and grows any row whose text wraps — and a name that wraps to two
  // lines is 0.45in against the 0.32in it was budgeted. Eighteen rows of that
  // put the last of them 1.6in below the bottom of the slide, still present in
  // the file and readable by any rule that greps the XML, and invisible to the
  // reader. Since this table is where every name the drawing shortened is
  // spelled out, a row off the page is a name lost after all. Measure the wrap
  // and pack rows by their real height.
  const FONT_PT = 12;
  const COL_W = [5.2, 3.9, 3.53];
  const cellsOf = (s: DeckService): string[] => [s.name, s.category || '—', s.group || '—'];
  const heightOf = (cells: string[], pt: number): number => tableRowHeightIn(cells, COL_W, pt);

  // A page's type shrinks only for the row that cannot otherwise fit, and never
  // below the deck's legibility floor. One name long enough to fill a page on
  // its own is not something an architecture produces, but a table that grows
  // silently past the slide is exactly the defect above, so it gets an answer
  // rather than an assumption.
  const pageFontFor = (services: DeckService[]): number => {
    let pt = FONT_PT;
    while (pt > LEGIBLE_TILE_PT) {
      const header = heightOf(['Service', 'Category', 'Zone / Group'], pt);
      const tallest = Math.max(...services.map((s) => heightOf(cellsOf(s), pt)));
      if (header + tallest <= BODY_H) break;
      pt -= 0.5;
    }
    return Math.max(LEGIBLE_TILE_PT, pt);
  };

  const pageSlices: DeckService[][] = [];
  let current: DeckService[] = [];
  let used = heightOf(['Service', 'Category', 'Zone / Group'], FONT_PT);
  for (const service of o.services) {
    const h = heightOf(cellsOf(service), FONT_PT);
    if (current.length > 0 && used + h > BODY_H) {
      pageSlices.push(current);
      current = [];
      used = heightOf(['Service', 'Category', 'Zone / Group'], FONT_PT);
    }
    current.push(service);
    used += h;
  }
  if (current.length > 0) pageSlices.push(current);
  const pages = pageSlices.length;

  for (let page = 0; page < pages; page += 1) {
    const shown = pageSlices[page];
    const pagePt = pageFontFor(shown);
    const headerH = heightOf(['Service', 'Category', 'Zone / Group'], pagePt);
    const slide = pptx.addSlide();
    addChrome(
      pptx, slide, t,
      pages > 1
        ? `Services  ·  ${o.services.length} components  (${page + 1} / ${pages})`
        : `Services  ·  ${o.services.length} components`,
    );

    const header = [
      { text: 'Service', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Category', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Zone / Group', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
    ];
    const rows = shown.map((s) => [
      { text: s.name, options: { color: t.titleText } },
      { text: s.category || '—', options: { color: t.metaText } },
      { text: s.group || '—', options: { color: t.metaText } },
    ]);
    const rowHeights = [headerH, ...shown.map((s) => heightOf(cellsOf(s), pagePt))];
    slide.addTable([header, ...rows], {
      x: 0.35, y: BODY_TOP, w: W - 0.7,
      h: rowHeights.reduce((sum, h) => sum + h, 0),
      colW: COL_W,
      fontSize: pagePt, fontFace: 'Yu Gothic UI',
      border: { type: 'solid', color: t.headerBg, pt: 1 },
      valign: 'middle', rowH: rowHeights,
    });
  }
}

/** Slide 4a — WAF executive summary (only when validation provided). */
function addValidationSummarySlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const v = o.validation;
  if (!v) return;
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, 'Well-Architected review', v.modelUsed ? `Assessed with ${v.modelUsed}` : undefined);

  // Big score block (left)
  slide.addShape(pptx.ShapeType.roundRect, { x: 0.35, y: BODY_TOP, w: 3.0, h: 2.35, fill: { color: t.headerBg }, line: { color: scoreColor(v.overallScore), width: 2 }, rectRadius: 0.08 });
  slide.addText(`${Math.round(v.overallScore)}`, { x: 0.35, y: BODY_TOP + 0.2, w: 3.0, h: 1.25, fontSize: 58, bold: true, color: scoreColor(v.overallScore), align: 'center', fontFace: 'Yu Gothic UI' });
  slide.addText('/ 100', { x: 0.35, y: BODY_TOP + 1.42, w: 3.0, h: 0.32, fontSize: 13, color: t.metaText, align: 'center', fontFace: 'Yu Gothic UI' });
  slide.addText(v.overallLabel || scoreBand(v.overallScore), { x: 0.3, y: BODY_TOP + 1.78, w: 3.1, h: 0.5, fontSize: 13, bold: true, color: scoreColor(v.overallScore), align: 'center', valign: 'middle', fontFace: 'Yu Gothic UI', wrap: true });

  // Executive summary text (right of score block)
  if (v.summary) {
    // 620 characters is about seven lines of English in this column and
    // thirteen of Japanese, and the box holds ten — so the tail of a Japanese
    // assessment printed straight over the "Pillar maturity" heading below it.
    // Shrink the type to what the box can hold rather than cap the characters:
    // a character cap discharges nothing, and it is wrong by roughly 2x
    // depending on the script.
    const SUMMARY_H = 2.35;
    const summaryW = W - 3.7 - 0.35 - 0.2; // less the box's own insets
    const text = `Assessment.  ${v.summary}`;
    let summaryPt = 12.5;
    while (summaryPt > LEGIBLE_TILE_PT
      && wrappedLineCount(text, summaryW, summaryPt) * summaryPt * 1.35 / 72 + 0.1 > SUMMARY_H) {
      summaryPt -= 0.5;
    }
    // Only if it still will not fit at the floor is anything cut, and then by
    // measurement rather than by a character count.
    let body = v.summary;
    while (body.length > 40
      && wrappedLineCount(`Assessment.  ${body}`, summaryW, summaryPt) * summaryPt * 1.35 / 72 + 0.1 > SUMMARY_H) {
      body = body.slice(0, Math.max(40, Math.floor(body.length * 0.9)));
    }
    slide.addText([
      { text: 'Assessment.  ', options: { bold: true, color: t.titleText } },
      { text: body === v.summary ? body : `${body.trimEnd()}…`, options: { color: t.metaText } },
    ], { x: 3.7, y: BODY_TOP, w: W - 3.7 - 0.35, h: SUMMARY_H, fontSize: summaryPt, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
  }

  // Pillar maturity table (full width, below)
  const pillars = v.pillars.slice(0, 5);
  const pTop = BODY_TOP + 2.65;
  slide.addText('Pillar maturity', { x: 0.35, y: pTop, w: 6, h: 0.3, fontSize: 13, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI' });
  const header = [
    { text: 'Pillar', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
    { text: 'Maturity', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
    { text: 'Score', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
  ];
  const PILLAR_COL_W = [4.2, 6.6, 1.83];
  const pillarCells = pillars.map((p) => [
    p.pillar, p.maturity || scoreBand(p.score), `${Math.round(p.score)} / 100`,
  ]);
  // Same contract as every other table in this deck: no autofit is declared, so
  // a long maturity label wraps and grows its row. This one starts 2.65in down
  // the body, so it has that much less room to grow into.
  const pillarFit = fitTableRows(
    pillarCells, ['Pillar', 'Maturity', 'Score'], PILLAR_COL_W,
    Math.max(TABLE_MIN_ROW_H * 2, FOOTER_Y - (pTop + 0.36) - 0.1), 12,
  );
  const rows = pillars.slice(0, pillarFit.rows).map((p) => [
    { text: p.pillar, options: { color: t.titleText } },
    { text: p.maturity || scoreBand(p.score), options: { color: t.metaText } },
    { text: `${Math.round(p.score)} / 100`, options: { color: scoreColor(p.score), bold: true, align: 'right' as const } },
  ]);
  slide.addTable([header, ...rows], {
    x: 0.35, y: pTop + 0.36, w: W - 0.7,
    colW: PILLAR_COL_W,
    fontSize: pillarFit.pt, fontFace: 'Yu Gothic UI',
    border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: TABLE_MIN_ROW_H,
  });
}

/** Slide 4b — WAF key findings & recommendations (only when findings exist). */
function addValidationFindingsSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const v = o.validation;
  if (!v || !v.findings.length) return;

  // Every other prose surface in this deck is paginated on measured height.
  // This one divided the body by a row *count* and capped its contents by
  // character count — and 170 characters is about two lines of English but
  // four of Japanese, which the model is explicitly instructed to write. At the
  // five findings the app always sends, the issue painted over its own Fix, the
  // Fix over the next issue, and the last Fix over the footer: nine collisions
  // on one slide, in the only case that occurs.
  const ISSUE_PT = 12;
  const FIX_PT = 11;
  const COL_W = W - 1.95 - 0.2; // the box, less its own left/right insets
  const GAP_IN = 0.12;
  const INSET_V_IN = 0.1;
  const available = FOOTER_Y - BODY_TOP - 0.1;
  const blockFor = (f: DeckFinding, issuePt: number, fixPt: number) => {
    const issueH = wrappedLineCount(`${f.category}. ${f.issue}`, COL_W, issuePt)
      * issuePt * 1.35 / 72 + INSET_V_IN;
    const fixH = f.recommendation
      ? wrappedLineCount(`→ Fix:  ${f.recommendation}`, COL_W, fixPt) * fixPt * 1.35 / 72 + INSET_V_IN
      : 0;
    return { issueH, fixH, total: issueH + fixH + GAP_IN };
  };
  // Shrink before splitting, so a page holds as many findings as it can read.
  let issuePt = ISSUE_PT;
  let fixPt = FIX_PT;
  while (issuePt > LEGIBLE_TILE_PT
    && Math.max(...v.findings.map((f) => blockFor(f, issuePt, fixPt).total)) > available) {
    issuePt -= 0.5;
    fixPt = Math.max(LEGIBLE_TILE_PT, fixPt - 0.5);
  }

  // A finding that still will not fit a whole page at the legible floor is the
  // one case shrinking cannot answer, and the packer below always places at
  // least one block per page — so without this it would be placed anyway and
  // painted straight off the slide. Cut it physically, by measurement, and only
  // as far as it takes: a character cap is what this slide was just fixed for.
  const shortened = new Map<DeckFinding, DeckFinding>();
  const fitToPage = (f: DeckFinding): DeckFinding => {
    if (blockFor(f, issuePt, fixPt).total <= available) return f;
    let issue = f.issue;
    let recommendation = f.recommendation ?? '';
    while (blockFor({ ...f, issue, recommendation }, issuePt, fixPt).total > available) {
      // Take from whichever half is longer, so neither is starved to keep the
      // other whole.
      if (recommendation.length > issue.length && recommendation.length > 24) {
        recommendation = recommendation.slice(0, Math.floor(recommendation.length * 0.9));
      } else if (issue.length > 24) {
        issue = issue.slice(0, Math.floor(issue.length * 0.9));
      } else break;
    }
    return {
      ...f,
      issue: issue === f.issue ? issue : `${issue.trimEnd()}…`,
      recommendation: recommendation === (f.recommendation ?? '') ? f.recommendation : `${recommendation.trimEnd()}…`,
    };
  };
  for (const f of v.findings) shortened.set(f, fitToPage(f));

  // Pack by measured height rather than dropping the tail. The app sends the
  // top six by severity and this slide drew five; the sixth appeared nowhere
  // else in the deck, which is the same silent truncation the cost table and
  // the services table were both fixed for.
  const pages: DeckFinding[][] = [];
  let current: DeckFinding[] = [];
  let used = 0;
  for (const authored of v.findings) {
    const f = shortened.get(authored)!;
    const h = blockFor(f, issuePt, fixPt).total;
    if (current.length > 0 && used + h > available) {
      pages.push(current);
      current = [];
      used = 0;
    }
    current.push(f);
    used += h;
  }
  if (current.length > 0) pages.push(current);

  pages.forEach((shown, page) => {
    const slide = pptx.addSlide();
    addChrome(
      pptx, slide, t,
      pages.length > 1 ? `Key findings & recommendations (${page + 1} / ${pages.length})` : 'Key findings & recommendations',
      `${v.findings.length} finding${v.findings.length === 1 ? '' : 's'}`,
    );
    let y = BODY_TOP;
    shown.forEach((f) => {
      const block = blockFor(f, issuePt, fixPt);
      // Severity chip
      slide.addText(f.severity.toUpperCase(), { x: 0.35, y: y + 0.05, w: 1.05, h: 0.34, fontSize: 9, bold: true, color: 'ffffff', fill: { color: severityColor(f.severity) }, align: 'center', valign: 'middle', fontFace: 'Yu Gothic UI' });
      // Issue + recommendation
      slide.addText([
        { text: `${f.category}. `, options: { bold: true, color: t.titleText } },
        { text: f.issue, options: { color: t.metaText } },
      ], { x: 1.55, y, w: W - 1.95, h: block.issueH, fontSize: issuePt, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
      if (f.recommendation) {
        slide.addText([
          { text: '→ Fix:  ', options: { bold: true, color: t.accent } },
          { text: f.recommendation, options: { color: t.metaText, italic: true } },
        ], { x: 1.55, y: y + block.issueH, w: W - 1.95, h: block.fixH, fontSize: fixPt, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
      }
      y += block.total;
    });
  });
}

/** Slide 5a — cost overview (only when cost provided). */
function addCostOverviewSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const c = o.cost;
  if (!c) return;
  const slide = pptx.addSlide();
  const meta = [
    c.term,
    c.region,
    c.pricesAsOf ? `prices as of ${c.pricesAsOf}` : undefined,
    c.oldestMeterAsOf ? `unchanged since ${c.oldestMeterAsOf}` : undefined,
  ].filter(Boolean).join('  ·  ');
  addChrome(pptx, slide, t, 'Estimated cost', meta || undefined);

  // Headline monthly + annual
  slide.addText(money(c.totalMonthly, c.currency), { x: 0.35, y: BODY_TOP, w: 5.2, h: 0.95, fontSize: 46, bold: true, color: t.accent, fontFace: 'Yu Gothic UI' });
  slide.addText('per month (estimate)', { x: 0.37, y: BODY_TOP + 0.95, w: 5.2, h: 0.32, fontSize: 12, color: t.metaText, fontFace: 'Yu Gothic UI' });
  if (c.annual) {
    slide.addText(`≈ ${money(c.annual, c.currency)} / year`, { x: 0.37, y: BODY_TOP + 1.3, w: 5.2, h: 0.35, fontSize: 15, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI' });
  }

  // Fixed vs usage-based split
  if (c.fixedCost != null && c.usageCost != null && c.totalMonthly > 0) {
    const fixedPct = Math.round((c.fixedCost / c.totalMonthly) * 100);
    const usagePct = 100 - fixedPct;
    const barY = BODY_TOP + 1.95;
    const barW = 5.0;
    slide.addText('Cost predictability', { x: 0.37, y: barY, w: 5.2, h: 0.3, fontSize: 12, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI' });
    // stacked bar
    slide.addShape(pptx.ShapeType.rect, { x: 0.37, y: barY + 0.35, w: barW, h: 0.28, fill: { color: t.headerBg }, line: { width: 0 } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.37, y: barY + 0.35, w: Math.max(0.02, barW * fixedPct / 100), h: 0.28, fill: { color: t.accent }, line: { width: 0 } });
    slide.addText([
      { text: `Fixed ${money(c.fixedCost, c.currency)} (${fixedPct}%)`, options: { color: t.accent, bold: true } },
      { text: `    ·    Usage-based ${money(c.usageCost, c.currency)} (${usagePct}%) — varies`, options: { color: t.metaText } },
    ], { x: 0.37, y: barY + 0.68, w: 5.2, h: 0.35, fontSize: 10, fontFace: 'Yu Gothic UI' });
  }

  if (c.regionComparisonIncomplete) {
    const unavailable = c.unavailableRegions?.map(item => item.split(':', 1)[0]).join(', ') || 'one or more regions';
    const notice = fitNotice(
      'Regional comparison is partial because selected SKUs are unavailable in: ',
      unavailable,
      '. No cheapest-region recommendation is made.',
      5.2, 0.75, 10,
    );
    slide.addText(
      notice.text,
      { x: 0.37, y: BODY_TOP + 3.15, w: 5.2, h: 0.75, fontSize: notice.fontSize, bold: true, color: 'b45309', fontFace: 'Yu Gothic UI', wrap: true, valign: 'top' },
    );
  }

  slide.addText('Estimate only — not a quote. Excludes taxes, egress, support plans and reservations unless modeled.', { x: 0.37, y: FOOTER_Y - 0.5, w: 5.2, h: 0.45, fontSize: 9, italic: true, color: t.footerText, fontFace: 'Yu Gothic UI', wrap: true });

  // Top cost drivers table (right)
  const svcs = c.topServices.slice(0, 10);
  if (svcs.length) {
    const COST_COL_W = [3.3, 1.85, 1.15, 0.78];
    const headerCells = ['Top cost drivers', 'Tier', 'Monthly', 'Share'];
    const rowCells = svcs.map((s) => [
      s.serviceName,
      s.tier || '—',
      money(s.cost, c.currency),
      s.percentage != null ? `${Math.round(s.percentage)}%` : '—',
    ]);
    // Ten rows at a declared 0.32in is a row *count*, and this table declares
    // no autofit — a service name that wraps in a 3.3in column grows its row,
    // and ten of those ran two inches past the bottom of the slide. Measure the
    // wrap, shrink the type, and only then show fewer drivers.
    const fitted = fitTableRows(rowCells, headerCells, COST_COL_W, BODY_H, 12);
    const shown = svcs.slice(0, fitted.rows);
    const header = [
      { text: `Top cost drivers${shown.length < svcs.length ? ` (top ${shown.length})` : ''}`, options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Tier', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Monthly', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
      { text: 'Share', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    ];
    const rows = shown.map((s) => [
      { text: s.serviceName, options: { color: t.titleText } },
      { text: s.tier || '—', options: { color: t.metaText } },
      { text: money(s.cost, c.currency), options: { color: t.metaText, align: 'right' as const } },
      { text: s.percentage != null ? `${Math.round(s.percentage)}%` : '—', options: { color: t.metaText, align: 'right' as const } },
    ]);
    slide.addTable([header, ...rows], {
      x: 5.9, y: BODY_TOP, w: W - 5.9 - 0.35,
      colW: COST_COL_W,
      fontSize: fitted.pt, fontFace: 'Yu Gothic UI',
      border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: TABLE_MIN_ROW_H,
    });
  }
}

/** Slide 5b — multi-region cost comparison (only when >1 region computed). */
function addCostRegionsSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const c = o.cost;
  if (!c || !c.regions || c.regions.length < 2) return;
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, 'Regional cost comparison', c.term || undefined);

  const comparisonComplete = !c.regionComparisonIncomplete;
  const lowestShown = c.regions[0];
  const cheapest = comparisonComplete
    ? c.regions.find(r => r.isCheapest) || lowestShown
    : undefined;
  const current = c.regions.find(r => r.isCurrent);
  if (!comparisonComplete) {
    const unavailable = c.unavailableRegions?.map(item => item.split(':', 1)[0]).join(', ') || 'one or more regions';
    const notice = fitNotice(
      'Partial comparison — unavailable: ',
      unavailable,
      '. Values below cover comparable regions only; no global cheapest or savings claim is shown.',
      W - 0.7, 0.55, 12,
    );
    slide.addText(
      notice.text,
      { x: 0.35, y: BODY_TOP, w: W - 0.7, h: 0.55, fontSize: notice.fontSize, bold: true, color: 'b45309', fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true },
    );
  } else if (cheapest) {
    const onCheapest = current && current.name === cheapest.name;
    const msg = onCheapest
      ? `Already on the cheapest region — ${cheapest.flag || ''} ${cheapest.name} at ${money(cheapest.monthly, c.currency)}/mo.`
      : `Cheapest region: ${cheapest.flag || ''} ${cheapest.name} at ${money(cheapest.monthly, c.currency)}/mo` +
        (current ? `  ·  potential saving ${money(current.monthly - cheapest.monthly, c.currency)}/mo` : '');
    slide.addText(msg, { x: 0.35, y: BODY_TOP, w: W - 0.7, h: 0.45, fontSize: 13, bold: true, color: t.accent, fontFace: 'Yu Gothic UI', valign: 'middle' });
  }

  const rows = c.regions.slice(0, 8).map((r) => {
    const tag = comparisonComplete && r.isCheapest ? '  ★' : r.isCurrent ? '  (current)' : '';
    const vsBaseline = lowestShown && lowestShown.monthly > 0
      ? (r.monthly === lowestShown.monthly ? 'baseline' : `+${(((r.monthly - lowestShown.monthly) / lowestShown.monthly) * 100).toFixed(1)}%`)
      : '—';
    return [
      { text: `${r.flag || ''} ${r.name}${tag}`, options: { color: t.titleText, bold: !!r.isCheapest } },
      { text: money(r.monthly, c.currency), options: { color: t.metaText, align: 'right' as const } },
      { text: money(r.annual, c.currency), options: { color: t.metaText, align: 'right' as const } },
      { text: vsBaseline, options: { color: t.metaText, align: 'right' as const } },
    ];
  });
  const regionCells = rows.map((r) => r.map((cell) => cell.text));
  const header = [
    { text: 'Region', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
    { text: 'Monthly', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    { text: 'Annual', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    { text: comparisonComplete ? 'vs cheapest' : 'vs lowest shown', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
  ];
  const REGION_COL_W = [6.13, 2.0, 2.0, 1.5];
  const regionFit = fitTableRows(
    regionCells, ['Region', 'Monthly', 'Annual', 'vs cheapest'], REGION_COL_W,
    Math.max(TABLE_MIN_ROW_H * 2, FOOTER_Y - (BODY_TOP + 0.6) - 0.1), 12,
  );
  slide.addTable([header, ...rows.slice(0, regionFit.rows)], {
    x: 0.35, y: BODY_TOP + 0.6, w: W - 0.7,
    colW: REGION_COL_W,
    fontSize: regionFit.pt, fontFace: 'Yu Gothic UI',
    border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: TABLE_MIN_ROW_H,
  });
}

/**
 * Assemble the multi-slide, customer-ready deck for the current architecture:
 * title, diagram, numbered workflow, services, and (when available) a
 * Well-Architected review (summary + findings) and a cost estimate (overview +
 * regional comparison). Split from the download so tests can inspect the deck.
 */
/**
 * Collapse a name onto one line.
 *
 * A newline survives the XML sanitiser and pptxgenjs turns each one into a real
 * paragraph, so a service name pasted out of a spreadsheet cell ("Azure SQL
 * Managed Instance\nProduction ring\nEast US 2") draws as four lines wherever
 * it appears — doubling a table past the bottom of the slide, squeezing the
 * icon off its tile, and painting a breadcrumb over the row beneath it. Prose
 * fields (summary, issue, recommendation, step description) mean their line
 * breaks and are measured and paginated as written; a *name* is an identifier
 * and reads as one line, wrapping only because the column is narrow.
 */
function singleLine(text: string): string {
  // Composed, for the same reason the drawing composes: the deck writes a
  // name in two places, and if the two spellings differ the reader searching
  // the inventory for the name on the tile finds nothing. The tile side goes
  // through `singleLineName`, which composes; this side did not, so seven
  // Vietnamese and Turkish services were drawn on the diagram slide and
  // absent from the inventory that is supposed to spell them out in full.
  return text.replace(/[\r\n\t\v\f\u2028\u2029]+/g, ' ').replace(/ {2,}/g, ' ').trim().normalize('NFC');
}

/**
 * Prose, composed but otherwise untouched.
 *
 * A sentence keeps its line breaks - that is the whole difference between it
 * and a name - but it must not keep a decomposed spelling, because every
 * measurement in this file prices a combining mark at nothing and the two
 * spellings would then be drawn at two different widths. Composing here, at
 * the entry point, means every box in the deck is measured against the string
 * that will actually be written into it.
 */
function composedProse(text: string): string {
  return text.normalize('NFC');
}

/**
 * Every name-shaped field the deck prints, collapsed onto one line. Applied
 * once at the entry point rather than at each of the twenty-odd draw sites, so
 * a slide added later cannot reintroduce the defect by forgetting to call it.
 */
function withSingleLineNames(o: ArchitectureDeckOptions): ArchitectureDeckOptions {
  const cost = o.cost
    ? {
      ...o.cost,
      byCategory: o.cost.byCategory.map((c) => ({ ...c, category: singleLine(c.category) })),
      topServices: o.cost.topServices.map((s) => ({
        ...s,
        serviceName: singleLine(s.serviceName),
        ...(s.tier ? { tier: singleLine(s.tier) } : {}),
      })),
      ...(o.cost.regions ? { regions: o.cost.regions.map((r) => ({ ...r, name: singleLine(r.name) })) } : {}),
      ...(o.cost.term ? { term: singleLine(o.cost.term) } : {}),
      ...(o.cost.region ? { region: singleLine(o.cost.region) } : {}),
      ...(o.cost.unavailableRegions
        ? { unavailableRegions: o.cost.unavailableRegions.map(singleLine) }
        : {}),
    }
    : o.cost;
  const validation = o.validation
    ? {
      ...o.validation,
      ...(o.validation.overallLabel ? { overallLabel: singleLine(o.validation.overallLabel) } : {}),
      pillars: o.validation.pillars.map((p) => ({
        ...p,
        pillar: singleLine(p.pillar),
        ...(p.maturity ? { maturity: singleLine(p.maturity) } : {}),
      })),
      findings: o.validation.findings.map((f) => ({
        ...f,
        severity: singleLine(f.severity),
        category: singleLine(f.category),
        issue: composedProse(f.issue),
        ...(f.recommendation ? { recommendation: composedProse(f.recommendation) } : {}),
      })),
      ...(o.validation.summary ? { summary: composedProse(o.validation.summary) } : {}),
      ...(o.validation.modelUsed ? { modelUsed: singleLine(o.validation.modelUsed) } : {}),
    }
    : o.validation;
  return {
    ...o,
    diagramName: singleLine(o.diagramName),
    // The other two thirds of the header triple. Left uncomposed they were
    // invisible only because nothing fitted them; the cover title now has a
    // fitter, so an NFD name measures differently from the NFC one that means
    // the same thing and the two shrink to different sizes.
    author: singleLine(o.author),
    date: singleLine(o.date),
    ...(o.prompt ? { prompt: composedProse(o.prompt) } : {}),
    ...(o.model ? { model: singleLine(o.model) } : {}),
    services: o.services.map((s) => ({
      ...s,
      name: singleLine(s.name),
      ...(s.category ? { category: singleLine(s.category) } : {}),
      ...(s.group ? { group: singleLine(s.group) } : {}),
    })),
    // `description` is prose and keeps the line breaks it was written with;
    // `services` is a breadcrumb of names drawn on one line.
    ...(o.workflow
      ? {
        workflow: o.workflow.map((w) => ({
          ...w,
          description: composedProse(w.description),
          ...(w.services ? { services: w.services.map(singleLine) } : {}),
        })),
      }
      : {}),
    cost,
    validation,
  };
}

export async function buildArchitectureDeckPptx(
  imageDataUrl: string,
  rawOptions: ArchitectureDeckOptions,
): Promise<PptxGenJS> {
  const options = withSingleLineNames(rawOptions);
  const t = options.isDarkMode ? DARK_THEME : LIGHT_THEME;

  const pptx = newDeck();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = options.author;
  pptx.title = options.diagramName;
  pptx.subject = 'Azure Architecture Review';
  pptx.company = 'Microsoft Azure';

  addTitleSlide(pptx, t, options);
  await addDiagramSlide(pptx, t, imageDataUrl, options);
  addWorkflowSlide(pptx, t, options);
  addServicesSlide(pptx, t, options);
  addValidationSummarySlide(pptx, t, options);
  addValidationFindingsSlide(pptx, t, options);
  addCostOverviewSlide(pptx, t, options);
  addCostRegionsSlide(pptx, t, options);

  return pptx;
}

/** Build the deck and download it. Returns the generated filename. */
export async function exportArchitectureDeck(
  imageDataUrl: string,
  options: ArchitectureDeckOptions,
): Promise<string> {
  const pptx = await buildArchitectureDeckPptx(imageDataUrl, options);
  const fileName = generateModelFilename('azure-architecture-deck', 'pptx');
  await downloadNativePptx(pptx, fileName);
  return fileName;
}


