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

import { generateModelFilename } from '../utils/modelNaming';
import { rasterizeIcons, type RasterizedIcon } from '../utils/exportIconRaster';
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
  metaSubline,
  partitionBoxes,
  stripHash,
  truncateLabel,
  usedConnectionLegend,
  workflowListFromEdges,
  narrateEdgeCallouts,
  zoneStyleFor,
  readableTextOn,
  carriesWording,
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
const MAX_TILED_CELLS = 4096;

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
function planDiagramWindows(
  bounds: Bounds,
  services: ExportBox[],
  frame: DiagramFrame,
  options: { mustTile?: boolean } = {},
): { windows: DiagramWindow[]; legible: boolean } {
  const contentW = Math.max(1, bounds.maxX - bounds.minX);
  const contentH = Math.max(1, bounds.maxY - bounds.minY);
  const whole = { windows: [] as DiagramWindow[], legible: true };
  if (services.length === 0 || frame.w <= 0 || frame.h <= 0) return whole;

  const shortest = Math.min(...services.map((box) => box.h).filter((h) => h > 0));
  if (!Number.isFinite(shortest) || shortest <= 0) return whole;

  // Inches-per-pixel needed for the shortest tile to keep a readable label.
  const legibleScale = LEGIBLE_TILE_PT / 12 / shortest;
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
    const holds = (col: number, row: number): boolean => {
      const cell = { ...cellX(col), ...cellY(row) };
      return services.some((box) => windowOwnsPoint(cell, bounds, box.x + box.w / 2, box.y + box.h / 2));
    };

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
    while (c * r > 1 && slidesFor(c, r) > MAX_LEGIBLE_TILED_SLIDES) ({ c, r } = drop(c, r));
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

  const comfortable = gridFor(COMFORTABLE_TILE_PT / 12 / shortest);
  const floor = gridFor(legibleScale);
  const worthIt = comfortable
    && comfortable.slides <= MAX_DIAGRAM_SLIDES
    && services.length / comfortable.slides >= MIN_SERVICES_PER_SLIDE;
  const grid = worthIt ? comfortable : floor;
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
  if (cols * rows > MAX_TILED_SLIDES) return { windows: [], legible: false };
  if (
    cols * rows > MAX_DIAGRAM_SLIDES
    && services.length / (cols * rows) < MIN_SERVICES_PER_TILED_SLIDE
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
  return planDiagramWindows(bounds, services, frame, { mustTile: true }).windows;
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
  const frameFor = (pageW: number, pageH: number): DiagramFrame => {
    const footer = pageH - FOOTER_H - 0.08;
    return { x: IMAGE_X, y: IMAGE_Y, w: pageW - IMAGE_X * 2, h: footer - IMAGE_Y - 0.1 - legendH };
  };
  let w = BASE_W;
  let h = BASE_H;
  let overflow = false;
  let outliersClamped = false;

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
      const standard = planDiagramWindows(bounds, services, frameFor(BASE_W, BASE_H));
      // `legible: false` is a request to grow the page, not a verdict that the
      // drawing cannot be tiled. A sparse architecture — the hub-and-spoke
      // every Architecture Center reference draws — defeats the
      // services-per-slide floors purely by having whitespace between its
      // parts, and was handed a 31x32in plotter page nobody can open. Ask the
      // planner for the finest grid the slide budget allows before giving up on
      // ordinary slides; the 7pt floor still governs whether the result is
      // readable, and every part still shares one scale.
      const forced = standard.legible ? null : planDiagramWindows(bounds, services, frameFor(BASE_W, BASE_H), { mustTile: true });
      if (standard.legible) {
        windows = standard.windows;
      } else if (forced && forced.legible && forced.windows.length > 1) {
        windows = forced.windows;
      } else {
        // Only a genuinely enormous drawing gets here — one that cannot be read
        // on nine standard slides. Grow the page for it, then tile that page
        // too: a 56in sheet split into three readable parts still beats one
        // 56in sheet at 4.9pt, which is not a diagram, it is a smudge.
        overflow = wantW > MAX_SLIDE_IN || wantH > MAX_SLIDE_IN;
        w = clamp(wantW, BASE_W, MAX_SLIDE_IN);
        h = clamp(wantH, BASE_H, MAX_SLIDE_IN);
        const grown = planDiagramWindows(bounds, services, frameFor(w, h));
        windows = grown.windows;
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
  };
}

/**
 * Approximate rendered width of a string in inches. CJK characters occupy a
 * full em, Latin about 0.54 em in Yu Gothic UI — good enough to size a chip so
 * its text is not clipped.
 */
function estimateTextWidthIn(text: string, fontSizePt: number): number {
  let units = 0;
  for (const character of text) {
    units += /[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]/.test(character) ? 1 : 0.54;
  }
  return (units * fontSizePt) / 72;
}

/**
 * As much of `text` as will fit in `widthIn` at `fontSizePt`, with an ellipsis
 * for the rest. Used where the tile is too small for the whole name and the
 * only alternatives are unreadable type or an empty box.
 */
function fitLabelToBox(text: string, widthIn: number, fontSizePt: number): string {
  if (estimateTextWidthIn(text, fontSizePt) <= widthIn) return text;
  const budget = widthIn - estimateTextWidthIn('…', fontSizePt);
  let out = '';
  for (const character of text) {
    if (estimateTextWidthIn(out + character, fontSizePt) > budget) break;
    out += character;
  }
  return out.trimEnd() === '' ? '…' : `${out.trimEnd()}…`;
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
  const text = bundle?.badgesOnly ? '' : truncateLabel(route.label, 42);

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
  const badgeD = route.stepNumber === undefined ? 0 : clamp(0.26 * px, 0.18, 0.42);  // A muted rung carries no wording, so it is exactly its callout. Reserving an
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
    const asIs = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / Math.max(maxW - 0.12, 0.05)));
    // Only when the chip does not already fit the band. Widening one that does
    // buys nothing and costs a lean on the columns either side.
    if (asIs * lineH0 + 0.06 > gap) {
      const fits = Math.max(1, Math.floor((gap - 0.06) / lineH0));
      const needed = estimateTextWidthIn(text, fontSize) / fits + 0.12;
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
      const lineCount = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / Math.max(narrowed - 0.12, 0.05)));
      if (lineCount * lineH0 + 0.06 <= 0.9 * px) roomW = narrowed;
    }
  }
  const w = bare
    ? badgeD
    : clamp(naturalW <= roomW ? naturalW : roomW, Math.min(0.34 * px, roomW), roomW);
  const perLine = Math.max(w - 0.12, 0.05);
  const lineH = (fontSize * 1.3) / 72;

  const lines = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / perLine));
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
    for (let ring = 1; bestScore > 0 && ring <= 10; ring += 1) {
      for (let a = -ring; a <= ring && bestScore > 0; a += 1) {
        for (let b = -ring; b <= ring; b += 1) {
          if (Math.max(Math.abs(a), Math.abs(b)) !== ring) continue;
          const candidate = place(stagger + a * (stepOut / 2), b * (w / 2 + 0.06));
          if (!inReach(candidate) || !attributable(candidate)) continue;
          const cost = score(candidate);
          if (cost < bestScore) {
            best = candidate;
            bestScore = cost;
          }
          if (bestScore <= 0) break;
        }
      }
    }
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
    margin: 0.02,
    wrap: true,
    objectName: `connector-label-${route.id}`,
  });
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
): { x: number; y: number; d: number } | null {
  if (route.stepNumber === undefined) return null;

  // No chip to hang off: either an unlabelled but numbered hop, or one whose
  // wording was muted because it had nowhere legible to stand. The anchor is
  // the natural home, but it is the middle of the arrow, which on a dense
  // drawing is routinely the middle of a tile — and a number printed over an
  // icon is the one thing on the slide the workflow list cannot survive
  // without. So walk outwards for a clear slot the way a chip does.
  const anchor = toInches(route.labelAnchor, transform);
  const d = clamp(0.26 * px, 0.18, 0.42);
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
    fontSize,
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
): {
  /** The box the service NAME is drawn in, not the room left over for it. */
  caption: { x: number; y: number; w: number; h: number } | null;
  /** The box the SKU / region / price sub-line is drawn in, when it is shown. */
  meta: { x: number; y: number; w: number; h: number } | null;
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
  const fontSize = clamp(h * 12, 4, 13);
  // At 72 services the overview clamps this to 4pt, which is not small type —
  // it is grey ink the reader cannot resolve, and it makes the thumbnail
  // harder to read rather than more informative. The overview exists to show
  // the shape of the architecture before the reader pans through the readable
  // slices of it, and every one of those names appears in full on the slice
  // that follows, so below the resolvable floor the thumbnail draws the icon
  // and the tile and leaves the naming to them.
  const named = !thumbnail || h * 12 >= OVERVIEW_LEGIBLE_PT;
  // Giving up the name only works when the icon is left to carry the tile. A
  // service with no icon would otherwise be drawn as an empty grey box, which
  // says strictly less than type that is merely small. So the name comes back
  // at exactly the floor, cut to what the tile can hold: a short legible word
  // beats both an empty box and a paragraph of grey mush.
  const stub = !named && !icon;
  const drawnFont = named ? fontSize : OVERVIEW_LEGIBLE_PT;
  const meta = metaSubline(box);
  const metaFontSize = clamp(fontSize - 2, 3.5, 9);
  // Sized from the sub-line's own font, not the name's. Deriving the band from
  // `fontSize` reserved 0.232in for a line needing 0.117in on every tile in the
  // corpus, and on a tight deck that phantom 0.05-0.09in was the whole reason
  // the icon did not fit and was dropped.
  let metaBand = named && showsMeta(h, px) && !!meta ? metaFontSize * 1.55 / 72 + 0.03 : 0;

  const innerW = Math.max(0.05, w - 0.06);
  // How much of the name the tile can actually hold, rather than a flat 40
  // cells. The flat cap clipped names a three-line tile had ample room for —
  // "Azure Database for PostgreSQL フレキシ…" on a tile that fits the whole
  // thing — and what it cut was not written down anywhere, so the reader had
  // no way to recover it. Cut to the tile, and only when the tile is really
  // too small.
  const nameLines = Math.max(1, Math.floor((h - pad * 2 - metaBand) / ((fontSize * 1.22) / 72)));
  const full = fitLabelToBox(box.label, innerW * nameLines, fontSize);
  // A stub gets as much of the name as fits the tile at the floor size, on as
  // many lines as the tile is tall enough for, and an ellipsis for the rest.
  const stubLines = stub
    ? Math.max(1, Math.floor((h - pad * 2) / ((OVERVIEW_LEGIBLE_PT * 1.22) / 72)))
    : 0;
  const label = stub ? fitLabelToBox(full, innerW * stubLines, OVERVIEW_LEGIBLE_PT) : full;
  const labelLines = named ? Math.max(1, Math.ceil(estimateTextWidthIn(label, fontSize) / innerW)) : 0;
  const labelBlockH = (labelLines * fontSize * 1.22) / 72;

  // Which of the three things a tile carries yields when it cannot hold all
  // three. The subline used to win by default — it is subtracted before the
  // icon is measured — and on a 1.36x0.68in tile with a two-line name that
  // left 0.06in for the icon, under the legibility floor, so the icon was
  // dropped. Twenty-five tiles then shipped as grey boxes of type. That is the
  // wrong way round: on the Architecture Center the icon is what says which
  // service this is, and the SKU and the price are supplementary. A tile larger
  // than one that comfortably draws an icon must not lose it to a subline.
  const iconFloor = 0.08 * px;
  const roomFor = (band: number): number =>
    Math.min(h * 0.42, w * 0.34, Math.max(0, h - pad * 2 - band - labelBlockH - 0.02));
  if (metaBand > 0 && icon && named && roomFor(metaBand) < iconFloor && roomFor(0) >= iconFloor) {
    metaBand = 0;
  }

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
  if (named || stub) {
    const boxY = stub ? topLeft.y + pad : textTop;
    const boxH = stub ? Math.max(0.08, h - pad * 2) : textHeight;
    // The box the words are drawn in, not the room left over for them. With no
    // icon the leftover room is nearly the whole tile, and a caption box that
    // claims the whole tile tells every later pass nothing: a chip weighed
    // against it is weighed against the tile it already knew about, and a rule
    // measuring "how much of the name is covered" is really measuring the tile.
    // Vertically centred text inside a shrunk, centred box draws in exactly the
    // same place, so this describes the caption without moving it.
    const drawnH = Math.min(boxH, Math.max(0.08, (Math.max(1, labelLines) * drawnFont * 1.22) / 72));
    const topAligned = !stub && iconSize > 0;
    captionBand = {
      x: topLeft.x + 0.03,
      y: topAligned ? boxY : boxY + (boxH - drawnH) / 2,
      w: innerW,
      h: drawnH,
    };
    slide.addText(label, {
      x: topLeft.x + 0.03,
      y: boxY,
      w: innerW,
      h: boxH,
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
    const drawnH = Math.min(metaBand, (metaPt * 1.22) / 72);
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
  return { caption: captionBand, meta: metaBandRect };
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
): void {
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
  const titleW = Math.max(0.4, w - 0.12);
  // A zone cut to a sliver at the frame edge has less width than its own title
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
  const part = (share: number): number => Math.max(0.4, w * share - 0.12);
  // The bands the tiles actually left free, rather than the fractions of the
  // width somebody guessed at. Fixed shares only work while the row has a
  // quarter of itself spare: fill a subnet — which is what a subnet drawn to
  // scale looks like — and every band on offer, full width, half or third,
  // lands on the same tiles, so a rule that fails a title at 25% coverage had
  // no legal placement left to choose. Reading the gaps finds the one the
  // author left, wherever it happens to be.
  const runs = (y: number): Array<{ x: number; y: number; w: number }> => {
    const lo = topLeft.x + 0.06;
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
  const inside = [
    { x: topLeft.x + 0.06, y: top, w: titleW },
    { x: topLeft.x + 0.06, y: foot, w: titleW },
    ...runs(top),
    ...runs(foot),
    { x: topLeft.x + w - part(0.5) - 0.06, y: top, w: part(0.5) },
    { x: topLeft.x + 0.06, y: top, w: part(0.5) },
    { x: topLeft.x + w - part(0.34) - 0.06, y: top, w: part(0.34) },
    { x: topLeft.x + 0.06, y: top, w: part(0.34) },
    { x: topLeft.x + w - part(0.34) - 0.06, y: foot, w: part(0.34) },
    { x: topLeft.x + 0.06, y: foot, w: part(0.34) },
  ];
  // A fragment is the one exception. Its drawn rectangle is not the zone — it
  // is whatever survived the window cut — so there may be no room inside it at
  // all, and the band just outside the cut is still inside the zone the reader
  // is being shown.
  const outside = clipped
    ? [
      { x: topLeft.x + 0.06, y: topLeft.y - titleH - 0.02, w: titleW },
      { x: topLeft.x + 0.06, y: topLeft.y + h + 0.02, w: titleW },
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
  const score = (c: { x: number; y: number; w: number }): number => {
    const area = Math.max(1e-6, c.w * titleH);
    return cover(c) / area + 2 * (trespass(c) / area);
  };
  let title = candidates[0];
  let best = score(title);
  for (const candidate of candidates.slice(1)) {
    if (best <= 0.01) break;
    const next = score(candidate);
    if (next < best - 1e-6) {
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
    const lo = topLeft.x + 0.06;
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
  slide.addText(truncateLabel(box.label, 60) + fragment, {
    x: title.x,
    y: title.y,
    w: title.w,
    h: bandH,
    fontSize: clamp(Math.round(h * 5), 8, 12),
    bold: true,
    color: labelColor,
    fontFace: 'Yu Gothic UI',
    align: 'left',
    valign: 'top',
    wrap: true,
    lineSpacingMultiple: 0.9,
    objectName: `zone-label-${box.id}`,
  });
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
   * This is the whole drawing shown small ahead of the readable slices of it,
   * so anything that would land under the resolvable floor is left to them.
   */
  thumbnail = false,
  presetIcons?: Map<string, RasterizedIcon>,
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
  const transform = computeFitTransform(drawnView, frame, { maxScale: 1 / PX_PER_IN });
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
  const placedZones = new Map(shownGroups.map((group) => [group.id, placeBox(group, transform, clampTo, true)]));
  shownGroups.forEach((group) => addGroupShape(
    pptx, slide, group, groups.indexOf(group), transform, clampTo,
    { here: zoneMembers(group, shownServices), all: zoneMembers(group, services) },
    placedTiles,
    shownGroups.filter((other) => other !== group).map((other) => placedZones.get(other.id)!),
  ));
  const captionBands: Obstacle[] = [];
  for (const service of shownServices) {
    const bands = addNodeShape(pptx, slide, service, transform, service.iconPath ? icons.get(service.iconPath) : undefined, px, clampTo, thumbnail);
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
  const labelFontSize = clamp(9 * px, 4, 10);
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
      // A chip that still lands on a service name or on another callout after
      // the walk has done its best is worse than no chip: it is drawn at 92%
      // opacity over the one thing that says which service this is, or over a
      // step number. On a slide this crowded the Architecture Center leaves a
      // numbered callout on the arrow and puts the sentence in the step list,
      // which is exactly the trade `mutedWording` already implements — so make
      // it per route, not only per fan.
      if (box && route.stepNumber !== undefined && narratedSteps.has(route.stepNumber)) {
        const b = box.block;
        let spoiled = 0;
        // A chip standing on a service that is not at either end of its own
        // arrow is the same failure by a different route: the reader reads it
        // as that service's caption and the hop it actually describes goes
        // unlabelled. The walk cannot always avoid it — on `meta-subline` and
        // `workflow-prose` every slot within reach laps a bystander, and
        // weighting them twelve times a tile moved the chip not at all — so
        // the wording is handed to the step list exactly as it is when a
        // caption is in the way.
        //
        // Priced separately and much more loosely than the caption bar above.
        // A chip may brush a bystander's rim; the export audit allows a
        // fiftieth of a tile, so this bar sits at exactly that and hands the
        // wording over the moment the drawing would fail the gate.
        // Priced as a FRACTION of the bystander, exactly as the export audit
        // prices it, so the exporter and the gate can never disagree about
        // what counts as standing on a stranger. A flat area bar cannot match
        // it: a fiftieth of a small tile and of a large one are different
        // numbers of square inches, and the difference is what let a 2% lap
        // ship while a 13% one was muted.
        const STRANGER_TILE_FRACTION = 0.02;
        const ownEnds = new Set([route.sourceId, route.targetId]);
        let onStrangers = 0;
        for (const o of chipObstacles) {
          const stranger = !o.annotation && !o.caption
            && o.node !== undefined && !ownEnds.has(o.node);
          if (!o.annotation && !o.caption && !stranger) continue;
          if (o.owner !== undefined && o.owner === bundleKey(route)) continue;
          const dx = Math.min(b.x + b.w, o.x + o.w) - Math.max(b.x, o.x);
          const dy = Math.min(b.y + b.h, o.y + o.h) - Math.max(b.y, o.y);
          if (dx > 0 && dy > 0) {
            if (stranger) onStrangers = Math.max(onStrangers, (dx * dy) / Math.max(1e-6, o.w * o.h));
            else spoiled += dx * dy;
          }
        }
        if (spoiled > SPOILED_CHIP_SQ_IN || onStrangers > STRANGER_TILE_FRACTION) {
          if (!thumbnail && route.label
            && !carriesWording(narratedRows.get(route.stepNumber) ?? '', route.label)) {
            mutedWording.set(route.stepNumber, route.label);
          }
          box = null;
        }
      }
      chips.set(route.id, box);
      if (box) chipObstacles.push(box.block);
      const badge = stepBadgeBox(
        route, transform, px, labelFrame, box, chipObstacles,
        ownGapFor(route), foreignGapFor(route),
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
        else if (rivals.length > 0) {
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
      const slot = chipObstacles.indexOf(box.block);
      if (slot >= 0) chipObstacles[slot] = moved.block;
      chips.set(route.id, moved);
      badges.set(route.id, stepBadgeBox(
        route, transform, px, labelFrame, moved, chipObstacles,
        ownGapFor(route), foreignGapFor(route),
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
        const moved = connectorLabelBox(
          route, transform, labelFontSize, px, labelFrame, pool, bundle,
          undefined, foreignGapFor(route),
        );
        chips.set(route.id, moved);
        badges.set(route.id, stepBadgeBox(
          route, transform, px, labelFrame, moved, pool,
          ownGapFor(route), foreignGapFor(route),
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
  const { diagramName, author, date, isDarkMode } = options;
  const t = isDarkMode ? DARK_THEME : LIGHT_THEME;
  // Number the callouts before anything measures them, so the drawing, the
  // badges and the workflow list are all built from the same edges.
  const diagram = options.diagram
    ? { ...options.diagram, edges: narrateEdgeCallouts(options.diagram.edges ?? []) }
    : options.diagram;

  const pptx = new PptxCtor();
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

  for (const [index, window] of windows.entries()) {
    const slide = pptx.addSlide();
    slide.background = { color: t.bg };
    const partOf = parts.length === 0
      ? ''
      : index === 0
        ? '  (Overview)'
        : `  (${index} / ${parts.length})`;

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
    slide.addText(`${diagramName}${partOf}`, {
      x: 0.35, y: ACCENT_H + 0.05, w: Math.max(3, W - 3.85), h: HEADER_H - 0.1,
      fontSize: 24,
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
      ? await addEditableDiagram(pptx, slide, diagram, geom.frame, isDarkMode, window, mutedWording, window === undefined && parts.length > 0, options.presetIcons)
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
  if (workflow.length > 0) {
    // Rows stop shrinking at a legible minimum, so a long workflow continues on
    // another slide. Dropping the tail would leave badges on the drawing whose
    // sentence appears nowhere in the deck.
    const listTop = IMAGE_Y + 0.1;
    const available = Math.max(MIN_WORKFLOW_ROW_IN, geom.footerY - 0.1 - listTop);
    // The sentence column, measured against the widest the badge is ever
    // allowed to be so the estimate is never optimistic.
    const rowTextW = Math.max(1, W - (0.42 + 0.34 + 0.16) - 0.42);
    const rowHeightIn = (text: string, pt: number): number => {
      const lines = Math.max(1, Math.ceil(estimateTextWidthIn(text, pt) / rowTextW));
      return lines * pt * 1.25 / 72;
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
function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n - 1).trimEnd() + '…' : s;
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
  slide.addText(title, { x: 0.35, y: ACCENT_H + 0.05, w: 9.5, h: HEADER_H - 0.1, fontSize: 22, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true });
  if (meta) {
    slide.addText(meta, { x: 9.9, y: ACCENT_H + 0.05, w: 3.08, h: HEADER_H - 0.1, fontSize: 10, color: t.metaText, fontFace: 'Yu Gothic UI', align: 'right', valign: 'middle' });
  }
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: HEADER_END, w: W, h: SEP_H, fill: { color: t.accent }, line: { color: t.accent, width: 0 } });
  slide.addText('Generated by Microsoft Product Architecture Diagram Builder  ·  Swarm Data SE, Jiayi Yang', { x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H, fontSize: 8, color: t.footerText, fontFace: 'Yu Gothic UI', valign: 'middle' });
}

/** Slide 1 — title / cover. */
function addTitleSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const slide = pptx.addSlide();
  slide.background = { color: t.headerBg };
  // Left accent band
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.28, h: 7.5, fill: { color: t.accent }, line: { color: t.accent, width: 0 } });
  slide.addText('AZURE ARCHITECTURE', { x: 0.9, y: 1.5, w: 11.5, h: 0.4, fontSize: 14, bold: true, color: t.accent, fontFace: 'Yu Gothic UI', charSpacing: 3 });
  slide.addText(o.diagramName, { x: 0.9, y: 2.0, w: 11.5, h: 1.6, fontSize: 40, bold: true, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
  slide.addText(`${o.author}   ·   ${o.date}`, { x: 0.9, y: 3.7, w: 11.5, h: 0.4, fontSize: 14, color: t.metaText, fontFace: 'Yu Gothic UI' });
  if (o.prompt) {
    slide.addText([
      { text: 'Brief:  ', options: { bold: true, color: t.metaText } },
      { text: truncate(o.prompt, 420), options: { color: t.metaText } },
    ], { x: 0.9, y: 4.35, w: 11.5, h: 1.7, fontSize: 13, fontFace: 'Yu Gothic UI', valign: 'top', italic: true });
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
  const frame = { x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H };
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
 */
function addWorkflowSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const steps = (o.workflow ?? [])
    .filter((entry) => entry && readStepValue(entry.step) !== undefined && !!entry.description)
    .sort((a, b) => a.step - b.step);
  if (steps.length === 0) return;

  const MAX_ROWS = 12;
  const shown = steps.slice(0, MAX_ROWS);
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, 'Workflow', `${steps.length} step${steps.length === 1 ? '' : 's'}`);

  // The "+ N more" note is drawn last and used to paint over the final row, so
  // the rows only ever get the space left after reserving it.
  const overflowH = steps.length > MAX_ROWS ? 0.34 : 0;
  const rowH = Math.min(0.62, (BODY_H - overflowH) / shown.length);
  // Long descriptions have to shrink rather than overflow the slide.
  const fontSize = rowH >= 0.5 ? 13 : rowH >= 0.38 ? 11 : 10;
  const badgeD = Math.min(0.34, rowH - 0.06);

  shown.forEach((entry, index) => {
    const y = BODY_TOP + index * rowH;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.4, y: y + (rowH - badgeD) / 2, w: badgeD, h: badgeD,
      fill: { color: t.accent }, line: { color: t.accent, width: 0 },
    });
    slide.addText(String(entry.step), {
      x: 0.4, y: y + (rowH - badgeD) / 2, w: badgeD, h: badgeD,
      fontSize: Math.max(8, Math.round(fontSize * 0.8)), bold: true, color: 'ffffff',
      fontFace: 'Yu Gothic UI', align: 'center', valign: 'middle',
    });
    // A wrapped two-line description used to run straight through the services
    // strip, so the strip is reserved out of the description box's height.
    const services = (entry.services ?? []).filter(Boolean);
    const showsServices = services.length > 0 && rowH >= 0.5;
    slide.addText(truncate(entry.description, 240), {
      x: 0.4 + badgeD + 0.16, y, w: W - 1.1 - badgeD, h: showsServices ? rowH - 0.2 : rowH,
      fontSize, color: t.titleText, fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true,
    });
    if (showsServices) {
      slide.addText(services.join('  →  '), {
        x: 0.4 + badgeD + 0.16, y: y + rowH - 0.2, w: W - 1.1 - badgeD, h: 0.18,
        fontSize: 9, color: t.metaText, fontFace: 'Yu Gothic UI', valign: 'middle',
      });
    }
  });

  if (steps.length > MAX_ROWS) {
    slide.addText(`+ ${steps.length - MAX_ROWS} more steps`, {
      x: 0.4, y: BODY_TOP + shown.length * rowH + 0.04, w: 6, h: 0.3,
      fontSize: 10, italic: true, color: t.footerText, fontFace: 'Yu Gothic UI',
    });
  }
}

/** Slide 4 — service inventory. */
function addServicesSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  if (!o.services.length) return;
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, `Services  ·  ${o.services.length} components`);

  const MAX_ROWS = 20;
  const shown = o.services.slice(0, MAX_ROWS);
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
  slide.addTable([header, ...rows], {
    x: 0.35, y: BODY_TOP, w: W - 0.7, h: BODY_H,
    colW: [5.2, 3.9, 3.53],
    fontSize: 12, fontFace: 'Yu Gothic UI',
    border: { type: 'solid', color: t.headerBg, pt: 1 },
    valign: 'middle', rowH: 0.32,
  });
  if (o.services.length > MAX_ROWS) {
    slide.addText(`+ ${o.services.length - MAX_ROWS} more services`, { x: 0.35, y: FOOTER_Y - 0.32, w: 6, h: 0.3, fontSize: 10, italic: true, color: t.footerText, fontFace: 'Yu Gothic UI' });
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
    slide.addText([
      { text: 'Assessment.  ', options: { bold: true, color: t.titleText } },
      { text: truncate(v.summary, 620), options: { color: t.metaText } },
    ], { x: 3.7, y: BODY_TOP, w: W - 3.7 - 0.35, h: 2.35, fontSize: 12.5, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
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
  const rows = pillars.map((p) => [
    { text: p.pillar, options: { color: t.titleText } },
    { text: p.maturity || scoreBand(p.score), options: { color: t.metaText } },
    { text: `${Math.round(p.score)} / 100`, options: { color: scoreColor(p.score), bold: true, align: 'right' as const } },
  ]);
  slide.addTable([header, ...rows], {
    x: 0.35, y: pTop + 0.36, w: W - 0.7,
    colW: [4.2, 6.6, 1.83],
    fontSize: 12, fontFace: 'Yu Gothic UI',
    border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: 0.34,
  });
}

/** Slide 4b — WAF key findings & recommendations (only when findings exist). */
function addValidationFindingsSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const v = o.validation;
  if (!v || !v.findings.length) return;
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, 'Key findings & recommendations');

  const findings = v.findings.slice(0, 5);
  const rowH = (FOOTER_Y - BODY_TOP - 0.1) / findings.length;
  findings.forEach((f, i) => {
    const y = BODY_TOP + i * rowH;
    // Severity chip
    slide.addText(f.severity.toUpperCase(), { x: 0.35, y: y + 0.05, w: 1.05, h: 0.34, fontSize: 9, bold: true, color: 'ffffff', fill: { color: severityColor(f.severity) }, align: 'center', valign: 'middle', fontFace: 'Yu Gothic UI' });
    // Issue + recommendation
    slide.addText([
      { text: `${f.category}. `, options: { bold: true, color: t.titleText } },
      { text: truncate(f.issue, 170), options: { color: t.metaText } },
    ], { x: 1.55, y: y + 0.02, w: W - 1.95, h: rowH * 0.5, fontSize: 12, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
    if (f.recommendation) {
      slide.addText([
        { text: '→ Fix:  ', options: { bold: true, color: t.accent } },
        { text: truncate(f.recommendation, 220), options: { color: t.metaText, italic: true } },
      ], { x: 1.55, y: y + rowH * 0.5, w: W - 1.95, h: rowH * 0.46, fontSize: 11, fontFace: 'Yu Gothic UI', valign: 'top', wrap: true });
    }
  });
}

/** Slide 5a — cost overview (only when cost provided). */
function addCostOverviewSlide(pptx: PptxGenJS, t: SlideTheme, o: ArchitectureDeckOptions): void {
  const c = o.cost;
  if (!c) return;
  const slide = pptx.addSlide();
  const meta = [c.term, c.region, c.pricesAsOf ? `prices as of ${c.pricesAsOf}` : undefined].filter(Boolean).join('  ·  ');
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
    slide.addText(
      `Regional comparison is partial because selected SKUs are unavailable in: ${truncate(unavailable, 120)}. No cheapest-region recommendation is made.`,
      { x: 0.37, y: BODY_TOP + 3.15, w: 5.2, h: 0.75, fontSize: 10, bold: true, color: 'b45309', fontFace: 'Yu Gothic UI', wrap: true, valign: 'top' },
    );
  }

  slide.addText('Estimate only — not a quote. Excludes taxes, egress, support plans and reservations unless modeled.', { x: 0.37, y: FOOTER_Y - 0.5, w: 5.2, h: 0.45, fontSize: 9, italic: true, color: t.footerText, fontFace: 'Yu Gothic UI', wrap: true });

  // Top cost drivers table (right)
  const svcs = c.topServices.slice(0, 10);
  if (svcs.length) {
    const header = [
      { text: 'Top cost drivers', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Tier', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
      { text: 'Monthly', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
      { text: 'Share', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    ];
    const rows = svcs.map((s) => [
      { text: s.serviceName, options: { color: t.titleText } },
      { text: s.tier || '—', options: { color: t.metaText } },
      { text: money(s.cost, c.currency), options: { color: t.metaText, align: 'right' as const } },
      { text: s.percentage != null ? `${Math.round(s.percentage)}%` : '—', options: { color: t.metaText, align: 'right' as const } },
    ]);
    slide.addTable([header, ...rows], {
      x: 5.9, y: BODY_TOP, w: W - 5.9 - 0.35,
      colW: [3.3, 1.85, 1.15, 0.78],
      fontSize: 12, fontFace: 'Yu Gothic UI',
      border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: 0.32,
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
    slide.addText(
      `Partial comparison — unavailable: ${truncate(unavailable, 150)}. Values below cover comparable regions only; no global cheapest or savings claim is shown.`,
      { x: 0.35, y: BODY_TOP, w: W - 0.7, h: 0.55, fontSize: 12, bold: true, color: 'b45309', fontFace: 'Yu Gothic UI', valign: 'middle', wrap: true },
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
  const header = [
    { text: 'Region', options: { bold: true, color: 'ffffff', fill: { color: t.accent } } },
    { text: 'Monthly', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    { text: 'Annual', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
    { text: comparisonComplete ? 'vs cheapest' : 'vs lowest shown', options: { bold: true, color: 'ffffff', fill: { color: t.accent }, align: 'right' as const } },
  ];
  slide.addTable([header, ...rows], {
    x: 0.35, y: BODY_TOP + 0.6, w: W - 0.7,
    colW: [6.13, 2.0, 2.0, 1.5],
    fontSize: 12, fontFace: 'Yu Gothic UI',
    border: { type: 'solid', color: t.headerBg, pt: 1 }, valign: 'middle', rowH: 0.36,
  });
}

/**
 * Assemble the multi-slide, customer-ready deck for the current architecture:
 * title, diagram, numbered workflow, services, and (when available) a
 * Well-Architected review (summary + findings) and a cost estimate (overview +
 * regional comparison). Split from the download so tests can inspect the deck.
 */
export async function buildArchitectureDeckPptx(
  imageDataUrl: string,
  options: ArchitectureDeckOptions,
): Promise<PptxGenJS> {
  const t = options.isDarkMode ? DARK_THEME : LIGHT_THEME;

  const pptx = new PptxCtor();
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


