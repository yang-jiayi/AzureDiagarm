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
import type { Edge, Node } from 'reactflow';

/**
 * Interop guard: bundlers hand back the class directly, while Node resolving
 * the CommonJS build hands back `{ default: PptxGenJS }`. Unit tests import
 * this module under Node, so normalise the constructor once.
 */
const PptxCtor = (PptxGenJS as unknown as { default?: typeof PptxGenJS }).default ?? PptxGenJS;

import { generateModelFilename } from '../utils/modelNaming';
import { rasterizeIcons, type RasterizedIcon } from '../utils/exportIconRaster';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  computeBounds,
  computeContentBounds,
  computeFitTransform,
  metaSubline,
  partitionBoxes,
  stripHash,
  truncateLabel,
  usedConnectionLegend,
  workflowListFromEdges,
  zoneStyleFor,
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
  footerText: '475569', // slate-600
};

const LIGHT_THEME: SlideTheme = {
  bg: 'f8fafc',       // slate-50
  headerBg: 'e2e8f0', // slate-200
  accent: '0078d4',   // Azure blue
  titleText: '0f172a', // slate-900
  metaText: '475569',  // slate-600
  footerText: '94a3b8', // slate-400
};

// ─── Slide layout (inches) ───────────────────────────────────────────────────
//
// The deck is normally 16:9 (13.333" × 7.5"). A diagram that would have to be
// squeezed below legible size on that canvas gets a larger custom slide
// instead — PowerPoint accepts any page up to 56", and a bigger page keeps the
// shapes at their true 96 dpi size so labels stay readable and editable.

const PX_PER_IN = 96;
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
   * Horizontal slices of the drawing, one per slide. A single entry (the usual
   * case) means the whole architecture fits on one legible page. More than one
   * means the drawing was too wide to stay readable and was split the way a
   * printed Azure Architecture Center diagram is continued across pages.
   */
  windows: Bounds[];
}

/**
 * Smallest tile label PowerPoint may render. `addNodeShape` derives the label
 * size from the tile height (`h * 12`), so a legible tile needs at least
 * `LEGIBLE_TILE_PT / 12` inches of height. Anything below this is unreadable on
 * a projector and forces the recipient to redraw the deck by hand.
 */
const LEGIBLE_TILE_PT = 7;

/** Never explode one architecture into an unreviewable pile of slides. */
const MAX_DIAGRAM_SLIDES = 6;

/**
 * Shortest workflow row that still fits a 12 pt sentence next to a badge. Rows
 * stop shrinking here and the list continues on another slide instead.
 */
const MIN_WORKFLOW_ROW_IN = 0.34;

/**
 * Split the drawing into as few horizontal bands as keep the tiles legible.
 *
 * PowerPoint allows exactly one page size per deck, so every band shares the
 * same slide geometry and the reader pans left to right across the slides.
 * Returns a single full-width window whenever the diagram already fits, which
 * keeps the common path byte-identical to the previous behaviour.
 */
function planDiagramWindows(
  bounds: Bounds,
  services: ExportBox[],
  frame: DiagramFrame,
): Bounds[] {
  const contentW = Math.max(1, bounds.maxX - bounds.minX);
  const contentH = Math.max(1, bounds.maxY - bounds.minY);
  const whole: Bounds[] = [bounds];
  if (services.length === 0 || frame.w <= 0 || frame.h <= 0) return whole;

  const shortest = Math.min(...services.map((box) => box.h).filter((h) => h > 0));
  if (!Number.isFinite(shortest) || shortest <= 0) return whole;

  // Inches-per-pixel needed for the shortest tile to keep a readable label.
  const legibleScale = LEGIBLE_TILE_PT / 12 / shortest;
  // Height is shared by every band, so it caps the scale no matter how thin the
  // bands get. When even a single-column band cannot be legible, splitting only
  // multiplies the slides without fixing anything.
  const heightScale = frame.h / contentH;
  if (heightScale < legibleScale) return whole;
  if (Math.min(frame.w / contentW, heightScale) >= legibleScale) return whole;

  const bandWidth = frame.w / legibleScale;
  const count = Math.ceil(contentW / bandWidth);
  if (count <= 1 || count > MAX_DIAGRAM_SLIDES) return whole;

  const step = contentW / count;
  return Array.from({ length: count }, (_, index) => ({
    minX: bounds.minX + step * index,
    maxX: index === count - 1 ? bounds.maxX : bounds.minX + step * (index + 1),
    minY: bounds.minY,
    maxY: bounds.maxY,
  }));
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
 * Pick the slide size. Grows the page (never the shrink factor) so a wide
 * architecture keeps 1 : 1 geometry; only diagrams larger than the 56" page
 * limit are scaled down, and then every dimension scales together.
 */
function planSlideGeometry(diagram?: DiagramShapeSource | null): SlideGeometry {
  const chrome = { top: IMAGE_Y, bottom: FOOTER_H + 0.18 + 0.1 };
  let w = BASE_W;
  let h = BASE_H;
  let overflow = false;
  let outliersClamped = false;

  const nodes = diagram?.nodes ?? [];
  let windows: Bounds[] = [];
  if (nodes.length > 0) {
    const boxes = collectExportBoxes(nodes);
    if (boxes.size > 0) {
      const { bounds, clamped } = chooseExportBounds(boxes.values());
      outliersClamped = clamped;
      const contentW = Math.max(1, bounds.maxX - bounds.minX) / PX_PER_IN;
      const contentH = Math.max(1, bounds.maxY - bounds.minY) / PX_PER_IN;
      // Room for the connection legend plus breathing space around the drawing.
      const wantW = contentW + IMAGE_X * 2 + 0.5;
      const wantH = contentH + chrome.top + chrome.bottom + 0.5;
      overflow = wantW > MAX_SLIDE_IN || wantH > MAX_SLIDE_IN;
      w = clamp(wantW, BASE_W, MAX_SLIDE_IN);
      h = clamp(wantH, BASE_H, MAX_SLIDE_IN);
      const footer = h - FOOTER_H - 0.08;
      windows = planDiagramWindows(
        bounds,
        partitionBoxes(boxes).services,
        { x: IMAGE_X, y: IMAGE_Y, w: w - IMAGE_X * 2, h: footer - IMAGE_Y - 0.1 },
      );
      // Splitting restores legibility, so the "scaled down to fit" warning no
      // longer applies — the drawing is now at its readable size.
      if (windows.length > 1) overflow = false;
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
    frame: { x: IMAGE_X, y: IMAGE_Y, w: w - IMAGE_X * 2, h: footerY - IMAGE_Y - 0.1 },
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

function addConnector(
  pptx: PptxGenJS,
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  clampTo?: DiagramFrame,
): void {
  const points = route.points
    .map((point) => toInches(point, transform))
    .map((point) => (clampTo
      ? {
        x: clamp(point.x, clampTo.x, clampTo.x + clampTo.w),
        y: clamp(point.y, clampTo.y, clampTo.y + clampTo.h),
      }
      : point));
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
  fontSize: number,
  px: number,
  clampTo?: DiagramFrame,
): { x: number; y: number; w: number; h: number; text: string } | null {
  if (!route.label) return null;
  const anchor = toInches(route.labelAnchor, transform);
  const text = truncateLabel(route.label, 42);

  // Size the chip from the text it actually carries, capped so it can never
  // dwarf the service tiles it sits between (a 150 px tile is 1.56" at 1 : 1)
  // nor overrun the gap between the two tiles it connects — a long label on a
  // short hop used to be drawn straight across both endpoints.
  const first = toInches(route.points[0] ?? route.labelAnchor, transform);
  const last = toInches(route.points[route.points.length - 1] ?? route.labelAnchor, transform);
  const span = Math.max(Math.abs(last.x - first.x), Math.abs(last.y - first.y));
  const maxW = Math.max(0.34 * px, Math.min(1.5 * px, span > 0 ? span - 0.08 : 1.5 * px));
  const naturalW = estimateTextWidthIn(text, fontSize) + 0.14;
  const w = clamp(naturalW <= maxW ? naturalW : maxW, Math.min(0.34 * px, maxW), maxW);
  const lines = Math.max(1, Math.ceil(estimateTextWidthIn(text, fontSize) / Math.max(w - 0.12, 0.05)));
  const h = Math.max(0.16 * px, (lines * fontSize * 1.3) / 72 + 0.06);

  // De-collide parallel-edge chips: stagger each ordinal clear of the previous
  // one, using the chip's own height so they never overlap.
  const stagger = route.ordinal === 0
    ? 0
    : (route.ordinal % 2 === 1 ? 1 : -1) * Math.ceil(route.ordinal / 2) * (h + 0.04);

  let x = anchor.x - w / 2;
  let y = anchor.y - h / 2 + stagger;
  if (clampTo) {
    x = clamp(x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - w));
    y = clamp(y, clampTo.y, Math.max(clampTo.y, clampTo.y + clampTo.h - h));
  }
  return { x, y, w, h, text };
}

function addConnectorLabel(
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  fontSize: number,
  px: number,
  clampTo?: DiagramFrame,
): void {
  const box = connectorLabelBox(route, transform, fontSize, px, clampTo);
  if (!box) return;

  slide.addText(box.text, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    shape: 'roundRect',
    rectRadius: 0.03,
    fill: { color: 'FEF9C3', transparency: 8 },
    line: { color: 'FDE68A', width: 0.5 },
    color: 'B45309',
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
function addStepBadge(
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  fontSize: number,
  px: number,
  clampTo?: DiagramFrame,
): void {
  if (route.stepNumber === undefined) return;
  const anchor = toInches(route.labelAnchor, transform);
  const d = clamp(0.26 * px, 0.18, 0.42);

  // Sit fully clear of the label chip, measured from the chip's own box: a
  // wrapped CJK label is several lines tall, so a fixed offset used to leave
  // the badge sitting on top of the text.
  const chip = connectorLabelBox(route, transform, fontSize, px, clampTo);
  let x = anchor.x - d / 2;
  let y = chip ? chip.y + chip.h + 0.03 : anchor.y - d / 2;
  if (chip) x = chip.x + chip.w / 2 - d / 2;
  if (clampTo) {
    const bottom = clampTo.y + clampTo.h - d;
    // Clamping the badge down-position back onto the page would push it into
    // the chip it was just placed below — and the chip has itself already been
    // clamped to the page bottom, so the badge would land squarely on the
    // number's own label. Flip it above the chip instead.
    if (chip && y > bottom) {
      const above = chip.y - d - 0.03;
      y = above >= clampTo.y ? above : bottom;
    }
    x = clamp(x, clampTo.x, Math.max(clampTo.x, clampTo.x + clampTo.w - d));
    y = clamp(y, clampTo.y, Math.max(clampTo.y, bottom));
  }

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
): void {
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
  const meta = metaSubline(box);
  const metaFontSize = clamp(fontSize - 2, 3.5, 9);
  const metaBand = showsMeta(h, px) && !!meta ? fontSize * 1.55 / 72 + 0.03 : 0;

  const innerW = Math.max(0.05, w - 0.06);
  const label = truncateLabel(box.label, 40);
  const labelLines = Math.max(1, Math.ceil(estimateTextWidthIn(label, fontSize) / innerW));
  const labelBlockH = (labelLines * fontSize * 1.22) / 72;

  // Fit the icon into whatever vertical room the label does not need, instead
  // of forcing a minimum that pushes the text out of the tile.
  const available = h - pad * 2 - metaBand;
  let iconSize = 0;
  if (icon) {
    iconSize = clamp(Math.min(h * 0.42, w * 0.34, Math.max(0, available - labelBlockH - 0.02)), 0, 0.6);
    if (iconSize < 0.08 * px) iconSize = 0; // too small to read — drop it and keep the words
  }

  if (iconSize > 0 && icon) {
    slide.addImage({
      data: icon.dataUrl,
      x: topLeft.x + (w - iconSize) / 2,
      y: topLeft.y + pad,
      w: iconSize,
      h: iconSize,
      objectName: `icon-${box.id}`,
    });
  }

  const textTop = iconSize > 0 ? topLeft.y + pad + iconSize + 0.02 : topLeft.y + pad;
  const textHeight = Math.max(0.08, topLeft.y + h - pad - metaBand - textTop);

  slide.addText(label, {
    x: topLeft.x + 0.03,
    y: textTop,
    w: innerW,
    h: textHeight,
    fontSize,
    color: '1F2937',
    fontFace: 'Yu Gothic UI',
    align: 'center',
    valign: iconSize > 0 ? 'top' : 'middle',
    margin: 0,
    lineSpacingMultiple: 0.9,
    wrap: true,
    objectName: `service-label-${box.id}`,
  });
  if (metaBand > 0 && meta) {
    slide.addText(truncateLabel(meta, 44), {
      x: topLeft.x + 0.03,
      y: topLeft.y + h - pad - metaBand,
      w: innerW,
      h: metaBand,
      fontSize: metaFontSize,
      color: '64748B',
      fontFace: 'Yu Gothic UI',
      align: 'center',
      valign: 'bottom',
      margin: 0,
      wrap: false,
      objectName: `service-meta-${box.id}`,
    });
  }
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
): void {
  const topLeft = placeBox(box, transform, clampTo, true);
  const w = topLeft.w;
  const h = topLeft.h;
  const palette = zoneStyleFor(box, index);
  const bg = stripHash(palette.bg);
  const border = stripHash(palette.border);

  slide.addShape(pptx.ShapeType.roundRect, {
    x: topLeft.x,
    y: topLeft.y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: bg, transparency: 15 },
    line: { color: border, width: 1, dashType: 'dash' },
    objectName: `zone-${box.id}`,
  });
  // Let a long zone title wrap to two lines instead of clipping at a fixed band.
  const titleH = clamp(h * 0.16, 0.24, 0.5);
  slide.addText(truncateLabel(box.label, 60), {
    x: topLeft.x + 0.06,
    y: topLeft.y + 0.04,
    w: Math.max(0.4, w - 0.12),
    h: titleH,
    fontSize: clamp(Math.round(h * 5), 8, 12),
    bold: true,
    color: border,
    fontFace: 'Yu Gothic UI',
    align: 'left',
    valign: 'top',
    wrap: true,
    lineSpacingMultiple: 0.9,
    objectName: `zone-label-${box.id}`,
  });
}

/** Small colour key so the deck agrees with the PNG's connection legend. */
function addConnectionLegend(
  pptx: PptxGenJS,
  slide: Slide,
  edges: Edge[],
  frame: DiagramFrame,
): void {
  const entries = usedConnectionLegend(edges);
  if (entries.length === 0) return;

  const rowH = 0.2;
  const swatchW = 0.34;
  const boxW = 2.35;
  const boxH = rowH * entries.length + 0.16;
  const x = frame.x + 0.05;
  const y = frame.y + frame.h - boxH - 0.02;

  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: boxW, h: boxH,
    rectRadius: 0.04,
    fill: { color: 'FFFFFF', transparency: 8 },
    line: { color: 'CBD5E1', width: 0.5 },
    objectName: 'connection-legend',
  });
  entries.forEach((entry, i) => {
    const ly = y + 0.08 + i * rowH;
    slide.addShape(pptx.ShapeType.line, {
      x: x + 0.1, y: ly + rowH / 2 - 0.02, w: swatchW, h: 0,
      line: {
        color: stripHash(entry.color),
        width: 1.5,
        dashType: entry.dashed ? (entry.type === 'telemetry' ? 'dashDot' : entry.type === 'async' ? 'dash' : 'sysDot') : 'solid',
      },
    });
    slide.addText(entry.label, {
      x: x + 0.1 + swatchW + 0.08, y: ly - 0.02, w: boxW - swatchW - 0.3, h: rowH,
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
  frame: DiagramFrame,
  _isDarkMode: boolean,
  window?: Bounds,
): Promise<boolean> {
  const boxes = collectExportBoxes(diagram.nodes ?? []);
  if (boxes.size === 0) return false;
  const { groups, services } = partitionBoxes(boxes);
  if (services.length === 0) return false;

  // Size and draw from the SAME bounds. Sizing the page for the dense cluster
  // while drawing every box is what silently pushed far-placed services off
  // the slide, so when outliers are excluded from the fit they are clamped
  // back onto the page instead of being drawn into the void.
  const { bounds, clamped } = chooseExportBounds(boxes.values());
  // A banded slide is sized from its own slice, which is what buys back the
  // legible scale; the slice is then clamped so a shape straddling the seam is
  // cut at the page edge instead of spilling into the void.
  const banded = !!window && (window.minX > bounds.minX + 0.5 || window.maxX < bounds.maxX - 0.5);
  const fitBounds = window ?? bounds;
  const transform = computeFitTransform(fitBounds, frame, { maxScale: 1 / PX_PER_IN });
  const clampTo = clamped || banded ? frame : undefined;
  const px = transform.scale * PX_PER_IN;
  const routes = buildExportRoutes(diagram.edges ?? [], boxes);
  const isFirstBand = !banded || fitBounds.minX <= bounds.minX + 0.5;
  const isLastBand = !banded || fitBounds.maxX >= bounds.maxX - 0.5;

  // A band owns whatever falls inside it, so a shape straddling a seam is drawn
  // once instead of twice. Strays that the outlier trim pushed outside `bounds`
  // sit in no band at all under a plain range test and vanish from the deck
  // entirely, so the outer bands claim everything beyond them and `clampTo`
  // pulls those back onto the page exactly as it does on an unbanded slide.
  const owns = (x: number): boolean => {
    if (!banded) return true;
    if (x < fitBounds.minX) return isFirstBand;
    if (x > fitBounds.maxX) return isLastBand;
    // Windows meet exactly at a seam, so a point landing on one is inside both.
    // Half-open ranges hand it to the later band and nothing is drawn twice.
    return x < fitBounds.maxX || isLastBand;
  };
  const visibleBox = (box: ExportBox): boolean => owns(box.x + box.w / 2);
  // A zone is routinely wider than a whole band, so centre-ownership would
  // print the boundary and its name on one slide and leave the services on the
  // other slides floating with no container. Unlike a service tile, a zone is
  // continued on every band it overlaps — the palette index below is already
  // stable across slices, and a partial rectangle reads as a boundary that
  // carries on, which is exactly what it does.
  const visibleGroup = (box: ExportBox): boolean => {
    if (!banded) return true;
    return (box.x + box.w >= fitBounds.minX || isFirstBand)
      && (box.x <= fitBounds.maxX || isLastBand);
  };
  // A connector is continued on every band it crosses so the reader can follow
  // where it goes; only the band holding its anchor draws the chip and number.
  const visibleRoute = (route: ExportRoute): boolean => {
    if (!banded) return true;
    const xs = route.points.map((point) => point.x);
    if (xs.length === 0) return owns(route.labelAnchor.x);
    return (Math.max(...xs) >= fitBounds.minX || isFirstBand)
      && (Math.min(...xs) <= fitBounds.maxX || isLastBand);
  };
  const shownGroups = groups.filter(visibleGroup);
  const shownServices = services.filter(visibleBox);
  const shownRoutes = routes.filter(visibleRoute);
  const annotatedRoutes = shownRoutes.filter((route) => owns(route.labelAnchor.x));
  const icons = await rasterizeIcons(shownServices.map((service) => service.iconPath), 128);

  // Index by the full group list so a zone keeps its palette colour on every
  // slice it appears on.
  shownGroups.forEach((group) => addGroupShape(pptx, slide, group, groups.indexOf(group), transform, clampTo));
  for (const service of shownServices) {
    addNodeShape(pptx, slide, service, transform, service.iconPath ? icons.get(service.iconPath) : undefined, px, clampTo);
  }

  for (const route of shownRoutes) addConnector(pptx, slide, route, transform, clampTo);
  // Labels are drawn after every connector so a chip is never hidden by a line
  // that is rendered later.
  const labelFontSize = clamp(9 * px, 4, 10);
  for (const route of annotatedRoutes) addConnectorLabel(slide, route, transform, labelFontSize, px, clampTo);
  for (const route of annotatedRoutes) addStepBadge(slide, route, transform, labelFontSize, px, clampTo);

  // Colour key so the deck's connectors agree with the PNG legend.
  addConnectionLegend(pptx, slide, diagram.edges ?? [], frame);

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

  const pptx = new PptxCtor();
  const geom = planSlideGeometry(options.diagram);
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

  const windows = geom.windows.length > 1 ? geom.windows : [undefined];
  let renderedNatively = false;

  for (const [index, window] of windows.entries()) {
    const slide = pptx.addSlide();
    slide.background = { color: t.bg };
    const partOf = windows.length > 1 ? `  (${index + 1} / ${windows.length})` : '';

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
    renderedNatively = options.diagram
      ? await addEditableDiagram(pptx, slide, options.diagram, geom.frame, isDarkMode, window)
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
    const note = windows.length > 1
      ? `This architecture is too wide for one readable page, so it continues across ${windows.length} slides — this is part ${index + 1}. Export to Visio (.vsdx) for the whole drawing on a single sheet.`
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
        color: 'B45309',
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
  const workflow = workflowListFromEdges(options.diagram?.edges ?? []);
  if (workflow.length > 0) {
    // Rows stop shrinking at a legible minimum, so a long workflow continues on
    // another slide. Dropping the tail would leave badges on the drawing whose
    // sentence appears nowhere in the deck.
    const listTop = IMAGE_Y + 0.1;
    const available = Math.max(MIN_WORKFLOW_ROW_IN, geom.footerY - 0.1 - listTop);
    const perSlide = Math.max(1, Math.floor(available / MIN_WORKFLOW_ROW_IN));
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
      const rowGap = Math.min(0.62, Math.max(MIN_WORKFLOW_ROW_IN, available / rows.length));
      const badge = Math.min(0.34, rowGap - 0.06);
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
          fontSize: 12, color: t.titleText, fontFace: 'Yu Gothic UI',
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
  await pptx.writeFile({ fileName });
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

/** Slide 2 — the diagram, drawn with native (editable) PowerPoint shapes. */
async function addDiagramSlide(pptx: PptxGenJS, t: SlideTheme, imageDataUrl: string, o: ArchitectureDeckOptions): Promise<void> {
  const slide = pptx.addSlide();
  addChrome(pptx, slide, t, o.diagramName, `${o.author}  ·  ${o.date}`);
  const renderedNatively = o.diagram
    ? await addEditableDiagram(
      pptx,
      slide,
      o.diagram,
      { x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H },
      o.isDarkMode,
    )
    : false;
  if (!renderedNatively) {
    slide.addImage({ data: imageDataUrl, x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H, sizing: { type: 'contain', w: IMAGE_W, h: IMAGE_H } });
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
  await pptx.writeFile({ fileName });
  return fileName;
}
