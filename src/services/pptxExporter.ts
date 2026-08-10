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
  collectExportBoxes,
  computeBounds,
  computeFitTransform,
  partitionBoxes,
  type ExportBox,
  type ExportRoute,
  type FitTransform,
  type Point,
} from './diagramExportGeometry';

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

// ─── Slide layout constants (inches, LAYOUT_WIDE = 13.33" × 7.5") ────────────

const W = 13.33;  // slide width
const ACCENT_H = 0.07;
const HEADER_H = 0.83;
const HEADER_END = ACCENT_H + HEADER_H;   // 0.9"
const SEP_H = 0.04;
const IMAGE_Y = HEADER_END + SEP_H + 0.06; // ~1.0"
const IMAGE_X = 0.2;
const IMAGE_W = W - IMAGE_X * 2;           // 12.93"
const FOOTER_H = 0.28;
const FOOTER_Y = 7.5 - FOOTER_H - 0.08;   // ~7.14"
const IMAGE_H = FOOTER_Y - IMAGE_Y - 0.1; // remaining body height

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

/** Category tints reused from the interactive HTML export for visual parity. */
const CATEGORY_STYLES: Record<string, { bg: string; border: string }> = {
  'ai + machine learning': { bg: 'E8F0FE', border: '4285F4' },
  'app services': { bg: 'E8F4FD', border: '0078D4' },
  compute: { bg: 'E8F4FD', border: '0078D4' },
  databases: { bg: 'E6F4EA', border: '0B8043' },
  storage: { bg: 'E6F4EA', border: '137333' },
  networking: { bg: 'FFF3E0', border: 'E65100' },
  analytics: { bg: 'F3E8FD', border: '7B1FA2' },
  containers: { bg: 'E0F7FA', border: '00838F' },
  integration: { bg: 'FCE4EC', border: 'C62828' },
  identity: { bg: 'FFF8E1', border: 'F9A825' },
  'management + governance': { bg: 'F1F8E9', border: '558B2F' },
  iot: { bg: 'E0F2F1', border: '00695C' },
  monitor: { bg: 'EDE7F6', border: '5E35B1' },
  security: { bg: 'FFEBEE', border: 'C62828' },
  web: { bg: 'E3F2FD', border: '1565C0' },
  other: { bg: 'FFFFFF', border: '94A3B8' },
};

const GROUP_PALETTE = [
  { bg: 'F0F6FF', border: '0078D4' },
  { bg: 'F0FFF4', border: '00B294' },
  { bg: 'FFFBEB', border: 'D97706' },
  { bg: 'F8F0FF', border: '8764B8' },
  { bg: 'FFF1F2', border: 'D13438' },
  { bg: 'ECFEFF', border: '038387' },
];

function styleForBox(box: ExportBox): { bg: string; border: string } {
  return CATEGORY_STYLES[box.category] ?? CATEGORY_STYLES.other;
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

function addConnector(
  pptx: PptxGenJS,
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  color: string,
): void {
  const points = route.points.map((point) => toInches(point, transform));
  if (points.length < 2) return;

  const lineProps = {
    color,
    width: 1.25,
    dashType: route.dashed ? ('dash' as const) : ('solid' as const),
    endArrowType: 'triangle' as const,
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

function addConnectorLabel(
  slide: Slide,
  route: ExportRoute,
  transform: FitTransform,
  fontSize: number,
): void {
  if (!route.label) return;
  const anchor = toInches(route.labelAnchor, transform);
  const text = truncate(route.label, 42);
  const w = clamp(text.length * fontSize * 0.0105 + 0.16, 0.5, 1.9);
  const h = text.length > 26 ? 0.34 : 0.22;
  slide.addText(text, {
    x: anchor.x - w / 2,
    y: anchor.y - h / 2,
    w,
    h,
    shape: 'roundRect',
    rectRadius: 0.03,
    fill: { color: 'FEF9C3' },
    line: { color: 'FDE68A', width: 0.5 },
    color: 'B45309',
    fontSize,
    fontFace: 'Yu Gothic UI',
    align: 'center',
    valign: 'middle',
    wrap: true,
    objectName: `connector-label-${route.id}`,
  });
}

function addNodeShape(
  pptx: PptxGenJS,
  slide: Slide,
  box: ExportBox,
  transform: FitTransform,
  icon: RasterizedIcon | undefined,
): void {
  const topLeft = toInches({ x: box.x, y: box.y }, transform);
  const w = box.w * transform.scale;
  const h = box.h * transform.scale;
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
  let textTop = topLeft.y + pad;
  let textHeight = h - pad * 2;

  if (icon) {
    const iconSize = clamp(Math.min(h * 0.46, w * 0.36), 0.16, 0.5);
    slide.addImage({
      data: icon.dataUrl,
      x: topLeft.x + (w - iconSize) / 2,
      y: topLeft.y + pad,
      w: iconSize,
      h: iconSize,
      objectName: `icon-${box.id}`,
    });
    textTop = topLeft.y + pad + iconSize + 0.02;
    textHeight = Math.max(0.14, topLeft.y + h - pad - textTop);
  }

  const fontSize = clamp(Math.round(h * 12), 6, 11);
  slide.addText(truncate(box.label, 46), {
    x: topLeft.x + 0.03,
    y: textTop,
    w: w - 0.06,
    h: textHeight,
    fontSize,
    color: '1F2937',
    fontFace: 'Yu Gothic UI',
    align: 'center',
    valign: icon ? 'top' : 'middle',
    wrap: true,
    objectName: `service-label-${box.id}`,
  });
}

function addGroupShape(
  pptx: PptxGenJS,
  slide: Slide,
  box: ExportBox,
  index: number,
  transform: FitTransform,
): void {
  const topLeft = toInches({ x: box.x, y: box.y }, transform);
  const w = box.w * transform.scale;
  const h = box.h * transform.scale;
  const palette = GROUP_PALETTE[index % GROUP_PALETTE.length];

  slide.addShape(pptx.ShapeType.roundRect, {
    x: topLeft.x,
    y: topLeft.y,
    w,
    h,
    rectRadius: 0.06,
    fill: { color: palette.bg, transparency: 15 },
    line: { color: palette.border, width: 1, dashType: 'dash' },
    objectName: `zone-${box.id}`,
  });
  slide.addText(truncate(box.label, 48), {
    x: topLeft.x + 0.06,
    y: topLeft.y + 0.04,
    w: Math.max(0.4, w - 0.12),
    h: 0.24,
    fontSize: clamp(Math.round(h * 5), 8, 12),
    bold: true,
    color: palette.border,
    fontFace: 'Yu Gothic UI',
    align: 'left',
    valign: 'middle',
    objectName: `zone-label-${box.id}`,
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
  isDarkMode: boolean,
): Promise<boolean> {
  const boxes = collectExportBoxes(diagram.nodes ?? []);
  if (boxes.size === 0) return false;
  const { groups, services } = partitionBoxes(boxes);
  if (services.length === 0) return false;

  const bounds = computeBounds(boxes.values());
  const transform = computeFitTransform(bounds, frame, { maxScale: 1 / 96 });
  const routes = buildExportRoutes(diagram.edges ?? [], boxes);
  const icons = await rasterizeIcons(services.map((service) => service.iconPath), 128);

  groups.forEach((group, index) => addGroupShape(pptx, slide, group, index, transform));
  for (const service of services) {
    addNodeShape(pptx, slide, service, transform, service.iconPath ? icons.get(service.iconPath) : undefined);
  }

  const connectorColor = isDarkMode ? '94A3B8' : '64748B';
  for (const route of routes) addConnector(pptx, slide, route, transform, connectorColor);
  // Labels are drawn after every connector so a chip is never hidden by a line
  // that is rendered later.
  const labelFontSize = clamp(Math.round(transform.scale * 850), 7, 10);
  for (const route of routes) addConnectorLabel(slide, route, transform, labelFontSize);

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
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = author;
  pptx.title = diagramName;
  pptx.subject = 'Azure Architecture Diagram';
  pptx.company = 'Microsoft Azure';

  const slide = pptx.addSlide();

  // ── Background ──────────────────────────────────────────────────────────────
  slide.background = { color: t.bg };

  // ── Top accent bar (Azure blue) ─────────────────────────────────────────────
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: W, h: ACCENT_H,
    fill: { color: t.accent },
    line: { color: t.accent, width: 0 },
  });

  // ── Header strip ────────────────────────────────────────────────────────────
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: ACCENT_H, w: W, h: HEADER_H,
    fill: { color: t.headerBg },
    line: { color: t.headerBg, width: 0 },
  });

  // ── Diagram title ────────────────────────────────────────────────────────────
  slide.addText(diagramName, {
    x: 0.35, y: ACCENT_H + 0.05, w: 9.5, h: HEADER_H - 0.1,
    fontSize: 24,
    bold: true,
    color: t.titleText,
    fontFace: 'Yu Gothic UI',
    valign: 'middle',
    wrap: true,
  });

  // ── Author + date (right side of header) ────────────────────────────────────
  slide.addText(`${author}  ·  ${date}`, {
    x: 9.9, y: ACCENT_H + 0.05, w: 3.08, h: HEADER_H - 0.1,
    fontSize: 10,
    color: t.metaText,
    fontFace: 'Yu Gothic UI',
    align: 'right',
    valign: 'middle',
  });

  // ── Thin separator between header and image ──────────────────────────────────
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: HEADER_END, w: W, h: SEP_H,
    fill: { color: t.accent },
    line: { color: t.accent, width: 0 },
  });

  // ── Diagram body — native shapes when available, captured PNG otherwise ─────
  const renderedNatively = options.diagram
    ? await addEditableDiagram(
      pptx,
      slide,
      options.diagram,
      { x: IMAGE_X, y: IMAGE_Y, w: IMAGE_W, h: IMAGE_H },
      isDarkMode,
    )
    : false;

  if (!renderedNatively) {
    slide.addImage({
      data: imageDataUrl,
      x: IMAGE_X,
      y: IMAGE_Y,
      w: IMAGE_W,
      h: IMAGE_H,
      sizing: { type: 'contain', w: IMAGE_W, h: IMAGE_H },
    });
  }

  // ── Footer text ──────────────────────────────────────────────────────────────
  slide.addText('Generated by Azure Architecture Diagram Builder  ·  microsoft.com/azure', {
    x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H,
    fontSize: 8,
    color: t.footerText,
    fontFace: 'Yu Gothic UI',
    valign: 'middle',
  });

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

export interface ArchitectureDeckOptions extends PptxExportOptions {
  /** The original natural-language prompt ("the napkin"). */
  prompt?: string;
  /** Model that generated the architecture. */
  model?: string;
  /** Flat service inventory derived from the diagram nodes. */
  services: DeckService[];
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
  slide.addText('Generated by Azure Architecture Diagram Builder  ·  microsoft.com/azure', { x: 0.35, y: FOOTER_Y, w: W - 0.7, h: FOOTER_H, fontSize: 8, color: t.footerText, fontFace: 'Yu Gothic UI', valign: 'middle' });
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

/** Slide 3 — service inventory. */
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
 * Build and download a multi-slide, customer-ready PPTX deck for the current
 * architecture: title, diagram, services, and (when available) a Well-Architected
 * review (summary + findings) and a cost estimate (overview + regional
 * comparison). Returns the generated filename.
 */
export async function exportArchitectureDeck(
  imageDataUrl: string,
  options: ArchitectureDeckOptions,
): Promise<string> {
  const t = options.isDarkMode ? DARK_THEME : LIGHT_THEME;

  const pptx = new PptxCtor();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = options.author;
  pptx.title = options.diagramName;
  pptx.subject = 'Azure Architecture Review';
  pptx.company = 'Microsoft Azure';

  addTitleSlide(pptx, t, options);
  await addDiagramSlide(pptx, t, imageDataUrl, options);
  addServicesSlide(pptx, t, options);
  addValidationSummarySlide(pptx, t, options);
  addValidationFindingsSlide(pptx, t, options);
  addCostOverviewSlide(pptx, t, options);
  addCostRegionsSlide(pptx, t, options);

  const fileName = generateModelFilename('azure-architecture-deck', 'pptx');
  await pptx.writeFile({ fileName });
  return fileName;
}
