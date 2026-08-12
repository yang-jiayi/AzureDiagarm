// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Shared, pure geometry helpers for the "native shape" exporters
 * (PowerPoint / Visio).
 *
 * The canvas stores React Flow coordinates in pixels with a top-left origin and
 * child nodes positioned relative to their parent group. Office formats need
 * absolute coordinates in inches, so every exporter needs the same three steps:
 *
 *   1. flatten node positions to an absolute pixel box,
 *   2. fit the drawing into the target frame (slide body / Visio page),
 *   3. route each edge as an orthogonal polyline between box anchors.
 *
 * Keeping this DOM-free makes it unit-testable and keeps the two exporters
 * visually consistent.
 */

import type { Edge, Node } from 'reactflow';
import { readStepNumber as readStepValue } from '../utils/workflowStepMapping';
import {
  getConnectionPresentation,
  normalizeConnectionType,
  type DiagramConnectionType,
} from '../utils/edgePresentation';

export type { DiagramConnectionType } from '../utils/edgePresentation';

export const DEFAULT_SERVICE_W = 150;
export const DEFAULT_SERVICE_H = 75;
export const DEFAULT_GROUP_W = 400;
export const DEFAULT_GROUP_H = 300;

/**
 * How far the outermost arrow of a parallel fan may sit from the straight line
 * between its two services, in diagram pixels. Half a service tile's height:
 * past that the arrow — and the label that follows it — is closer to the row
 * above than to the hop it belongs to.
 */
const MAX_FAN_SPREAD = DEFAULT_SERVICE_H / 2;

export interface ExportBox {
  id: string;
  kind: 'group' | 'service';
  label: string;
  iconPath?: string;
  /**
   * Azure icon-folder category (lower-cased, e.g. `databases`). Drives the
   * colour coding in every export. Resolved from the node's own `category`
   * first because AI-generated services keep their category even when no icon
   * file matched and `iconPath` is empty.
   */
  category: string;
  /** Canonical service name when it differs from the user-editable label. */
  serviceName?: string;
  /**
   * User-picked zone colour (GroupNode `data.customColor`). When present it
   * overrides the index-based fallback palette so a red zone stays red in every
   * export. `bg` may be an `rgba(...)` string; exporters derive a solid tint.
   */
  customColor?: { bg?: string; border?: string; header?: string };
  /**
   * Optional SKU / region / monthly-cost annotation mirrored from the canvas
   * tile, so native exports can render the same sub-line instead of dropping it.
   */
  meta?: BoxMeta;
  /**
   * The zone this box declares as its container (`Node.parentNode`), when it
   * has one. Membership has to be declared, not inferred from geometry: an
   * Architecture Center security diagram routinely draws a compliance boundary
   * straight across a drawing, and a purely geometric test let that boundary
   * claim — and drag into the margin — half a grid of services that belong to
   * a different container entirely.
   */
  parent?: string;
  /** Absolute pixel geometry (top-left origin). */
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface BoxMeta {
  /** SKU / tier, e.g. `Premium P1v3`. */
  sku?: string;
  /** Azure region, e.g. `eastus` (omitted when unknown). */
  region?: string;
  /** Total monthly cost in USD, when priced. */
  cost?: number;
  /** Pre-formatted cost label, e.g. `$120/mo` / `Free`. */
  costLabel?: string;
}

export interface Point {
  x: number;
  y: number;
}

export interface ExportRoute {
  id: string;
  sourceId: string;
  targetId: string;
  label: string;
  /** Canonical semantic type; drives colour + dash in every exporter. */
  connectionType: DiagramConnectionType;
  /** Canonical stroke colour (`#rrggbb`) for this connection type. */
  color: string;
  dashed: boolean;
  /** SVG-style dash pattern (e.g. `6, 5`) when the type/edge is dashed. */
  dashPattern?: string;
  /** Line opacity (optional/faded connectors render < 1). */
  opacity: number;
  /** 0-based index among parallel edges sharing the same endpoint pair. */
  ordinal: number;
  /**
   * Perpendicular offset, in pixels, this route was fanned by. The fan
   * alternates about the centre (0, +16, -16 …), so a label ladder ordered by
   * `ordinal` runs in a different order from the arrows themselves and every
   * callout ends up beside the wrong one. Rank by this instead.
   */
  fanOffset: number;
  /**
   * True when the edge was drawn with arrowheads at both ends, so exporters
   * add a start arrow as well. `sourceId`/`targetId` are already oriented to
   * the drawn arrow, so a one-way head at the target end is always correct.
   */
  bidirectional: boolean;
  /** True for a source===target edge rendered as a small loop stub. */
  isSelfLoop: boolean;
  /** Orthogonal polyline in absolute pixels; first point starts at the source. */
  points: Point[];
  /** Preferred label anchor (centre of the middle segment). */
  labelAnchor: Point;
  /**
   * Workflow step this arrow carries, when the diagram declares one.
   *
   * Every reference architecture on the Azure Architecture Center numbers its
   * arrows and repeats the numbers in the prose beneath, so exporters draw a
   * numbered callout here to keep the picture and the narrative linked.
   */
  stepNumber?: number;
}

export interface Bounds {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
}

export interface FitTransform {
  /** Multiply a pixel value by this to get inches. */
  scale: number;
  /** Inches added after scaling (already includes the frame origin). */
  offsetX: number;
  offsetY: number;
}

function readSize(node: Node): { w: number; h: number } {
  const anyNode = node as any;
  const styleW = typeof anyNode.style?.width === 'number' ? anyNode.style.width : undefined;
  const styleH = typeof anyNode.style?.height === 'number' ? anyNode.style.height : undefined;
  if (node.type === 'groupNode') {
    return {
      w: styleW ?? anyNode.width ?? DEFAULT_GROUP_W,
      h: styleH ?? anyNode.height ?? DEFAULT_GROUP_H,
    };
  }
  return {
    w: anyNode.width ?? styleW ?? DEFAULT_SERVICE_W,
    h: anyNode.height ?? styleH ?? DEFAULT_SERVICE_H,
  };
}

/**
 * Resolve every node's absolute canvas position by walking the full
 * `parentNode` chain. Groups can be nested (a zone inside a zone), and React
 * Flow stores each child relative to its immediate parent, so a single-level
 * lookup would drop nested content near the canvas origin. The `visiting` set
 * makes a corrupt cyclic parent chain fall back to the relative position
 * instead of recursing forever.
 */
function resolveAbsolutePositions(nodes: Node[]): Map<string, Point> {
  const nodesById = new Map(nodes.map((node) => [node.id, node]));
  const resolved = new Map<string, Point>();

  const resolve = (node: Node, visiting: Set<string>): Point => {
    const cached = resolved.get(node.id);
    if (cached) return cached;

    const own: Point = { x: node.position?.x ?? 0, y: node.position?.y ?? 0 };
    if (visiting.has(node.id)) return own;

    const parentId = (node as { parentNode?: string }).parentNode;
    const parent = parentId ? nodesById.get(parentId) : undefined;
    let position = own;
    if (parent) {
      const nextVisiting = new Set(visiting);
      nextVisiting.add(node.id);
      const parentPosition = resolve(parent, nextVisiting);
      position = { x: parentPosition.x + own.x, y: parentPosition.y + own.y };
    }
    resolved.set(node.id, position);
    return position;
  };

  for (const node of nodes) resolve(node, new Set());
  return resolved;
}

/**
 * Normalize an Azure icon-folder name into a category key.
 * Icon paths look like `/Azure_Public_Service_Icons/Icons/<folder>/<file>.svg`.
 */
export function categoryFromIconPath(iconPath?: string): string | undefined {
  const match = iconPath?.match(/\/Icons\/([^/]+)\//i);
  return match ? match[1].replace(/[-_]+/g, ' ').trim().toLowerCase() : undefined;
}

/**
 * Resolve the category used for export colour coding. The node's own category
 * wins: AI-generated nodes are stored with `iconPath: icon?.path || ''`, so a
 * service with no matching icon file would otherwise lose its colour and be
 * rendered as an unstyled grey tile.
 */
function resolveCategory(data: Record<string, unknown>, iconPath?: string): string {
  const explicit = typeof data.category === 'string' ? data.category.replace(/[-_]+/g, ' ').trim().toLowerCase() : '';
  return explicit || categoryFromIconPath(iconPath) || 'other';
}

// ─── Canonical colour palette (single source of truth for every exporter) ────

export interface CategoryStyle {
  /** Light fill tint, `#rrggbb`. */
  bg: string;
  /** Accent / border colour, `#rrggbb`. */
  border: string;
  /** Readable text colour, `#rrggbb`. */
  text: string;
}

/**
 * Azure-category → colour. Hoisted here so PowerPoint, Visio, Draw.io and HTML
 * all tint the same service identically (they had drifted into three different
 * maps). Keys are the normalised icon-folder categories.
 */
export const CATEGORY_STYLES: Record<string, CategoryStyle> = {
  'ai + machine learning': { bg: '#E8F0FE', border: '#4285F4', text: '#1A73E8' },
  'app services': { bg: '#E8F4FD', border: '#0078D4', text: '#004578' },
  compute: { bg: '#E8F4FD', border: '#0078D4', text: '#004578' },
  databases: { bg: '#E6F4EA', border: '#0B8043', text: '#0B6B3A' },
  storage: { bg: '#E6F4EA', border: '#137333', text: '#0B6B3A' },
  networking: { bg: '#FFF3E0', border: '#E65100', text: '#BF360C' },
  analytics: { bg: '#F3E8FD', border: '#7B1FA2', text: '#6A1B9A' },
  containers: { bg: '#E0F7FA', border: '#00838F', text: '#006064' },
  integration: { bg: '#FCE4EC', border: '#C62828', text: '#B71C1C' },
  identity: { bg: '#FFF8E1', border: '#F9A825', text: '#F57F17' },
  'management + governance': { bg: '#F1F8E9', border: '#558B2F', text: '#33691E' },
  iot: { bg: '#E0F2F1', border: '#00695C', text: '#004D40' },
  monitor: { bg: '#EDE7F6', border: '#5E35B1', text: '#4527A0' },
  security: { bg: '#FFEBEE', border: '#C62828', text: '#B71C1C' },
  web: { bg: '#E3F2FD', border: '#1565C0', text: '#0D47A1' },
  other: { bg: '#F5F5F5', border: '#616161', text: '#424242' },
};

export const DEFAULT_CATEGORY_STYLE: CategoryStyle = CATEGORY_STYLES.other;

/** Resolve the colour for a service category, falling back to a neutral tint. */
export function categoryStyle(category?: string): CategoryStyle {
  const key = typeof category === 'string' ? category.trim().toLowerCase() : '';
  return CATEGORY_STYLES[key] ?? DEFAULT_CATEGORY_STYLE;
}

export interface ZoneStyle {
  bg: string;
  border: string;
  text: string;
}

/** Fallback zone palette, cycled by group index and identical across formats. */
export const ZONE_PALETTE: ZoneStyle[] = [
  { bg: '#F0F6FF', border: '#0078D4', text: '#12395B' },
  { bg: '#F0FFF4', border: '#00B294', text: '#04463A' },
  { bg: '#FFFBEB', border: '#D97706', text: '#5A3200' },
  { bg: '#F8F0FF', border: '#8764B8', text: '#3B2557' },
  { bg: '#FFF1F2', border: '#D13438', text: '#5A1417' },
  { bg: '#ECFEFF', border: '#038387', text: '#023B3D' },
];

// ─── Colour helpers ──────────────────────────────────────────────────────────

/** Normalise `#rgb` / `#rrggbb` / `rgb[a](...)` into `#rrggbb`, else undefined. */
export function normalizeHex(color?: string): string | undefined {
  if (typeof color !== 'string') return undefined;
  const value = color.trim();
  const six = /^#?([0-9a-fA-F]{6})$/.exec(value);
  if (six) return `#${six[1].toLowerCase()}`;
  const three = /^#?([0-9a-fA-F]{3})$/.exec(value);
  if (three) {
    const [r, g, b] = three[1].split('');
    return `#${r}${r}${g}${g}${b}${b}`.toLowerCase();
  }
  const rgb = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value);
  if (rgb) {
    const toHex = (n: string) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, '0');
    return `#${toHex(rgb[1])}${toHex(rgb[2])}${toHex(rgb[3])}`.toLowerCase();
  }
  return undefined;
}

/** Drop a leading `#` — PowerPoint/pptxgenjs colours are bare hex. */
export function stripHash(color: string): string {
  return color.replace(/^#/, '');
}

/** Blend `hex` toward white by `weight` (0..1 = share of the original colour). */
export function mixWithWhite(hex: string, weight: number): string {
  const base = normalizeHex(hex) ?? '#f5f5f5';
  const r = parseInt(base.slice(1, 3), 16);
  const g = parseInt(base.slice(3, 5), 16);
  const b = parseInt(base.slice(5, 7), 16);
  const mix = (v: number) => Math.round(v * weight + 255 * (1 - weight));
  const toHex = (v: number) => mix(v).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

/**
 * sRGB relative luminance (WCAG 2.1) of a `#rrggbb` colour.
 */
export function relativeLuminance(hex: string): number {
  const base = normalizeHex(hex) ?? '#000000';
  const channel = (i: number) => {
    const c = parseInt(base.slice(i, i + 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  };
  return 0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5);
}

/** WCAG 2.1 contrast ratio between two colours, 1..21. */
export function contrastRatio(a: string, b: string): number {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/**
 * The nearest version of `color` that is legible as text on `background`.
 *
 * Scales the colour toward black or toward white — whichever direction the
 * background allows — until it clears `target`, keeping the hue so the drawing
 * still reads as the user's chosen palette. Falls back to plain black or white
 * when even the extreme does not reach the target.
 */
export function readableTextOn(color: string, background: string, target = 4.5): string {
  const base = normalizeHex(color) ?? '#000000';
  const bg = normalizeHex(background) ?? '#ffffff';
  if (contrastRatio(base, bg) >= target) return base;
  const r = parseInt(base.slice(1, 3), 16);
  const g = parseInt(base.slice(3, 5), 16);
  const b = parseInt(base.slice(5, 7), 16);
  // Which extreme to walk toward. Black and white give equal contrast at a
  // background luminance of 0.179, not 0.5 — the WCAG ratio is not linear in
  // luminance. Asking which extreme actually wins states that without the
  // magic number, and makes the fallback below correct by construction: on a
  // mid-grey (0.18 < L < 0.5) the old test walked toward black and then
  // returned black, when white was the readable choice.
  const toward = contrastRatio('#000000', bg) >= contrastRatio('#ffffff', bg) ? 0 : 255;
  for (let step = 1; step <= 20; step += 1) {
    const t = step / 20;
    const mix = (v: number) => Math.round(v * (1 - t) + toward * t).toString(16).padStart(2, '0');
    const candidate = `#${mix(r)}${mix(g)}${mix(b)}`;
    if (contrastRatio(candidate, bg) >= target) return candidate;
  }
  return toward === 0 ? '#000000' : '#ffffff';
}

/**
 * Resolve the colour for a zone. The user's picked colour wins; otherwise the
 * shared palette is cycled by group index so the fallback is identical in every
 * export.
 */
export function zoneStyleFor(box: ExportBox, index: number): ZoneStyle {
  const border = normalizeHex(box.customColor?.border) ?? normalizeHex(box.customColor?.header);
  if (border) {
    const bg = mixWithWhite(border, 0.14);
    // A picked colour is chosen to look right as a border, where contrast does
    // not matter. Reused verbatim as label text it routinely lands at 2-3:1 —
    // the zone title, which names the tier, becomes the least readable words on
    // the slide. Darken it until it clears AA against its own panel.
    return { bg, border, text: readableTextOn(border, bg) };
  }
  return ZONE_PALETTE[index % ZONE_PALETTE.length];
}

// ─── Connection styling (mirrors the on-canvas / PNG legend) ─────────────────

export interface ConnectionStyle {
  type: DiagramConnectionType;
  /** Stroke colour, `#rrggbb`. */
  color: string;
  dashed: boolean;
  /** SVG dash pattern when dashed, e.g. `6, 5`. */
  dashPattern?: string;
  opacity: number;
}

/**
 * Canonical connection-type → colour + dash, taken straight from the shared
 * {@link getConnectionPresentation} the canvas uses, so an exported deck can
 * never contradict the PNG's colour-coded legend.
 */
export function connectionStyleFor(type: DiagramConnectionType): ConnectionStyle {
  const presentation = getConnectionPresentation(type);
  return {
    type: presentation.type,
    color: presentation.stroke,
    dashed: !!presentation.strokeDasharray,
    dashPattern: presentation.strokeDasharray,
    opacity: presentation.opacity ?? 1,
  };
}

export interface ConnectionLegendEntry {
  type: DiagramConnectionType;
  label: string;
  description: string;
  color: string;
  dashed: boolean;
  dashPattern?: string;
}

/** The five connection types, in canvas order, for a colour key in exports. */
export const CONNECTION_LEGEND: ConnectionLegendEntry[] = (
  ['sync', 'async', 'optional', 'security', 'telemetry'] as DiagramConnectionType[]
).map((type) => {
  const style = connectionStyleFor(type);
  const meta: Record<DiagramConnectionType, { label: string; description: string }> = {
    sync: { label: 'Synchronous', description: 'Real-time, request-response (HTTP, SQL)' },
    async: { label: 'Asynchronous', description: 'Message-based, event-driven (queues, events)' },
    optional: { label: 'Optional', description: 'Conditional, fallback paths' },
    security: { label: 'Security', description: 'Identity, trust, and policy enforcement' },
    telemetry: { label: 'Telemetry', description: 'Metrics, logs, traces, and diagnostics' },
  };
  return {
    type,
    label: meta[type].label,
    description: meta[type].description,
    color: style.color,
    dashed: style.dashed,
    dashPattern: style.dashPattern,
  };
});

/** Which of the five connection types actually appear on these edges. */
export function usedConnectionLegend(edges: Edge[]): ConnectionLegendEntry[] {
  const used = new Set<DiagramConnectionType>(
    edges.map((edge) => normalizeConnectionType((edge.data as { connectionType?: unknown } | undefined)?.connectionType)),
  );
  return CONNECTION_LEGEND.filter((entry) => used.has(entry.type));
}

/** One numbered row of the Azure Architecture Center style workflow list. */
export interface WorkflowListEntry {
  step: number;
  description: string;
}

/**
 * Collect the numbered workflow narration carried by the edges.
 *
 * Every Azure Architecture Center diagram pairs its numbered callouts with a
 * numbered prose list; a badge with no matching sentence tells the reader
 * nothing. Exporters call this so the list is built from the same edge data
 * that produced the badges and can never drift from them.
 *
 * Duplicate step numbers keep the first description, and the result is sorted
 * so the list reads in flow order regardless of edge ordering.
 */
export function workflowListFromEdges(edges: Edge[]): WorkflowListEntry[] {
  const byStep = new Map<number, string>();
  for (const edge of edges) {
    const data = edge.data as { stepNumber?: unknown; stepDescription?: unknown } | undefined;
    const step = readStepValue(data?.stepNumber);
    if (step === undefined) continue;
    const description = typeof data?.stepDescription === 'string' ? data.stepDescription.trim() : '';
    if (!description || byStep.has(step)) continue;
    byStep.set(step, description);
  }
  return [...byStep.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([step, description]) => ({ step, description }));
}

/**
 * Make sure every numbered callout the drawing will show has a sentence to
 * point at, and that every label in a deep parallel fan carries a number.
 *
 * A model asked for "the app talks to the database" happily emits five or ten
 * separate connections between the same pair, each with its own sentence. Ten
 * sentences cannot be written along one arrow, so the Architecture Center draws
 * the numbers on the drawing and the sentences in the workflow list. That only
 * works if each of those arrows actually carries a number: the step mapper
 * gives at most one number per node pair, so the rest of the fan would have no
 * callout and no row, and dropping its chip would delete the author's words.
 *
 * So: any labelled edge that already has a number but no sentence is given its
 * own label as the sentence — a badge nobody can look up is worse than no badge
 * — and the un-numbered members of a deep fan are handed their own steps,
 * continuing after the highest number already in use so nothing is renumbered.
 */
/**
 * Whether a workflow row already says what an arrow's own label says. Compared
 * with punctuation and case folded away, because the same wording routinely
 * differs by a trailing full stop or a capital between the two fields.
 *
 * Used when a chip is muted: the trade the exporter makes is "the workflow
 * slide carries this wording instead", which is only a trade if the row
 * actually contains it.
 */
export function carriesWording(description: string, label: string): boolean {
  const fold = (s: string): string => s
    .toLowerCase()
    .replace(/[\s\u3000]+/g, '')
    .replace(/[.,;:!?、。（）()[\]「」"'`´’‘“”-]/g, '');
  const needle = fold(label);
  return needle.length === 0 || fold(description).includes(needle);
}

export function narrateEdgeCallouts(edges: Edge[], minFanSize = 5): Edge[] {
  const labelled = edges.filter((edge) => readEdgeLabel(edge) !== '');

  const pairKey = (edge: Edge): string => (edge.source < edge.target
    ? `${edge.source}|${edge.target}`
    : `${edge.target}|${edge.source}`);
  const fanSize = new Map<string, number>();
  let nextStep = 0;
  for (const edge of labelled) {
    const key = pairKey(edge);
    fanSize.set(key, (fanSize.get(key) ?? 0) + 1);
  }
  for (const edge of edges) {
    const step = readStepValue((edge.data as { stepNumber?: unknown } | undefined)?.stepNumber);
    if (step !== undefined) nextStep = Math.max(nextStep, step);
  }

  let changed = false;
  // Two arrows may not carry the same number. The workflow list is keyed by
  // number, so the second and later sentences were dropped without a trace
  // while every one of those badges still read the same digit — five callouts
  // all saying "3" and one row to look them up in. The first occurrence keeps
  // the author's number; the rest continue after the highest already in use.
  const used = new Set<number>();
  const renumbered = new Map<Edge, number>();
  for (const edge of edges) {
    if (readEdgeLabel(edge) === '') continue;
    const step = readStepValue((edge.data as { stepNumber?: unknown } | undefined)?.stepNumber);
    if (step === undefined) continue;
    if (!used.has(step)) {
      used.add(step);
      continue;
    }
    nextStep += 1;
    used.add(nextStep);
    renumbered.set(edge, nextStep);
  }

  const result = edges.map((edge) => {
    const label = readEdgeLabel(edge);
    if (!label) return edge;
    const data = edge.data as { stepNumber?: unknown; stepDescription?: unknown } | undefined;
    const step = readStepValue(data?.stepNumber);
    const description = typeof data?.stepDescription === 'string' ? data.stepDescription.trim() : '';
    const fresh = renumbered.get(edge);
    if (fresh !== undefined) {
      changed = true;
      return {
        ...edge,
        data: { ...(data ?? {}), stepNumber: fresh, stepDescription: description || label },
      };
    }
    if (step !== undefined && description) return edge;
    // A number with no sentence is as useless as a sentence with no number: the
    // badge lands on the drawing and the workflow list has no row to read it
    // from. Only a fan is deep enough to be worth numbering from scratch.
    if (step === undefined && (fanSize.get(pairKey(edge)) ?? 0) < minFanSize) return edge;
    changed = true;
    if (step === undefined) nextStep += 1;
    return {
      ...edge,
      data: { ...(data ?? {}), stepNumber: step ?? nextStep, stepDescription: description || label },
    };
  });
  return changed ? result : edges;
}

// ─── Truncation / wrapping policy (single, wide-character aware) ──────────────

/** Display width of a character: CJK / full-width glyphs count as two cells. */
function charWidth(codePoint: number): number {
  if (
    (codePoint >= 0x1100 && codePoint <= 0x115f) || // Hangul Jamo
    (codePoint >= 0x2e80 && codePoint <= 0x303e) || // CJK radicals, Kangxi
    (codePoint >= 0x3041 && codePoint <= 0x33ff) || // Hiragana, Katakana, CJK symbols
    (codePoint >= 0x3400 && codePoint <= 0x4dbf) || // CJK Ext A
    (codePoint >= 0x4e00 && codePoint <= 0x9fff) || // CJK Unified
    (codePoint >= 0xa000 && codePoint <= 0xa4cf) || // Yi
    (codePoint >= 0xac00 && codePoint <= 0xd7a3) || // Hangul syllables
    (codePoint >= 0xf900 && codePoint <= 0xfaff) || // CJK compatibility
    (codePoint >= 0xfe30 && codePoint <= 0xfe4f) || // CJK compatibility forms
    (codePoint >= 0xff00 && codePoint <= 0xff60) || // full-width forms
    (codePoint >= 0xffe0 && codePoint <= 0xffe6) ||
    (codePoint >= 0x1f300 && codePoint <= 0x1faff) // emoji / pictographs
  ) {
    return 2;
  }
  return 1;
}

/** Wide-character-aware display width, in cells. */
export function textWidth(text: string): number {
  let width = 0;
  for (const char of text) width += charWidth(char.codePointAt(0) ?? 0);
  return width;
}

/**
 * Truncate to `maxWidth` display cells (CJK aware), appending an ellipsis. One
 * policy shared by every exporter so a Japanese label is cut identically.
 */
export function truncateLabel(text: string, maxWidth: number): string {
  if (maxWidth <= 0 || textWidth(text) <= maxWidth) return text;
  let width = 0;
  let result = '';
  const budget = Math.max(1, maxWidth - 1); // leave room for the ellipsis
  for (const char of text) {
    const next = width + charWidth(char.codePointAt(0) ?? 0);
    if (next > budget) break;
    result += char;
    width = next;
  }
  return `${result.trimEnd()}…`;
}

/** Wrap into lines of at most `maxWidth` cells, capped at `maxLines`. */
export function wrapLabel(text: string, maxWidth: number, maxLines = 3): string[] {
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let current = '';
  const pushWord = (word: string) => {
    const candidate = current ? `${current} ${word}` : word;
    if (textWidth(candidate) <= maxWidth || !current) {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  };
  for (const word of words) {
    if (textWidth(word) > maxWidth) {
      // Hard-break an over-long token (common with CJK strings that lack spaces).
      if (current) { lines.push(current); current = ''; }
      let chunk = '';
      for (const char of word) {
        if (textWidth(chunk + char) > maxWidth) { lines.push(chunk); chunk = char; }
        else chunk += char;
      }
      current = chunk;
    } else {
      pushWord(word);
    }
  }
  if (current) lines.push(current);
  if (lines.length <= maxLines) return lines.length ? lines : [''];
  const kept = lines.slice(0, maxLines);
  kept[maxLines - 1] = truncateLabel(`${kept[maxLines - 1]} ${lines.slice(maxLines).join(' ')}`, maxWidth);
  return kept;
}

// ─── Metadata (SKU · region · cost) ──────────────────────────────────────────

function firstString(...values: unknown[]): string | undefined {
  for (const value of values) {
    if (typeof value === 'string' && value.trim()) return value.trim();
  }
  return undefined;
}

/** Format a monthly cost the same way the canvas cost badge does. */
export function formatCost(amount: number): string {
  if (!Number.isFinite(amount)) return '';
  if (amount === 0) return 'Free';
  if (amount < 1) return `$${amount.toFixed(3)}/mo`;
  return `$${amount.toFixed(2)}/mo`;
}

function readMeta(data: Record<string, unknown>): BoxMeta | undefined {
  const pricing = (data.pricing ?? undefined) as Record<string, unknown> | undefined;
  const sku = firstString(
    data.sku,
    data.tier,
    pricing?.tier,
    pricing?.skuName,
  );
  const regionRaw = firstString(data.region, pricing?.region);
  const region = regionRaw && regionRaw.toLowerCase() !== 'unknown' ? regionRaw : undefined;

  let cost: number | undefined;
  if (
    pricing
    && typeof pricing.estimatedCost === 'number'
    && Number.isFinite(pricing.estimatedCost)
    && pricing.estimatedCost >= 0
  ) {
    const quantity = typeof pricing.quantity === 'number' && Number.isFinite(pricing.quantity) && pricing.quantity > 0
      ? pricing.quantity
      : 1;
    const total = pricing.estimatedCost * quantity;
    if (Number.isFinite(total)) cost = total;
  }

  if (sku === undefined && region === undefined && cost === undefined) return undefined;
  return { sku, region, cost, costLabel: cost !== undefined ? formatCost(cost) : undefined };
}

/** `SKU · region · $X/mo` sub-line for a box (empty when nothing is known). */
export function metaSubline(box: ExportBox): string {
  const meta = box.meta;
  if (!meta) return '';
  const parts: string[] = [];
  if (meta.sku) parts.push(meta.sku);
  if (meta.region) parts.push(meta.region);
  if (meta.costLabel) parts.push(meta.costLabel);
  return parts.join(' · ');
}

function readCustomColor(data: Record<string, unknown>): ExportBox['customColor'] {
  const raw = data.customColor;
  if (!raw || typeof raw !== 'object') return undefined;
  const value = raw as Record<string, unknown>;
  const border = typeof value.border === 'string' ? value.border : undefined;
  const bg = typeof value.bg === 'string' ? value.bg : undefined;
  const header = typeof value.header === 'string' ? value.header : undefined;
  if (!border && !bg && !header) return undefined;
  return { border, bg, header };
}

/** Coerce a possibly non-finite number to a safe default (guards fix 5). */
function finiteOr(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

/**
 * Flatten React Flow nodes into absolute pixel boxes. Children — including
 * nested groups — are offset by their whole ancestor chain.
 */
export function collectExportBoxes(nodes: Node[]): Map<string, ExportBox> {
  const positions = resolveAbsolutePositions(nodes);

  const boxes = new Map<string, ExportBox>();
  for (const node of nodes) {
    const { w, h } = readSize(node);
    const rawPosition = positions.get(node.id) ?? { x: node.position?.x ?? 0, y: node.position?.y ?? 0 };
    const position = { x: finiteOr(rawPosition.x, 0), y: finiteOr(rawPosition.y, 0) };
    const data = (node.data ?? {}) as Record<string, unknown>;
    const iconPath = typeof data.iconPath === 'string' && data.iconPath ? data.iconPath : undefined;
    const isGroup = node.type === 'groupNode';
    boxes.set(node.id, {
      id: node.id,
      kind: isGroup ? 'group' : 'service',
      label: typeof data.label === 'string' && data.label.trim()
        ? data.label
        : isGroup ? 'Group' : 'Service',
      iconPath,
      category: resolveCategory(data, iconPath),
      serviceName: typeof data.serviceName === 'string' && data.serviceName.trim()
        ? data.serviceName
        : undefined,
      customColor: isGroup ? readCustomColor(data) : undefined,
      meta: isGroup ? undefined : readMeta(data),
      parent: typeof node.parentNode === 'string' && node.parentNode ? node.parentNode : undefined,
      x: position.x,
      y: position.y,
      w: Math.max(1, finiteOr(w, isGroup ? DEFAULT_GROUP_W : DEFAULT_SERVICE_W)),
      h: Math.max(1, finiteOr(h, isGroup ? DEFAULT_GROUP_H : DEFAULT_SERVICE_H)),
    });
  }
  return boxes;
}

export function computeBounds(boxes: Iterable<ExportBox>): Bounds {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;
  let seen = false;
  for (const box of boxes) {
    seen = true;
    minX = Math.min(minX, box.x);
    minY = Math.min(minY, box.y);
    maxX = Math.max(maxX, box.x + box.w);
    maxY = Math.max(maxY, box.y + box.h);
  }
  if (!seen) return { minX: 0, minY: 0, maxX: 1, maxY: 1 };
  return { minX, minY, maxX, maxY };
}

/** Linear-interpolated quartiles (Q1, Q3) of an ascending numeric array. */

/**
 * How far apart two parts of a drawing have to be, centre to centre across an
 * otherwise empty band, before the space between them stops being spacing and
 * becomes a void. Sixteen inches of blank paper says nothing that three inches
 * does not, and it costs the whole drawing its scale.
 *
 * Measured centre to centre, not edge to edge, deliberately: a service tile is
 * 150 x 75, so the same radius leaves a 75px wider gap horizontally than
 * vertically, and an edge-to-edge bar therefore closed one axis of a symmetric
 * drawing and not the other. A hub-and-spoke at radius 1700 came out with its
 * north and south arms 4.1in long and its east and west arms 17.7in, for four
 * hops the author drew identical. Centre distance is the radius on both axes,
 * so the decision is the same on both.
 *
 * The bar itself sits above the widest separation an author's own layout
 * produces — the 1400px radius the Architecture Center draws hub-and-spoke at —
 * and is exactly the 16in the audit rejects a void at, so there is no band the
 * exporter refuses to close that the gate then fails.
 */
const MAX_VOID_SPAN_PX = 1536;
/** What a closed void is replaced by — still a clear separation, not a join. */
const VOID_GUTTER_PX = 320;

/**
 * Close empty bands that run the full height (or width) of the drawing.
 *
 * The AI generator and hand-editing both produce canvases where two parts of
 * an architecture sit thousands of pixels apart with nothing in between: a DR
 * region 6000px east of the primary made a 72in Visio sheet, 50in of which was
 * blank. Trimming cannot help — a third of the boxes is not an outlier, it is
 * a region — and packing must not, because moving a region into a margin strip
 * throws away the one thing its position was saying.
 *
 * Closing the void keeps every semantic the author drew: order along both axes
 * is preserved exactly, nothing changes side, and relative spacing within each
 * part is untouched. Only genuinely empty space is removed, so no shape can
 * collide with another.
 *
 * Emptiness is judged by where the *services* are. A zone is a claim about a
 * region of the canvas rather than a thing occupying it, and one rectangle
 * drawn around the whole architecture — an "Azure" frame, a subscription or
 * tenant boundary, the most ordinary annotation in the Architecture Center —
 * spans every void there is. Counting it as content meant a five-region
 * drawing kept all 256in of its empty space, and the gate that should have
 * caught that was blinded by the same rectangle. A zone that straddles a
 * closed void is shrunk by what was removed from underneath it, so it still
 * ends exactly where its contents do.
 */
export function compactEmptyGutters(boxes: Map<string, ExportBox>): Map<string, ExportBox> {
  const all = [...boxes.values()];
  const parented = new Set(all.map((box) => box.parent).filter((id): id is string => !!id));
  const services = all.filter((box) => box.kind !== 'group');
  // A zone is a claim about a region whenever anything lives in that region,
  // whether by declaration or by geometry, and a claim must not be allowed to
  // bridge a void: a frame drawn around the whole drawing would otherwise
  // report the drawing as gapless and export a sheet nine tenths blank.
  //
  // A zone with nothing in it is the opposite. It is a label for a piece of
  // empty canvas — the corridor between two regions marked "ExpressRoute
  // circuit", "Internet", "On-premises", which is one click away in the editor
  // and among the commonest annotations the Architecture Center draws — and the
  // space it names is the whole point of it. Treating those as claims deleted
  // them: a 900px corridor box came out 1px wide, a vertical line where the
  // author had drawn a labelled band.
  const holdsNothing = (zone: ExportBox): boolean => (
    !parented.has(zone.id)
    && !services.some((box) => (
      box.x < zone.x + zone.w && box.x + box.w > zone.x
      && box.y < zone.y + zone.h && box.y + box.h > zone.y
    ))
  );
  if (all.length < 2 || services.length < 2) return boxes;

  /**
   * The boxes that define emptiness on one axis.
   *
   * An empty zone names the space it stands in, so on the axis where it stands
   * *between* things it is content and the band must not be closed underneath
   * it. On an axis where it merely stretches *over* the drawing it is not: a
   * caption band across the top of an architecture would otherwise report every
   * horizontal gap as occupied and switch compaction off entirely, which is how
   * two clusters 5,450px apart stayed 5,450px apart under a rectangle that
   * covered neither of them. The distinction is per-axis because the same
   * rectangle is usually both: the sovereign band stands above the drawing on y
   * and spans it on x.
   *
   * "Stretches over" is measured against each service's own size so it scales
   * with the drawing: a corridor whose edge merely abuts a region is still
   * standing between them, while one that swallows a tile whole is not.
   */
  const occupyingOn = (start: (b: ExportBox) => number, size: (b: ExportBox) => number): ExportBox[] => {
    const spans = (zone: ExportBox): boolean => services.some((box) => {
      const over = Math.min(start(box) + size(box), start(zone) + size(zone)) - Math.max(start(box), start(zone));
      return over > size(box) / 2;
    });
    return all.filter((box) => box.kind !== 'group' || (holdsNothing(box) && !spans(box)));
  };

  /** Voids on one axis, as [start, gap, amount-to-remove] in ascending order. */
  const voids = (start: (b: ExportBox) => number, size: (b: ExportBox) => number): [number, number, number][] => {
    const occupying = occupyingOn(start, size);
    if (occupying.length < 2) return [];
    const spans = occupying
      .map((box) => [start(box), start(box) + size(box), size(box)] as [number, number, number])
      .sort((a, b) => a[0] - b[0]);
    const found: [number, number, number][] = [];
    let reach = spans[0][1];
    let reachSize = spans[0][2];
    for (const [from, to, ownSize] of spans) {
      const gap = from - reach;
      if (gap > 0 && gap + (reachSize + ownSize) / 2 > MAX_VOID_SPAN_PX) {
        found.push([reach, gap, gap - VOID_GUTTER_PX]);
      }
      if (to > reach) {
        reach = to;
        reachSize = ownSize;
      }
    }
    return found;
  };
  // Continuous and non-decreasing, which is what keeps a rectangle a
  // rectangle: a coordinate on the near lip of a void does not move, one past
  // the far lip moves by the whole amount, and one inside moves by how far in
  // it is. A step function instead of a ramp was worse than it sounds — a tile
  // whose right edge landed exactly on the lip had that edge moved while its
  // left edge stayed, so the tile collapsed to a hairline and took its label
  // with it. Two of the twelve services in an ordinary two-region drawing end
  // on the lip, because the lip is by definition where the last of them ends.
  const shift = (found: [number, number, number][], at: number): number => {
    let total = 0;
    for (const [from, , amount] of found) {
      if (at <= from) break;
      total += Math.min(at - from, amount);
    }
    return total;
  };

  const xVoids = voids((b) => b.x, (b) => b.w);
  const yVoids = voids((b) => b.y, (b) => b.h);
  if (xVoids.length === 0 && yVoids.length === 0) return boxes;

  const out = new Map<string, ExportBox>();
  for (const [id, box] of boxes) {
    // Both edges are mapped, so a rectangle that spans a void loses exactly the
    // emptiness that was under it and keeps everything else.
    const x = box.x - shift(xVoids, box.x);
    const y = box.y - shift(yVoids, box.y);
    const right = box.x + box.w - shift(xVoids, box.x + box.w);
    const bottom = box.y + box.h - shift(yVoids, box.y + box.h);
    // Nothing may be compacted out of existence. Removing space a shape stands
    // in is not closing a void, it is deleting the shape — and a rectangle
    // reduced to a hairline is worse than one left too wide, because the reader
    // cannot see that anything is missing. Whatever the mapping says, every box
    // keeps a usable fraction of the size the author gave it.
    const floor = (drawn: number, was: number): number => Math.max(drawn, Math.min(was, VOID_GUTTER_PX / 2));
    out.set(id, {
      ...box,
      x,
      y,
      w: floor(Math.max(1, right - x), box.w),
      h: floor(Math.max(1, bottom - y), box.h),
    });
  }
  return out;
}

/**
 * Squeeze the empty space between shapes, proportionally, until the drawing
 * fits inside a hard limit — leaving every shape its own size.
 *
 * This is the answer to a format that simply refuses a drawing: Visio will not
 * open a page over 200in on a side, and a 27-service cascade at the author's
 * own spacing is 247in wide. Every other lever makes the export worse. Scaling
 * the drawing shrinks the type with it, and the type is already sitting on the
 * legibility floor, so the file would open and be unreadable. Cropping loses
 * services. Refusing to export tells the user nothing they can act on.
 *
 * Tightening the gaps is what a person does when a drawing runs off the paper,
 * and it is the only transform here that costs nothing but whitespace: shapes
 * keep their size, text keeps its point size, reading order and relative
 * position are preserved, and what is lost is the distance between things,
 * uniformly, so nothing about the architecture is misrepresented. Gaps that
 * were larger give up proportionally more, because that is where the space is.
 *
 * `compactEmptyGutters` runs first and removes whole empty bands; this handles
 * what is left when a drawing is genuinely, evenly too big.
 */
export function fitBoxesWithin(
  boxes: Map<string, ExportBox>,
  maxW: number,
  maxH: number,
): Map<string, ExportBox> {
  const all = [...boxes.values()];
  if (all.length < 2) return boxes;
  // The union is taken over the *services*, for the same reason
  // `compactEmptyGutters` closes voids by them. A subscription rectangle drawn
  // around the whole architecture is one span covering the axis, so counting it
  // as solid made the union the entire drawing, left no whitespace to spend and
  // returned the identity map — an ordinary annotation switched the fit off and
  // handed the sheet to the uniform scaler, which takes the tiles down while
  // the label point size stays where it is. A 40-service cascade lost 47% of
  // its tile width to a rectangle drawn around it.
  //
  // `origin` and `end` still come from every box, because the frame is part of
  // what has to fit; only the question "which parts cannot be given up" is
  // answered by the shapes.
  const solid = all.filter((b) => b.kind !== 'group');
  const dense = solid.length >= 2 ? solid : all;

  const axis = (
    lo: (b: ExportBox) => number,
    size: (b: ExportBox) => number,
    limit: number,
  ): ((at: number) => number) | null => {
    const spans = all.map((b) => [lo(b), lo(b) + size(b)] as [number, number])
      .sort((a, b) => a[0] - b[0]);
    const origin = spans[0][0];
    const end = Math.max(...spans.map((s) => s[1]));
    if (end - origin <= limit) return null;
    // The union of the shapes is the part that cannot be given up.
    const merged: Array<[number, number]> = [];
    for (const [from, to] of dense.map((b) => [lo(b), lo(b) + size(b)] as [number, number]).sort((a, b) => a[0] - b[0])) {
      const last = merged[merged.length - 1];
      if (last && from <= last[1]) last[1] = Math.max(last[1], to);
      else merged.push([from, to]);
    }
    const covered = merged.reduce((sum, [from, to]) => sum + (to - from), 0);
    const empty = (end - origin) - covered;
    // Shapes alone over the limit: there is no whitespace left to spend, so
    // close every gap and let the caller deal with what remains. Squeezing to
    // touching is still the least-lossy thing available.
    const keep = empty <= 0 ? 0 : Math.max(0, Math.min(1, (limit - covered) / empty));
    return (at: number): number => {
      let mapped = origin;
      let cursor = origin;
      for (const [from, to] of merged) {
        if (at <= cursor) return mapped;
        mapped += Math.min(at - cursor, from - cursor) * keep;
        cursor = from;
        if (at <= cursor) return mapped;
        mapped += Math.min(at - cursor, to - from);
        cursor = to;
      }
      return mapped + Math.max(0, at - cursor) * keep;
    };
  };

  const mapX = axis((b) => b.x, (b) => b.w, maxW);
  const mapY = axis((b) => b.y, (b) => b.h, maxH);
  if (!mapX && !mapY) return boxes;

  const out = new Map<string, ExportBox>();
  for (const [id, box] of boxes) {
    const x = mapX ? mapX(box.x) : box.x;
    const y = mapY ? mapY(box.y) : box.y;
    // Widths come from the map as well, so a zone still ends where its last
    // member ends; a shape can only ever keep or lose gap, never lose itself.
    const w = mapX ? Math.max(box.w > 0 ? 1 : 0, mapX(box.x + box.w) - x) : box.w;
    const h = mapY ? Math.max(box.h > 0 ? 1 : 0, mapY(box.y + box.h) - y) : box.h;
    out.set(id, { ...box, x, y, w: Math.max(w, Math.min(box.w, VOID_GUTTER_PX / 2)), h: Math.max(h, Math.min(box.h, VOID_GUTTER_PX / 2)) });
  }
  return out;
}

/**
 * Scale a drawing down until it fits a hard limit — the last resort, used only
 * where the format refuses the file outright.
 *
 * {@link fitBoxesWithin} is always tried first because it costs nothing but
 * distance. It has nothing left to give when the shapes themselves are over the
 * limit: a single row of more than 127 tiles, or a zone rectangle drawn across
 * the whole drawing, which it has to count as solid because it cannot know the
 * rectangle is mostly air. Scaling takes the type down with the tiles and is
 * genuinely worse — but a page Visio will not open is not an export at all, so
 * a small drawing beats no drawing. Returns the map untouched when it fits.
 */
export function scaleBoxesWithin(
  boxes: Map<string, ExportBox>,
  maxW: number,
  maxH: number,
): Map<string, ExportBox> {
  if (boxes.size === 0) return boxes;
  const bounds = computeBounds(boxes.values());
  const w = bounds.maxX - bounds.minX;
  const h = bounds.maxY - bounds.minY;
  const scale = Math.min(1, w > 0 ? maxW / w : 1, h > 0 ? maxH / h : 1);
  if (scale >= 0.999) return boxes;
  const out = new Map<string, ExportBox>();
  for (const [id, box] of boxes) {
    out.set(id, {
      ...box,
      x: bounds.minX + (box.x - bounds.minX) * scale,
      y: bounds.minY + (box.y - bounds.minY) * scale,
      w: Math.max(1, box.w * scale),
      h: Math.max(1, box.h * scale),
    });
  }
  return out;
}

function quartiles(sorted: number[]): [number, number] {
  const at = (q: number) => {
    const pos = (sorted.length - 1) * q;
    const lo = Math.floor(pos);
    const hi = Math.ceil(pos);
    if (lo === hi) return sorted[lo];
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
  };
  return [at(0.25), at(0.75)];
}

/**
 * Bounds of the *dense* cluster, ignoring extreme outliers that would otherwise
 * collapse the fit scale (a single node at (8000, 6000) shrinks readable tiles
 * to specks). Boxes past `k`×IQR from the centre quartiles are dropped from the
 * bounds only — they are still exported, just not allowed to dictate the scale.
 * With no outliers this returns exactly {@link computeBounds}, so normal
 * diagrams are unchanged.
 */
export function computeContentBounds(
  boxes: Iterable<ExportBox>,
  options: { iqrMultiplier?: number } = {},
): Bounds {
  const all = [...boxes].filter(
    (box) => Number.isFinite(box.x) && Number.isFinite(box.y) && Number.isFinite(box.w) && Number.isFinite(box.h),
  );
  if (all.length <= 3) return computeBounds(all);

  const k = options.iqrMultiplier ?? 3;
  // The clear majority of the diagram has to survive, or the "outliers" are the
  // diagram; anything less and the trim is abandoned rather than hiding real
  // content.
  const majority = Math.max(2, Math.ceil(all.length * 0.6));

  // Peeled repeatedly, because outliers contaminate the quartiles that are
  // supposed to find them. Four strays off three corners of a twelve-box
  // drawing dragged the upper quartile out to 2820px, so the fence swallowed
  // three of the four and the "dense cluster" came back 99in wide — a 56in
  // slide and a 100in Visio sheet carrying eight tiles. Each pass recomputes
  // the quartiles from the survivors, so the fence tightens onto the pack;
  // a drawing with no gap in it is stable on the first pass and unchanged.
  let kept = all;
  for (let pass = 0; pass < 8; pass += 1) {
    const [qx1, qx3] = quartiles(kept.map((box) => box.x + box.w / 2).sort((a, b) => a - b));
    const [qy1, qy3] = quartiles(kept.map((box) => box.y + box.h / 2).sort((a, b) => a - b));
    // One fence width for both axes, because a diagram has one scale — and
    // because a degenerate axis was otherwise vetoing the other. When every
    // service sits on one row the vertical quartile range is zero, the fence
    // has zero width, the single service on a second row counts as an outlier
    // and gets parked in the margin away from the neighbours it is wired to.
    // Sharing the wider of the two ranges can only ever keep more boxes than
    // per-axis fences, so it cannot over-trim.
    const fence = k * Math.max(qx3 - qx1, qy3 - qy1);
    const next = kept.filter((box) => {
      const cx = box.x + box.w / 2;
      const cy = box.y + box.h / 2;
      return cx >= qx1 - fence && cx <= qx3 + fence && cy >= qy1 - fence && cy <= qy3 + fence;
    });
    if (next.length === kept.length) break;
    if (next.length < majority) break;
    kept = next;
  }

  if (kept.length >= majority && kept.length < all.length) return computeBounds(kept);
  return computeBounds(all);
}

/**
 * Fit `bounds` (pixels) inside `frame` (inches) preserving the aspect ratio and
 * centring the result. `maxScale` prevents a two-node diagram from being blown
 * up to absurd tile sizes.
 */
export function computeFitTransform(
  bounds: Bounds,
  frame: { x: number; y: number; w: number; h: number },
  options: { maxScale?: number } = {},
): FitTransform {
  const safe = {
    minX: finiteOr(bounds.minX, 0),
    minY: finiteOr(bounds.minY, 0),
    maxX: finiteOr(bounds.maxX, 1),
    maxY: finiteOr(bounds.maxY, 1),
  };
  const frameW = finiteOr(frame.w, 1) || 1;
  const frameH = finiteOr(frame.h, 1) || 1;
  const contentW = Math.max(1, safe.maxX - safe.minX);
  const contentH = Math.max(1, safe.maxY - safe.minY);
  const maxScale = options.maxScale ?? 1 / 96; // never larger than 96 px per inch
  const rawScale = Math.min(frameW / contentW, frameH / contentH, maxScale);
  const scale = Number.isFinite(rawScale) && rawScale > 0 ? rawScale : maxScale;
  const offsetX = finiteOr(frame.x, 0) + (frameW - contentW * scale) / 2 - safe.minX * scale;
  const offsetY = finiteOr(frame.y, 0) + (frameH - contentH * scale) / 2 - safe.minY * scale;
  return { scale, offsetX, offsetY };
}

export function applyTransform(point: Point, transform: FitTransform): Point {
  return {
    x: point.x * transform.scale + transform.offsetX,
    y: point.y * transform.scale + transform.offsetY,
  };
}

function centre(box: ExportBox): Point {
  return { x: box.x + box.w / 2, y: box.y + box.h / 2 };
}

/** Does an axis-aligned segment come within `margin` px of a box? */
function segmentHitsBox(a: Point, b: Point, box: ExportBox, margin: number): boolean {
  const minX = Math.min(a.x, b.x);
  const maxX = Math.max(a.x, b.x);
  const minY = Math.min(a.y, b.y);
  const maxY = Math.max(a.y, b.y);
  return (
    minX <= box.x + box.w + margin
    && maxX >= box.x - margin
    && minY <= box.y + box.h + margin
    && maxY >= box.y - margin
  );
}

/** Count polyline segments that clip any obstacle box. */
function countBlocked(points: Point[], obstacles: ExportBox[], margin: number): number {
  let blocked = 0;
  for (let i = 1; i < points.length; i += 1) {
    for (const box of obstacles) {
      if (segmentHitsBox(points[i - 1], points[i], box, margin)) { blocked += 1; break; }
    }
  }
  return blocked;
}

export interface RouteOptions {
  /** Perpendicular offset (px) so parallel edges don't overlap. */
  offset?: number;
  /** Boxes to route around (source/target should be excluded by the caller). */
  obstacles?: ExportBox[];
  /**
   * Lateral shift (px) of the anchor on the source / target side, so a hop that
   * arrives around a corner does not land on the same point of the same tile as
   * one that arrives head-on. Applied only to the elbow form: a straight hop
   * keeps the centre of the side, which is the anchor a reader expects and the
   * one that must not move if the line is to stay straight.
   */
  sourceShift?: number;
  targetShift?: number;
  /**
   * False when the hop is one of several drawn between the same pair of tiles.
   *
   * A bundle is held apart by nothing but the lane each member was given, so
   * any rule that moves a lane for its own reasons has to leave a fan alone —
   * including the middle member of an odd fan, whose offset is zero and which
   * therefore looks exactly like a hop standing on its own.
   */
  solo?: boolean;
}

/**
 * How far outside a tile a shifted hop turns off the centre line.
 *
 * The anchor itself must stay on the middle of the side: that is the only point
 * PowerPoint recognises as a connection site, and an arrow that leaves it stops
 * being glued to the service it joins. So the hop leaves head-on, takes a short
 * step, and only then moves to its own lane. The stub is what remains lying on
 * the shared centre line, so it is kept to a fraction of the run and capped —
 * long enough to read as a deliberate jog, far too short to hide a hop.
 */
function stubLength(from: number, mid: number): number {
  return Math.min(Math.abs(mid - from) * 0.35, 14);
}

/**
 * Drop duplicate and collinear points, so a route built from the general
 * eight-point form collapses to exactly the four-point Z it used to be
 * whenever no lane shift applies.
 */
function simplifyPath(points: Point[]): Point[] {
  const out: Point[] = [];
  for (const p of points) {
    const last = out[out.length - 1];
    if (last && Math.abs(last.x - p.x) < 0.01 && Math.abs(last.y - p.y) < 0.01) continue;
    out.push(p);
  }
  for (let i = out.length - 2; i >= 1; i -= 1) {
    const a = out[i - 1];
    const b = out[i];
    const c = out[i + 1];
    const collinear = (Math.abs(a.x - b.x) < 0.01 && Math.abs(b.x - c.x) < 0.01)
      || (Math.abs(a.y - b.y) < 0.01 && Math.abs(b.y - c.y) < 0.01);
    if (collinear) out.splice(i, 1);
  }
  return out;
}

/**
 * Right-angle "Z" route between two boxes. The exit/entry anchors sit on the
 * facing edges so the arrow never starts inside a tile, and the elbow is placed
 * on the mid-line so labels stay clear of both shapes. With `offset`, parallel
 * edges are separated; with `obstacles`, the route detours around intervening
 * tiles. Called with no options it is byte-for-byte the original 2-bend route.
 */
export function routeOrthogonal(
  source: ExportBox,
  target: ExportBox,
  options: RouteOptions = {},
): { points: Point[]; labelAnchor: Point } {
  const offset = options.offset ?? 0;
  const obstacles = options.obstacles ?? [];
  const sc = centre(source);
  const tc = centre(target);
  const dx = tc.x - sc.x;
  const dy = tc.y - sc.y;
  const horizontal = Math.abs(dx) >= Math.abs(dy);

  let base: { points: Point[]; labelAnchor: Point };
  let ends: { lead: Point[]; tail: Point[] } | undefined;
  const sourceShift = options.sourceShift ?? 0;
  const targetShift = options.targetShift ?? 0;
  // The parallel fan used to be added to the endpoints themselves, which took
  // every member of a fan off its tile's connection site and made it impossible
  // for PowerPoint to glue — on the "parallel" fixture only one arrow in six
  // could be attached, so the deck came apart the moment a reader moved a tile.
  // The separation is now a jog, exactly like the per-side port fan: the line
  // leaves the site, steps into its own lane, and comes back to a site.
  const shifted = Math.abs(sourceShift) > 0.01 || Math.abs(targetShift) > 0.01 || Math.abs(offset) > 0.01;
  if (horizontal) {
    const startX = dx >= 0 ? source.x + source.w : source.x;
    const endX = dx >= 0 ? target.x : target.x + target.w;
    const startY = sc.y;
    const endY = tc.y;
    if (Math.abs(startY - endY) < 0.5 && Math.abs(offset) <= 0.01) {
      base = {
        points: [{ x: startX, y: startY }, { x: endX, y: endY }],
        labelAnchor: { x: (startX + endX) / 2, y: startY },
      };
    } else if (Math.abs(startY - endY) < 0.5) {
      // Aligned tiles. Without a lane every member of the fan would be drawn on
      // the one centre line and the reader would see a single arrow.
      const stub = stubLength(startX, endX);
      const sStubX = startX + Math.sign(endX - startX) * stub;
      const eStubX = endX + Math.sign(startX - endX) * stub;
      base = {
        points: simplifyPath([
          { x: startX, y: startY },
          { x: sStubX, y: startY },
          { x: sStubX, y: startY + offset },
          { x: eStubX, y: endY + offset },
          { x: eStubX, y: endY },
          { x: endX, y: endY },
        ]),
        labelAnchor: { x: (startX + endX) / 2, y: startY + offset },
      };
      ends = {
        lead: [{ x: startX, y: startY }, { x: sStubX, y: startY }, { x: sStubX, y: startY + offset }],
        tail: [{ x: eStubX, y: endY + offset }, { x: eStubX, y: endY }, { x: endX, y: endY }],
      };
    } else {
      const midX = (startX + endX) / 2;
      const shiftedStartY = startY + offset + sourceShift;
      const shiftedEndY = endY + offset + targetShift;
      const stub = stubLength(startX, midX);
      const sStubX = startX + Math.sign(midX - startX) * stub;
      const eStubX = endX + Math.sign(midX - endX) * stub;
      base = {
        points: simplifyPath([
          { x: startX, y: startY },
          { x: sStubX, y: startY },
          { x: sStubX, y: shiftedStartY },
          { x: midX, y: shiftedStartY },
          { x: midX, y: shiftedEndY },
          { x: eStubX, y: shiftedEndY },
          { x: eStubX, y: endY },
          { x: endX, y: endY },
        ]),
        // The anchor stays on the UNSHIFTED mid line. The per-side fan is a
        // short jog near the endpoints that keeps two hops off one another; it
        // is not meant to move the label, and letting it do so dragged a chip
        // up to 26px onto a bystanding tile.
        labelAnchor: { x: midX, y: (startY + endY) / 2 + offset },
      };
      if (shifted) {
        ends = {
          lead: [{ x: startX, y: startY }, { x: sStubX, y: startY }, { x: sStubX, y: shiftedStartY }],
          tail: [{ x: eStubX, y: shiftedEndY }, { x: eStubX, y: endY }, { x: endX, y: endY }],
        };
      }
    }
  } else {
    const startY = dy >= 0 ? source.y + source.h : source.y;
    const endY = dy >= 0 ? target.y : target.y + target.h;
    const startX = sc.x;
    const endX = tc.x;
    if (Math.abs(startX - endX) < 0.5 && Math.abs(offset) <= 0.01) {
      base = {
        points: [{ x: startX, y: startY }, { x: endX, y: endY }],
        labelAnchor: { x: startX, y: (startY + endY) / 2 },
      };
    } else if (Math.abs(startX - endX) < 0.5) {
      const stub = stubLength(startY, endY);
      const sStubY = startY + Math.sign(endY - startY) * stub;
      const eStubY = endY + Math.sign(startY - endY) * stub;
      base = {
        points: simplifyPath([
          { x: startX, y: startY },
          { x: startX, y: sStubY },
          { x: startX + offset, y: sStubY },
          { x: endX + offset, y: eStubY },
          { x: endX, y: eStubY },
          { x: endX, y: endY },
        ]),
        labelAnchor: { x: startX + offset, y: (startY + endY) / 2 },
      };
      ends = {
        lead: [{ x: startX, y: startY }, { x: startX, y: sStubY }, { x: startX + offset, y: sStubY }],
        tail: [{ x: endX + offset, y: eStubY }, { x: endX, y: eStubY }, { x: endX, y: endY }],
      };
    } else {
      const midY = (startY + endY) / 2;
      const shiftedStartX = startX + offset + sourceShift;
      const shiftedEndX = endX + offset + targetShift;
      const stub = stubLength(startY, midY);
      const sStubY = startY + Math.sign(midY - startY) * stub;
      const eStubY = endY + Math.sign(midY - endY) * stub;
      base = {
        points: simplifyPath([
          { x: startX, y: startY },
          { x: startX, y: sStubY },
          { x: shiftedStartX, y: sStubY },
          { x: shiftedStartX, y: midY },
          { x: shiftedEndX, y: midY },
          { x: shiftedEndX, y: eStubY },
          { x: endX, y: eStubY },
          { x: endX, y: endY },
        ]),
        labelAnchor: { x: (startX + endX) / 2 + offset, y: midY },
      };
      if (shifted) {
        ends = {
          lead: [{ x: startX, y: startY }, { x: startX, y: sStubY }, { x: shiftedStartX, y: sStubY }],
          tail: [{ x: shiftedEndX, y: eStubY }, { x: endX, y: eStubY }, { x: endX, y: endY }],
        };
      }
    }
  }

  if (obstacles.length === 0) return base;
  const margin = 6;
  if (countBlocked(base.points, obstacles, margin) === 0) return base;
  const detourable = (options.solo ?? true) && Math.abs(offset) <= 0.01;
  return bestDetour(
    base,
    obstacles,
    horizontal,
    margin,
    ends,
    source,
    target,
    detourable,
    offset,
  );
}

/**
 * Midpoints of the clear bands between obstacle spans, nearest the direct line
 * first. `bestDetour` derives its lanes from the bounding box of every blocked
 * obstacle at once, so when a route is blocked in two separate places — the row
 * it starts in and the row it ends in, which is what a wrap-around hop always
 * does — every lane it can name lies outside both, and the one that actually
 * works, the gutter between them, is inside the box and never offered.
 */
function clearLanes(spans: Array<[number, number]>, from: number, to: number, margin: number): number[] {
  if (spans.length === 0) return [];
  const merged: Array<[number, number]> = [];
  for (const [lo, hi] of [...spans].sort((a, b) => a[0] - b[0])) {
    const last = merged[merged.length - 1];
    if (last && lo <= last[1] + 2 * margin) last[1] = Math.max(last[1], hi);
    else merged.push([lo, hi]);
  }
  const lanes: number[] = [];
  for (let i = 1; i < merged.length; i += 1) {
    const gapLo = merged[i - 1][1];
    const gapHi = merged[i][0];
    if (gapHi - gapLo > 2 * margin + 4) lanes.push((gapLo + gapHi) / 2);
  }
  const mid = (from + to) / 2;
  return lanes.sort((a, b) => Math.abs(a - mid) - Math.abs(b - mid));
}

/**
 * The four points PowerPoint and Visio recognise as connection sites.
 *
 * `routeOrthogonal` picks the pair of sides from the dominant axis alone, so a
 * target whose facing side abuts a neighbour can never be reached: every route
 * has to finish inside the neighbour. Re-anchoring to a perpendicular side is
 * the only way out, and because these are sites too the arrow stays glued.
 */
type BoxSide = 'left' | 'right' | 'top' | 'bottom';

const BOX_SIDES: BoxSide[] = ['right', 'left', 'bottom', 'top'];

function sitePoint(box: ExportBox, side: BoxSide): Point {
  switch (side) {
    case 'left': return { x: box.x, y: box.y + box.h / 2 };
    case 'right': return { x: box.x + box.w, y: box.y + box.h / 2 };
    case 'top': return { x: box.x + box.w / 2, y: box.y };
    default: return { x: box.x + box.w / 2, y: box.y + box.h };
  }
}

function stepOutward(point: Point, side: BoxSide, distance: number): Point {
  switch (side) {
    case 'left': return { x: point.x - distance, y: point.y };
    case 'right': return { x: point.x + distance, y: point.y };
    case 'top': return { x: point.x, y: point.y - distance };
    default: return { x: point.x, y: point.y + distance };
  }
}

const isVerticalSide = (side: BoxSide): boolean => side === 'top' || side === 'bottom';

/** -1 when the side faces up or left, +1 when it faces down or right. */
const outwardSign = (side: BoxSide): number => (side === 'top' || side === 'left' ? -1 : 1);

/**
 * Does an axis-aligned segment pass through the *interior* of a box?
 *
 * `segmentHitsBox` inflates the box and compares bounding boxes, which is the
 * right test for "keep clear of a bystander" and the wrong one for "did this
 * arrow re-enter the tile it just left" — every stub touches its own tile by
 * construction. The inset keeps a segment lying exactly on an edge outside.
 */
function segmentEntersBox(a: Point, b: Point, box: ExportBox, inset = 1): boolean {
  return (
    Math.max(a.x, b.x) > box.x + inset
    && Math.min(a.x, b.x) < box.x + box.w - inset
    && Math.max(a.y, b.y) > box.y + inset
    && Math.min(a.y, b.y) < box.y + box.h - inset
  );
}

function pathEntersBox(points: Point[], box: ExportBox): boolean {
  for (let i = 1; i < points.length; i += 1) {
    if (segmentEntersBox(points[i - 1], points[i], box)) return true;
  }
  return false;
}

/**
 * Orthogonal routes that leave `source` and arrive at `target` head-on, one per
 * plausible shape for the given pair of sides.
 *
 * Several shapes rather than one because the obvious construction — join the
 * two stubs on their mid-line — doubles back through the tile it just left
 * whenever the stubs point away from each other. Leaving the top of a service
 * to reach something *below* it put the arrow straight back down through its
 * own source, and the caller could not see it: the boxes a route connects are
 * excluded from its obstacle list, so the collision was invisible to every
 * check in this file. The lane is therefore constrained to the side of each
 * tile its own stub points at, and when that is infeasible the route jogs
 * sideways instead. The caller picks the first candidate that is clear.
 */
function routesBetweenSides(
  source: ExportBox,
  sourceSide: BoxSide,
  target: ExportBox,
  targetSide: BoxSide,
  gap: number,
  // Coordinates on the perpendicular axis that are known to be clear of every
  // obstacle. Without these the only crossing lane on offer is derived from the
  // two tiles alone, which on a packed grid runs straight down the middle of
  // the column between them.
  crossLanes: number[] = [],
): Array<{ points: Point[]; labelAnchor: Point }> {
  const head = sitePoint(source, sourceSide);
  const tail = sitePoint(target, targetSide);
  const headStub = stepOutward(head, sourceSide, gap);
  const tailStub = stepOutward(tail, targetSide, gap);
  const sourceVertical = isVerticalSide(sourceSide);
  const targetVertical = isVerticalSide(targetSide);

  const shapes: Point[][] = [];
  if (sourceVertical === targetVertical) {
    // Both stubs run along the same axis, so they are joined by one crossing
    // lane. It has to sit beyond each stub in the direction that stub points,
    // or the run back to it re-enters the tile.
    const along = (p: Point): number => (sourceVertical ? p.y : p.x);
    let lo = -Infinity;
    let hi = Infinity;
    for (const [side, stub] of [[sourceSide, headStub], [targetSide, tailStub]] as const) {
      if (outwardSign(side) < 0) hi = Math.min(hi, along(stub));
      else lo = Math.max(lo, along(stub));
    }
    if (lo <= hi) {
      const mid = (along(headStub) + along(tailStub)) / 2;
      const feasible = [mid, ...crossLanes.filter((lane) => lane >= lo && lane <= hi)];
      for (const raw of feasible) {
        const lane = Math.min(Math.max(raw, lo), hi);
        shapes.push(sourceVertical
          ? [head, headStub, { x: headStub.x, y: lane }, { x: tailStub.x, y: lane }, tailStub, tail]
          : [head, headStub, { x: lane, y: headStub.y }, { x: lane, y: tailStub.y }, tailStub, tail]);
      }
    }
    // Stubs pointing the same way, or away from each other, need a jog on the
    // perpendicular axis: out along each stub, across in the clear, and back.
    const near = sourceVertical
      ? [source.x + source.w, target.x, target.x + target.w, source.x]
      : [source.y + source.h, target.y, target.y + target.h, source.y];
    const spans: number[] = [
      (near[0] + near[1]) / 2,
      (near[2] + near[3]) / 2,
      Math.max(near[0], near[2]) + gap,
      Math.min(near[1], near[3]) - gap,
      ...crossLanes,
    ];
    for (const jog of spans) {
      shapes.push(sourceVertical
        ? [head, headStub, { x: jog, y: headStub.y }, { x: jog, y: tailStub.y }, tailStub, tail]
        : [head, headStub, { x: headStub.x, y: jog }, { x: tailStub.x, y: jog }, tailStub, tail]);
    }
  } else {
    // Perpendicular sides meet at a single bend. Put it on each stub's own
    // outward coordinate, so neither leg ever runs back across its own tile.
    const safe = sourceVertical
      ? { x: tailStub.x, y: headStub.y }
      : { x: headStub.x, y: tailStub.y };
    const other = sourceVertical
      ? { x: headStub.x, y: tailStub.y }
      : { x: tailStub.x, y: headStub.y };
    shapes.push([head, headStub, safe, tailStub, tail], [head, headStub, other, tailStub, tail]);
    // Two bends through a clear lane, for when one bend has to be placed inside
    // a neighbour because the tiles are packed against each other.
    for (const lane of crossLanes) {
      shapes.push(sourceVertical
        ? [head, headStub, { x: headStub.x, y: lane }, { x: tailStub.x, y: lane }, tailStub, tail]
        : [head, headStub, { x: lane, y: headStub.y }, { x: lane, y: tailStub.y }, tailStub, tail]);
    }
  }

  return shapes.map((raw) => {
    const points = simplifyPath(raw);
    const mid = points[Math.max(1, Math.floor(points.length / 2) - 1)];
    const next = points[Math.max(1, Math.floor(points.length / 2))];
    return {
      points,
      labelAnchor: { x: (mid.x + next.x) / 2, y: (mid.y + next.y) / 2 },
    };
  });
}

/**
 * Try a handful of deterministic detours (shift the connecting line past the
 * blocking cluster, or take a clear perpendicular gutter) and keep the first
 * fully-clear candidate — otherwise the least-blocked one. Cheap and stable.
 */
function bestDetour(
  base: { points: Point[]; labelAnchor: Point },
  obstacles: ExportBox[],
  horizontal: boolean,
  margin: number,
  // The fixed head and tail of the route, when the hop leaves or arrives on its
  // own lane rather than the centre of the side. Every candidate below is built
  // from two endpoints and a lane, so without this a blocked hop silently threw
  // its lane away and went back to lying on top of its neighbour — which is
  // every wrap-around hop, the ones most likely to share a side in the first
  // place. Omitted when nothing was shifted, so an unshifted route is rebuilt
  // exactly as it always was.
  ends?: { lead: Point[]; tail: Point[] },
  // The tiles being joined, so a route with no clear lane can re-anchor onto a
  // different connection site instead of finishing inside a bystander.
  source?: ExportBox,
  target?: ExportBox,
  // False for one member of a parallel fan, whose separation comes from its
  // lane: re-anchoring would put every member of the bundle back on one site.
  allowResite = false,
  // This hop's place in its bundle. Every lane below is rebuilt from the two
  // endpoints, which throws the fan spread away: a blocked bundle collapsed
  // onto one lane, drawn as a single arrow with every sibling hidden underneath
  // it. Carried into the lane so a detoured fan is still a fan.
  laneOffset = 0,
): { points: Point[]; labelAnchor: Point } {
  const lead = ends?.lead ?? [base.points[0]];
  const tail = ends?.tail ?? [base.points[base.points.length - 1]];
  const start = lead[lead.length - 1];
  const end = tail[0];
  const mk = (mid: Point[]): Point[] => (ends
    ? simplifyPath([...lead, ...mid, ...tail])
    : [start, ...mid, end]);
  const blocking = obstacles.filter((box) =>
    base.points.some((_, i) => i > 0 && segmentHitsBox(base.points[i - 1], base.points[i], box, margin)),
  );
  const cluster = blocking.length ? blocking : obstacles;
  const minBX = Math.min(...cluster.map((b) => b.x));
  const maxBX = Math.max(...cluster.map((b) => b.x + b.w));
  const minBY = Math.min(...cluster.map((b) => b.y));
  const maxBY = Math.max(...cluster.map((b) => b.y + b.h));
  const gap = 18;
  const spread = (lane: number): number => lane + laneOffset;

  type Candidate = { points: Point[]; labelAnchor: Point };
  const candidates: Candidate[] = [];
  if (horizontal) {
    // Route the vertical connector just past the cluster, on the roomier side.
    for (const laneX of [spread(maxBX + gap), spread(minBX - gap)]) {
      candidates.push({
        points: mk([{ x: laneX, y: start.y }, { x: laneX, y: end.y }]),
        labelAnchor: { x: laneX, y: (start.y + end.y) / 2 },
      });
    }
    // Take a clear horizontal gutter above / below the cluster.
    for (const laneY of [spread(minBY - gap), spread(maxBY + gap)]) {
      candidates.push({
        points: mk([{ x: start.x, y: laneY }, { x: end.x, y: laneY }]),
        labelAnchor: { x: (start.x + end.x) / 2, y: laneY },
      });
    }
  } else {
    for (const laneY of [spread(maxBY + gap), spread(minBY - gap)]) {
      candidates.push({
        points: mk([{ x: start.x, y: laneY }, { x: end.x, y: laneY }]),
        labelAnchor: { x: (start.x + end.x) / 2, y: laneY },
      });
    }
    for (const laneX of [spread(minBX - gap), spread(maxBX + gap)]) {
      candidates.push({
        points: mk([{ x: laneX, y: start.y }, { x: laneX, y: end.y }]),
        labelAnchor: { x: laneX, y: (start.y + end.y) / 2 },
      });
    }
  }

  let best = base;
  let bestBlocked = countBlocked(base.points, obstacles, margin);
  const consider = (candidate: Candidate): boolean => {
    const blocked = countBlocked(candidate.points, obstacles, margin);
    if (blocked === 0) return true;
    if (blocked < bestBlocked) { best = candidate; bestBlocked = blocked; }
    return false;
  };
  for (const candidate of candidates) if (consider(candidate)) return candidate;
  // Nothing past the cluster worked. Fall through to the gutters between the
  // obstacle rows and columns — appended after the originals on purpose, so a
  // route that already had a clear detour keeps exactly the one it had and only
  // the previously-unsolvable cases change.
  const gutters: Candidate[] = [];
  for (const lane of clearLanes(obstacles.map((b) => [b.y, b.y + b.h] as [number, number]), start.y, end.y, margin)) {
    const laneY = spread(lane);
    gutters.push({
      points: mk([{ x: start.x, y: laneY }, { x: end.x, y: laneY }]),
      labelAnchor: { x: (start.x + end.x) / 2, y: laneY },
    });
  }
  for (const lane of clearLanes(obstacles.map((b) => [b.x, b.x + b.w] as [number, number]), start.x, end.x, margin)) {
    const laneX = spread(lane);
    gutters.push({
      points: mk([{ x: laneX, y: start.y }, { x: laneX, y: end.y }]),
      labelAnchor: { x: laneX, y: (start.y + end.y) / 2 },
    });
  }
  for (const candidate of gutters) if (consider(candidate)) return candidate;

  // Every lane above keeps the sides the dominant axis chose. When the facing
  // side of a tile is flush against a neighbour — a service on the seam of a
  // grid, the shape a wide estate always takes — no lane can reach it, because
  // the last stretch is inside the neighbour whatever route precedes it. So try
  // the other connection sites. Accepted only when the result is completely
  // clear, so a route that already had a clear lane keeps exactly the one it
  // had and only the previously-unsolvable cases change.
  if (allowResite && bestBlocked > 0 && source && target) {
    // Lanes at a relaxed margin as well as the comfortable one. The 6px
    // clearance is a preference, not a requirement: on a grid whose gutters are
    // 10px wide it merges every column into a single span and offers no lane at
    // all, so the router could not see the one route that crosses nothing and
    // drew an arrow the full height of three tiles it does not connect. A line
    // 5px from a tile edge is a far better drawing than a line through it.
    const spansX = obstacles.map((b) => [b.x, b.x + b.w] as [number, number]);
    const spansY = obstacles.map((b) => [b.y, b.y + b.h] as [number, number]);
    const relaxed = (spans: Array<[number, number]>, from: number, to: number): number[] => {
      for (const m of [margin, 2, 0]) {
        const lanes = clearLanes(spans, from, to, m);
        if (lanes.length > 0) return lanes;
      }
      return [];
    };
    const lanesX = relaxed(spansX, start.x, end.x);
    const lanesY = relaxed(spansY, start.y, end.y);
    // A stub has to fit the gutter it stands in. At 18px it lands inside the
    // neighbour on a packed grid, which makes every route from that side wrong
    // before it has gone anywhere.
    let strict: Candidate | null = null;
    for (const stub of [gap, 6, 3]) {
      for (const targetSide of BOX_SIDES) {
        for (const sourceSide of BOX_SIDES) {
          const vertical = isVerticalSide(sourceSide) === isVerticalSide(targetSide)
            ? !isVerticalSide(sourceSide)
            : true;
          const lanes = vertical ? lanesY : lanesX;
          for (const candidate of routesBetweenSides(source, sourceSide, target, targetSide, stub, lanes)) {
            if (pathEntersBox(candidate.points, source) || pathEntersBox(candidate.points, target)) continue;
            if (countBlocked(candidate.points, obstacles, margin) === 0) return candidate;
            // Second tier: clear of every tile's interior, even though it does
            // not keep the full comfort margin. Held back rather than returned
            // so a route that can have the margin still gets it.
            if (!strict && !obstacles.some((box) => pathEntersBox(candidate.points, box))) strict = candidate;
          }
        }
      }
    }
    if (strict) return strict;
  }
  return best;
}

/**
 * A small loop stub for a self-referencing edge (source === target). React Flow
 * hid these entirely; here they bump off the right edge so the connector — and
 * its label — stay visible. Parallel self-loops grow their extension.
 */
export function selfLoopRoute(box: ExportBox, ordinal = 0): { points: Point[]; labelAnchor: Point } {
  const extension = Math.min(box.w, box.h) * 0.5 + 14 + ordinal * 10;
  const topY = box.y + box.h * 0.32;
  const bottomY = box.y + box.h * 0.68;
  const rightX = box.x + box.w;
  const outerX = rightX + extension;
  return {
    points: [
      { x: rightX, y: topY },
      { x: outerX, y: topY },
      { x: outerX, y: bottomY },
      { x: rightX, y: bottomY },
    ],
    labelAnchor: { x: outerX, y: (topY + bottomY) / 2 },
  };
}

function readEdgeLabel(edge: Edge): string {
  const dataLabel = (edge.data as { label?: unknown } | undefined)?.label;
  if (typeof dataLabel === 'string' && dataLabel.trim()) return dataLabel.trim();
  if (typeof edge.label === 'string' && edge.label.trim()) return edge.label.trim();
  return '';
}

function isDashed(edge: Edge): boolean {
  if (edge.animated) return true;
  const dash = (edge.style as { strokeDasharray?: unknown } | undefined)?.strokeDasharray;
  return typeof dash === 'string' && dash.trim().length > 0 && dash !== 'none' && dash !== '0';
}

function edgeConnectionType(edge: Edge): DiagramConnectionType {
  return normalizeConnectionType((edge.data as { connectionType?: unknown } | undefined)?.connectionType);
}

/**
 * An edge stores the endpoints in the order they were connected, but an edge
 * with `direction: 'reverse'` moves only its arrowhead — the tuple is left
 * alone. Every exporter draws a single head at the target end, so following
 * the tuple pointed the arrow the opposite way to the canvas. Orienting the
 * route here fixes PPTX, VSDX, Draw.io and HTML at once and keeps the
 * polyline running from the arrow's tail, which the exporters rely on.
 */
function orientEdge(edge: Edge): { fromId: string; toId: string; bidirectional: boolean } {
  const direction = (edge.data as { direction?: unknown } | undefined)?.direction;
  if (direction === 'reverse') {
    return { fromId: edge.target, toId: edge.source, bidirectional: false };
  }
  return { fromId: edge.source, toId: edge.target, bidirectional: direction === 'bidirectional' };
}

/**
 * Alternating perpendicular offset for the Nth parallel edge: 0, +16, -16…
 *
 * The spread is capped, because a fan is drawn between two services that have
 * neighbours: ten edges stepping 16 px apart reach 80 px out, far enough for
 * the outer arrows to run through the row above and below and collect their
 * labels. Past the cap the arrows keep their order and simply crowd together —
 * which is what a cable bundle looks like anyway, and each one still carries
 * its own numbered callout.
 */
function parallelOffset(ordinal: number, siblings = 0): number {
  if (ordinal <= 0) return 0;
  const step = 16;
  const rungs = Math.ceil(ordinal / 2);
  const deepest = Math.max(rungs, Math.ceil(Math.max(siblings - 1, 0) / 2));
  const magnitude = deepest * step <= MAX_FAN_SPREAD
    ? rungs * step
    : (rungs * MAX_FAN_SPREAD) / deepest;
  return ordinal % 2 === 1 ? magnitude : -magnitude;
}

function pairKey(a: string, b: string): string {
  return a < b ? `${a}|${b}` : `${b}|${a}`;
}

/**
 * Build orthogonal routes for every edge whose endpoints exist on the canvas.
 * Self-loops become a visible stub, parallel edges are fanned out by ordinal,
 * and each route carries its canonical connection colour/dash so PPTX, VSDX,
 * Draw.io and HTML stay in agreement with the PNG legend.
 */
/**
 * Workflow step declared on an edge, if any. Delegates to the single shared
 * predicate so the geometry, the mapper, the canvas and the deck's numbered
 * list can never disagree about which steps are usable.
 */
function readStepNumber(edge: Edge): number | undefined {
  return readStepValue((edge.data as { stepNumber?: unknown } | undefined)?.stepNumber);
}

export function buildExportRoutes(
  edges: Edge[],
  boxes: Map<string, ExportBox>,
  options: { obstacles?: ExportBox[] } = {},
): ExportRoute[] {
  const routes: ExportRoute[] = [];
  const obstacles = options.obstacles
    ?? [...boxes.values()].filter((box) => box.kind === 'service');
  const ordinals = new Map<string, number>();
  // How deep each fan is, so the whole bundle can be scaled to a spread that
  // does not reach into its neighbours' rows.
  const fanSizes = new Map<string, number>();
  for (const edge of edges) {
    const { fromId, toId } = orientEdge(edge);
    if (!boxes.has(fromId) || !boxes.has(toId) || fromId === toId) continue;
    const key = pairKey(fromId, toId);
    fanSizes.set(key, (fanSizes.get(key) ?? 0) + 1);
  }

  // Which point of which side each hop attaches to.
  //
  // The router is a pure function of one edge, so two hops that meet the same
  // side of the same service are handed the identical anchor — the centre of
  // that side. On a chain laid out in rows that is guaranteed: hop n leaves a
  // tile head-on along the row centre line, and hop n-1 arrives at the same
  // tile around a corner, whose final leg runs along that same centre line.
  // The two are then drawn on top of each other and the shorter one is simply
  // not on the slide. Reference architectures never do this; every arrow into
  // a box lands on its own point.
  //
  // Only the elbow form is moved. A straight hop keeps the centre of the side,
  // because that is the anchor a reader expects and the only one that keeps the
  // line straight; the hop that already turns a corner can afford to turn it
  // slightly earlier. Offsets are dealt from ±1, ±2 … so no elbow ever lands on
  // the centre, and no two elbows on one side land on each other.
  const elbowEnds = new Map<string, { edgeId: string; end: 'source' | 'target'; along: number }[]>();
  const shifts = new Map<string, number>();
  for (const edge of edges) {
    const { fromId, toId } = orientEdge(edge);
    const source = boxes.get(fromId);
    const target = boxes.get(toId);
    if (!source || !target || fromId === toId) continue;
    const sc = centre(source);
    const tc = centre(target);
    const dx = tc.x - sc.x;
    const dy = tc.y - sc.y;
    const horizontal = Math.abs(dx) >= Math.abs(dy);
    const straight = horizontal ? Math.abs(sc.y - tc.y) < 0.5 : Math.abs(sc.x - tc.x) < 0.5;
    if (straight) continue;
    const sourceSide = horizontal ? (dx >= 0 ? 'E' : 'W') : (dy >= 0 ? 'S' : 'N');
    const targetSide = horizontal ? (dx >= 0 ? 'W' : 'E') : (dy >= 0 ? 'N' : 'S');
    const add = (nodeId: string, side: string, end: 'source' | 'target', along: number): void => {
      const key = `${nodeId}#${side}`;
      const list = elbowEnds.get(key) ?? [];
      list.push({ edgeId: String(edge.id), end, along });
      elbowEnds.set(key, list);
    };
    add(source.id, sourceSide, 'source', horizontal ? tc.y : tc.x);
    add(target.id, targetSide, 'target', horizontal ? sc.y : sc.x);
  }
  for (const [key, ends] of elbowEnds) {
    const box = boxes.get(key.slice(0, key.lastIndexOf('#')));
    if (!box) continue;
    const side = key.slice(key.lastIndexOf('#') + 1);
    const extent = side === 'N' || side === 'S' ? box.w : box.h;
    // Rank by where the far end of the hop sits, so the arrows leave the side
    // in the same order they arrive at their destinations and never cross each
    // other on the tile itself. Ties break on edge id so exports stay
    // deterministic.
    const sorted = [...ends].sort((a, b) => (a.along - b.along) || a.edgeId.localeCompare(b.edgeId));
    const step = Math.min(extent / (2 * (sorted.length + 1)), 26);
    // Monotone along the side. Dealing ±1, ±2 … by rank instead put the k-th
    // nearest destination alternately above and below the centre, so a six-way
    // fan out of a front door left the side in the order 4 2 0 1 3 5 and drew
    // itself as a braid. An odd fan is nudged half a step so nothing lands on
    // the centre, which belongs to the straight hops.
    const middle = (sorted.length - 1) / 2;
    const offCentre = sorted.length % 2 === 1 ? 0.5 : 0;
    sorted.forEach((entry, i) => {
      shifts.set(`${entry.edgeId}#${entry.end}`, (i - middle + offCentre) * step);
    });
  }

  for (const edge of edges) {
    const { fromId, toId, bidirectional } = orientEdge(edge);
    const source = boxes.get(fromId);
    const target = boxes.get(toId);
    if (!source || !target) continue;
    const type = edgeConnectionType(edge);
    const style = connectionStyleFor(type);
    const dashed = style.dashed || isDashed(edge);
    const label = readEdgeLabel(edge);
    const key = pairKey(source.id, target.id);
    const ordinal = ordinals.get(key) ?? 0;
    ordinals.set(key, ordinal + 1);

    const fanOffset = source.id === target.id ? 0 : parallelOffset(ordinal, fanSizes.get(key) ?? 1);
    const geometry = source.id === target.id
      ? selfLoopRoute(source, ordinal)
      : routeOrthogonal(source, target, {
        offset: fanOffset,
        obstacles: obstacles.filter((box) => box.id !== source.id && box.id !== target.id),
        sourceShift: shifts.get(`${String(edge.id)}#source`) ?? 0,
        targetShift: shifts.get(`${String(edge.id)}#target`) ?? 0,
        solo: (fanSizes.get(key) ?? 1) <= 1,
      });

    routes.push({
      id: edge.id,
      sourceId: fromId,
      targetId: toId,
      label,
      connectionType: type,
      color: style.color,
      dashed,
      dashPattern: dashed ? (style.dashPattern ?? '6, 4') : undefined,
      opacity: style.opacity,
      ordinal,
      fanOffset,
      bidirectional,
      isSelfLoop: source.id === target.id,
      points: geometry.points,
      labelAnchor: geometry.labelAnchor,
      ...(readStepNumber(edge) !== undefined ? { stepNumber: readStepNumber(edge) } : {}),
    });
  }
  return routes;
}

/**
 * Split nodes into painting order: groups first (they sit behind), then
 * services. Both lists keep their canvas order so exports stay deterministic.
 */
export function partitionBoxes(boxes: Map<string, ExportBox>): {
  groups: ExportBox[];
  services: ExportBox[];
} {
  const groups: ExportBox[] = [];
  const services: ExportBox[] = [];
  for (const box of boxes.values()) {
    if (box.kind === 'group') groups.push(box);
    else services.push(box);
  }
  return { groups, services };
}

/**
 * Where the far-placed nodes go on a drawing whose fit had to trim them.
 *
 * Returns the boxes with the strays parked and the bounds that now contain
 * them. Clamping is a layout decision, made once here: when each slice clamped
 * independently, in inches and against its own window's transform, a stray
 * landed somewhere different on every slide, the hop touching it was planned
 * for none of them, and — because that clamp targets the printable frame,
 * which is larger than the drawing sitting centred inside it — the stray's
 * coordinates stayed outside the drawing, so no window ever claimed the arrow.
 *
 * The strays are PACKED into a strip in the margin beside the drawing, keeping
 * their reading order but discarding the empty space between them. Translating
 * the cloud rigidly instead — preserving the author's spacing — made the parked
 * drawing *larger* than never trimming at all: two strays 9000px either side of
 * a 8.4in cluster produced a 198in drawing, a 199in Visio sheet (Visio refuses
 * anything past 200in) and 4pt type on the fixed-size deck. Whatever the
 * spacing between two far-flung nodes was meant to convey, it is not worth 190
 * inches of paper.
 *
 * Each stray is parked with the zone that contains it, because a group moved
 * away from its children stops being their container. Parking inside the
 * content bounds is worse than the margin: on a full grid every corner is
 * occupied, so the stray is simply drawn on top of a service.
 *
 * Trimming outliers exists to make the drawing smaller, and packing keeps that
 * promise on its own for strays that are loose boxes. A zone is not bounded
 * that way, so the invariant is still asserted at the end and the untrimmed
 * layout returned when parking would grow the drawing.
 */
export function clampedBoxes(
  boxes: Map<string, ExportBox>,
  bounds: Bounds,
): { boxes: Map<string, ExportBox>; bounds: Bounds } {
  const strays: [string, ExportBox][] = [];
  const clipped: [string, ExportBox][] = [];
  for (const [id, box] of boxes) {
    const overlaps = box.x < bounds.maxX && box.x + box.w > bounds.minX
      && box.y < bounds.maxY && box.y + box.h > bounds.minY;
    const outside = box.x < bounds.minX || box.y < bounds.minY
      || box.x + box.w > bounds.maxX || box.y + box.h > bounds.maxY;
    if (!outside) continue;
    // A zone that starts inside the drawing and runs out of it is not a stray,
    // it is a band — a compliance scope drawn across the architecture to one
    // remote service. Moving it destroys its meaning and, because it is wider
    // than the whole drawing, packing it is the rigid translation this function
    // exists to avoid: an 8800px scope band parked as a unit produced a 101.7in
    // drawing where never trimming at all gave 95.2in. Clip it to the drawing
    // instead. It still contains the services it did before; only the empty
    // reach towards the parked member is given up.
    if (overlaps && box.kind === 'group') clipped.push([id, box]);
    else if (!overlaps) strays.push([id, box]);
  }
  if (strays.length === 0 && clipped.length === 0) return { boxes: new Map(boxes), bounds };

  // Membership is what the author declared, never what happens to overlap: a
  // compliance band drawn across an architecture owns nothing it crosses, and
  // reading containment geometrically once let one claim half a grid that
  // belonged to another zone, tearing the grid down the middle and moving 55%
  // of the drawing — past the 40% the outlier trim's own majority floor allows.
  //
  // An earlier version of this comment argued the two definitions had become
  // indistinguishable from outside, on the grounds that anything geometrically
  // inside a parked zone is itself already a stray being parked. The premise
  // is true and the conclusion does not follow: gathering a box changes which
  // *cluster* it lands in, and clusters are packed into separate slots. A
  // policy service standing inside a parked region's rectangle is carried with
  // that region under geometric containment and packed into a slot of its own
  // under declared, so the two produce visibly different sheets — the service
  // ends up outside the boundary it was drawn inside. `pipeline-region` is the
  // fixture that shows it. Declared is still the right definition, because it
  // is what the drawing says rather than what its coordinates imply, but it is
  // a choice this code makes and not a distinction without a difference.
  const holds = (parent: ExportBox, child: ExportBox): boolean => parent !== child && child.parent === parent.id;

  // A stray zone takes everything it contains with it, stray or not: the zone
  // is what left the drawing, and a frame parked away from its services is
  // drawn as an empty box next to an unexplained cluster.
  const strayZones = strays
    .filter(([, box]) => box.kind === 'group')
    .map(([, box]) => box)
    .sort((a, b) => b.w * b.h - a.w * a.h);
  const claimed = new Set<string>();
  const clusters: Array<{ ids: string[]; x: number; y: number; w: number; h: number }> = [];
  const cluster = (ids: string[]): void => {
    const members = ids.map((id) => boxes.get(id)!);
    const x = Math.min(...members.map((b) => b.x));
    const y = Math.min(...members.map((b) => b.y));
    clusters.push({
      ids,
      x,
      y,
      w: Math.max(...members.map((b) => b.x + b.w)) - x,
      h: Math.max(...members.map((b) => b.y + b.h)) - y,
    });
  };
  for (const zone of strayZones) {
    if (claimed.has(zone.id)) continue;
    const ids = [zone.id];
    claimed.add(zone.id);
    for (const [id, box] of boxes) {
      if (claimed.has(id) || !holds(zone, box)) continue;
      claimed.add(id);
      ids.push(id);
    }
    cluster(ids);
  }
  for (const [id] of strays) {
    if (claimed.has(id)) continue;
    claimed.add(id);
    cluster([id]);
  }

  const contentW = Math.max(1, bounds.maxX - bounds.minX);
  const contentH = Math.max(1, bounds.maxY - bounds.minY);
  const GAP = 60;
  const PITCH = 40;
  // Parked on the side they drifted off, so the hop to a stray still runs the
  // way the author drew it and the reader's mental map survives the move.
  const cloudX = clusters.length > 0 ? clusters.reduce((sum, c) => sum + c.x + c.w / 2, 0) / clusters.length : bounds.maxX;
  const cloudY = clusters.length > 0 ? clusters.reduce((sum, c) => sum + c.y + c.h / 2, 0) / clusters.length : bounds.maxY;
  const awayX = cloudX - (bounds.minX + bounds.maxX) / 2;
  const awayY = cloudY - (bounds.minY + bounds.maxY) / 2;
  const column = Math.abs(awayX) >= Math.abs(awayY);

  // Reading order down the column or along the row, so the strip lists the
  // strays in the order the drawing did.
  clusters.sort((a, b) => (column ? a.y - b.y || a.x - b.x : a.x - b.x || a.y - b.y));

  // Pack into the strip, wrapping onto a second lane rather than running past
  // the drawing it sits beside.
  const limit = column
    ? Math.max(contentH, ...clusters.map((c) => c.h))
    : Math.max(contentW, ...clusters.map((c) => c.w));
  const slots: Array<{ dx: number; dy: number }> = [];
  let lane = 0;
  let laneDepth = 0;
  let along = 0;
  let stripDepth = 0;
  let stripAlong = 0;
  for (const c of clusters) {
    const size = column ? c.h : c.w;
    const depth = column ? c.w : c.h;
    if (along > 0 && along + size > limit) {
      lane += laneDepth + PITCH;
      laneDepth = 0;
      along = 0;
    }
    slots.push(column ? { dx: lane, dy: along } : { dx: along, dy: lane });
    laneDepth = Math.max(laneDepth, depth);
    along += size + PITCH;
    stripDepth = Math.max(stripDepth, lane + laneDepth);
    stripAlong = Math.max(stripAlong, along - PITCH);
  }
  const stripW = column ? stripDepth : stripAlong;
  const stripH = column ? stripAlong : stripDepth;
  const origin = column
    ? { x: awayX >= 0 ? bounds.maxX + GAP : bounds.minX - GAP - stripW, y: bounds.minY }
    : { x: bounds.minX, y: awayY >= 0 ? bounds.maxY + GAP : bounds.minY - GAP - stripH };

  const moved = new Map<string, ExportBox>(boxes);
  clusters.forEach((c, index) => {
    const dx = origin.x + slots[index].dx - c.x;
    const dy = origin.y + slots[index].dy - c.y;
    for (const id of c.ids) {
      const box = boxes.get(id)!;
      moved.set(id, { ...box, x: box.x + dx, y: box.y + dy });
    }
  });

  const parked = {
    minX: Math.min(bounds.minX, origin.x),
    minY: Math.min(bounds.minY, origin.y),
    maxX: Math.max(bounds.maxX, origin.x + stripW),
    maxY: Math.max(bounds.maxY, origin.y + stripH),
  };
  for (const [id, box] of clipped) {
    const x = Math.max(parked.minX, box.x);
    const y = Math.max(parked.minY, box.y);
    moved.set(id, {
      ...box,
      x,
      y,
      w: Math.max(1, Math.min(parked.maxX, box.x + box.w) - x),
      h: Math.max(1, Math.min(parked.maxY, box.y + box.h) - y),
    });
  }
  // Trimming outliers exists to make the drawing smaller, so parking can never
  // legitimately return bounds larger than not trimming at all. Packing alone
  // upholds that on every layout whose strays are loose boxes, but a zone can
  // be arbitrarily large, and the case that proves the guard is not decoration
  // is a compliance band drawn across the drawing: parked 101.7in against a
  // 95.2in untrimmed original.
  const full = computeBounds(boxes.values());
  if (
    parked.maxX - parked.minX > full.maxX - full.minX + 1e-6
    || parked.maxY - parked.minY > full.maxY - full.minY + 1e-6
  ) {
    return { boxes: new Map(boxes), bounds: full };
  }
  return { boxes: moved, bounds: parked };
}