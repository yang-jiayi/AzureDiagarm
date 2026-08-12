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
  const centresX = all.map((box) => box.x + box.w / 2).sort((a, b) => a - b);
  const centresY = all.map((box) => box.y + box.h / 2).sort((a, b) => a - b);
  const [qx1, qx3] = quartiles(centresX);
  const [qy1, qy3] = quartiles(centresY);
  const iqrX = qx3 - qx1;
  const iqrY = qy3 - qy1;
  const loX = qx1 - k * iqrX;
  const hiX = qx3 + k * iqrX;
  const loY = qy1 - k * iqrY;
  const hiY = qy3 + k * iqrY;

  const kept = all.filter((box) => {
    const cx = box.x + box.w / 2;
    const cy = box.y + box.h / 2;
    return cx >= loX && cx <= hiX && cy >= loY && cy <= hiY;
  });

  // Only trust the trimmed set when it keeps the clear majority of the diagram;
  // otherwise fall back to the full bounds to avoid hiding real content.
  if (kept.length >= 2 && kept.length >= Math.ceil(all.length * 0.6)) {
    return computeBounds(kept);
  }
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
  if (horizontal) {
    const startX = dx >= 0 ? source.x + source.w : source.x;
    const endX = dx >= 0 ? target.x : target.x + target.w;
    const startY = sc.y + offset;
    const endY = tc.y + offset;
    if (Math.abs(startY - endY) < 0.5) {
      base = {
        points: [{ x: startX, y: startY }, { x: endX, y: endY }],
        labelAnchor: { x: (startX + endX) / 2, y: startY },
      };
    } else {
      const midX = (startX + endX) / 2;
      base = {
        points: [
          { x: startX, y: startY },
          { x: midX, y: startY },
          { x: midX, y: endY },
          { x: endX, y: endY },
        ],
        labelAnchor: { x: midX, y: (startY + endY) / 2 },
      };
    }
  } else {
    const startY = dy >= 0 ? source.y + source.h : source.y;
    const endY = dy >= 0 ? target.y : target.y + target.h;
    const startX = sc.x + offset;
    const endX = tc.x + offset;
    if (Math.abs(startX - endX) < 0.5) {
      base = {
        points: [{ x: startX, y: startY }, { x: endX, y: endY }],
        labelAnchor: { x: startX, y: (startY + endY) / 2 },
      };
    } else {
      const midY = (startY + endY) / 2;
      base = {
        points: [
          { x: startX, y: startY },
          { x: startX, y: midY },
          { x: endX, y: midY },
          { x: endX, y: endY },
        ],
        labelAnchor: { x: (startX + endX) / 2, y: midY },
      };
    }
  }

  if (obstacles.length === 0) return base;
  const margin = 6;
  if (countBlocked(base.points, obstacles, margin) === 0) return base;
  return bestDetour(base, obstacles, horizontal, margin);
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
 * Try a handful of deterministic detours (shift the connecting line past the
 * blocking cluster, or take a clear perpendicular gutter) and keep the first
 * fully-clear candidate — otherwise the least-blocked one. Cheap and stable.
 */
function bestDetour(
  base: { points: Point[]; labelAnchor: Point },
  obstacles: ExportBox[],
  horizontal: boolean,
  margin: number,
): { points: Point[]; labelAnchor: Point } {
  const start = base.points[0];
  const end = base.points[base.points.length - 1];
  const blocking = obstacles.filter((box) =>
    base.points.some((_, i) => i > 0 && segmentHitsBox(base.points[i - 1], base.points[i], box, margin)),
  );
  const cluster = blocking.length ? blocking : obstacles;
  const minBX = Math.min(...cluster.map((b) => b.x));
  const maxBX = Math.max(...cluster.map((b) => b.x + b.w));
  const minBY = Math.min(...cluster.map((b) => b.y));
  const maxBY = Math.max(...cluster.map((b) => b.y + b.h));
  const gap = 18;

  const candidates: Array<{ points: Point[]; labelAnchor: Point }> = [];
  if (horizontal) {
    // Route the vertical connector just past the cluster, on the roomier side.
    for (const laneX of [maxBX + gap, minBX - gap]) {
      const points = [start, { x: laneX, y: start.y }, { x: laneX, y: end.y }, end];
      candidates.push({ points, labelAnchor: { x: laneX, y: (start.y + end.y) / 2 } });
    }
    // Take a clear horizontal gutter above / below the cluster.
    for (const laneY of [minBY - gap, maxBY + gap]) {
      const points = [start, { x: start.x, y: laneY }, { x: end.x, y: laneY }, end];
      candidates.push({ points, labelAnchor: { x: (start.x + end.x) / 2, y: laneY } });
    }
  } else {
    for (const laneY of [maxBY + gap, minBY - gap]) {
      const points = [start, { x: start.x, y: laneY }, { x: end.x, y: laneY }, end];
      candidates.push({ points, labelAnchor: { x: (start.x + end.x) / 2, y: laneY } });
    }
    for (const laneX of [minBX - gap, maxBX + gap]) {
      const points = [start, { x: laneX, y: start.y }, { x: laneX, y: end.y }, end];
      candidates.push({ points, labelAnchor: { x: laneX, y: (start.y + end.y) / 2 } });
    }
  }

  let best = base;
  let bestBlocked = countBlocked(base.points, obstacles, margin);
  for (const candidate of candidates) {
    const blocked = countBlocked(candidate.points, obstacles, margin);
    if (blocked === 0) return candidate;
    if (blocked < bestBlocked) { best = candidate; bestBlocked = blocked; }
  }
  // Nothing past the cluster worked. Fall through to the gutters between the
  // obstacle rows and columns — appended after the originals on purpose, so a
  // route that already had a clear detour keeps exactly the one it had and only
  // the previously-unsolvable cases change.
  const gutters: Array<{ points: Point[]; labelAnchor: Point }> = [];
  for (const laneY of clearLanes(obstacles.map((b) => [b.y, b.y + b.h] as [number, number]), start.y, end.y, margin)) {
    gutters.push({
      points: [start, { x: start.x, y: laneY }, { x: end.x, y: laneY }, end],
      labelAnchor: { x: (start.x + end.x) / 2, y: laneY },
    });
  }
  for (const laneX of clearLanes(obstacles.map((b) => [b.x, b.x + b.w] as [number, number]), start.x, end.x, margin)) {
    gutters.push({
      points: [start, { x: laneX, y: start.y }, { x: laneX, y: end.y }, end],
      labelAnchor: { x: laneX, y: (start.y + end.y) / 2 },
    });
  }
  for (const candidate of gutters) {
    const blocked = countBlocked(candidate.points, obstacles, margin);
    if (blocked === 0) return candidate;
    if (blocked < bestBlocked) { best = candidate; bestBlocked = blocked; }
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
