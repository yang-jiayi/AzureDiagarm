// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Native-vector SVG export.
 *
 * The SVG this app produced was a screenshot in disguise. `html-to-image`'s
 * `toSvg` takes the live DOM and wraps the whole of it in a single
 * `<foreignObject>` holding XHTML. A browser renders that, so the file looked
 * perfect on the one surface anybody checked it on. `<foreignObject>` is
 * optional in the SVG spec and essentially nothing outside a browser engine
 * implements it, so the same file opened blank in Inkscape, blank in
 * Illustrator, blank through librsvg, blank in Office's Insert > Picture and
 * blank in macOS Preview. An SVG that only a browser can open is a PNG with
 * extra steps, and a user who picks "SVG" is almost always picking it *because*
 * they intend to open it somewhere else and edit it.
 *
 * This module draws the diagram out of the same geometry model that already
 * feeds PowerPoint, Visio and Draw.io — `collectExportBoxes` / `buildExportRoutes`
 * — as real `<rect>`, `<path>`, `<text>` and nested `<svg>` elements. Nothing in
 * the output needs an HTML engine, so every tool opens it and every shape is
 * selectable and editable. Reusing the shared geometry is also what keeps the
 * five formats agreeing with each other and with the canvas.
 *
 * Two output details are load-bearing and must not be "tidied":
 *
 *  - Edge paths carry `class="react-flow__edge-path"` and sit inside a group
 *    with `data-testid="rf__edge-<id>"`. `animateEdgeFlow` and
 *    `sequenceWorkflowSvg` are text transforms that find edges by exactly those
 *    two tokens. Renaming either silently produces an animated SVG with no
 *    animation in it.
 *  - The root element uses integer `width`/`height` and a `viewBox` anchored at
 *    `0 0`. `sequenceWorkflowSvg` grows the canvas for its caption strip with
 *    `/(<svg\b[^>]*\bheight=")\d+(")/`, which does not match a decimal.
 */

import type { Node, Edge } from 'reactflow';

import { readTextAsset } from '../utils/assetSource';
import { categoryAccent, rgba, shade, tint } from '../utils/canvasPalette';
import type { ExportBackground } from '../utils/captureCanvas';
import {
  advanceWidthIn,
  buildExportRoutes,
  collectExportBoxes,
  compactEmptyGutters,
  computeBounds,
  metaSubline,
  partitionBoxes,
  readableTextOn,
  zoneStyleFor,
  type ExportBox,
  type ExportRoute,
} from './diagramExportGeometry';

export interface VectorSvgOptions {
  /** Paint the dark-mode surface set instead of the light one. */
  isDarkMode?: boolean;
  /**
   * The background pattern the user picked on the export panel.
   *
   * `'plain'` is a flat page. `'dots'` and `'grid'` are the same two CSS
   * gradients `applyExportBackground` paints on the capture, reproduced as
   * native `<pattern>` tiles — a gradient-based background is exactly the kind
   * of thing that survives a screenshot and vanishes from a hand-built file if
   * nobody carries it across.
   */
  background?: ExportBackground | boolean;
  /** Margin around the drawing, in px. */
  padding?: number;
  /** Diagram name, emitted as the accessible `<title>`. */
  title?: string;
  /**
   * Icon SVG source keyed by `iconPath`, bypassing the loader.
   *
   * The loader resolves URLs through Vite's asset graph, which does not exist
   * outside a bundle, so a headless renderer or an audit script has no way to
   * get real artwork into the file. Every other exporter already takes the same
   * escape hatch under the same name.
   */
  presetIcons?: Map<string, string>;
}

interface Theme {
  pageBg: string;
  surface: string;
  hairline: string;
  text: string;
  mutedInk: string;
  chipFill: string;
  chipStroke: string;
  chipText: string;
}

/**
 * Surfaces the canvas actually paints, not an approximation of them.
 *
 * `#d8e1ea` is the tile hairline in *both* themes because the canvas sets it as
 * an inline style on the tile, where no stylesheet can reach it. Reproducing
 * the bug-shaped truth is the point: export fidelity means matching what was on
 * screen, including the parts that are arguably wrong.
 */
const LIGHT: Theme = {
  pageBg: '#f8fafc',
  surface: '#ffffff',
  hairline: '#d8e1ea',
  text: '#1f2937',
  mutedInk: '#64748b',
  chipFill: '#ffffff',
  chipStroke: '#cbd5e1',
  chipText: '#334155',
};

const DARK: Theme = {
  pageBg: '#1e293b',
  surface: '#27333d',
  hairline: '#d8e1ea',
  text: '#e5edf7',
  mutedInk: '#94a3b8',
  chipFill: '#0f172a',
  chipStroke: '#475569',
  chipText: '#e2e8f0',
};

const FONT_STACK =
  "'Segoe UI','Yu Gothic UI','Hiragino Sans','Noto Sans JP',system-ui,-apple-system,sans-serif";

const LABEL_PX = 12;
const META_PX = 9;
const ZONE_LABEL_PX = 14;
const CHIP_PX = 10;
const BADGE_R = 9;

/** 2dp keeps the file small and byte-stable between runs of the same diagram. */
const f = (n: number): number => +n.toFixed(2);

function esc(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Width of `text` in px when set at `px`.
 *
 * Goes through the shared metric table rather than guessing a flat per-character
 * advance, so a label wraps in the SVG at the same word the deck wraps it at.
 * `advanceWidthIn` is defined against points; CSS px are 3/4 of a point.
 */
function widthPx(text: string, px: number): number {
  return advanceWidthIn(text, px * 0.75) * 96;
}

/** Greedy wrap that also breaks a single over-long word, which CJK needs. */
function wrapToWidth(text: string, maxPx: number, px: number, maxLines: number): string[] {
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let line = '';

  const pushChunked = (word: string) => {
    let chunk = '';
    for (const char of word) {
      if (chunk && widthPx(chunk + char, px) > maxPx) {
        lines.push(chunk);
        chunk = char;
        if (lines.length >= maxLines) return;
      } else {
        chunk += char;
      }
    }
    line = chunk;
  };

  for (const word of words) {
    if (lines.length >= maxLines) break;
    const candidate = line ? `${line} ${word}` : word;
    if (widthPx(candidate, px) <= maxPx) {
      line = candidate;
      continue;
    }
    if (line) {
      lines.push(line);
      line = '';
      if (lines.length >= maxLines) break;
    }
    if (widthPx(word, px) > maxPx) {
      pushChunked(word);
    } else {
      line = word;
    }
  }
  if (line && lines.length < maxLines) lines.push(line);

  if (!lines.length) return [];

  // Ellipsise rather than silently dropping the tail: a truncated name a reader
  // can see is truncated beats a name that looks complete and is not.
  if (lines.length >= maxLines) {
    const consumed = lines.join(' ').length;
    if (consumed < text.replace(/\s+/g, ' ').length) {
      let last = lines[lines.length - 1];
      while (last && widthPx(`${last}…`, px) > maxPx) last = last.slice(0, -1);
      lines[lines.length - 1] = `${last}…`;
    }
  }
  return lines;
}

/**
 * Inline an icon as a nested `<svg>` with every internal id rewritten.
 *
 * Embedding via `<image href="data:image/svg+xml;base64,…">` would be less code
 * and is what the Draw.io exporter does, because Draw.io wants it that way. It
 * is the wrong choice here: Inkscape has never reliably rasterised an SVG
 * nested inside an `<image>`, and even where it renders the artwork arrives as
 * one opaque object a designer cannot recolour — which defeats the reason to
 * pick SVG over PNG.
 *
 * Inlining means the ids inside the icon join the host document's single id
 * namespace. Azure's icons name their gradients with UUIDs, so two *different*
 * icons will not collide, but the same service placed on two tiles emits the
 * same UUID twice. Browsers resolve a duplicate `url(#id)` to the first match,
 * which happens to be an identical gradient, so it renders — until a tool that
 * validates ids refuses the file, or someone edits one of the two gradients and
 * watches the other change with it. Prefixing per tile costs one regex pass.
 */
function inlineIcon(svgSource: string, uid: string, x: number, y: number, size: number): string {
  let body = svgSource
    .replace(/<\?xml[\s\S]*?\?>/g, '')
    .replace(/<!DOCTYPE[\s\S]*?>/g, '')
    .trim();

  const open = /<svg\b([^>]*)>/i.exec(body);
  if (!open) return '';

  const viewBox = /viewBox\s*=\s*["']([^"']+)["']/i.exec(open[1])?.[1];
  const rawW = parseFloat(/\bwidth\s*=\s*["']([\d.]+)/i.exec(open[1])?.[1] ?? '');
  const rawH = parseFloat(/\bheight\s*=\s*["']([\d.]+)/i.exec(open[1])?.[1] ?? '');
  const box =
    viewBox ??
    (Number.isFinite(rawW) && Number.isFinite(rawH) ? `0 0 ${rawW} ${rawH}` : '0 0 18 18');

  const inner = body.slice(open.index + open[0].length).replace(/<\/svg\s*>\s*$/i, '');

  const ids = new Set<string>();
  for (const match of inner.matchAll(/\sid\s*=\s*["']([^"']+)["']/g)) ids.add(match[1]);

  let scoped = inner;
  for (const id of ids) {
    const q = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    scoped = scoped
      .replace(new RegExp(`(\\sid\\s*=\\s*["'])${q}(["'])`, 'g'), `$1${uid}-${id}$2`)
      .replace(new RegExp(`url\\((['"]?)#${q}\\1\\)`, 'g'), `url(#${uid}-${id})`)
      .replace(new RegExp(`((?:xlink:)?href\\s*=\\s*["'])#${q}(["'])`, 'g'), `$1#${uid}-${id}$2`);
  }

  // Class names in an icon's own <style> block share the document namespace too,
  // and Azure's newer icons use `.cls-1` — a name every one of them reuses.
  scoped = scoped.replace(/\bcls-(\d+)\b/g, `${uid}-cls-$1`);

  return (
    `<svg x="${f(x)}" y="${f(y)}" width="${f(size)}" height="${f(size)}" ` +
    `viewBox="${esc(box)}" overflow="visible">${scoped}</svg>`
  );
}

/**
 * The dotted / ruled canvas background, as a tile the file carries itself.
 *
 * Colours, radius and the 20px gap are the values `applyExportBackground` sets
 * on the capture, so the two exports of the same diagram put the dots in the
 * same places.
 */
function patternFor(background: ExportBackground | boolean | undefined): { id: string; def: string } | null {
  if (background === 'dots') {
    return {
      id: 'bg-dots',
      def:
        `<pattern id="bg-dots" width="20" height="20" patternUnits="userSpaceOnUse">` +
        `<circle cx="10" cy="10" r="1.2" fill="rgba(96, 165, 250, 0.32)"/></pattern>`,
    };
  }
  if (background === 'grid') {
    return {
      id: 'bg-grid',
      def:
        `<pattern id="bg-grid" width="20" height="20" patternUnits="userSpaceOnUse">` +
        `<path d="M 20 0 L 0 0 0 20" fill="none" stroke="rgba(96, 165, 250, 0.24)" ` +
        `stroke-width="1"/></pattern>`,
    };
  }
  return null;
}

async function loadIconSource(iconPath: string): Promise<string | null> {  if (!iconPath) return null;
  try {
    const { loadIcon } = await import('../utils/iconLoader');
    const url = await loadIcon(iconPath);
    if (!url) return null;
    return await readTextAsset(url);
  } catch (error) {
    console.warn(`[SVG Export] Could not load icon ${iconPath}:`, error);
    return null;
  }
}

function zoneMarkup(box: ExportBox, theme: Theme): string {
  const style = zoneStyleFor(box);
  const accent = style.border;
  const label = box.label?.trim();
  const parts = [
    `<rect x="${f(box.x)}" y="${f(box.y)}" width="${f(box.w)}" height="${f(box.h)}" rx="12" ` +
      `fill="${style.bg}" stroke="${accent}" stroke-width="1.5"/>`,
  ];
  if (label) {
    const ink = readableTextOn(style.text, theme.pageBg);
    const room = Math.max(24, box.w - 24);
    const shown = wrapToWidth(label, room, ZONE_LABEL_PX, 1)[0] ?? label;
    parts.push(
      `<text x="${f(box.x + 14)}" y="${f(box.y + 22)}" font-family="${FONT_STACK}" ` +
        `font-size="${ZONE_LABEL_PX}" font-weight="600" fill="${ink}">${esc(shown)}</text>`,
    );
  }
  return `<g data-zone="${esc(box.id)}">${parts.join('')}</g>`;
}

function serviceMarkup(box: ExportBox, iconSvg: string | null, theme: Theme, uid: string): string {
  const accent = categoryAccent(box.category);
  const radius = Math.min(8, box.w / 4, box.h / 4);
  // Same rule the deck uses: a 4px rule on a 12px-wide tile is not a rule, it
  // is a fill, and a fill is the treatment this stripe exists to replace.
  const stripe = Math.min(4, box.w / 4);

  const parts = [
    `<rect x="${f(box.x)}" y="${f(box.y)}" width="${f(box.w)}" height="${f(box.h)}" ` +
      `rx="${f(radius)}" fill="${theme.surface}" stroke="${theme.hairline}" stroke-width="1"/>`,
    `<rect x="${f(box.x)}" y="${f(box.y + radius)}" width="${f(stripe)}" ` +
      `height="${f(Math.max(0, box.h - radius * 2))}" fill="${accent}"/>`,
  ];

  const meta = metaSubline(box);
  const metaH = meta ? META_PX + 3 : 0;
  const iconSize = Math.min(36, box.w * 0.4, box.h * 0.42);
  const textTop = box.y + 8 + iconSize + 4;
  const textRoom = Math.max(16, box.w - stripe - 14);

  if (iconSvg) {
    parts.push(inlineIcon(iconSvg, uid, box.x + (box.w - iconSize) / 2, box.y + 8, iconSize));
  }

  const available = box.y + box.h - 6 - metaH - textTop;
  const maxLines = Math.max(1, Math.min(2, Math.floor(available / (LABEL_PX + 2))));
  const lines = wrapToWidth(box.label ?? '', textRoom, LABEL_PX, maxLines);
  const cx = box.x + stripe / 2 + box.w / 2;
  lines.forEach((line, i) => {
    parts.push(
      `<text x="${f(cx)}" y="${f(textTop + LABEL_PX + i * (LABEL_PX + 2))}" ` +
        `font-family="${FONT_STACK}" font-size="${LABEL_PX}" text-anchor="middle" ` +
        `fill="${theme.text}">${esc(line)}</text>`,
    );
  });

  if (meta) {
    const ink = readableTextOn(theme.mutedInk, theme.surface);
    // `Standard · japaneast · $128/mo` cut to width reads `Standard · japaneast ·…`,
    // which spends four characters telling the reader something was dropped
    // without saying what, and leaves a dangling separator. Drop whole segments
    // from the end instead, so whatever is shown is complete and true.
    const segments = meta.split(' · ');
    let shown = '';
    for (let take = segments.length; take > 0; take -= 1) {
      const candidate = segments.slice(0, take).join(' · ');
      if (widthPx(candidate, META_PX) <= textRoom) {
        shown = candidate;
        break;
      }
    }
    if (!shown) shown = wrapToWidth(segments[0] ?? meta, textRoom, META_PX, 1)[0] ?? '';
    if (shown) {
      parts.push(
        `<text x="${f(cx)}" y="${f(box.y + box.h - 6)}" font-family="${FONT_STACK}" ` +
          `font-size="${META_PX}" text-anchor="middle" fill="${ink}">${esc(shown)}</text>`,
      );
    }
  }

  return `<g data-service="${esc(box.id)}">${parts.join('')}</g>`;
}

function pathData(route: ExportRoute): string {
  const points = route.points;
  if (!points.length) return '';
  const head = `M ${f(points[0].x)} ${f(points[0].y)}`;
  const rest = points.slice(1).map((p) => `L ${f(p.x)} ${f(p.y)}`);
  return [head, ...rest].join(' ');
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function overlaps(a: Rect, b: Rect, pad = 0): boolean {
  return (
    a.x - pad < b.x + b.w &&
    a.x + a.w + pad > b.x &&
    a.y - pad < b.y + b.h &&
    a.y + a.h + pad > b.y
  );
}

/** Point at `distance` along an orthogonal polyline, clamped to its ends. */
function pointAtLength(points: readonly { x: number; y: number }[], distance: number) {
  if (!points.length) return { x: 0, y: 0 };
  let remaining = Math.max(0, distance);
  for (let i = 1; i < points.length; i += 1) {
    const a = points[i - 1];
    const b = points[i];
    const seg = Math.hypot(b.x - a.x, b.y - a.y);
    if (seg <= 0) continue;
    if (remaining <= seg) {
      const t = remaining / seg;
      return { x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t };
    }
    remaining -= seg;
  }
  return points[points.length - 1];
}

function polylineLength(points: readonly { x: number; y: number }[]): number {
  let total = 0;
  for (let i = 1; i < points.length; i += 1) {
    total += Math.hypot(points[i].x - points[i - 1].x, points[i].y - points[i - 1].y);
  }
  return total;
}

interface Chip {
  route: ExportRoute;
  rect: Rect;
}

function chipFor(route: ExportRoute): Chip | null {
  if (!route.label) return null;
  const w = widthPx(route.label, CHIP_PX) + 10;
  const h = CHIP_PX + 8;
  return {
    route,
    rect: { x: route.labelAnchor.x - w / 2, y: route.labelAnchor.y - h / 2, w, h },
  };
}

/**
 * Where a numbered callout can stand without landing on something else.
 *
 * Placing it at a fixed *fraction* along the arrow, which is the obvious thing
 * to do, puts it exactly where the label chip is on any edge short enough for
 * 30% and 50% to be within a badge-width of each other — and short edges are
 * the common case in a grid layout. The rendered proof read "1|TTPS" with the
 * disc sitting on the letter H.
 *
 * So walk the arrow by real arc length instead and take the first station that
 * clears every chip on the drawing and every badge already placed. Testing only
 * against the edge's *own* chip is not enough: a fanned arrow starts inside its
 * neighbour's label.
 */
function badgeCentre(
  route: ExportRoute,
  chips: readonly Chip[],
  placed: readonly Rect[],
): { x: number; y: number } {
  const total = polylineLength(route.points);
  const start = BADGE_R + 8;
  const box = (p: { x: number; y: number }): Rect => ({
    x: p.x - BADGE_R,
    y: p.y - BADGE_R,
    w: BADGE_R * 2,
    h: BADGE_R * 2,
  });
  const clear = (p: { x: number; y: number }): boolean => {
    const b = box(p);
    for (const chip of chips) if (overlaps(b, chip.rect, 3)) return false;
    for (const other of placed) if (overlaps(b, other, 3)) return false;
    return true;
  };

  let best = pointAtLength(route.points, Math.min(start, total * 0.3));
  if (clear(best)) return best;

  const step = 6;
  for (let d = start; d <= Math.max(start, total - start); d += step) {
    const candidate = pointAtLength(route.points, d);
    if (clear(candidate)) return candidate;
  }
  // Nothing along the arrow is clear. Step off it rather than stack discs: an
  // unreadable number is worse than one a reader has to glance sideways for.
  const anchor = pointAtLength(route.points, Math.min(start, total / 2));
  for (const dy of [-1, 1]) {
    for (let k = 1; k <= 3; k += 1) {
      const candidate = { x: anchor.x, y: anchor.y + dy * k * (BADGE_R * 2 + 3) };
      if (clear(candidate)) return candidate;
    }
  }
  return best;
}

function edgeMarkup(
  route: ExportRoute,
  theme: Theme,
  markerId: (color: string) => string,
  chip: Chip | null,
  badge: { x: number; y: number } | null,
): string {
  const d = pathData(route);
  if (!d) return '';

  const dash = route.dashed && route.dashPattern ? ` stroke-dasharray="${route.dashPattern}"` : '';
  const start = route.bidirectional
    ? ` marker-start="url(#${markerId(route.color)}-start)"`
    : '';

  // `data-flow-color` is read by animateEdgeFlow to colour the flowing dot. A
  // captured ReactFlow SVG encodes the same thing inside its marker-end url; a
  // native marker id cannot carry a `#`, so it is stated explicitly instead.
  const path =
    `<path class="react-flow__edge-path" d="${d}" fill="none" stroke="${route.color}" ` +
    `stroke-width="2" stroke-opacity="${f(route.opacity)}" stroke-linecap="round" ` +
    `stroke-linejoin="round"${dash} data-flow-color="${route.color}"${start} ` +
    `marker-end="url(#${markerId(route.color)})"/>`;

  const extras: string[] = [];

  if (chip) {
    const { x, y, w, h } = chip.rect;
    extras.push(
      `<rect x="${f(x)}" y="${f(y)}" width="${f(w)}" height="${f(h)}" rx="4" ` +
        `fill="${theme.chipFill}" fill-opacity="0.94" stroke="${theme.chipStroke}" ` +
        `stroke-width="1"/>`,
      `<text x="${f(x + w / 2)}" y="${f(y + h / 2 + CHIP_PX / 2 - 1)}" ` +
        `font-family="${FONT_STACK}" font-size="${CHIP_PX}" text-anchor="middle" ` +
        `fill="${theme.chipText}">${esc(route.label)}</text>`,
    );
  }

  if (badge && route.stepNumber != null) {
    extras.push(
      `<circle cx="${f(badge.x)}" cy="${f(badge.y)}" r="${BADGE_R}" fill="${route.color}" ` +
        `stroke="${theme.surface}" stroke-width="1.5"/>`,
      `<text x="${f(badge.x)}" y="${f(badge.y + 3.5)}" font-family="${FONT_STACK}" ` +
        `font-size="10" font-weight="700" text-anchor="middle" ` +
        `fill="${readableTextOn('#ffffff', route.color)}">${route.stepNumber}</text>`,
    );
  }

  return `<g class="react-flow__edge" data-testid="rf__edge-${esc(route.id)}">${path}${extras.join('')}</g>`;
}

/**
 * Render the diagram as a standalone, tool-agnostic vector SVG.
 */
export async function exportToSvg(
  nodes: Node[],
  edges: Edge[],
  options: VectorSvgOptions = {},
): Promise<string> {
  const theme = options.isDarkMode ? DARK : LIGHT;
  const pad = options.padding ?? 48;

  const boxes = compactEmptyGutters(collectExportBoxes(nodes));
  const { groups, services } = partitionBoxes(boxes);
  const routes = buildExportRoutes(edges, boxes);
  const bounds = computeBounds(boxes.values());

  const dx = pad - bounds.minX;
  const dy = pad - bounds.minY;
  const width = Math.max(1, Math.ceil(bounds.maxX - bounds.minX + pad * 2));
  const height = Math.max(1, Math.ceil(bounds.maxY - bounds.minY + pad * 2));

  const iconSources = await Promise.all(
    services.map((service) => {
      if (!service.iconPath) return null;
      const preset = options.presetIcons?.get(service.iconPath);
      return preset ?? loadIconSource(service.iconPath);
    }),
  );

  const colors = new Map<string, string>();
  const markerId = (color: string): string => {
    const existing = colors.get(color);
    if (existing) return existing;
    const id = `arrow-${colors.size}`;
    colors.set(color, id);
    return id;
  };
  for (const route of routes) markerId(route.color);

  const markers = [...colors.entries()]
    .map(
      ([color, id]) =>
        `<marker id="${id}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" ` +
        `markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" ` +
        `fill="${color}"/></marker>` +
        `<marker id="${id}-start" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" ` +
        `markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" ` +
        `fill="${color}"/></marker>`,
    )
    .join('');

  const zoneMarkupParts = groups.map((group) => zoneMarkup(group, theme));
  const serviceParts = services.map((service, index) =>
    serviceMarkup(service, iconSources[index], theme, `s${index}`),
  );

  // Chips are sized before any badge is placed, because a badge has to know
  // about every chip on the drawing — including the ones belonging to other
  // arrows — before it can pick a spot that is clear of all of them.
  const chips = routes.map((route) => chipFor(route));
  const knownChips = chips.filter((chip): chip is Chip => chip !== null);
  const placed: Rect[] = [];
  const badges = routes.map((route) => {
    if (route.stepNumber == null) return null;
    const centre = badgeCentre(route, knownChips, placed);
    placed.push({
      x: centre.x - BADGE_R,
      y: centre.y - BADGE_R,
      w: BADGE_R * 2,
      h: BADGE_R * 2,
    });
    return centre;
  });

  const edgeParts = routes.map((route, index) =>
    edgeMarkup(route, theme, markerId, chips[index], badges[index]),
  );

  const background = options.background === false
    ? ''
    : `<rect x="0" y="0" width="${width}" height="${height}" fill="${theme.pageBg}"/>`;
  const pattern = patternFor(options.background);
  const overlay = pattern
    ? `<rect x="0" y="0" width="${width}" height="${height}" fill="url(#${pattern.id})"/>`
    : '';
  const title = options.title?.trim();

  return (
    `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
    `width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" ` +
    `role="img" aria-label="${esc(title || 'Architecture diagram')}">` +
    (title ? `<title>${esc(title)}</title>` : '') +
    `<defs>${markers}${pattern?.def ?? ''}</defs>` +
    background +
    overlay +
    `<g class="react-flow__viewport" transform="translate(${f(dx)},${f(dy)})">` +
    `<g class="zones">${zoneMarkupParts.join('')}</g>` +
    `<g class="react-flow__edges">${edgeParts.join('')}</g>` +
    `<g class="services">${serviceParts.join('')}</g>` +
    `</g></svg>`
  );
}

/** Unused re-exports kept out on purpose; these two are for tests only. */
export const __testing = { wrapToWidth, inlineIcon, widthPx, tint, shade, rgba };
