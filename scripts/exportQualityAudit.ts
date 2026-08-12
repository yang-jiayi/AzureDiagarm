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
import JSZip from 'jszip';
import type { Edge, Node } from 'reactflow';
import { buildDiagramSlidePptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';
import { WRAP_TRIGGER_RATIO } from '../src/utils/serpentineWrap.ts';
import { narrateEdgeCallouts, CATEGORY_STYLES } from '../src/services/diagramExportGeometry.ts';

const OUT = path.join(process.cwd(), 'tmp-export-audit');
const EMU_PER_INCH = 914400;

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
  fontSize: number | null;
  /** Vertical anchor of the text body: 't', 'ctr', or 'b'. */
  anchor: string | null;
  /** Shape fill as RRGGBB, or null when the shape declares no fill. */
  fill: string | null;
  /** Fill opacity, 0..1. A translucent chip shows what is underneath it. */
  fillAlpha: number;
  /** Every drawn text run, with the colour and size it is drawn at. */
  runs: { color: string | null; sizePt: number; bold: boolean; text: string }[];
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
function textWidthIn(text: string, fontSizePt: number): number {
  let units = 0;
  for (const ch of text) {
    units += /[\u3000-\u9fff\uff00-\uffef]/.test(ch) ? 1 : 0.54;
  }
  return (units * fontSizePt) / 72;
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
    const texts = [...body.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((t) => t[1]);
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
    const runs = [...body.matchAll(/<a:r>([\s\S]*?)<\/a:r>/g)].map((r) => {
      const rb = r[1];
      const rpr = /<a:rPr[^>]*>([\s\S]*?)<\/a:rPr>/.exec(rb);
      return {
        color: /<a:srgbClr val="([0-9A-Fa-f]{6})"/.exec(rpr?.[1] ?? '')?.[1]?.toLowerCase() ?? null,
        sizePt: (+(/<a:rPr[^>]*\bsz="(\d+)"/.exec(rb)?.[1] ?? 0) || 1800) / 100,
        bold: /<a:rPr[^>]*\bb="1"/.test(rb),
        text: (/<a:t>([\s\S]*?)<\/a:t>/.exec(rb)?.[1] ?? '').trim(),
      };
    });
    shapes.push({
      name,
      x,
      y,
      w,
      h,
      text: texts.join('').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
      fontSize: sz ? +sz[1] / 100 : null,
      anchor: /<a:bodyPr[^>]*\banchor="([^"]+)"/.exec(body)?.[1] ?? null,
      fill,
      fillAlpha,
      runs,
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

function drawnTextRect(shape: Shape, singleLine = false): { x: number; y: number; w: number; h: number } | null {
  const text = shape.text.trim();
  if (text === '' || !shape.fontSize) return null;
  // A `wrap="none"` run is drawn on one line whatever its width, and a centred
  // one that outgrows its box overflows equally on both sides rather than
  // wrapping. Modelling it as wrapped would claim rows above it that hold
  // nothing and report chips that never touched a glyph.
  const lines = singleLine ? 1 : Math.max(1, Math.ceil(textWidthIn(text, shape.fontSize) / shape.w));
  const h = Math.min(shape.h, (lines * shape.fontSize * 1.22) / 72);
  const w = lines > 1 ? shape.w : (singleLine ? textWidthIn(text, shape.fontSize) : Math.min(shape.w, textWidthIn(text, shape.fontSize)));
  const x = shape.x + (shape.w - w) / 2;
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
function overlapArea(a: Shape, b: Shape): number {
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

function grp(id: string, label: string, x: number, y: number, w: number, h: number): Node {
  return { id, type: 'groupNode', position: { x, y }, style: { width: w, height: h }, data: { label } } as Node;
}

/**
 * Real Architecture-Center step prose, long enough that every row wraps. The
 * whole corpus used `step N` and one-clause labels, so the workflow list was
 * only ever measured with sentences that fit on one line — and pagination
 * assumed exactly that.
 */
function workflowProseScenario(): Scenario {
  const sentences = [
    'The client sends the request to Azure Front Door, which terminates TLS at the edge and applies the WAF ruleset before anything reaches the origin.',
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

interface Report { scenario: string; format: string; issues: string[]; metrics: Record<string, number> }

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
function auditNativeConversion(rawSlides: readonly string[]): { issues: string[]; glued: number; ungluable: number; groups: number } {
  const issues: string[] = [];
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
      if (!after.includes(escapeXml(label.text))) {
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

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
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

async function auditPptx(scenario: Scenario): Promise<Report> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Audit',
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

  const issues: string[] = [];
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
  const native = auditNativeConversion(allSlides);
  issues.push(...native.issues);
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
  const labels = shapes.filter((s) => s.name.startsWith('service-label-'));
  const chips = shapes.filter((s) => s.name.startsWith('connector-label-'));

  const minTileW = Math.min(...tiles.map((t) => t.w));
  const minFont = Math.min(...labels.map((l) => l.fontSize ?? 99));

  for (const label of labels) {
    const font = label.fontSize ?? 11;
    const needed = textWidthIn(label.text, font);
    const lines = Math.ceil(needed / Math.max(label.w, 0.01));
    const charsPerLine = label.w / (font / 72);
    if (charsPerLine < 4) {
      issues.push(`label "${label.text}" has room for only ${charsPerLine.toFixed(1)} chars/line (w=${label.w.toFixed(2)}in @ ${font}pt)`);
    }
    const lineHeight = (font * 1.25) / 72;
    if (lines * lineHeight > label.h + 0.02) {
      issues.push(`label "${label.text}" needs ${lines} lines (${(lines * lineHeight).toFixed(2)}in) but has ${label.h.toFixed(2)}in`);
    }
  }
  for (const chip of chips) {
    if (minTileW > 0 && chip.w > minTileW) {
      issues.push(`edge chip "${chip.text}" is ${(chip.w / minTileW).toFixed(1)}x wider than the smallest node tile`);
    }
  }
  // Collisions are only real between shapes printed on the same sheet. Reading
  // every slide as one pile reported a chip on part 1 as covering a tile on
  // part 3, which turned every tiled deck into a wall of phantom issues and
  // hid whatever was genuinely wrong.
  for (const slideShapes of perSlide) {
    const slideTiles = slideShapes.filter((s) => s.name.startsWith('service-') && !s.name.includes('label') && !s.name.includes('meta'));
    const slideChips = slideShapes.filter((s) => s.name.startsWith('connector-label-'));
    // A label may lean on the two services its own arrow connects. The reader
    // still attributes it correctly — it is touching the very icons it is about
    // — and on a hop shorter than the label there is nowhere else for it to go.
    // Leaning on a THIRD service is a different thing entirely: it hides an
    // unrelated icon and reads as that service's caption. So the bar for a
    // stranger's tile stays at a couple of percent, and an endpoint of the
    // arrow itself is allowed a tenth of its area before it counts as hidden.
    const membersOfZone = new Map<string, number>();
    for (const node of scenario.nodes) {
      if (!node.parentNode) continue;
      membersOfZone.set(node.parentNode, (membersOfZone.get(node.parentNode) ?? 0) + 1);
    }
    const endpointsOf = new Map<string, Set<string>>();
    for (const edge of scenario.edges) {
      endpointsOf.set(edge.id, new Set([`service-${edge.source}`, `service-${edge.target}`]));
    }
    const tileBudget = (annotation: string, tile: Shape): number => {
      const routeId = annotation.replace(/^connector-(label|step)-/, '');
      return endpointsOf.get(routeId)?.has(tile.name) ? 0.1 : 0.02;
    };
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
    // drawn across it.
    for (const title of slideShapes.filter((s) => s.name.startsWith('zone-label-'))) {
      let covered = 0;
      for (const tile of slideTiles) covered += overlapArea(title, tile);
      if (covered > 0.25 * title.w * title.h) {
        const pct = ((covered / (title.w * title.h)) * 100).toFixed(0);
        issues.push(`zone title "${title.text}" is ${pct}% covered by the tiles inside it`);
      }
    }
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
      if (buried > 0.9 * badge.w * badge.h) {
        const worst = slideTiles.reduce((a, b) => (overlapArea(badge, b) > overlapArea(badge, a) ? b : a), slideTiles[0]);
        issues.push(`step badge "${badge.name}" is ${((buried / (badge.w * badge.h)) * 100).toFixed(0)}% buried inside "${worst?.name}"`);
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
      if (!/^(connector-label-|connector-step-|service-)/.test(other.name)) continue;
      const hit = overlapArea(legend, other);
      if (hit <= 0.001) continue;
      issues.push(`connection legend covers ${((hit / Math.max(other.w * other.h, 1e-6)) * 100).toFixed(0)}% of "${other.name}"`);
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
    const label = typeof edge.label === 'string' ? edge.label.trim() : '';
    if (!label || drawnChips.has(edge.id)) continue;
    const badge = drawnBadges.get(edge.id);
    if (badge !== undefined && explained.has(badge)) {
      if (!deckWording.includes(foldWording(label))) {
        issues.push(`edge "${edge.id}" was muted to callout ${badge}, but its wording "${label}" appears nowhere in the deck`);
      }
      continue;
    }
    issues.push(
      badge === undefined
        ? `edge "${edge.id}" is labelled "${label}" but the deck has neither a chip nor a callout for it`
        : `edge "${edge.id}" lost its label "${label}" to callout ${badge}, which no workflow row explains`,
    );
  }
  // Truncation is only acceptable when the full wording survives somewhere the
  // reader can reach. A chip clipped to 42 cells with no workflow row carrying
  // the rest has silently thrown away what the author wrote.
  const truncated = shapes.filter((s) => s.text.includes('…'));
  const stranded = truncated.filter((s) => {
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
    (e) => Number.isInteger((e.data as { stepNumber?: number } | undefined)?.stepNumber),
  );
  const badges = shapes.filter((s) => s.name.startsWith('connector-step-'));
  if (badges.length !== numberedEdges.length) {
    issues.push(`${badges.length} step badges drawn for ${numberedEdges.length} numbered connectors`);
  }
  // Membership alone is permutation-blind: swapping every badge onto the wrong
  // arrow would pass. The object name carries the route id, so check the exact
  // arrow-to-number correspondence instead.
  const expectedByRoute = new Map(
    numberedEdges.map((e) => [e.id, String((e.data as { stepNumber: number }).stepNumber)]),
  );
  for (const badge of badges) {
    const routeId = badge.name.replace(/^connector-step-/, '');
    const want = expectedByRoute.get(routeId);
    if (want === undefined) {
      issues.push(`step badge "${badge.name}" does not belong to any numbered connector`);
    } else if (badge.text !== want) {
      issues.push(`connector ${routeId} is numbered "${badge.text}" but its workflow step is ${want}`);
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
      const lines = Math.max(1, Math.ceil(textWidthIn(row.text.trim(), row.fontSize) / row.w));
      const needed = (lines * row.fontSize * 1.22) / 72;
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
      .filter((d): d is string => !!d),
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
  const serviceIds = scenario.nodes.filter((n) => n.type === 'azureNode').map((n) => n.id);
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
  for (const [name, count] of countByName(badges)) {
    if (count > 1) issues.push(`step badge "${name}" is drawn ${count} times`);
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
  const pagePart = pkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  const media = pkg.parts.filter((p) => /\/media\//i.test(p.path));
  const serviceCount = scenario.nodes.filter((n) => n.type !== 'groupNode').length;
  if (iconPaths.size > 0 && media.length === 0) {
    issues.push(`no embedded icon media parts (expected ~${serviceCount})`);
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

  for (const size of xml.matchAll(/<Cell N="Size" V="([\d.]+)"\/>/g)) {
    minFontIn = Math.min(minFontIn, +size[1]);
  }
  const minFontPt = +(minFontIn * 72).toFixed(2);
  // The floor is PowerPoint's, deliberately. Both exporters draw the same
  // drawing at the same scale, so type that is unreadable in the deck is
  // unreadable on the sheet, and the two must not disagree about where the
  // limit is.
  if (minFontPt < 7) issues.push(`smallest Visio font is ${minFontPt}pt (below the 7pt floor the deck enforces)`);

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
    .map((m) => /<Text>([\s\S]*?)<\/Text>/.exec(m[0])?.[1] ?? '')
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
    const label = typeof edge.label === 'string' ? edge.label.trim() : '';
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

  // Workflow numbering must survive into Visio too, or the same drawing tells
  // a different story in PowerPoint and in Visio. Measured against the repaired
  // edges, which is what both exporters draw from.
  const numberedEdges = narrateEdgeCallouts(scenario.edges).filter(
    (e) => Number.isInteger((e.data as { stepNumber?: number } | undefined)?.stepNumber),
  );
  const badgeBlocks = [...xml.matchAll(/<Shape [^>]*NameU="StepBadge\.\d+"[\s\S]*?<\/Shape>/g)].map((m) => m[0]);
  if (badgeBlocks.length !== numberedEdges.length) {
    issues.push(`${badgeBlocks.length} Visio step badges for ${numberedEdges.length} numbered connectors`);
  }
  const expectedNumbers = new Set(
    numberedEdges.map((e) => String((e.data as { stepNumber: number }).stepNumber)),
  );
  // Service boxes in page coordinates, so a badge that lands on one is caught.
  const serviceBoxes: Array<{ x: number; y: number; w: number; h: number }> = [];
  for (const m of xml.matchAll(
    /NameU="Service\.\d+"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )) {
    const [, pinX, pinY, w, h] = m;
    serviceBoxes.push({ x: +pinX - +w / 2, y: +pinY - +h / 2, w: +w, h: +h });
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
        const inside = serviceBoxes.some(
          (box) => at.x > box.x + 0.02 && at.x < box.x + box.w - 0.02 && at.y > box.y + 0.02 && at.y < box.y + box.h - 0.02,
        );
        if (inside) through += len / steps;
      }
    }
    if (through > 0.2) {
      const name = /NameU="(Connector\.\d+)"/.exec(shape)?.[1] ?? 'connector';
      issues.push(`Visio ${name} is drawn through a service for ${through.toFixed(2)}in`);
    }
  }
  for (const block of badgeBlocks) {
    const shown = /<Text>([^<]*)<\/Text>/.exec(block)?.[1] ?? '';
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
  const labelBoxes: Array<{ text: string; x: number; y: number; w: number; h: number }> = [];
  for (const block of xml.matchAll(/<Shape [^>]*NameU="Connector\.\d+"[\s\S]*?<\/Shape>/g)) {
    const shape = block[0];
    const shown = /<Text>([^<]*)<\/Text>/.exec(shape)?.[1] ?? '';
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
    labelBoxes.push({ text: shown, x: cx - +txt[3] / 2, y: cy - +txt[4] / 2, w: +txt[3], h: +txt[4] });
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
  if (serviceCount >= 4 && density < 0.005) {
    issues.push(`page is ${(density * 100).toFixed(2)}% full — a stray node blew the sheet up`);
  }

  return {
    scenario: scenario.id,
    format: 'vsdx',
    issues,
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

async function main(): Promise<void> {
  mkdirSync(OUT, { recursive: true });
  const base = [
    compactScenario(), wideScenario(), oversizeScenario(), outlierScenario(),
    bandedScenario(), narrativeScenario(), barbellScenario(), hubFanScenario(), sharedServiceScenario(), tightGridScenario(), parallelScenario(),
    ladderInGridScenario(), twinLaddersScenario(), strayLadderScenario(), legendCornerScenario(), duplicateStepsScenario(), denseZoneScenario(),
    metaChipScenario(), gridFanScenario(), gridFan3Scenario(), fan8Tight5x5Scenario(), metaSublineScenario(), grid5x5CaptionScenario(), longNameGridScenario(), longLabelGridScenario(), metaTightScenario(),     longNameFanScenario(), estateChainScenario(), chain24Scenario(), tripleMutedScenario(), estate72Scenario(), workflowProseScenario(), workflowLongProseScenario(), allCategoriesScenario(),
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
  if (only.length > 0 && selected.length === 0) {
    throw new Error(`no scenario matched ${only.join(', ')}; known: ${scenarios.map((s) => s.id).join(', ')}`);
  }
  for (const scenario of selected) {
    reports.push(await auditPptx(scenario));
    reports.push(await auditVsdx(scenario));
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

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
