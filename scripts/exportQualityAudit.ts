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
import { buildDiagramSlidePptx, buildArchitectureDeckPptx } from '../src/services/pptxExporter.ts';
import { nativizeSlideXml } from '../src/services/pptxNativeShapes.ts';
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';
import { WRAP_TRIGGER_RATIO } from '../src/utils/serpentineWrap.ts';

import { narrateEdgeCallouts, CATEGORY_STYLES } from '../src/services/diagramExportGeometry.ts';

const OUT = path.join(process.cwd(), 'tmp-export-audit');
const EMU_PER_INCH = 914400;
/** Layout pixels per inch — the scale both exporters draw the canvas at. */
const PX_PER_IN = 96;
/** The standard 16:9 slide both decks start from, before any page growth. */
const BASE_SLIDE_W_IN = 13.333;
const BASE_SLIDE_H_IN = 7.5;
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

function xmlWellFormednessIssues(parts: Array<{ path: string; text: string }>, prefix: string): string[] {  const issues: string[] = [];
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
  const names = ['Azure App Service', 'Azure SQL Database', 'Azure Key Vault', 'Azure Functions', 'Azure Cache for Redis'];
  const nodes: Node[] = [];
  const edges: Edge[] = [];
  for (let s = 0; s < 480; s += 1) {
    const originX = s * 260;
    nodes.push(grp(`sub-${s}`, `Perimeter network subnet ${s}`, originX, 0, 200, 130));
    nodes.push(svc(`z-${s}`, names[s % names.length], originX + 25, 40));
    if (s > 0) {
      edges.push({ id: `ze-${s}`, source: `z-${s - 1}`, target: `z-${s}`, label: 'Peers', data: { stepNumber: s } } as Edge);
    }
  }
  return { id: 'scaled-zone-row', nodes, edges };
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
  const pptx = await buildArchitectureDeckPptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Audit',
    date: '2026-08-10',
    isDarkMode: scenario.dark === true,
    diagram: { nodes: scenario.nodes, edges: scenario.edges },
    presetIcons: synthesisedIcons(scenario),
    services: deckServices,
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
    for (const match of xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)) deckText.add(match[1].trim());
  }
  const stranded: string[] = [];
  for (const node of scenario.nodes) {
    if ((node.type ?? '') === 'groupNode') continue;
    // Trimmed on both sides of the comparison. A label carrying a forbidden
    // code point comes back from `auditStrip` with the gap it left, and the
    // deck writes the same name without it; that is the sanitiser working, not
    // a name gone missing.
    const name = auditStrip(String((node.data as { label?: string } | undefined)?.label ?? '')).trim();
    if (name && !deckText.has(name)) stranded.push(name);
  }
  if (stranded.length > 0) {
    issues.push(
      `customer deck: ${stranded.length} service name(s) appear nowhere in full — e.g. `
      + `${stranded.slice(0, 3).map((n) => `"${n}"`).join(', ')}`,
    );
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
  );
}
async function zipXmlParts(zip: JSZip): Promise<Array<{ path: string; text: string }>> {
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir && /\.(xml|rels)$/i.test(n));
  return Promise.all(names.map(async (path) => ({ path, text: await zip.files[path].async('string') })));
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
  issues.push(...xmlWellFormednessIssues(await zipXmlParts(zip), ''));
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
    // Both sides sanitised: the deck's chips and badges are keyed by the name
    // the exporter wrote, and its wording is the sanitised wording.
    const eid = auditStrip(String(edge.id));
    const label = typeof edge.label === 'string' ? auditStrip(edge.label).trim() : '';
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
    if (typeof label === 'string' && label) authoredNames.set(auditStrip(String(node.id)), label);
  }
  // Truncation is only acceptable when the full wording survives somewhere the
  // reader can reach. A chip clipped to 42 cells with no workflow row carrying
  // the rest has silently thrown away what the author wrote.
  const indexed = new Set(
    shapes.filter((s) => s.name.startsWith('index-name-')).map((s) => s.text.trim()),
  );
  const truncated = shapes.filter((s) => s.text.includes('…'));
  const stranded = truncated.filter((s) => {
    const svcId = /^service-label-(.*)$/.exec(s.name)?.[1];
    if (svcId !== undefined) {
      // The deck's index slide spells this one out, so the drawing is free to
      // abbreviate it.
      const name = authoredNames.get(svcId);
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
    numberedEdges.map((e) => [auditStrip(String(e.id)), String((e.data as { stepNumber: number }).stepNumber)]),
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
  for (const slideShapes of perSlide) {
    const slideTiles = slideShapes.filter(
      (s) => s.name.startsWith('service-')
        && !s.name.startsWith('service-label-')
        && !s.name.startsWith('service-meta-')
        && s.w > 0,
    );
    if (slideTiles.length !== 1) continue;
    const authoredIn = (authoredWidths.get(slideTiles[0].name.replace(/^service-/, '')) ?? 150) / 96;
    if (slideTiles[0].w >= authoredIn * 0.99) lonelyWindows += 1;
  }
  // One lone tile is ordinary — a drawing whose last window holds the remainder
  // will have exactly that. A deck built out of them is the defect.
  if (lonelyWindows > 2 && lonelyWindows > slideCount * 0.5) {
    issues.push(
      `${lonelyWindows} of ${slideCount} slides carry a single service tile already at natural `
      + `width — the deck split past the point where splitting can enlarge anything`,
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
    return n.type !== 'groupNode'
      && Boolean((n.data as { iconPath?: string } | undefined)?.iconPath)
      && (n.height ?? styled ?? 75) >= ICON_MIN_PX;
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
  const allowedW = Math.max(11, span.w + PAGE_CHROME_SLACK_IN);
  const allowedH = Math.max(8.5, span.h + PAGE_CHROME_SLACK_IN + bandIn);
  if (pkg.pageWidthIn > allowedW) {
    issues.push(`Visio sheet is ${pkg.pageWidthIn.toFixed(1)}in wide for a drawing that spans ${span.w.toFixed(1)}in — trimming outliers must shrink the sheet, never grow it`);
  }
  if (pkg.pageHeightIn > allowedH) {
    issues.push(`Visio sheet is ${pkg.pageHeightIn.toFixed(1)}in tall for a drawing that spans ${span.h.toFixed(1)}in — trimming outliers must shrink the sheet, never grow it`);
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
  for (const group of xml.matchAll(/NameU="Service\.\d+" Name="([^"]*)"[\s\S]*?<Cell N="Width" V="([\d.]+)"[\s\S]*?<Cell N="Size" V="([\d.]+)"/g)) {
    const label = group[1];
    const tileIn = +group[2];
    const fontIn = +group[3];
    if (!label || tileIn <= 0 || fontIn <= 0) continue;
    const ratio = fontIn / tileIn;
    const natural = NATURAL_LABEL_IN / NATURAL_TILE_IN;
    if (ratio > natural * 1.05) {
      issues.push(`service name "${label}" is set at ${(fontIn * 72).toFixed(2)}pt on a ${tileIn.toFixed(2)}in tile `
        + `— ${(ratio / natural).toFixed(1)}x the type-to-tile ratio the sheet draws at full size, `
        + `so the name wraps past the room the tile has for it`);
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
  const badges = [...xml.matchAll(/NameU="StepBadge\.\d+"[\s\S]*?<Cell N="Width" V="([\d.]+)"/g)].map((m) => +m[1]);
  if (badges.length > 0 && tileWidths.length > 0) {
    const widest = Math.max(...badges);
    const tile = Math.min(...tileWidths);
    if (widest > tile * 0.55) {
      issues.push(`a step badge is ${widest.toFixed(3)}in across on a ${tile.toFixed(3)}in tile `
        + `— ${((widest / tile) * 100).toFixed(0)}% of the service it is calling out`);
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
    const label = typeof edge.label === 'string' ? auditStrip(edge.label).trim() : '';
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
    text: m[7].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&'),
  }));
  const spilling = workflowRows
    .map((row) => {
      const lines = Math.max(1, Math.ceil(textWidthIn(row.text.trim(), row.pt) / row.w));
      return { row, needed: (lines * row.pt * 1.22) / 72, lines };
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
  const serviceBoxes: Array<{ x: number; y: number; w: number; h: number; name?: string }> = [];
  for (const m of xml.matchAll(
    /NameU="(Service\.\d+)"[\s\S]*?<Cell N="PinX" V="([\d.-]+)"\/>\s*<Cell N="PinY" V="([\d.-]+)"\/>\s*<Cell N="Width" V="([\d.-]+)"\/>\s*<Cell N="Height" V="([\d.-]+)"\/>/g,
  )) {
    const [, name, pinX, pinY, w, h] = m;
    serviceBoxes.push({ x: +pinX - +w / 2, y: +pinY - +h / 2, w: +w, h: +h, name });
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
  const labelBoxes: Array<{ text: string; edge: string; x: number; y: number; w: number; h: number }> = [];
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
    labelBoxes.push({
      text: shown,
      edge: /Name="edge-([^"]*)"/.exec(shape)?.[1] ?? '',
      x: cx - +txt[3] / 2,
      y: cy - +txt[4] / 2,
      w: +txt[3],
      h: +txt[4],
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
      badgeItems.push({ id: `callout ${m[4].trim()}`, edge: m[1], at: { x: +m[2], y: +m[3] } });
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
    metrics: Object.fromEntries(seen.map((s) => [`n${s.n}`, `${s.slides}sl/${s.font}pt`])),
  };
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
    corridorZoneScenario(),
    ladderInGridScenario(), twinLaddersScenario(), strayLadderScenario(), legendCornerScenario(), duplicateStepsScenario(), denseZoneScenario(),
    metaChipScenario(), gridFanScenario(), gridFan3Scenario(), fan8Tight5x5Scenario(), metaSublineScenario(), grid5x5CaptionScenario(), longNameGridScenario(), longLabelGridScenario(), metaTightScenario(),     longNameFanScenario(), estateChainScenario(), chain24Scenario(), tripleMutedScenario(), estate72Scenario(),     workflowProseScenario(), workflowLongProseScenario(), workflowFanScenario(), workflowWideBandScenario(), allCategoriesScenario(), controlCharScenario(), shortServiceGridScenario(),
    cascadeScenario(),
    sharedPrefixEstateScenario(),
    shortTileEstateScenario(),
    compactEstateScenario(),
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
    reports.push(await auditPptx(scenario));
    reports.push(await auditVsdx(scenario));
  }
  if (growth) reports.push(await auditDeckGrowth());
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
