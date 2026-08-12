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
import { buildVsdxPackage } from '../src/services/visioVsdxExporter.ts';
import { WRAP_TRIGGER_RATIO } from '../src/utils/serpentineWrap.ts';
import { narrateEdgeCallouts } from '../src/services/diagramExportGeometry.ts';

const OUT = path.join(process.cwd(), 'tmp-export-audit');
const EMU_PER_INCH = 914400;

const PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

interface Shape {
  name: string;
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  fontSize: number | null;
  path?: { x: number; y: number }[];
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
    shapes.push({
      name,
      x,
      y,
      w,
      h,
      text: texts.join('').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
      fontSize: sz ? +sz[1] / 100 : null,
      path,
    });
  }
  return shapes;
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
}

function svc(id: string, label: string, x: number, y: number, parent?: string, icon = true): Node {
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
 * Two dense clusters joined by one long bridge, so the middle of the grid
 * holds nothing. A part that owns only its own fitted cell leaves the bridge's
 * label and callout belonging to no slide at all: the arrow is drawn, the
 * number is missing, and the workflow list still cites it.
 */
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

async function auditPptx(scenario: Scenario): Promise<Report> {
  const pptx = await buildDiagramSlidePptx(PIXEL_PNG, {
    diagramName: 'Contoso Platform',
    author: 'Audit',
    date: '2026-08-10',
    isDarkMode: false,
    diagram: { nodes: scenario.nodes, edges: scenario.edges },
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
    for (const badge of slideShapes.filter((s) => s.name.startsWith('connector-step-'))) {
      for (const tile of slideTiles) {
        if (overlapArea(badge, tile) > tileBudget(badge.name, tile) * tile.w * tile.h) {
          issues.push(`step badge "${badge.name}" covers node "${tile.name}" by ${((overlapArea(badge, tile)/(tile.w*tile.h))*100).toFixed(0)}% (badge area ${((overlapArea(badge, tile)/(badge.w*badge.h))*100).toFixed(0)}%)`);
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
        issues.push(`edge chip "${chip.text}" is ${pathGap(nearest, at).toFixed(2)}in from ${nearest.name} but ${mine.toFixed(2)}in from its own arrow`);
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
  for (const edge of scenario.edges) {
    const label = typeof edge.label === 'string' ? edge.label.trim() : '';
    if (!label || drawnChips.has(edge.id)) continue;
    const badge = drawnBadges.get(edge.id);
    if (badge !== undefined && explained.has(badge)) continue;
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
    const lost = [...authored].filter((d) => ![...rowText].some((r) => r.length > 0 && d.startsWith(r)));
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
      chips: chips.length,
      maxChipWidthIn: chips.length ? +Math.max(...chips.map((c) => c.w)).toFixed(3) : 0,
      stepBadges: badges.length,
      fillPct: +(density * 100).toFixed(2),
    },
  };
}

async function auditVsdx(scenario: Scenario): Promise<Report> {
  const pkg = await buildVsdxPackage(scenario.nodes, scenario.edges, 'Contoso Platform');
  const issues: string[] = [];
  const pagePart = pkg.parts.find((p) => /page1\.xml$/i.test(p.path));
  const media = pkg.parts.filter((p) => /\/media\//i.test(p.path));
  const serviceCount = scenario.nodes.filter((n) => n.type !== 'groupNode').length;
  // Icon rasterisation needs a DOM; under Node it always yields zero media
  // parts, so icon coverage is asserted by the Playwright probe instead.
  const canRasterize = typeof document !== 'undefined';
  if (canRasterize && media.length === 0) issues.push(`no embedded icon media parts (expected ~${serviceCount})`);
  const xml = typeof pagePart?.data === 'string' ? pagePart.data : '';
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
  if (minFontPt < 5.5) issues.push(`smallest Visio font is ${minFontPt}pt (below the 5.5pt floor)`);

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
  const scenarios = [
    compactScenario(), wideScenario(), oversizeScenario(), outlierScenario(),
    bandedScenario(), narrativeScenario(), barbellScenario(), parallelScenario(),
    ladderInGridScenario(), twinLaddersScenario(), strayLadderScenario(), legendCornerScenario(), duplicateStepsScenario(), denseZoneScenario(),
    gridFanScenario(), estateChainScenario(),
    await generatedScenario(), await groupedGeneratedScenario(),
  ];
  const reports: Report[] = [];
  for (const scenario of scenarios) {
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
