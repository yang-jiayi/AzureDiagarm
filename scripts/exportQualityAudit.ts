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
    shapes.push({
      name,
      x: +off[1] / EMU_PER_INCH,
      y: +off[2] / EMU_PER_INCH,
      w: +ext[1] / EMU_PER_INCH,
      h: +ext[2] / EMU_PER_INCH,
      text: texts.join('').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'),
      fontSize: sz ? +sz[1] / 100 : null,
    });
  }
  return shapes;
}

function overlapArea(a: Shape, b: Shape): number {
  const w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
  const h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
  return w > 0 && h > 0 ? w * h : 0;
}

interface Scenario { id: string; nodes: Node[]; edges: Edge[] }

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
      ...(i < 6 ? { data: { stepNumber: i + 1 } } : {}),
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

/** A dense cluster plus one far-placed node: nothing may fall off the page. */
function outlierScenario(): Scenario {
  const nodes: Node[] = [];
  for (let i = 0; i < 8; i += 1) {
    nodes.push(svc(`c-${i}`, i % 2 ? 'Azure Functions' : 'Azure SQL Database', (i % 4) * 220, Math.floor(i / 4) * 180));
  }
  nodes.push(svc('outlier', 'Copilot Studio', 9000, 4000));
  return { id: 'outlier', nodes, edges: [] };
}

interface Report { scenario: string; format: string; issues: string[]; metrics: Record<string, number> }

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
  const xml = await Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')))
      .map((name) => zip.file(name)!.async('string')),
  );
  const slideCount = xml.length;
  const shapes = xml.flatMap((slideXml) => parseShapes(slideXml));

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
    for (const tile of tiles) {
      const area = overlapArea(chip, tile);
      if (area > 0.02 * tile.w * tile.h) {
        issues.push(`edge chip "${chip.text}" overlaps node "${tile.name}" by ${((area / (tile.w * tile.h)) * 100).toFixed(0)}%`);
      }
    }
  }
  const truncated = shapes.filter((s) => s.text.includes('…'));
  if (truncated.length) issues.push(`${truncated.length} shapes carry truncated "…" text`);

  // Workflow numbering: an arrow that the AI numbered must carry its callout,
  // and the callout must not sit on top of a node or its own label chip —
  // either way the reader cannot match the arrow to the workflow prose.
  const numberedEdges = scenario.edges.filter(
    (e) => Number.isInteger((e.data as { stepNumber?: number } | undefined)?.stepNumber),
  );
  const badges = shapes.filter((s) => s.name.startsWith('connector-step-'));
  if (badges.length !== numberedEdges.length) {
    issues.push(`${badges.length} step badges drawn for ${numberedEdges.length} numbered connectors`);
  }
  const expectedNumbers = new Set(
    numberedEdges.map((e) => String((e.data as { stepNumber: number }).stepNumber)),
  );
  for (const badge of badges) {
    if (!expectedNumbers.has(badge.text)) {
      issues.push(`step badge shows "${badge.text}", which is not a workflow step number`);
    }
    for (const tile of tiles) {
      if (overlapArea(badge, tile) > 0.02 * tile.w * tile.h) {
        issues.push(`step badge "${badge.name}" covers node "${tile.name}"`);
      }
    }
    for (const chip of chips) {
      if (overlapArea(badge, chip) > 0.25 * badge.w * badge.h) {
        issues.push(`step badge "${badge.name}" collides with edge chip "${chip.text}"`);
      }
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
  // a different story in PowerPoint and in Visio.
  const numberedEdges = scenario.edges.filter(
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
  const scenarios = [compactScenario(), wideScenario(), oversizeScenario(), outlierScenario()];
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
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
