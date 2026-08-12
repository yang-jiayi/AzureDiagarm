// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Visio VSDX Exporter
 * -------------------
 * Generates a native Visio drawing (.vsdx) — the modern Open Packaging
 * Conventions (OPC) format that opens directly in BOTH desktop Visio and
 * Visio for the web, and can also be imported by diagrams.net.
 *
 * A .vsdx is a ZIP package of XML parts (2012 Visio schema). We assemble the
 * minimal set Visio requires:
 *   [Content_Types].xml
 *   _rels/.rels
 *   docProps/core.xml, docProps/app.xml
 *   visio/document.xml            (DocumentSettings, FaceNames, StyleSheets)
 *   visio/_rels/document.xml.rels (-> pages, windows)
 *   visio/windows.xml
 *   visio/pages/pages.xml         (Page + PageSheet with layers, -> page1)
 *   visio/pages/_rels/pages.xml.rels
 *   visio/pages/page1.xml         (Shapes + Connects)
 *
 * Shape model
 * -----------
 * - Zones  → rounded, dashed 2-D rectangles with a top-left title.
 * - Services → a Visio **group** containing the tile rectangle and the embedded
 *   Azure icon, so dragging the service in Visio moves the icon with it.
 * - Edges  → real **1-D connectors** (ObjType 2, Begin/End cells) that are
 *   glued to the shapes they join through the `<Connects>` table, so rerouting
 *   or moving a service keeps the diagram wired up. The edge label is the
 *   connector's own text, which means it follows the line.
 *
 * Coordinates convert React Flow px / top-left (Y down) → Visio inches /
 * bottom-left (Y up) at 96 px per inch.
 */

import JSZip from 'jszip';
import type { Node, Edge } from 'reactflow';
import { rasterizeIcons, type RasterizedIcon } from '../utils/exportIconRaster';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  compactEmptyGutters,
  fitBoxesWithin,
  clampedBoxes,
  computeBounds,
  computeContentBounds,
  metaSubline,
  narrateEdgeCallouts,
  truncateLabel,
  partitionBoxes,
  readableTextOn,
  usedConnectionLegend,
  workflowListFromEdges,
  carriesWording,
  zoneStyleFor,
  type ConnectionLegendEntry,
  type WorkflowListEntry,
  type ExportBox,
  type ExportRoute,
  type Point,
} from './diagramExportGeometry';

const PX_PER_INCH = 96;
const PAGE_PADDING_IN = 0.6;
const MIN_PAGE_W_IN = 11;
const MIN_PAGE_H_IN = 8.5;
/**
 * Visio itself allows 200", but a page much beyond a plotter sheet is only ever
 * produced by a stray far-placed node; past this the drawing is re-fitted to the
 * dense cluster and the strays are clamped onto the page so nothing is lost.
 */
const MAX_USEFUL_PAGE_IN = 60;
/**
 * And the limit the format itself imposes. Visio will not open a page beyond
 * this, so it is not a preference — a drawing that needs more has to be made
 * to fit before it is written, or the file is unopenable.
 */
const MAX_VISIO_PAGE_IN = 200;
const CORNER_ROUNDING_IN = 0.08;

/**
 * Visio font sizes are inches (1 pt = 1/72"). These match the PowerPoint export
 * at 1 : 1 — a 150 px tile is 1.56" wide, so the label reads at ~7.6 pt and the
 * SKU sub-line at ~7 pt instead of the previous near-illegible 6.5/5 pt.
 *
 * Nothing here may fall below the legibility floor the deck enforces. The two
 * exporters draw the same drawing at the same scale, so a size that is
 * unreadable in one is unreadable in the other, and the sub-line used to be set
 * at 5.98 pt — a whole point under the deck's own floor for exactly the same
 * words on exactly the same tile.
 */
const LEGIBLE_PT = 7;
const LEGIBLE_IN = LEGIBLE_PT / 72;
const LABEL_FONT_IN = 0.105;
const META_FONT_IN = Math.max(0.083, LEGIBLE_IN);
const CONNECTOR_FONT_IN = Math.max(0.1, LEGIBLE_IN);
const LEGEND_FONT_IN = Math.max(0.1, LEGIBLE_IN);

/**
 * Approximate rendered width in inches. CJK glyphs occupy a full em, Latin
 * about 0.54 em — enough to decide how many lines a label needs.
 */
function estimateTextWidthIn(text: string, fontSizeIn: number): number {
  let units = 0;
  for (const character of text) {
    units += /[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]/.test(character) ? 1 : 0.54;
  }
  return units * fontSizeIn;
}

const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Layer indexes declared on the page sheet. */
const LAYER_ZONES = 0;
const LAYER_SERVICES = 1;
const LAYER_CONNECTIONS = 2;

function esc(value: string): string {
  return (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const f = (n: number) => +n.toFixed(4);

// ─── Palette (single source of truth in the shared geometry layer) ───────────

interface Palette { fill: string; line: string; text: string }

const CONNECTOR_TEXT = '#374151';

/** Adapt the shared category style ({bg,border,text}) to Visio's cell names. */
function paletteForService(box: ExportBox): Palette {
  const style = categoryStyle(box.category);
  return { fill: style.bg, line: style.border, text: style.text };
}

/** Adapt the shared zone style (honours `data.customColor`) to Visio's cells. */
function paletteForZone(box: ExportBox, index: number): Palette {
  const style = zoneStyleFor(box, index);
  return { fill: style.bg, line: style.border, text: style.text };
}

/**
 * Visio built-in line pattern for a canonical connection type. Solid=1,
 * dashed=2, dotted=3, dash-dot=4 — mirroring the PNG legend's dash coding.
 */
function visioLinePattern(route: ExportRoute): number {
  if (!route.dashed) return 1;
  switch (route.connectionType) {
    case 'optional':
    case 'security':
      return 3;
    case 'telemetry':
      return 4;
    default:
      return 2;
  }
}

// ─── OPC static parts ────────────────────────────────────────────────────────

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>
  <Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>
  <Override PartName="/visio/pages/page1.xml" ContentType="application/vnd.ms-visio.page+xml"/>
  <Override PartName="/visio/windows.xml" ContentType="application/vnd.ms-visio.windows+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/document" Target="visio/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages" Target="pages/pages.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/windows" Target="windows.xml"/>
</Relationships>`;

const PAGES_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page1.xml"/>
</Relationships>`;

function windowsXml(pageWidthIn: number, pageHeightIn: number): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Windows xmlns="${VISIO_NS}" xmlns:r="${REL_NS}" ClientWidth="1200" ClientHeight="720">
  <Window ID="0" WindowType="Drawing" WindowState="1073741824" WindowLeft="0" WindowTop="0" WindowWidth="1200" WindowHeight="720" ContainerType="Page" Page="0" ViewScale="-1" ViewCenterX="${f(pageWidthIn / 2)}" ViewCenterY="${f(pageHeightIn / 2)}"/>
</Windows>`;
}

/**
 * document.xml — DocumentSettings, the Segoe UI face name used by every shape,
 * and the "No Style" stylesheet (ID 0) that pages and shapes reference.
 */
const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<VisioDocument xmlns="${VISIO_NS}" xmlns:r="${REL_NS}">
  <DocumentSettings TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0" DefaultGuideStyle="0">
    <GlueSettings>9</GlueSettings>
    <SnapSettings>65847</SnapSettings>
    <SnapExtensions>34</SnapExtensions>
    <DynamicGridEnabled>1</DynamicGridEnabled>
    <ProtectStyles>0</ProtectStyles>
    <ProtectShapes>0</ProtectShapes>
    <ProtectMasters>0</ProtectMasters>
    <ProtectBkgnds>0</ProtectBkgnds>
  </DocumentSettings>
  <FaceNames>
    <FaceName ID="1" NameU="Segoe UI" UnicodeRanges="-1 -369098753 63 0" CharSets="536871423 0" Panos="2 11 5 2 4 2 4 2 2 3" Flags="325"/>
  </FaceNames>
  <StyleSheets>
    <StyleSheet ID="0" NameU="No Style" Name="No Style">
      <Cell N="EnableLineProps" V="1"/>
      <Cell N="EnableFillProps" V="1"/>
      <Cell N="EnableTextProps" V="1"/>
      <Cell N="LineWeight" V="0.01041666666666667"/>
      <Cell N="LineColor" V="0"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="FillForegnd" V="1"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="TextBkgnd" V="0"/>
      <Cell N="Font" V="1"/>
      <Cell N="Size" V="0.1111111111111111"/>
    </StyleSheet>
  </StyleSheets>
</VisioDocument>`;

function coreXml(title: string): string {
  const now = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${esc(title)}</dc:title>
  <dc:creator>Microsoft Product Architecture Diagram Builder</dc:creator>
  <dc:description>Generated by Microsoft Product Architecture Diagram Builder — customization by Swarm Data SE, Jiayi Yang.</dc:description>
  <cp:lastModifiedBy>Microsoft Product Architecture Diagram Builder</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Visio</Application>
  <AppVersion>16.0000</AppVersion>
</Properties>`;

// ─── Shape builders ──────────────────────────────────────────────────────────

interface Rect { x: number; y: number; w: number; h: number }

function layerRow(index: number, name: string): string {
  return `      <Row IX="${index}"><Cell N="Name" V="${esc(name)}"/><Cell N="NameUniv" V="${esc(name)}"/><Cell N="Color" V="255"/><Cell N="Status" V="0"/><Cell N="Visible" V="1"/><Cell N="Print" V="1"/><Cell N="Active" V="0"/><Cell N="Lock" V="0"/><Cell N="Snap" V="1"/><Cell N="Glue" V="1"/><Cell N="ColorTrans" V="0"/></Row>`;
}

/** Rounded-rectangle geometry rows, relative to the shape's own width/height. */
function roundedRectGeometry(): string {
  return `      <Section N="Geometry" IX="0">
        <Cell N="NoFill" V="0"/>
        <Cell N="NoLine" V="0"/>
        <Cell N="NoShow" V="0"/>
        <Cell N="NoSnap" V="0"/>
        <Row T="RelMoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row>
        <Row T="RelLineTo" IX="2"><Cell N="X" V="1"/><Cell N="Y" V="0"/></Row>
        <Row T="RelLineTo" IX="3"><Cell N="X" V="1"/><Cell N="Y" V="1"/></Row>
        <Row T="RelLineTo" IX="4"><Cell N="X" V="0"/><Cell N="Y" V="1"/></Row>
        <Row T="RelLineTo" IX="5"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row>
      </Section>`;
}

function propertyRow(name: string, label: string, value: string, sortKey: number): string {
  return `        <Row N="${name}"><Cell N="Label" V="${esc(label)}"/><Cell N="Value" V="${esc(value)}"/><Cell N="Type" V="0"/><Cell N="SortKey" V="${sortKey}"/></Row>`;
}

/** Zone / group rectangle: dashed rounded outline with a top-left title. */
function zoneShapeXml(id: number, rect: Rect, label: string, palette: Palette): string {
  // Allow the title band to grow to two lines instead of clipping a long name.
  const titleH = Math.min(0.56, Math.max(0.24, rect.h * 0.22));
  const titleW = Math.max(0.4, rect.w - 0.24);
  return `    <Shape ID="${id}" NameU="Zone.${id}" Name="${esc(label)}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(rect.x + rect.w / 2)}"/>
      <Cell N="PinY" V="${f(rect.y + rect.h / 2)}"/>
      <Cell N="Width" V="${f(rect.w)}"/>
      <Cell N="Height" V="${f(rect.h)}"/>
      <Cell N="LocPinX" V="${f(rect.w / 2)}"/>
      <Cell N="LocPinY" V="${f(rect.h / 2)}"/>
      <Cell N="Angle" V="0"/>
      <Cell N="LayerMember" V="${LAYER_ZONES}"/>
      <Cell N="FillForegnd" V="${palette.fill}"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="${palette.line}"/>
      <Cell N="LineWeight" V="0.0125"/>
      <Cell N="LinePattern" V="2"/>
      <Cell N="Rounding" V="${f(CORNER_ROUNDING_IN * 1.5)}"/>
      <Cell N="TxtWidth" V="${f(titleW)}"/>
      <Cell N="TxtHeight" V="${f(titleH)}"/>
      <Cell N="TxtPinX" V="${f(rect.w / 2)}"/>
      <Cell N="TxtPinY" V="${f(rect.h - titleH / 2 - 0.06)}"/>
      <Cell N="TxtLocPinX" V="${f(titleW / 2)}"/>
      <Cell N="TxtLocPinY" V="${f(titleH / 2)}"/>
      <Cell N="TxtAngle" V="0"/>
${roundedRectGeometry()}
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${palette.text}"/><Cell N="Size" V="0.13"/><Cell N="Style" V="1"/></Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="0"/></Row>
      </Section>
      <Section N="Property">
${propertyRow('ZoneName', 'Zone', label, 1)}
      </Section>
      <Text>${esc(label)}</Text>
    </Shape>`;
}

/**
 * Service tile: a Visio group whose members are the rounded rectangle and the
 * Azure icon, so the two never drift apart when the shape is moved.
 */
function serviceGroupXml(
  ids: { group: number; rect: number; icon: number },
  rect: Rect,
  box: ExportBox,
  palette: Palette,
  iconRelId: string | null,
  properties: Array<{ name: string; label: string; value: string }>,
  meta: string,
): string {
  const textW = Math.max(0.3, rect.w - 0.12);
  // Give the label the room it actually needs and let the icon take the rest,
  // so a two-line service name is never clipped and the icon never vanishes.
  const labelLines = Math.max(1, Math.ceil(estimateTextWidthIn(box.label, LABEL_FONT_IN) / textW));
  const neededTextH = labelLines * LABEL_FONT_IN * 1.28 + (meta ? META_FONT_IN * 1.4 : 0) + 0.05;
  const maxIcon = iconRelId ? Math.min(rect.h * 0.46, rect.w * 0.5, 0.55) : 0;
  const minIcon = Math.min(maxIcon, 0.18);
  const room = Math.max(0.2, rect.h - 0.19);
  const textH = Math.max(0.16, Math.min(neededTextH, room - minIcon));
  const iconSizeIn = maxIcon > 0 ? Math.max(0, Math.min(maxIcon, room - textH)) : 0;
  const showIcon = iconRelId !== null && iconSizeIn >= 0.08;
  const iconChild = showIcon
    ? `
        <Shape ID="${ids.icon}" NameU="Icon.${ids.icon}" Type="Foreign" LineStyle="0" FillStyle="0" TextStyle="0">
          <Cell N="PinX" V="${f(rect.w / 2)}"/>
          <Cell N="PinY" V="${f(rect.h - iconSizeIn / 2 - 0.07)}"/>
          <Cell N="Width" V="${f(iconSizeIn)}"/>
          <Cell N="Height" V="${f(iconSizeIn)}"/>
          <Cell N="LocPinX" V="${f(iconSizeIn / 2)}"/>
          <Cell N="LocPinY" V="${f(iconSizeIn / 2)}"/>
          <Cell N="Angle" V="0"/>
          <Cell N="LayerMember" V="${LAYER_SERVICES}"/>
          <Cell N="ImgOffsetX" V="0"/>
          <Cell N="ImgOffsetY" V="0"/>
          <Cell N="ImgWidth" V="${f(iconSizeIn)}"/>
          <Cell N="ImgHeight" V="${f(iconSizeIn)}"/>
          <ForeignData ForeignType="Bitmap" CompressionType="PNG">
            <Rel r:id="${iconRelId}"/>
          </ForeignData>
        </Shape>`
    : '';

  // Second, smaller run carries the SKU · region · cost annotation so the tile
  // shows the same metadata as the canvas instead of dropping it.
  // The sub-line colour is resolved against the tile it sits on, not fixed. A
  // flat #64748B reads at 4.26:1 or worse on every lighter category fill, which
  // is under the WCAG AA bar — the PowerPoint path was corrected the same way.
  const metaColor = meta ? readableTextOn('#64748B', palette.fill) : '#64748B';
  // The tile name is drawn in the category's own accent, which on two of the
  // sixteen palettes is unreadable on that category's fill — ai + machine
  // learning at 3.93:1 and identity at 2.49:1. PowerPoint draws the same words
  // at a flat #1F2937, so the two exporters were disagreeing about what is
  // legible. Resolve the accent against the tile the same way the sub-line is.
  const nameColor = readableTextOn(palette.text, palette.fill);
  const characterRows = meta
    ? `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${nameColor}"/><Cell N="Size" V="${LABEL_FONT_IN}"/></Row>
        <Row IX="1"><Cell N="Font" V="1"/><Cell N="Color" V="${metaColor}"/><Cell N="Size" V="${META_FONT_IN}"/></Row>`
    : `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${nameColor}"/><Cell N="Size" V="${LABEL_FONT_IN}"/></Row>`;
  const textBody = meta
    ? `<cp IX="0"/>${esc(box.label)}\n<cp IX="1"/>${esc(meta)}`
    : esc(box.label);

  return `    <Shape ID="${ids.group}" NameU="Service.${ids.group}" Name="${esc(box.label)}" Type="Group" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(rect.x + rect.w / 2)}"/>
      <Cell N="PinY" V="${f(rect.y + rect.h / 2)}"/>
      <Cell N="Width" V="${f(rect.w)}"/>
      <Cell N="Height" V="${f(rect.h)}"/>
      <Cell N="LocPinX" V="${f(rect.w / 2)}"/>
      <Cell N="LocPinY" V="${f(rect.h / 2)}"/>
      <Cell N="Angle" V="0"/>
      <Cell N="LayerMember" V="${LAYER_SERVICES}"/>
      <Cell N="DontMoveChildren" V="0"/>
      <Cell N="IsDropTarget" V="0"/>
      <Cell N="IsTextEditTarget" V="1"/>
      <Cell N="TxtWidth" V="${f(textW)}"/>
      <Cell N="TxtHeight" V="${f(textH)}"/>
      <Cell N="TxtPinX" V="${f(rect.w / 2)}"/>
      <Cell N="TxtPinY" V="${f(0.06 + textH / 2)}"/>
      <Cell N="TxtLocPinX" V="${f(textW / 2)}"/>
      <Cell N="TxtLocPinY" V="${f(textH / 2)}"/>
      <Cell N="TxtAngle" V="0"/>
      <Section N="Character">
${characterRows}
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="1"/></Row>
      </Section>
      <Section N="Property">
${properties.map((property, index) => propertyRow(property.name, property.label, property.value, index + 1)).join('\n')}
      </Section>
      <Text>${textBody}</Text>
      <Shapes>
        <Shape ID="${ids.rect}" NameU="Tile.${ids.rect}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
          <Cell N="PinX" V="${f(rect.w / 2)}"/>
          <Cell N="PinY" V="${f(rect.h / 2)}"/>
          <Cell N="Width" V="${f(rect.w)}"/>
          <Cell N="Height" V="${f(rect.h)}"/>
          <Cell N="LocPinX" V="${f(rect.w / 2)}"/>
          <Cell N="LocPinY" V="${f(rect.h / 2)}"/>
          <Cell N="Angle" V="0"/>
          <Cell N="LayerMember" V="${LAYER_SERVICES}"/>
          <Cell N="FillForegnd" V="${palette.fill}"/>
          <Cell N="FillPattern" V="1"/>
          <Cell N="LineColor" V="${palette.line}"/>
          <Cell N="LineWeight" V="0.0125"/>
          <Cell N="LinePattern" V="1"/>
          <Cell N="Rounding" V="${CORNER_ROUNDING_IN}"/>
          <Cell N="ShdwPattern" V="1"/>
          <Cell N="ShdwForegnd" V="#D5DDE8"/>
${roundedRectGeometry()}
        </Shape>${iconChild}
      </Shapes>
    </Shape>`;
}

/**
 * 1-D connector glued to its endpoints. Local geometry is expressed along the
 * begin→end axis (Visio's convention for connectors) so multi-segment routes
 * survive rotation, and the label rides on the line as the shape's own text.
 */
function connectorShapeXml(
  id: number,
  points: Point[],
  label: string,
  color: string,
  linePattern: number,
  opacity: number,
  bidirectional = false,
  /** Where the label sits relative to the line, and how big it is. */
  text?: { drop: number; along: number; w: number; h: number },
): string {
  const begin = points[0];
  const end = points[points.length - 1];
  const dx = end.x - begin.x;
  const dy = end.y - begin.y;
  const length = Math.max(Math.hypot(dx, dy), 0.0001);
  const angle = Math.atan2(dy, dx);
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);

  const rows = points
    .map((point, index) => {
      const vx = point.x - begin.x;
      const vy = point.y - begin.y;
      const lx = vx * cos + vy * sin;
      const ly = -vx * sin + vy * cos;
      return `        <Row T="${index === 0 ? 'MoveTo' : 'LineTo'}" IX="${index + 1}"><Cell N="X" V="${f(lx)}"/><Cell N="Y" V="${f(ly)}"/></Row>`;
    })
    .join('\n');

  const textSections = label
    ? `
      <Cell N="TextBkgnd" V="2"/>${text
  ? `
      <Cell N="TxtPinX" V="${f(length / 2 + text.along)}"/>
      <Cell N="TxtPinY" V="${f(text.drop)}"/>
      <Cell N="TxtWidth" V="${f(text.w)}"/>
      <Cell N="TxtHeight" V="${f(text.h)}"/>
      <Cell N="TxtLocPinX" V="${f(text.w / 2)}"/>
      <Cell N="TxtLocPinY" V="${f(text.h / 2)}"/>
      <Cell N="TxtAngle" V="${f(-angle)}"/>`
  : ''}
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${CONNECTOR_TEXT}"/><Cell N="Size" V="${CONNECTOR_FONT_IN}"/></Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="1"/></Row>
      </Section>`
    : '';
  // Faded (optional) connectors get line transparency so they read as secondary.
  const transCell = opacity < 1 ? `\n      <Cell N="LineColorTrans" V="${f(1 - opacity)}"/>` : '';

  return `    <Shape ID="${id}" NameU="Connector.${id}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(begin.x + dx / 2)}"/>
      <Cell N="PinY" V="${f(begin.y + dy / 2)}"/>
      <Cell N="Width" V="${f(length)}"/>
      <Cell N="Height" V="0"/>
      <Cell N="LocPinX" V="${f(length / 2)}"/>
      <Cell N="LocPinY" V="0"/>
      <Cell N="Angle" V="${f(angle)}"/>
      <Cell N="BeginX" V="${f(begin.x)}"/>
      <Cell N="BeginY" V="${f(begin.y)}"/>
      <Cell N="EndX" V="${f(end.x)}"/>
      <Cell N="EndY" V="${f(end.y)}"/>
      <Cell N="ObjType" V="2"/>
      <Cell N="ShapeRouteStyle" V="16"/>
      <Cell N="ConFixedCode" V="0"/>
      <Cell N="ConLineRouteExt" V="0"/>
      <Cell N="LayerMember" V="${LAYER_CONNECTIONS}"/>
      <Cell N="LineColor" V="${color}"/>
      <Cell N="LineWeight" V="0.0125"/>
      <Cell N="LinePattern" V="${linePattern}"/>${transCell}
      <Cell N="Rounding" V="0.0625"/>
      <Cell N="BeginArrow" V="${bidirectional ? 4 : 0}"/>
      <Cell N="BeginArrowSize" V="2"/>
      <Cell N="EndArrow" V="4"/>
      <Cell N="EndArrowSize" V="2"/>${textSections}
      <Section N="Geometry" IX="0">
        <Cell N="NoFill" V="1"/>
        <Cell N="NoLine" V="0"/>
        <Cell N="NoShow" V="0"/>
        <Cell N="NoSnap" V="0"/>
${rows}
      </Section>
      <Text>${esc(label)}</Text>
    </Shape>`;
}

/**
 * Numbered callout drawn beside a connector, matching the workflow list.
 *
 * Reference architectures on the Azure Architecture Center number the arrows
 * and repeat those numbers in the prose. A real Visio ellipse (rather than
 * text baked into the connector) keeps the badge selectable and movable when
 * the reader edits the drawing.
 */
/** Badge diameter in inches, shared with the on-page clamp. */
const STEP_BADGE_IN = 0.24;

function stepBadgeXml(id: number, centre: Point, stepNumber: number): string {
  const d = STEP_BADGE_IN;
  return `    <Shape ID="${id}" NameU="StepBadge.${id}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(centre.x)}"/>
      <Cell N="PinY" V="${f(centre.y)}"/>
      <Cell N="Width" V="${f(d)}"/>
      <Cell N="Height" V="${f(d)}"/>
      <Cell N="LocPinX" V="${f(d / 2)}"/>
      <Cell N="LocPinY" V="${f(d / 2)}"/>
      <Cell N="Angle" V="0"/>
      <Cell N="LayerMember" V="${LAYER_CONNECTIONS}"/>
      <Cell N="FillForegnd" V="#1F2937"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="#FFFFFF"/>
      <Cell N="LineWeight" V="0.0125"/>
      <Cell N="LinePattern" V="1"/>
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="#FFFFFF"/><Cell N="Size" V="0.11"/><Cell N="Style" V="1"/></Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="1"/></Row>
      </Section>
      <Section N="Geometry" IX="0">
        <Cell N="NoFill" V="0"/>
        <Cell N="NoLine" V="0"/>
        <Row T="Ellipse" IX="1">
          <Cell N="X" V="${f(d / 2)}"/>
          <Cell N="Y" V="${f(d / 2)}"/>
          <Cell N="A" V="${f(d)}"/>
          <Cell N="B" V="${f(d / 2)}"/>
          <Cell N="C" V="${f(d / 2)}"/>
          <Cell N="D" V="${f(d)}"/>
        </Row>
      </Section>
      <Text>${esc(String(stepNumber))}</Text>
    </Shape>`;
}

function connectXml(connectorId: number, sourceId: number, targetId: number): string {
  return `    <Connect FromSheet="${connectorId}" FromCell="BeginX" FromPart="9" ToSheet="${sourceId}" ToCell="PinX" ToPart="3"/>
    <Connect FromSheet="${connectorId}" FromCell="EndX" FromPart="12" ToSheet="${targetId}" ToCell="PinX" ToPart="3"/>`;
}

/** A short coloured swatch line for one legend row. */
function legendSwatchXml(id: number, x: number, y: number, entry: ConnectionLegendEntry): string {
  const pattern = entry.dashed
    ? entry.type === 'telemetry' ? 4 : entry.type === 'async' ? 2 : 3
    : 1;
  return `    <Shape ID="${id}" NameU="LegendLine.${id}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(x + 0.16)}"/>
      <Cell N="PinY" V="${f(y)}"/>
      <Cell N="Width" V="0.32"/>
      <Cell N="Height" V="0"/>
      <Cell N="LocPinX" V="0.16"/>
      <Cell N="LocPinY" V="0"/>
      <Cell N="Angle" V="0"/>
      <Cell N="LineColor" V="${entry.color}"/>
      <Cell N="LineWeight" V="0.02"/>
      <Cell N="LinePattern" V="${pattern}"/>
      <Cell N="EndArrow" V="4"/>
      <Cell N="EndArrowSize" V="1"/>
      <Section N="Geometry" IX="0">
        <Cell N="NoFill" V="1"/>
        <Cell N="NoLine" V="0"/>
        <Cell N="NoShow" V="0"/>
        <Cell N="NoSnap" V="0"/>
        <Row T="MoveTo" IX="1"><Cell N="X" V="0"/><Cell N="Y" V="0"/></Row>
        <Row T="LineTo" IX="2"><Cell N="X" V="0.32"/><Cell N="Y" V="0"/></Row>
      </Section>
    </Shape>`;
}

/** The text label beside a legend swatch. */
function legendTextXml(id: number, x: number, y: number, width: number, text: string): string {
  return `    <Shape ID="${id}" NameU="LegendText.${id}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(x + width / 2)}"/>
      <Cell N="PinY" V="${f(y)}"/>
      <Cell N="Width" V="${f(width)}"/>
      <Cell N="Height" V="0.18"/>
      <Cell N="LocPinX" V="${f(width / 2)}"/>
      <Cell N="LocPinY" V="0.09"/>
      <Cell N="Angle" V="0"/>
      <Cell N="LinePattern" V="0"/>
      <Cell N="FillPattern" V="0"/>
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="#475569"/><Cell N="Size" V="${LEGEND_FONT_IN}"/></Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="0"/></Row>
      </Section>
      <Text>${esc(text)}</Text>
    </Shape>`;
}

/**
 * Emit a colour key so the Visio page agrees with the PNG's connection legend.
 * Rows stack upward from the bottom-left corner (Visio's Y-up origin).
 */
function buildConnectionLegend(
  startId: number,
  entries: ConnectionLegendEntry[],
  originX: number,
  originY: number,
): { shapes: string[]; nextId: number } {
  const shapes: string[] = [];
  let id = startId;
  const rowH = 0.24;
  const boxW = 2.4;
  const boxH = rowH * entries.length + 0.34;
  // Background panel.
  shapes.push(`    <Shape ID="${id++}" NameU="Legend.${startId}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(originX + boxW / 2)}"/>
      <Cell N="PinY" V="${f(originY + boxH / 2)}"/>
      <Cell N="Width" V="${f(boxW)}"/>
      <Cell N="Height" V="${f(boxH)}"/>
      <Cell N="LocPinX" V="${f(boxW / 2)}"/>
      <Cell N="LocPinY" V="${f(boxH / 2)}"/>
      <Cell N="Angle" V="0"/>
      <Cell N="FillForegnd" V="#FFFFFF"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="#CBD5E1"/>
      <Cell N="LineWeight" V="0.01"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="Rounding" V="0.04"/>
${roundedRectGeometry()}
    </Shape>`);
  // Title.
  shapes.push(legendTextXml(id++, originX + 0.12, originY + boxH - 0.18, boxW - 0.24, 'Connections'));
  entries.forEach((entry, index) => {
    const rowY = originY + 0.14 + index * rowH;
    shapes.push(legendSwatchXml(id++, originX + 0.14, rowY, entry));
    shapes.push(legendTextXml(id++, originX + 0.62, rowY, boxW - 0.72, entry.label));
  });
  return { shapes, nextId: id };
}

/**
 * Emit the numbered workflow narration beside the drawing.
 *
 * An Azure Architecture Center diagram never shows a numbered callout without
 * the sentence it refers to; on its own the badge is just an unexplained digit.
 * Rows stack downward from the top-left of the panel.
 */
function buildWorkflowPanel(
  startId: number,
  entries: WorkflowListEntry[],
  originX: number,
  topY: number,
  width: number,
): { shapes: string[]; nextId: number } {
  const shapes: string[] = [];
  let id = startId;
  const rowH = 0.26;
  const boxH = rowH * entries.length + 0.34;
  const originY = topY - boxH;
  shapes.push(`    <Shape ID="${id++}" NameU="Workflow.${startId}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(originX + width / 2)}"/>
      <Cell N="PinY" V="${f(originY + boxH / 2)}"/>
      <Cell N="Width" V="${f(width)}"/>
      <Cell N="Height" V="${f(boxH)}"/>
      <Cell N="LocPinX" V="${f(width / 2)}"/>
      <Cell N="LocPinY" V="${f(boxH / 2)}"/>
      <Cell N="Angle" V="0"/>
      <Cell N="FillForegnd" V="#FFFFFF"/>
      <Cell N="FillPattern" V="1"/>
      <Cell N="LineColor" V="#CBD5E1"/>
      <Cell N="LineWeight" V="0.01"/>
      <Cell N="LinePattern" V="1"/>
      <Cell N="Rounding" V="0.04"/>
${roundedRectGeometry()}
    </Shape>`);
  shapes.push(legendTextXml(id++, originX + 0.12, originY + boxH - 0.18, width - 0.24, 'Workflow'));
  entries.forEach((entry, index) => {
    const rowY = originY + boxH - 0.34 - index * rowH;
    shapes.push(legendTextXml(id++, originX + 0.14, rowY, 0.3, `${entry.step}.`));
    shapes.push(legendTextXml(id++, originX + 0.46, rowY, width - 0.6, entry.description));
  });
  return { shapes, nextId: id };
}

function pagesXml(pageWidthIn: number, pageHeightIn: number, title: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Pages xmlns="${VISIO_NS}" xmlns:r="${REL_NS}">
  <Page ID="0" NameU="${esc(title)}" Name="${esc(title)}" ViewScale="-1" ViewCenterX="${f(pageWidthIn / 2)}" ViewCenterY="${f(pageHeightIn / 2)}">
    <PageSheet LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PageWidth" V="${f(pageWidthIn)}"/>
      <Cell N="PageHeight" V="${f(pageHeightIn)}"/>
      <Cell N="ShdwOffsetX" V="0.06"/>
      <Cell N="ShdwOffsetY" V="-0.06"/>
      <Cell N="PageScale" V="1"/>
      <Cell N="DrawingScale" V="1"/>
      <Cell N="DrawingSizeType" V="3"/>
      <Cell N="DrawingScaleType" V="0"/>
      <Cell N="InhibitSnap" V="0"/>
      <Cell N="PageLockReplace" V="0"/>
      <Cell N="PageLockDuplicate" V="0"/>
      <Cell N="UIVisibility" V="0"/>
      <Section N="Layer">
${layerRow(LAYER_ZONES, 'Zones')}
${layerRow(LAYER_SERVICES, 'Azure services')}
${layerRow(LAYER_CONNECTIONS, 'Connections')}
      </Section>
    </PageSheet>
    <Rel r:id="rId1"/>
  </Page>
</Pages>`;
}

function pageContentsXml(shapes: string[], connects: string[]): string {
  const connectsBlock = connects.length
    ? `\n  <Connects>\n${connects.join('\n')}\n  </Connects>`
    : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<PageContents xmlns="${VISIO_NS}" xmlns:r="${REL_NS}">
  <Shapes>
${shapes.join('\n')}
  </Shapes>${connectsBlock}
</PageContents>`;
}

// ─── Package assembly ────────────────────────────────────────────────────────

export interface VsdxPackage {
  parts: Array<{ path: string; data: string | Uint8Array }>;
  pageWidthIn: number;
  pageHeightIn: number;
}

/**
 * Build every part of the .vsdx package. Split out from {@link buildVsdxBlob}
 * so the XML can be asserted in unit tests without a browser Blob.
 */
export async function buildVsdxPackage(
  nodes: Node[],
  edges: Edge[],
  diagramName = 'Azure Architecture',
  /**
   * Already-rendered icons, when the caller has them.
   *
   * Rasterisation needs a DOM, so under Node every icon resolves to nothing
   * and the drawing ships with no media and no page relationships at all —
   * which is the one configuration no user ever receives, and the one every
   * check was being run against. A caller that can supply the bitmaps gets the
   * real drawing measured instead.
   */
  presetIcons?: Map<string, RasterizedIcon>,
): Promise<VsdxPackage> {
  // Empty space is closed on both axes before the sheet is sized: a DR region
  // drawn 6000px east of the primary is a two-region architecture, not an
  // outlier to trim and not a stray to park, and exporting the void between
  // them cost 50in of a 72in sheet.
  // Visio refuses to open a page larger than 200in on a side, so a drawing
  // wider than that is not a big export, it is no export at all. Tighten the
  // gaps until it fits: every shape keeps its size and every label its point
  // size, and the only thing given up is distance — which is the one thing on
  // a sheet this size the reader was never going to use.
  const limitPx = (MAX_VISIO_PAGE_IN - PAGE_PADDING_IN * 2 - 0.5) * PX_PER_INCH;
  const raw = fitBoxesWithin(compactEmptyGutters(collectExportBoxes(nodes)), limitPx, limitPx);
  // Same narration the deck gets: only one hop between a given pair of services
  // is ever given a step number, so the other members of a fan carry a callout
  // that the panel never explains — or, once the fan drops its wording, say
  // nothing at all.
  const narrated = narrateEdgeCallouts(edges);
  // Match the PowerPoint strategy: draw 1 : 1 from the full bounds whenever the
  // page stays a sensible size, and only fall back to the dense-cluster bounds
  // (clamping the strays back on) when a far-placed node would otherwise blow
  // the page up into a near-empty plotter sheet.
  const fullBounds = computeBounds(raw.values());
  const fullWIn = Math.max(fullBounds.maxX - fullBounds.minX, 1) / PX_PER_INCH;
  const fullHIn = Math.max(fullBounds.maxY - fullBounds.minY, 1) / PX_PER_INCH;
  let bounds = fullBounds;
  let clampToPage = false;
  if (fullWIn > MAX_USEFUL_PAGE_IN || fullHIn > MAX_USEFUL_PAGE_IN) {
    const trimmed = computeContentBounds(raw.values());
    const trimmedWIn = Math.max(trimmed.maxX - trimmed.minX, 1) / PX_PER_INCH;
    const trimmedHIn = Math.max(trimmed.maxY - trimmed.minY, 1) / PX_PER_INCH;
    if (trimmedWIn < fullWIn * 0.8 || trimmedHIn < fullHIn * 0.8) {
      bounds = trimmed;
      clampToPage = true;
    }
  }
  // Pull the strays back onto the drawing once, in the drawing's own
  // coordinates, before anything is routed or measured. The clamp used to live
  // in `toRect`, in inches, which left the router planning hops to where a
  // stray used to be — on the "outlier" fixture one arrow finished at the page
  // corner, 0.3in clear of the tile the <Connects> table said it was glued to,
  // and Visio would have snapped that line across the page the first time the
  // reader moved anything. It also parks strays clear of each other: two nodes
  // far off the same corner clamp to the same corner, and the hop between them
  // was then drawn straight through the tile covering both.
  const parked = clampToPage ? clampedBoxes(raw, bounds) : { boxes: raw, bounds };
  const boxes = parked.boxes;
  bounds = parked.bounds;
  const { groups, services } = partitionBoxes(boxes);
  const routes = buildExportRoutes(narrated, boxes);

  const contentWIn = Math.max(bounds.maxX - bounds.minX, 1) / PX_PER_INCH;
  const contentHIn = Math.max(bounds.maxY - bounds.minY, 1) / PX_PER_INCH;
  // A numbered workflow gets its own band across the top of the sheet, so the
  // prose never lands on the drawing the way an overlaid panel would.
  const workflowEntries = workflowListFromEdges(narrated);
  const workflowBandIn = workflowEntries.length > 0 ? 0.26 * workflowEntries.length + 0.5 : 0;
  const pageWidthIn = f(Math.max(contentWIn + PAGE_PADDING_IN * 2, MIN_PAGE_W_IN));
  const pageHeightIn = f(Math.max(contentHIn + PAGE_PADDING_IN * 2 + workflowBandIn, MIN_PAGE_H_IN));

  // Centre the drawing on the page, converting to Visio's bottom-left origin.
  const offsetXIn = (pageWidthIn - contentWIn) / 2;
  const offsetYIn = (pageHeightIn - workflowBandIn - contentHIn) / 2 + workflowBandIn;
  const clampIn = (value: number, lo: number, hi: number) => Math.min(Math.max(value, lo), Math.max(lo, hi));
  // The workflow panel is opaque and is drawn after every service, so a stray
  // node clamped into its band would be painted over. Keep the clamp below it.
  const drawingTopIn = workflowBandIn + PAGE_PADDING_IN / 2;
  const toRect = (box: ExportBox): Rect => {
    const w = box.w / PX_PER_INCH;
    const h = box.h / PX_PER_INCH;
    let x = (box.x - bounds.minX) / PX_PER_INCH + offsetXIn;
    let topY = (box.y - bounds.minY) / PX_PER_INCH + offsetYIn;
    if (clampToPage) {
      x = clampIn(x, PAGE_PADDING_IN / 2, pageWidthIn - w - PAGE_PADDING_IN / 2);
      topY = clampIn(topY, drawingTopIn, pageHeightIn - h - PAGE_PADDING_IN / 2);
    }
    return { x, y: pageHeightIn - topY - h, w, h };
  };
  const toPoint = (point: Point): Point => {
    let x = (point.x - bounds.minX) / PX_PER_INCH + offsetXIn;
    let topY = (point.y - bounds.minY) / PX_PER_INCH + offsetYIn;
    if (clampToPage) {
      x = clampIn(x, 0, pageWidthIn);
      topY = clampIn(topY, drawingTopIn, pageHeightIn);
    }
    return { x, y: pageHeightIn - topY };
  };

  // Reroute is no longer needed here: `clampedBoxes` above moved the strays
  // before anything was routed, so the tile, the arrow aimed at it and the glue
  // record all agree. The inch-space clamp below stays as a backstop for the
  // workflow band, and is a no-op for a box already inside the drawing.

  const icons = presetIcons ?? await rasterizeIcons(services.map((box) => box.iconPath), 128);

  const shapes: string[] = [];
  const connects: string[] = [];
  const media: Array<{ file: string; bytes: Uint8Array }> = [];
  const pageRels: string[] = [];
  const shapeIdByNode = new Map<string, number>();
  let nextId = 1;
  let mediaIndex = 0;

  for (let zoneIndex = 0; zoneIndex < groups.length; zoneIndex += 1) {
    const zone = groups[zoneIndex];
    const id = nextId++;
    shapeIdByNode.set(zone.id, id);
    const palette = paletteForZone(zone, zoneIndex);
    shapes.push(zoneShapeXml(id, toRect(zone), zone.label, palette));
  }

  for (const service of services) {
    const rect = toRect(service);
    const groupId = nextId++;
    const rectId = nextId++;
    const iconId = nextId++;
    shapeIdByNode.set(service.id, groupId);

    const icon: RasterizedIcon | undefined = service.iconPath
      ? icons.get(service.iconPath)
      : undefined;
    let relId: string | null = null;
    if (icon) {
      mediaIndex += 1;
      const file = `image${mediaIndex}.png`;
      relId = `rId${mediaIndex}`;
      media.push({ file, bytes: icon.bytes });
      pageRels.push(
        `  <Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${file}"/>`,
      );
    }

    const meta = metaSubline(service);
    const properties = [
      {
        name: 'AzureService',
        label: 'Azure service',
        value: service.serviceName ?? service.label,
      },
      {
        name: 'Category',
        label: 'Category',
        value: service.category,
      },
      { name: 'NodeId', label: 'Diagram node ID', value: service.id },
    ];
    if (service.meta?.sku) {
      properties.push({ name: 'Sku', label: 'SKU', value: service.meta.sku });
    }
    if (service.meta?.region) {
      properties.push({ name: 'Region', label: 'Region', value: service.meta.region });
    }
    if (service.meta?.costLabel) {
      properties.push({ name: 'MonthlyCost', label: 'Monthly cost', value: service.meta.costLabel });
    }

    shapes.push(
      serviceGroupXml(
        { group: groupId, rect: rectId, icon: iconId },
        rect,
        service,
        paletteForService(service),
        relId,
        properties,
        meta,
      ),
    );
  }

  // Parallel hops between the same two services all put their text at the same
  // midpoint, so a fan of five wrote five sentences on top of each other and
  // the sheet was unreadable at exactly the place it had the most to say. Fan
  // the labels the way the arrows themselves are fanned, and hang each badge
  // off its own rung.
  const bundleOf = (route: ExportRoute): string => (route.sourceId < route.targetId
    ? `${route.sourceId}|${route.targetId}`
    : `${route.targetId}|${route.sourceId}`);
  const labelSize = (label: string): { w: number; h: number } => {
    const natural = estimateTextWidthIn(label, CONNECTOR_FONT_IN) + 0.08;
    const w = Math.min(Math.max(natural, 0.5), 1.7);
    const lines = Math.max(1, Math.ceil(estimateTextWidthIn(label, CONNECTOR_FONT_IN) / Math.max(w - 0.08, 0.1)));
    return { w, h: lines * CONNECTOR_FONT_IN * 1.3 + 0.05 };
  };
  // The same cut every other exporter makes. A 200-character sentence left
  // whole wraps into a twelve-line block that buries the services around it.
  // A muted hop is one whose fan could not be written out anywhere on the
  // sheet; it keeps its numbered callout and the workflow band explains it.
  const muted = new Set<string>();
  const labelOf = (route: ExportRoute): string => (route.label && !muted.has(route.id)
    ? truncateLabel(route.label, 42)
    : '');
  const ladder = new Map<string, { drop: number; along: number }>();
  const byBundle = new Map<string, ExportRoute[]>();
  for (const route of routes as ExportRoute[]) {
    // A numbered hop with no wording still puts a callout on the page, and that
    // callout has to be placed like everything else: leaving it out of the
    // search let it come to rest in the middle of a service tile.
    if (!route.label && route.stepNumber === undefined) continue;
    const key = bundleOf(route);
    const list = byBundle.get(key);
    if (list) list.push(route); else byBundle.set(key, [route]);
  }
  /** Unit normal of a route, oriented downward in page space. */
  const normalOf = (route: ExportRoute): Point => {
    const points = route.points.map(toPoint);
    const a = points[0] ?? { x: 0, y: 0 };
    const b = points[points.length - 1] ?? a;
    const len = Math.hypot(b.x - a.x, b.y - a.y) || 1;
    const n = { x: (b.y - a.y) / len, y: -(b.x - a.x) / len };
    return n.y > 0 ? { x: -n.x, y: -n.y } : n;
  };
  // A connector's local +Y is the line's LEFT normal, which is the opposite of
  // the page-down normal the badges use unless the line runs right to left.
  // Reading the sign back off the flipped normal cannot work — both cases come
  // out pointing down — so it is taken from the run direction itself.
  const localSignOf = (route: ExportRoute): number => {
    const points = route.points.map(toPoint);
    const a = points[0] ?? { x: 0, y: 0 };
    const b = points[points.length - 1] ?? a;
    return b.x - a.x < 0 ? 1 : -1;
  };
  /** Unit vector along the line, begin to end. */
  const directionOf = (route: ExportRoute): Point => {
    const points = route.points.map(toPoint);
    const a = points[0] ?? { x: 0, y: 0 };
    const b = points[points.length - 1] ?? a;
    const len = Math.hypot(b.x - a.x, b.y - a.y) || 1;
    return { x: (b.x - a.x) / len, y: (b.y - a.y) / len };
  };
  /** How long the connector is, so the text is never slid off its own arrow. */
  const runOf = (route: ExportRoute): number => {
    const points = route.points.map(toPoint);
    const a = points[0] ?? { x: 0, y: 0 };
    const b = points[points.length - 1] ?? a;
    return Math.hypot(b.x - a.x, b.y - a.y);
  };
  const midOf = (route: ExportRoute): Point => toPoint(route.labelAnchor);
  /**
   * Where Visio actually puts connector text: the shape's pin is the midpoint
   * of the begin→end chord, and TxtPin is measured from there. An elbowed hop's
   * label anchor sits on the polyline, which is somewhere else entirely, so a
   * search that reasoned about the anchor was placing labels it could not see.
   */
  const chordOf = (route: ExportRoute): Point => {
    const points = route.points.map(toPoint);
    const a = points[0] ?? { x: 0, y: 0 };
    const b = points[points.length - 1] ?? a;
    return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
  };
  /** The anchor expressed in the connector's own frame, relative to the chord. */
  const seatOf = (route: ExportRoute): { drop: number; along: number } => {
    const anchor = midOf(route);
    const chord = chordOf(route);
    const n = normalOf(route);
    const u = directionOf(route);
    const dx = anchor.x - chord.x;
    const dy = anchor.y - chord.y;
    return { drop: dx * n.x + dy * n.y, along: dx * u.x + dy * u.y };
  };
  // Where every label would land if nothing moved, so the ladders can be
  // stepped clear of the services and of each other. Visio has no autolayout
  // for connector text: whatever position the file carries is what the reader
  // opens, so the placement has to be settled here.
  const serviceRects = services.map((service) => {
    const rect = toRect(service);
    return { x: rect.x, y: rect.y, w: rect.w, h: rect.h };
  });
  const rectAt = (route: ExportRoute, drop: number, along: number): { x: number; y: number; w: number; h: number } => {
    const centre = chordOf(route);
    const n = normalOf(route);
    const u = directionOf(route);
    const size = labelSize(labelOf(route));
    return {
      x: centre.x + n.x * drop + u.x * along - size.w / 2,
      y: centre.y + n.y * drop + u.y * along - size.h / 2,
      w: size.w,
      h: size.h,
    };
  };
  const hit = (
    a: { x: number; y: number; w: number; h: number },
    b: { x: number; y: number; w: number; h: number },
  ): number => {
    const ow = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
    const oh = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return ow > 0 && oh > 0 ? ow * oh : 0;
  };
  // Every label starts where its own anchor put it, expressed in the frame the
  // drawing actually uses, and the ladder and the settle pass move it from
  // there.
  for (const [, members] of byBundle) {
    for (const member of members) ladder.set(member.id, seatOf(member));
  }
  /** Space the fan's text out along the normal, one rung per arrow. */
  const rungify = (members: ExportRoute[]): void => {
    if (members.length < 2) return;
    let rung = 0;
    for (const member of members) {
      const size = labelOf(member) ? labelSize(labelOf(member)).h : 0;
      rung = Math.max(rung, size + (member.stepNumber === undefined ? 0.05 : STEP_BADGE_IN + 0.07));
    }
    // Ranked by where the ARROW was fanned to, not by the order the edges were
    // declared, so rung n sits beside arrow n instead of crossing over it.
    const n = normalOf(members[0]);
    const ordered = [...members].sort((l, r) => {
      const lm = midOf(l);
      const rm = midOf(r);
      return (lm.x * n.x + lm.y * n.y) - (rm.x * n.x + rm.y * n.y);
    });
    ordered.forEach((member, index) => {
      const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
      ladder.set(member.id, { drop: (index - (ordered.length - 1) / 2) * rung, along: seat.along });
    });
  };
  for (const [, members] of byBundle) rungify(members);

  // Step each ladder — and each lone label — clear of the services and of the
  // labels already settled, and hold it on the sheet. Without this the rungs
  // land wherever the arrows happen to run: on a grid that is on top of the
  // tiles, on top of the next hop's sentence, or past the page edge, where
  // Visio quietly draws nothing at all.
  const placedLabels: Array<{ x: number; y: number; w: number; h: number }> = [];
  const settleOrder = [...byBundle.entries()].sort((l, r) => r[1].length - l[1].length);
  /**
   * Move one fan as a body and report where it came to rest. `blocked` is the
   * collision area alone: the drift terms only break ties, so mixing them into
   * the total would make a clean placement look like a dirty one and the
   * caller could never tell whether the fan actually fits.
   */
  const settle = (members: ExportRoute[]): { shift: number; slide: number; blocked: number; buried: number } => {
    const step = Math.max(
      0.14,
      ...members.map((member) => (labelOf(member) ? labelSize(labelOf(member)).h : STEP_BADGE_IN) + 0.05),
    );
    const badgeRoom = members.some((member) => member.stepNumber !== undefined) ? STEP_BADGE_IN + 0.06 : 0;
    const blockage = (shift: number, slide: number): number => {
      let cost = 0;
      for (const member of members) {
        const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
        const drop = seat.drop + shift;
        const along = seat.along + slide;
        const text = labelOf(member);
        const box = rectAt(member, drop, along);
        const parts = text ? [box] : [];
        if (member.stepNumber !== undefined && badgeRoom > 0) {
          const n = normalOf(member);
          const u = directionOf(member);
          const centre = chordOf(member);
          const away = drop + (text ? box.h / 2 + STEP_BADGE_IN / 2 + 0.03 : 0);
          parts.push({
            x: centre.x + n.x * away + u.x * along - STEP_BADGE_IN / 2,
            y: centre.y + n.y * away + u.y * along - STEP_BADGE_IN / 2,
            w: STEP_BADGE_IN,
            h: STEP_BADGE_IN,
          });
        }
        for (const part of parts) {
          if (part.x < 0.05 || part.y < 0.05
            || part.x + part.w > pageWidthIn - 0.05 || part.y + part.h > pageHeightIn - 0.05) {
            // Off the sheet is not a cost to be traded against, it is a deletion.
            cost += 100;
          }
          for (const rect of serviceRects) cost += hit(part, rect) * 4;
          for (const other of placedLabels) cost += hit(part, other) * 12;
        }
      }
      return cost;
    };
    // How much of the worst sentence a reader cannot see, as a fraction of the
    // sentence itself. `blockage` is a weighted score tuned for choosing
    // between positions, so its magnitude says nothing about legibility: a
    // label clipping a tile corner and a label buried under one both score
    // "greater than zero". Deciding whether wording survives needs the plain
    // question — what fraction of this text is covered?
    const buriedAt = (shift: number, slide: number): number => {
      let worst = 0;
      for (const member of members) {
        const text = labelOf(member);
        if (!text) continue;
        const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
        const box = rectAt(member, seat.drop + shift, seat.along + slide);
        const own = Math.max(box.w * box.h, 1e-6);
        if (box.x < 0.05 || box.y < 0.05
          || box.x + box.w > pageWidthIn - 0.05 || box.y + box.h > pageHeightIn - 0.05) {
          worst = 1;
          continue;
        }
        let covered = 0;
        for (const rect of serviceRects) covered += hit(box, rect);
        // Text over a tile is ugly but still readable — the tile is mostly
        // empty fill. Two sentences written on the same spot are both lost, and
        // no fraction of that is acceptable, so any contact a reader could
        // notice counts as fully buried. The threshold is the one the export
        // audit itself calls a defect, so the exporter and the rule that judges
        // it cannot disagree about what "unreadable" means.
        for (const other of placedLabels) {
          if (hit(box, other) > 0.01) return 1;
        }
        worst = Math.max(worst, Math.min(1, covered / own));
      }
      return worst;
    };
    // Sliding along the arrow is the cheaper move visually — the text still
    // reads as that arrow's — but both are a last resort against a label
    // sitting where it cannot be read.
    const score = (shift: number, slide: number): number => blockage(shift, slide)
      + Math.abs(shift) * 0.01 + Math.abs(slide) * 0.008;
    let best = { shift: 0, slide: 0 };
    let bestCost = score(0, 0);
    let bestBlocked = blockage(0, 0);
    // Far enough to cross the drawing. A fixed budget is a budget in units of
    // the label's own height, so on a big sheet a ladder could only ever travel
    // an inch and never reached the clear band one row away.
    const rings = Math.max(8, Math.min(60, Math.ceil(Math.max(pageWidthIn, pageHeightIn) / Math.max(step, 0.05))));
    // Sliding a label past its own endpoints parks it beside a service instead
    // of on the hop it belongs to, so the run is the limit.
    const slideStep = Math.max(0.12, Math.max(...members.map((member) => labelSize(labelOf(member)).w)) / 2);
    const slides = Math.max(0, Math.floor(Math.max(...members.map(runOf)) / (2 * slideStep)));
    const consider = (shift: number, slide: number): void => {
      const cost = score(shift, slide);
      if (cost < bestCost) {
        bestCost = cost;
        bestBlocked = blockage(shift, slide);
        best = { shift, slide };
      }
    };
    for (let ring = 1; bestBlocked > 0 && ring <= rings; ring += 1) {
      for (let s = -Math.min(slides, ring); s <= Math.min(slides, ring); s += 1) {
        consider(ring * step, s * slideStep);
        consider(-ring * step, s * slideStep);
        if (s !== 0) consider(0, s * slideStep);
      }
    }
    return { ...best, blocked: bestBlocked, buried: buriedAt(best.shift, best.slide) };
  };
  // How much of a sentence can be covered before dropping it outright is the
  // kinder outcome. A sliver off one corner still reads; a third of the words
  // hidden under a service tile does not, and a half-read sentence is worse
  // than a numbered callout the workflow band spells out in full.
  const BURIED_LIMIT = 0.35;
  // Wording a muted label handed to the workflow band, by step number. Nothing
  // may be muted unless it lands here or the row already says it, because the
  // Visio sheet has nowhere else for the sentence to go.
  const mutedWording = new Map<number, string>();
  const narratedRows = new Map(workflowEntries.map((entry) => [entry.step, entry.description]));
  for (const [, members] of settleOrder) {
    let placement = settle(members);
    // A fan whose sentences cannot be written anywhere clear keeps its numbers
    // and drops its wording, exactly as the deck does. Ten hops between one
    // pair of services need a ladder taller than the sheet, and a half-hidden
    // sentence is worse than a callout the workflow band spells out in full.
    //
    // A lone label has exactly the same problem and used to have no way out of
    // it: on a tight grid a sentence is wider than the lane between two columns
    // and there is no clear air anywhere on the sheet, so it simply shipped on
    // top of whatever it landed on. Depth was never the point — being unable to
    // read the words is.
    const explained = members.every(
      (member) => member.stepNumber !== undefined && narratedRows.has(member.stepNumber),
    );
    const stuck = members.length >= 2 ? placement.blocked > 0 : placement.buried >= BURIED_LIMIT;
    if (stuck && explained && members.some((member) => labelOf(member))) {
      const before = placement.blocked;
      const buriedBefore = placement.buried;
      const handed = members
        .filter((member) => member.stepNumber !== undefined && member.label
          && !carriesWording(narratedRows.get(member.stepNumber) ?? '', member.label))
        .map((member) => [member.stepNumber as number, member.label as string] as const);
      for (const member of members) muted.add(member.id);
      rungify(members);
      const retry = settle(members);
      if (retry.blocked < before || retry.buried < buriedBefore) {
        placement = retry;
        // Muting is only honest once the sentence has somewhere else to be
        // read, so the handover happens here — after the retry has proved the
        // trade was worth making — and never on the branch that puts the
        // wording back.
        for (const [step, label] of handed) mutedWording.set(step, label);
      } else {
        for (const member of members) muted.delete(member.id);
        rungify(members);
      }
    }
    const best = { shift: placement.shift, slide: placement.slide };
    for (const member of members) {
      const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
      const placed = { drop: seat.drop + best.shift, along: seat.along + best.slide };
      ladder.set(member.id, placed);
      if (labelOf(member)) placedLabels.push(rectAt(member, placed.drop, placed.along));
      if (member.stepNumber !== undefined) {
        const n = normalOf(member);
        const u = directionOf(member);
        const centre = chordOf(member);
        const box = rectAt(member, placed.drop, placed.along);
        const away = placed.drop + (labelOf(member) ? box.h / 2 + STEP_BADGE_IN / 2 + 0.03 : 0);
        placedLabels.push({
          x: centre.x + n.x * away + u.x * placed.along - STEP_BADGE_IN / 2,
          y: centre.y + n.y * away + u.y * placed.along - STEP_BADGE_IN / 2,
          w: STEP_BADGE_IN,
          h: STEP_BADGE_IN,
        });
      }
    }
  }

  for (const route of routes as ExportRoute[]) {
    const sourceId = shapeIdByNode.get(route.sourceId);
    const targetId = shapeIdByNode.get(route.targetId);
    if (sourceId === undefined || targetId === undefined) continue;
    const id = nextId++;
    const text = labelOf(route);
    const size = text ? labelSize(text) : { w: 0, h: 0 };
    const seat = ladder.get(route.id) ?? { drop: 0, along: 0 };
    // Local +Y is the line's left normal; the page-space normal used for the
    // badge is its opposite whenever the line runs left to right, so the sign
    // has to be carried across or the ladder and the badges fan apart.
    const pageNormal = normalOf(route);
    const localSign = localSignOf(route);
    shapes.push(
      connectorShapeXml(
        id,
        route.points.map(toPoint),
        text,
        route.color,
        visioLinePattern(route),
        route.opacity,
        route.bidirectional,
        text ? { drop: localSign * seat.drop, along: seat.along, w: size.w, h: size.h } : undefined,
      ),
    );
    connects.push(connectXml(id, sourceId, targetId));
    if (route.stepNumber !== undefined) {
      // The badge hangs off the label's own rung, in the same frame the label
      // is drawn in. Measuring it from the route's anchor instead put it inches
      // from its own sentence on any elbowed hop, because Visio pins connector
      // text to the begin→end chord and the anchor sits on the polyline.
      const anchor = chordOf(route);
      const { x: nx, y: ny } = pageNormal;
      const along = directionOf(route);
      // Each rung carries its own badge, so the badge steps off that rung and
      // not off the line: with a fan, one shared drop stacked every badge in
      // the bundle on the same spot.
      const drop = (text ? size.h / 2 + STEP_BADGE_IN / 2 + 0.03 : 0) + seat.drop;
      // A clamped connector can put its anchor at the very page edge, and the
      // normal offset then pushes the badge off the sheet, where Visio simply
      // does not draw it.
      const half = STEP_BADGE_IN / 2;
      shapes.push(
        stepBadgeXml(
          nextId++,
          {
            x: clampIn(anchor.x + nx * drop + along.x * seat.along, half, pageWidthIn - half),
            y: clampIn(anchor.y + ny * drop + along.y * seat.along, half, pageHeightIn - half),
          },
          route.stepNumber,
        ),
      );
    }
  }

  // Colour key so the Visio page can't contradict the PNG's connection legend.
  const legendEntries = usedConnectionLegend(edges);
  if (legendEntries.length > 0) {
    const legend = buildConnectionLegend(nextId, legendEntries, 0.35, 0.35);
    nextId = legend.nextId;
    shapes.push(...legend.shapes);
  }

  // The numbered prose that the connector badges point at.
  if (workflowEntries.length > 0) {
    const panel = buildWorkflowPanel(
      nextId,
      // A label that was muted traded its wording for this row, so the row has
      // to say it. The same parenthesis the deck uses, so the two exports of
      // one drawing read alike.
      workflowEntries.map((entry) => {
        const handed = mutedWording.get(entry.step);
        return handed ? { ...entry, description: `${entry.description}（${handed}）` } : entry;
      }),
      0.35,
      pageHeightIn - 0.2,
      Math.min(Math.max(pageWidthIn - 0.7, 2.4), 7.5),
    );
    nextId = panel.nextId;
    shapes.push(...panel.shapes);
  }

  const parts: Array<{ path: string; data: string | Uint8Array }> = [
    { path: '[Content_Types].xml', data: CONTENT_TYPES },
    { path: '_rels/.rels', data: ROOT_RELS },
    { path: 'docProps/core.xml', data: coreXml(diagramName) },
    { path: 'docProps/app.xml', data: APP_XML },
    { path: 'visio/document.xml', data: DOCUMENT_XML },
    { path: 'visio/_rels/document.xml.rels', data: DOCUMENT_RELS },
    { path: 'visio/windows.xml', data: windowsXml(pageWidthIn, pageHeightIn) },
    { path: 'visio/pages/pages.xml', data: pagesXml(pageWidthIn, pageHeightIn, diagramName) },
    { path: 'visio/pages/_rels/pages.xml.rels', data: PAGES_RELS },
    { path: 'visio/pages/page1.xml', data: pageContentsXml(shapes, connects) },
  ];

  if (media.length > 0) {
    parts.push({
      path: 'visio/pages/_rels/page1.xml.rels',
      data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n${pageRels.join('\n')}\n</Relationships>`,
    });
    for (const file of media) {
      parts.push({ path: `visio/media/${file.file}`, data: file.bytes });
    }
  }

  return { parts, pageWidthIn, pageHeightIn };
}

/**
 * Build a .vsdx package for the diagram and return it as a Blob.
 */
export async function buildVsdxBlob(
  nodes: Node[],
  edges: Edge[],
  diagramName = 'Azure Architecture',
): Promise<Blob> {
  const { parts } = await buildVsdxPackage(nodes, edges, diagramName);
  const zip = new JSZip();
  for (const part of parts) {
    zip.file(part.path, part.data);
  }
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.ms-visio.drawing',
    compression: 'DEFLATE',
  });
}
