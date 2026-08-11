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
  computeBounds,
  computeContentBounds,
  metaSubline,
  partitionBoxes,
  usedConnectionLegend,
  zoneStyleFor,
  type ConnectionLegendEntry,
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
const CORNER_ROUNDING_IN = 0.08;

/**
 * Visio font sizes are inches (1 pt = 1/72"). These match the PowerPoint export
 * at 1 : 1 — a 150 px tile is 1.56" wide, so the label reads at ~7.6 pt and the
 * SKU sub-line at ~6 pt instead of the previous near-illegible 6.5/5 pt.
 */
const LABEL_FONT_IN = 0.105;
const META_FONT_IN = 0.083;
const CONNECTOR_FONT_IN = 0.1;
const LEGEND_FONT_IN = 0.1;

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
  const characterRows = meta
    ? `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${palette.text}"/><Cell N="Size" V="${LABEL_FONT_IN}"/></Row>
        <Row IX="1"><Cell N="Font" V="1"/><Cell N="Color" V="#64748B"/><Cell N="Size" V="${META_FONT_IN}"/></Row>`
    : `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${palette.text}"/><Cell N="Size" V="${LABEL_FONT_IN}"/></Row>`;
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
      <Cell N="TextBkgnd" V="2"/>
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
): Promise<VsdxPackage> {
  const boxes = collectExportBoxes(nodes);
  const { groups, services } = partitionBoxes(boxes);
  const routes = buildExportRoutes(edges, boxes);
  // Match the PowerPoint strategy: draw 1 : 1 from the full bounds whenever the
  // page stays a sensible size, and only fall back to the dense-cluster bounds
  // (clamping the strays back on) when a far-placed node would otherwise blow
  // the page up into a near-empty plotter sheet.
  const fullBounds = computeBounds(boxes.values());
  const fullWIn = Math.max(fullBounds.maxX - fullBounds.minX, 1) / PX_PER_INCH;
  const fullHIn = Math.max(fullBounds.maxY - fullBounds.minY, 1) / PX_PER_INCH;
  let bounds = fullBounds;
  let clampToPage = false;
  if (fullWIn > MAX_USEFUL_PAGE_IN || fullHIn > MAX_USEFUL_PAGE_IN) {
    const trimmed = computeContentBounds(boxes.values());
    const trimmedWIn = Math.max(trimmed.maxX - trimmed.minX, 1) / PX_PER_INCH;
    const trimmedHIn = Math.max(trimmed.maxY - trimmed.minY, 1) / PX_PER_INCH;
    if (trimmedWIn < fullWIn * 0.8 || trimmedHIn < fullHIn * 0.8) {
      bounds = trimmed;
      clampToPage = true;
    }
  }

  const contentWIn = Math.max(bounds.maxX - bounds.minX, 1) / PX_PER_INCH;
  const contentHIn = Math.max(bounds.maxY - bounds.minY, 1) / PX_PER_INCH;
  const pageWidthIn = f(Math.max(contentWIn + PAGE_PADDING_IN * 2, MIN_PAGE_W_IN));
  const pageHeightIn = f(Math.max(contentHIn + PAGE_PADDING_IN * 2, MIN_PAGE_H_IN));

  // Centre the drawing on the page, converting to Visio's bottom-left origin.
  const offsetXIn = (pageWidthIn - contentWIn) / 2;
  const offsetYIn = (pageHeightIn - contentHIn) / 2;
  const clampIn = (value: number, lo: number, hi: number) => Math.min(Math.max(value, lo), Math.max(lo, hi));
  const toRect = (box: ExportBox): Rect => {
    const w = box.w / PX_PER_INCH;
    const h = box.h / PX_PER_INCH;
    let x = (box.x - bounds.minX) / PX_PER_INCH + offsetXIn;
    let topY = (box.y - bounds.minY) / PX_PER_INCH + offsetYIn;
    if (clampToPage) {
      x = clampIn(x, PAGE_PADDING_IN / 2, pageWidthIn - w - PAGE_PADDING_IN / 2);
      topY = clampIn(topY, PAGE_PADDING_IN / 2, pageHeightIn - h - PAGE_PADDING_IN / 2);
    }
    return { x, y: pageHeightIn - topY - h, w, h };
  };
  const toPoint = (point: Point): Point => {
    let x = (point.x - bounds.minX) / PX_PER_INCH + offsetXIn;
    let topY = (point.y - bounds.minY) / PX_PER_INCH + offsetYIn;
    if (clampToPage) {
      x = clampIn(x, 0, pageWidthIn);
      topY = clampIn(topY, 0, pageHeightIn);
    }
    return { x, y: pageHeightIn - topY };
  };

  const icons = await rasterizeIcons(services.map((box) => box.iconPath), 128);

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

  for (const route of routes as ExportRoute[]) {
    const sourceId = shapeIdByNode.get(route.sourceId);
    const targetId = shapeIdByNode.get(route.targetId);
    if (sourceId === undefined || targetId === undefined) continue;
    const id = nextId++;
    shapes.push(
      connectorShapeXml(
        id,
        route.points.map(toPoint),
        route.label,
        route.color,
        visioLinePattern(route),
        route.opacity,
        route.bidirectional,
      ),
    );
    connects.push(connectXml(id, sourceId, targetId));
  }

  // Colour key so the Visio page can't contradict the PNG's connection legend.
  const legendEntries = usedConnectionLegend(edges);
  if (legendEntries.length > 0) {
    const legend = buildConnectionLegend(nextId, legendEntries, 0.35, 0.35);
    nextId = legend.nextId;
    shapes.push(...legend.shapes);
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
