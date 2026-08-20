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
import { stripXmlForbidden } from '../utils/xmlText';
import {
  buildExportRoutes,
  categoryStyle,
  collectExportBoxes,
  compactEmptyGutters,
  fitBoxesWithin,
  scaleBoxesWithin,
  boxScaleWithin,
  clampedBoxes,
  computeBounds,
  computeContentBounds,
  fitLabelToLines,
  metaSubline,
  narrateEdgeCallouts,
  truncateLabel,
  drawableInColumn,
  advanceWidthIn,
  trailingWhitespaceIn,
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
 * The thinnest line Visio renders as a line rather than as nothing: its
 * hairline, 1/4 point.
 */
const HAIRLINE_IN = 0.0035;

/**
 * Pen width for drawing content, in inches, at a given sheet scale.
 *
 * The literals this replaces did not scale while everything around them did, so
 * on a deeply reduced sheet the ink stopped being an outline and became the
 * shape: at 900 stages a 0.0125in border was a quarter of the tile's height and
 * a connector stroke was over half the gap it crossed, which is why a large
 * Visio export read as a grey mat with nothing legible on it. Zooming in does
 * not recover it — the weight is stored in the file, not derived from the view.
 *
 * The floor is Visio's own hairline rather than zero. Below it the pen would go
 * on thinning past what the renderer can draw and the line would stop appearing
 * altogether; at the hairline it stays visible while taking as little of the
 * tile as the format allows.
 *
 * Page furniture — the legend and the workflow band — deliberately does not use
 * this. Those panels are drawn at natural size on the sheet whatever the drawing
 * inside is reduced to, so their pen is already in the right proportion to them.
 */
function penIn(natural: number, scale: number, maxIn = Infinity): number {
  return Math.max(HAIRLINE_IN, Math.min(natural * scale, maxIn));
}

/** Height of one numbered row in the workflow band. */
const WORKFLOW_ROW_IN = 0.26;
/**
 * How many columns the workflow band may be split into.
 *
 * Splitting trades column height for column width, and the sentences have to
 * fit the width: past a dozen columns on an 11in page a step is under an inch
 * wide and every row wraps to more lines than the split saved.
 */
const MAX_WORKFLOW_COLUMNS = 12;

/**
 * The narrowest column a workflow sentence may be set in.
 *
 * The column count is chosen to minimise the band's height, and height is the
 * only thing that score measures - so with brief descriptions, where splitting
 * always shortens the stack, it ran straight to the cap and set twelve steps in
 * columns 0.8583in wide. `workflowRowHeightIn` takes 0.6in of that for the
 * number and the gutters, which left 0.2583in of text column: about two
 * characters, wrapping "acknowledged" down twelve lines. Legibility is not
 * something the height score can see, so it has to be a bound on the search.
 * 1.6in leaves a 1.0in text column, which is where a short sentence still
 * reads as one.
 */
const MIN_WORKFLOW_COL_IN = 1.6;

/**
 * The other half of the same bound, because a width floor cannot see this one.
 *
 * The minimiser scores nothing but stack height, and on LONG descriptions the
 * height still falls with every split while the column stays comfortably above
 * the width floor - so it split to the cap anyway and set the median sentence
 * of `visio-token-workflow` in SIXTEEN wrapped lines. Sixteen lines of a column
 * two or three words wide is not a sentence a reader follows; it is a ribbon.
 *
 * Neither bound covers for the other. Twelve one-word steps shred to a hairline
 * column while each still sets in ONE line, which no line bound can see; twelve
 * long steps keep a wide column while every one wraps to eight, which no width
 * floor can see. `probe-brief-workflow` and `probe-shredded-workflow` hold the
 * two geometries so that neither can be deleted on the grounds of the other.
 *
 * Three, because two is the ordinary shape of a step description at this size
 * and four is where the eye starts losing its place returning to the margin.
 */
const MAX_WORKFLOW_MEDIAN_LINES = 3;

/**
 * How many times its unsplit height a single sentence may be folded.
 *
 * One extra fold is the ordinary price of a column; three times over is a
 * ribbon. Twice is the bar.
 */
const WORKFLOW_TAIL_MULTIPLE = 2;

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

/** Zone caption, and the number inside a step badge — both drawing content. */
const ZONE_TITLE_FONT_IN = 0.13;
const BADGE_FONT_IN = 0.11;

interface VisioFonts {
  /** How far the drawing was scaled to fit the page; 1 when it was not. */
  scale: number;
  label: number;
  meta: number;
  connector: number;
  zone: number;
  badge: number;
  /**
   * The heaviest a drawing pen may be on this sheet.
   *
   * Same blind spot as the badge, same cause: `scale` only knows about the
   * page fit. A tile drawn small by hand kept a 0.0125in border, and at that
   * size the border is the shape.
   */
  penMaxIn: number;
}

const NATURAL_FONTS: VisioFonts = {
  scale: 1,
  label: LABEL_FONT_IN,
  meta: META_FONT_IN,
  connector: CONNECTOR_FONT_IN,
  zone: ZONE_TITLE_FONT_IN,
  badge: BADGE_FONT_IN,
  // No tiles have been measured yet, so nothing is capped yet.
  penMaxIn: Infinity,
};

/**
 * Type for a drawing that had to be scaled down to fit the page Visio will
 * open. The floors above are a promise about an ordinary drawing, not a law of
 * geometry: once a thousand tiles are squeezed onto 200in a tile is an eighth
 * of an inch wide and no type is legible on it at any size. Holding the point
 * size fixed there does not rescue the words, it prints "Azure Front Door" over
 * seven of its neighbours and destroys the structure too — a reader can zoom
 * into small type but cannot untangle overlapping type.
 *
 * So the type comes down with the drawing, all the way. A floor was tried and
 * is wrong: below about a quarter size it has stopped rescuing anything a
 * reader can read — a name at 1.9pt is not legible at any zoom — and it is
 * still buying that nothing with the overlap it was meant to prevent, because
 * Visio wraps a name too big for its shape into more and more lines inside a
 * text block that is itself shrinking. Proportion is the only property left
 * worth keeping at that size.
 *
 * A zone caption and a numbered step badge are drawing content and are in
 * here. The legend and the workflow band are page furniture laid out in page
 * inches, so they are not, and keep their natural size.
 */
function fontsForScale(scale: number): VisioFonts {
  if (scale >= 0.999) return NATURAL_FONTS;
  // Only to keep the XML well formed; nothing is drawn at this size.
  const k = Math.max(0.001, scale);
  return {
    scale: k,
    label: LABEL_FONT_IN * k,
    meta: META_FONT_IN * k,
    connector: CONNECTOR_FONT_IN * k,
    zone: ZONE_TITLE_FONT_IN * k,
    badge: BADGE_FONT_IN * k,
    penMaxIn: Infinity,
  };
}

/**
 * The same fonts, with the badge and the pen held to the tiles they will be
 * drawn among.
 *
 * Separate from `fontsForScale` because it needs the laid-out boxes, which
 * that function has no business knowing. The typical tile is the median, not
 * the smallest: one deliberately tiny node should not thin every line on the
 * sheet, but a page made of tiny nodes should.
 *
 * `badgeBoxes` is narrower than `boxes` on purpose. A pen is drawn everywhere,
 * so its ceiling is a statement about the whole sheet; a badge is drawn on a
 * numbered arrow, so its ceiling is a statement about the tiles at the ends of
 * those arrows and about nothing else. Taking the badge median over every
 * shape left the statistic one node from wrong in the other direction: five
 * 200px tiles beside four slivers gives a 2.08in median and a correct 0.240in
 * disc, and moving a single node to make it five slivers gives a 0.146in
 * median, collapses the same seven discs to the 0.1119in floor, and does it to
 * a workflow that runs entirely among the 2in tiles.
 */
/**
 * The same fonts, with the pen held to the tiles it will be drawn among.
 *
 * Separate from `fontsForScale` because it needs the laid-out boxes, which
 * that function has no business knowing. The typical tile is the median, not
 * the smallest: one deliberately tiny node should not thin every line on the
 * sheet, but a page made of tiny nodes should.
 */
function withDrawingCeilings(
  fonts: VisioFonts,
  boxes: Iterable<{ w: number; h: number }>,
): VisioFonts {
  const all = [...boxes].filter((b) => b.w > 0 && b.h > 0);
  if (all.length === 0) return fonts;
  const median = (xs: number[]): number => {
    const s = [...xs].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  };
  const tileHIn = median(all.map((b) => b.h)) / PX_PER_INCH;
  return {
    ...fonts,
    // A sixteenth is where an outline stops being a border and starts being
    // the shape. `penIn` still applies Visio's hairline underneath this, so
    // the line never thins into nothing.
    //
    // The badge used to be capped here too, from the same box list. It is not
    // any more: a pen is drawn everywhere on the sheet, so a sheet-wide
    // statistic answers the question about it, and a badge is drawn on one
    // arrow, so no sheet-wide statistic ever could. See `badgeCeilingIn`.
    penMaxIn: tileHIn / 16,
  };
}

/**
 * Approximate rendered width in inches, from the measured Yu Gothic UI
 * advances shared with the PowerPoint exporter.
 *
 * The flat 0.54 em this used to charge is the average LOWERCASE advance, so
 * every title-cased service name measured narrow and the sheet believed a label
 * needed one line fewer than it draws. Sizes here are in INCHES, not points, so
 * the shared helper is called at 72pt and scaled.
 */
function estimateTextWidthIn(text: string, fontSizeIn: number): number {
  return advanceWidthIn(text, 72) * fontSizeIn;
}

const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** Layer indexes declared on the page sheet. */
const LAYER_ZONES = 0;
const LAYER_SERVICES = 1;
const LAYER_CONNECTIONS = 2;

/**
 * Escape a string for XML, after removing what XML cannot carry at all.
 *
 * The escaping and the stripping are separate problems and both are required.
 * `& < > " '` have encodings; the C0 controls do not — `&#11;` is exactly as
 * illegal as a raw U+000B — so a vertical tab pasted out of Word into a service
 * name produced a `.vsdx` that Visio refuses to open, with no error at export
 * time and nothing to see until the recipient double-clicks it.
 */
function esc(value: string): string {
  return stripXmlForbidden(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

const f = (n: number) => +n.toFixed(4);

/**
 * Font sizes round UP. `f`'s 4-decimal truncation takes the 7pt sub-line to
 * 0.0972in, which is 6.998pt — a hair under the floor both exporters promise,
 * and enough to fail the deck's own legibility gate on a size that was correct
 * before it was written down.
 */
const ff = (n: number) => +(Math.ceil(n * 1e6) / 1e6).toFixed(6);

// ─── Palette (single source of truth in the shared geometry layer) ───────────

interface Palette { fill: string; line: string; text: string }

const CONNECTOR_TEXT = '#374151';

/** Adapt the shared category style ({bg,border,text}) to Visio's cell names. */
function paletteForService(box: ExportBox): Palette {
  const style = categoryStyle(box.category);
  return { fill: style.bg, line: style.border, text: style.text };
}

/** Adapt the shared zone style (honours `data.customColor`) to Visio's cells. */
function paletteForZone(box: ExportBox): Palette {
  const style = zoneStyleFor(box);
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

/**
 * Content types, generated from the page count rather than patched into.
 *
 * A part with no content type is not a missing feature, it is a REJECTED FILE:
 * Visio refuses the whole package, so the reader loses the drawing as well as
 * the index. This used to be two constants, the second built by running
 * `.replace()` over the first with a literal marker that carried its own two
 * spaces of indentation. `String.replace` returns the subject UNCHANGED when
 * the pattern does not match, and says nothing. Reindent the block, reorder
 * its overrides, or rename the part, and the package ships page 2 present in
 * `parts`, listed in `pages.xml` and resolving through `pages.xml.rels` - with
 * no content type, and every gate green, because the file is still well formed.
 *
 * Generating from the page count removes the marker, and removes the ceiling
 * with it: a page 3 gets its override by existing rather than by someone
 * remembering to add a third constant.
 */
/**
 * Remove markup, and keep removing until there is none left. One pass of
 * `/<[^>]*>/g` leaves residue on nested brackets - `<<b>>` becomes `<>` - and
 * the caller then treats a string that still carries markup as a service name.
 */
function stripMarkup(text: string): string {
  let out = text;
  for (;;) {
    const next = out.replace(/<[^>]*>/g, '');
    if (next === out) return out;
    out = next;
  }
}

function contentTypesXml(pageCount: number): string {
  const pages = Array.from({ length: Math.max(1, pageCount) }, (_, i) => (
    `  <Override PartName="/visio/pages/page${i + 1}.xml" ContentType="application/vnd.ms-visio.page+xml"/>`
  )).join('\n');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/>
  <Override PartName="/visio/pages/pages.xml" ContentType="application/vnd.ms-visio.pages+xml"/>
${pages}
  <Override PartName="/visio/windows.xml" ContentType="application/vnd.ms-visio.windows+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}
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

const PAGES_RELS_WITH_INDEX = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/page" Target="page2.xml"/>
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
 *
 * THE DECLARED FACE IS SEGOE UI; THE WIDTH TABLE IS YU GOTHIC UI. That is not
 * a defect and it is not worth "fixing" - it is written down here so nobody
 * loses an afternoon to it a third time. The two faces are close enough that
 * the difference cannot reach a layout decision: over the ASCII range the mean
 * absolute difference is 0.0017 em, the Cyrillic and Greek advances are
 * bit-identical between them, and over the real service names this exporter
 * draws the two models land 0.00% apart. Yu Gothic UI is the table because it
 * is the face that has to carry the Japanese names, and a model that is right
 * about kanji and 0.0017 em out on Latin is the correct trade.
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
function zoneShapeXml(
  id: number,
  rect: Rect,
  label: string,
  palette: Palette,
  fonts: VisioFonts = NATURAL_FONTS,
): string {
  // A zone is drawing content: it goes through the same fit and scale as the
  // tiles inside it, so its caption has to come down with it. Held at its
  // natural size it was 4.4x the service names beside it on a deeply scaled
  // sheet, wrapping to five characters a line and overflowing both its own
  // text block and the zone onto the tiles it is supposed to contain.
  const k = fonts.scale;
  // Allow the title band to grow to two lines instead of clipping a long name,
  // but never past half the zone — a caption taller than the box it names has
  // stopped being a caption.
  const titleH = Math.min(rect.h * 0.5, Math.min(0.56 * k, Math.max(0.24 * k, rect.h * 0.22)));
  // The inset is a margin, not a minimum: floored at 0.4in it measured *wider*
  // than the zone once the zone was under 0.64in, so the text was laid out to
  // a width that does not exist.
  const titleW = Math.max(0.08, Math.min(rect.w * 0.92, rect.w - 0.24 * k));
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
      <Cell N="LineWeight" V="${f(penIn(0.0125, fonts.scale, fonts.penMaxIn))}"/>
      <Cell N="LinePattern" V="2"/>
      <Cell N="Rounding" V="${f(CORNER_ROUNDING_IN * 1.5 * fonts.scale)}"/>
      <Cell N="TxtWidth" V="${f(titleW)}"/>
      <Cell N="TxtHeight" V="${f(titleH)}"/>
      <Cell N="TxtPinX" V="${f(rect.w / 2)}"/>
      <Cell N="TxtPinY" V="${f(rect.h - titleH / 2 - 0.06 * fonts.scale)}"/>
      <Cell N="TxtLocPinX" V="${f(titleW / 2)}"/>
      <Cell N="TxtLocPinY" V="${f(titleH / 2)}"/>
      <Cell N="TxtAngle" V="0"/>
${roundedRectGeometry()}
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${palette.text}"/><Cell N="Size" V="${ff(fonts.zone)}"/><Cell N="Style" V="1"/></Row>
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
  fonts: VisioFonts = NATURAL_FONTS,
): string {
  // The column the tile actually has, and the band that is emitted for it.
  //
  // `Math.max(0.3 * scale, ...)` is a floor that keeps a band usable on an
  // ordinary tile, but on a hairline it hands the text a band WIDER than the
  // shape it sits in — a 0.08in tile was given a 0.30in band. Every later
  // question then got the wrong answer: the wrap was counted in a column the
  // tile does not have, and "is this still a name?" was asked of the floor
  // rather than of the tile. The band is now clamped to the shape, and the
  // decision below is taken on the real column.
  const textColumn = Math.max(0.01, rect.w - 0.12 * fonts.scale);
  const textW = Math.min(rect.w, Math.max(0.3 * fonts.scale, textColumn));
  // Give the label the room it actually needs and let the icon take the rest,
  // so a two-line service name is never clipped and the icon never vanishes.
  //
  // Every inch here is scaled, because the tile it divides up is. Written flat,
  // these floors kept their absolute size while the tile shrank, so the text
  // band reserved a constant 0.16in out of a tile that was 0.78125*scale inches
  // tall: below scale 0.5504 the arithmetic `0.78125s - 0.19 - 0.16 >= 0.08`
  // fails and `showIcon` goes false for *every* tile at once. Adding one
  // service to a 205-service pipeline took all 205 Azure icons off the sheet,
  // while the PowerPoint export of the same diagram drew them; below scale
  // 0.2048 the reserved band was taller than the whole tile. The equivalent
  // floor in the deck is written `0.08 * px` and never had the bug.
  // Measured, not `ceil(width / column)`. The ratio assumes text can break
  // anywhere, so it under-counts a name made of long tokens — and it never saw
  // a hard break at all, which is how a four-paragraph name pasted out of a
  // spreadsheet sized its band for one line. `wrappedLinesIn` breaks between
  // words and honours newlines, the way Visio actually lays the shape out.
  const maxIcon = iconRelId ? Math.min(rect.h * 0.46, rect.w * 0.5, 0.55 * fonts.scale) : 0;
  const minIcon = Math.min(maxIcon, 0.18 * fonts.scale);
  const room = Math.max(0.2 * fonts.scale, rect.h - 0.19 * fonts.scale);
  // What the band is allowed to become. Everything below fits *into* this;
  // nothing is permitted to be drawn past it.
  const bandRoom = Math.max(0.16 * fonts.scale, room - minIcon);
  // What the band may take when the comfortable budget is not enough.
  //
  // `minIcon` reserves 0.18in for the picture, and on a tile of the app's own
  // default proportions that leaves the caption a single line: the name is
  // then CUT, at full-size type, with an icon sitting above it at more than
  // twice the size below which this file would not draw an icon at all. The
  // cross-format rule caught it the day it could see truncation - the deck
  // spelled "Azure Synapse Analytics workspace" and the sheet spelled
  // "Azure Synaps...kspace", from the same nodes, on the same tile.
  //
  // The trade this file already argues for is icon over name, and that stands:
  // the band borrows only what the icon can give up while STILL BEING DRAWN,
  // and only after type has come all the way down to the legibility floor. So
  // the icon shrinks before a letter is deleted, and a letter is deleted only
  // when the icon has nothing left to lend.
  const iconFloor = maxIcon > 0 ? Math.min(maxIcon, 0.085 * fonts.scale) : 0;
  const bandMax = Math.max(bandRoom, room - iconFloor);
  const lineH = 1.3;
  const floorLabel = Math.min(fonts.label, LEGIBLE_IN * fonts.scale);
  // The sub-line yields before the name does. A band with room for one line of
  // type cannot hold a name *and* a SKU, and of the two it is the name that
  // identifies the tile; the deck makes the same call for the same reason.
  // The sub-line yields before the name does. A band with room for one line of
  // type cannot hold a name *and* a SKU, and of the two it is the name that
  // identifies the tile; the deck makes the same call for the same reason.
  //
  // And it is budgeted at ONE line because one line is all it may occupy. The
  // band reserved `meta * 1.4` unconditionally, so on a 0.30in column
  // "Premium - japaneast" wrapped to four lines inside a reservation for one
  // and painted 0.100in out through the bottom of a 1.25in shape. A SKU and a
  // region are a caption: if the tile is too narrow to set one on a line, the
  // tile is too narrow to carry it at all, and the whole string is still on
  // the shape's `Name` attribute and in its shape data, which is the recovery
  // route this file already documents for a cut name.
  const metaOneLine = !!meta && wrappedLinesIn(meta, textW, fonts.meta) <= 1;
  const showsMeta = metaOneLine && floorLabel * lineH + fonts.meta * 1.4 <= bandRoom;
  const metaBand = showsMeta ? fonts.meta * 1.4 : 0;
  //
  // `min(needed, room - minIcon)` was a clamp with nothing behind it. Visio
  // does not clip text to its text block — that is the premise the connector
  // chip was fixed on — so a name needing five lines in a band cut to three
  // simply drew the other two below the tile, through the icon and out of the
  // shape. At the app's own default node size (180x75, and `AzureNode` is not
  // resizable) a realistic 120-character service name overran by 0.271in and
  // escaped the tile by 0.076in, while the 77-character name the corpus
  // happened to exercise cleared it by two thousandths of an inch. The
  // PowerPoint tile has never had this bug: it measures the name, and what
  // will not fit is *cut*, visibly, with an ellipsis.
  //
  // So the name is made to fit rather than assumed to. Type comes down first,
  // as far as the drawing's own legibility floor and no further, because a
  // reader can zoom into small type but cannot recover a letter that was
  // deleted; only then is the name cut. The icon keeps its floor throughout —
  // the deck resolves the same squeeze by dropping the icon instead, and on
  // this input drops three of four, which is the worse of the two trades: the
  // icon is what says which service a tile is.
  let labelFont = fonts.label;
  let label = box.label;
  // Sized by the column the tile has, not only by the height it has.
  //
  // The loop below only shrinks when the name does not FIT, and on a tall
  // narrow tile it always fits: a 0.42 x 1.25in shape has room for eleven
  // lines, so the name kept full-size type and wrapped to nine of them, three
  // glyphs to a line. That is the shape the deck hit in round 53 and fixed by
  // deriving the size from the column as well as the height; doing only half
  // of it here is why the two formats disagreed about this tile at all.
  //
  // Four characters of the drawn size per line, the same bargain the deck
  // strikes, floored at the sheet's own legibility limit so this can shrink
  // type but never make it unreadable.
  const columnLabel = Math.max(floorLabel, Math.min(fonts.label, textW / 4));
  let bandUsed = bandRoom;
  for (const budget of [bandRoom, bandMax]) {
    bandUsed = budget;
    for (let step = 0; step < 6; step += 1) {
      const size = columnLabel - ((columnLabel - floorLabel) * step) / 5;
      labelFont = size;
      const fits = wrappedLinesIn(box.label, textW, size) * size * lineH + metaBand <= budget;
      if (fits) {
        label = box.label;
        break;
      }
      label = fitLabelToLines(
        box.label,
        textW,
        size,
        Math.floor((budget - metaBand) / (size * lineH)),
        wrappedLinesIn,
      );
    }
    if (label === box.label) break;
  }
  const labelLines = wrappedLinesIn(label, textW, labelFont);
  // A tile with no column to set the name in draws no name. Visio wraps text
  // inside the shape and does not clip it, so a 0.08in tile holding 7pt type
  // spells the name one letter per line straight out through the bottom of the
  // shape — the same artefact the deck's chips produced, from the same cause,
  // and the type-to-tile ratio here reached 17.4x what the sheet draws at full
  // size. Nothing is lost by refusing: the whole name stays on the shape's
  // `Name` attribute and in its shape data, so Drawing Explorer and Visio's own
  // search still find it, which is the recovery route this file already
  // documents for a cut name.
  //
  // Two of the widest glyph the name contains, which is the bar the deck now
  // holds itself to, so the two drawings cannot disagree about when a name has
  // stopped being a name.
  //
  // "Four characters" charged a flat em to every glyph. It survived here after
  // the deck replaced it, and the cross-format rule caught the result at once:
  // the same diagram named four services in the .pptx and three in the .vsdx.
  // A user who exports both and finds different services labelled in each has
  // been handed two drawings, not two renderings of one.
  // Measured at the LEGIBILITY FLOOR, not at the size this tile happened to
  // settle on. The question is whether the name can be drawn at all, and the
  // exporter is free to shrink to `floorLabel` before it gives up - so asking
  // it at a larger size refuses a name that the very next step would have set
  // comfortably. The deck asks at its floor, and the two answers diverged on a
  // 0.271in tile: `Camion logistica analisis` was named in the .pptx and drawn
  // on no shape at all in the .vsdx, because one `m` at 0.861 em vetoed it here
  // and nowhere else.
  // Asked about the column the text is ACTUALLY SET IN, which is `textW` - the
  // value written to `TxtWidth` a few lines below, and the column every other
  // question in this function is asked of: the fit loop wraps in it, the line
  // count measures in it, the sub-line tests against it.
  //
  // `textColumn` is that column minus a flat 0.12in inset, and a flat inset
  // against a variable tile is the same defect the audit's own ratio rule had:
  // on a 0.1458in tile it leaves 0.0258in, so the decision was taken about a
  // column 5.6x narrower than the one the shape actually gives the text. The
  // name was refused for want of room the shape had all along, and on a sheet
  // there is no index to recover it from - it survives only in `Name`, which
  // is metadata and is never printed. Four tiles across two corpus scenarios
  // drew a blank box under an icon for exactly this reason.
  const drawsName = drawableInColumn(label, floorLabel * 72, textW);
  if (!drawsName) label = '';
  const drawsMeta = showsMeta && drawsName;
  const neededTextH = drawsName
    ? labelLines * labelFont * lineH + metaBand + 0.05 * fonts.scale
    : 0;
  // Never taller than the tile it sits in. `0.16 * scale` is a floor so a band
  // is never degenerate, but a collapsed 12px node is 0.125in tall — shorter
  // than one line of legible type — and the floor then declared a band half as
  // tall again as the whole shape.
  const textH = Math.min(rect.h, Math.max(0.16 * fonts.scale, Math.min(neededTextH, bandUsed)));
  // Seated 0.06in above the tile's floor, but never pushed out through its
  // ceiling. Positioning the band by its bottom edge alone hung the name of
  // every collapsed tile 0.079in above the shape it names — 85 of them across
  // the corpus, each one a caption floating in the gap over its own box.
  //
  // SCALED, like every other inch in this function. Flat, this offset and the
  // icon's flat 0.07in below stayed put while the tile shrank around them, and
  // the two blocks overlapped by exactly `0.13 - 0.19 * scale` inches - so
  // below scale 0.6842 the icon is drawn ON TOP OF the name. A 260-service
  // pipeline scales to 0.436 and overlapped on 100% of tiles; at 440 services
  // the icon was entirely inside the text block on 340 of them. Scaled, the
  // gap is `0.06 * scale` and cannot close.
  const txtPinY = textH >= rect.h
    ? rect.h / 2
    : Math.min(0.06 * fonts.scale + textH / 2, rect.h - textH / 2);
  const iconSizeIn = maxIcon > 0 ? Math.max(0, Math.min(maxIcon, room - textH)) : 0;
  const showIcon = iconRelId !== null && iconSizeIn >= 0.08 * fonts.scale;
  const iconChild = showIcon
    ? `
        <Shape ID="${ids.icon}" NameU="Icon.${ids.icon}" Type="Foreign" LineStyle="0" FillStyle="0" TextStyle="0">
          <Cell N="PinX" V="${f(rect.w / 2)}"/>
          <Cell N="PinY" V="${f(rect.h - iconSizeIn / 2 - 0.07 * fonts.scale)}"/>
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
  const metaColor = showsMeta ? readableTextOn('#64748B', palette.fill) : '#64748B';
  // The tile name is drawn in the category's own accent, which on two of the
  // sixteen palettes is unreadable on that category's fill — ai + machine
  // learning at 3.93:1 and identity at 2.49:1. PowerPoint draws the same words
  // at a flat #1F2937, so the two exporters were disagreeing about what is
  // legible. Resolve the accent against the tile the same way the sub-line is.
  const nameColor = readableTextOn(palette.text, palette.fill);
  const characterRows = drawsMeta
    ? `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${nameColor}"/><Cell N="Size" V="${ff(labelFont)}"/></Row>
        <Row IX="1"><Cell N="Font" V="1"/><Cell N="Color" V="${metaColor}"/><Cell N="Size" V="${ff(fonts.meta)}"/></Row>`
    : `        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${nameColor}"/><Cell N="Size" V="${ff(labelFont)}"/></Row>`;
  const textBody = drawsMeta
    ? `<cp IX="0"/>${esc(label)}\n<cp IX="1"/>${esc(meta)}`
    : esc(label);

  // The shape keeps the *whole* name, always. Drawing Explorer lists it, the
  // shape data carries it, and search finds it — so a name the tile had to cut
  // is recoverable inside Visio rather than lost with the export.
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
      <Cell N="TxtPinY" V="${f(txtPinY)}"/>
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
          <Cell N="LineWeight" V="${f(penIn(0.0125, fonts.scale, fonts.penMaxIn))}"/>
          <Cell N="LinePattern" V="1"/>
          <Cell N="Rounding" V="${f(CORNER_ROUNDING_IN * fonts.scale)}"/>
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
  fonts: VisioFonts = NATURAL_FONTS,
  /** The edge this arrow draws, so the shape carries a name a reader knows. */
  edgeId?: string,
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
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="${CONNECTOR_TEXT}"/><Cell N="Size" V="${ff(fonts.connector)}"/></Row>
      </Section>
      <Section N="Paragraph">
        <Row IX="0"><Cell N="HorzAlign" V="1"/></Row>
      </Section>`
    : '';
  // Faded (optional) connectors get line transparency so they read as secondary.
  const transCell = opacity < 1 ? `\n      <Cell N="LineColorTrans" V="${f(1 - opacity)}"/>` : '';

  return `    <Shape ID="${id}" NameU="Connector.${id}"${edgeId ? ` Name="edge-${esc(edgeId)}"` : ''} Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
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
      <Cell N="LineWeight" V="${f(penIn(0.0125, fonts.scale, fonts.penMaxIn))}"/>
      <Cell N="LinePattern" V="${linePattern}"/>${transCell}
      <Cell N="Rounding" V="${f(0.0625 * fonts.scale)}"/>
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

/**
 * The 7pt floor, in the inches Visio's `Size` cell is written in. Both
 * exporters promise it and the export audit enforces it, so a badge that
 * shrinks to fit a gap may take its padding away but never its legibility.
 * 0.0973 rather than 0.0972 because 0.0972in is 6.998pt — a hair under.
 */
const MIN_BADGE_FONT_IN = 0.0973;

/**
 * The floor the badge digits may shrink to *on this sheet*.
 *
 * The shrink ladder is a natural-scale mechanism — it buys a badge a seat
 * between two tiles by giving up padding — and against it a fixed 7pt floor is
 * right. Against `fonts.scale` it is not: a scaled sheet draws everything,
 * including the service names, below 7pt on purpose, and holding the badge
 * alone at 7pt does not rescue it, it detaches it from the disc that backs it.
 * At `scaled-zone-row`'s 0.236 the digits came out 2.78x the width of their own
 * disc, so about two thirds of every number was white ink on white paper — a
 * dark speck with an invisible smear across it, and 479 of them to repair by
 * hand. `fontsForScale` already settled this argument for the other type on the
 * sheet: proportion is the only property worth keeping at that size. So the
 * floor may never raise the badge above what the sheet scale asks for.
 */
function badgeFontFloorIn(fonts: VisioFonts): number {
  return Math.min(MIN_BADGE_FONT_IN, fonts.badge);
}

/**
 * The smallest disc that still holds its own number at the legibility floor.
 * A badge is mostly padding, so squeezing it into a gap between two tiles is
 * mostly a matter of giving that padding up; what cannot be given up is the
 * digits, and a three-digit step needs visibly more room than a one-digit one.
 *
 * Measured on the diagonal, because the digits do not sit on the diameter. A
 * line of type centred in a circle occupies a chord, and the half-chord at the
 * height of the glyphs is shorter than the radius — so solving `width <= d`
 * leaves the ends of a two-digit number outside the disc even though the
 * arithmetic says it fits. The old form cleared its own test by 1.4%, which is
 * to say it did not clear the real one at all.
 *
 * And a tenth of the disc is kept as a ring, rather than a flat 0.02in. The
 * flat pad is a tenth of a one-digit disc and a twenty-fifth of a two-digit
 * one, so the moment a ceiling actually pushed a two-digit badge down to this
 * floor, all fourteen of them drew their number across 96% of the disc, into
 * the rim: white ink with nothing dark behind it. The ring has to scale with
 * the number it surrounds.
 */
function badgeMinDiameterIn(stepNumber: number, fonts: VisioFonts): number {
  const pt = badgeFontFloorIn(fonts);
  const digits = String(Math.max(1, Math.abs(Math.trunc(stepNumber)))).length;
  return Math.max(pt * 1.15, Math.hypot(digits * pt * 0.66, pt * 0.7) / 0.9);
}

/**
 * Grow a drawing whose tiles are too small to carry the callout it asks for.
 *
 * PowerPoint answers this by splitting: `markableTileWIn` raises the planner's
 * target to `floor / 0.55` and the deck spends windows until the tile reaches
 * it. A sheet has no windows to spend, and `fontsForScale` returns natural
 * fonts at or above 0.999, so this exporter could only ever shrink. On a
 * drawing of 14px tiles numbered past a hundred that left a disc 156% of the
 * service it points at - half again wider than the thing it calls out - while
 * the same input came out of PowerPoint at 66% and was reported. The silent
 * format was the one that could not split its way out.
 *
 * What a sheet has instead is paper. A Visio page is sized to its drawing, so
 * scaling the whole drawing up costs nothing but inches, and inches are the
 * one resource this format has. Everything moves together, so no relationship
 * in the picture changes: only the sheet it is printed on gets bigger.
 *
 * Bounded three ways. It never shrinks (`k >= 1`), so an ordinary drawing is
 * returned untouched and byte-identical - a 150px tile already clears the bar
 * for a three-digit callout by a factor of four. It never asks for more than
 * the callout needs. And it never grows the drawing past `MAX_USEFUL_PAGE_IN`,
 * which is the same limit the outlier trim uses, so a sheet that was already
 * at the edge of useful is left alone rather than pushed over it.
 *
 * The narrowest tile that carries a numbered hop, not the median of all
 * services. The median was borrowed from the PowerPoint planner, where chasing
 * an extremum costs SLIDES and `probe-whitespace` proves that trade is bad -
 * four slivers there drag their 160px neighbours to 2.3in each and put one
 * tile on a slide. On a sheet the cost is inches, `roomK` already bounds them,
 * and the argument does not transfer. Measured, the median declined an
 * available move on an ordinary seven-node diagram: six 150px services and one
 * 14px private DNS zone in the numbered flow left two discs at 77% of the zone
 * they point at with `k = 1`, when `k = 1.395` was free - an 11.1 x 8.5in sheet
 * becoming 15.5 x 11.9in against a 60in bound.
 *
 * Only tiles that carry a hop, because a callout sits on an arrow: a sliver
 * with no numbered edge touching it has no disc to be dwarfed by, and letting
 * it set the scale would be chasing an extremum for nothing.
 */
function magnifiedForCallouts(
  boxes: Map<string, ExportBox>,
  edges: readonly Edge[],
): Map<string, ExportBox> {
  const plan = calloutMagnificationPlan(boxes, edges);
  if (plan.k <= 1.001) return boxes;
  const bounds = computeBounds(boxes.values());
  const out = new Map<string, ExportBox>();
  for (const [id, box] of boxes) {
    out.set(id, {
      ...box,
      x: bounds.minX + (box.x - bounds.minX) * plan.k,
      y: bounds.minY + (box.y - bounds.minY) * plan.k,
      w: box.w * plan.k,
      h: box.h * plan.k,
    });
  }
  return out;
}

/**
 * What the magnifier will do and why, without doing it.
 *
 * Split out so that `wantedK` and `roomK` survive the decision. The audit has
 * to tell a callout that is disproportionate because the sheet ran out of paper
 * from one that is disproportionate because the exporter left paper unused, and
 * a plan that reports only the `k` it settled on cannot: reverting the target
 * from the narrowest badged tile to the median made the magnifier decline an
 * available move entirely, and an exemption keyed on the result rather than on
 * the bound went blind on exactly that.
 */
function calloutMagnificationPlan(
  boxes: Map<string, ExportBox>,
  edges: readonly Edge[],
): { k: number; wantedK: number; roomK: number; barIn: number } {
  const inert = { k: 1, wantedK: 1, roomK: Infinity, barIn: 0 };
  if (boxes.size === 0) return inert;
  let widestStep = 0;
  const badged = new Set<string>();
  for (const edge of edges) {
    const step = Number((edge as unknown as { data?: { stepNumber?: unknown } }).data?.stepNumber);
    if (!Number.isFinite(step) || step <= 0) continue;
    if (step > widestStep) widestStep = step;
    badged.add(edge.source);
    badged.add(edge.target);
  }
  if (widestStep <= 0) return inert;
  const services = [...boxes.entries()]
    .filter(([id, box]) => box.kind === 'service' && box.w > 0 && badged.has(id))
    .map(([, box]) => box);
  if (services.length === 0) return inert;
  const typicalW = Math.min(...services.map((box) => box.w));
  if (!typicalW || typicalW <= 0) return inert;
  const barIn = badgeMinDiameterIn(widestStep, NATURAL_FONTS) / 0.55;
  const bounds = computeBounds(boxes.values());
  const spanPx = Math.max(bounds.maxX - bounds.minX, bounds.maxY - bounds.minY, 1);
  const roomK = (MAX_USEFUL_PAGE_IN * PX_PER_INCH) / spanPx;
  const wantedK = (barIn * PX_PER_INCH) / typicalW;
  const k = Math.min(wantedK, roomK);
  return { k: Number.isFinite(k) && k > 1.001 ? k : 1, wantedK, roomK, barIn };
}

/**
 * What {@link magnifiedForCallouts} did to a drawing, and whether the sheet ran
 * out of paper before it was done.
 *
 * Exported for the export audit, which has to tell a callout that is
 * disproportionate because there was no move left from one that is
 * disproportionate because the exporter did not make the move it had. Asked
 * rather than replicated, for the reason `calloutBarClampedFor` is: a second
 * copy of this arithmetic in the gate would drift from this one, and did.
 */
export function calloutMagnificationFor(
  nodes: readonly Node[],
  edges: readonly Edge[],
): { k: number; paperBound: boolean } {
  const plan = calloutMagnificationPlan(
    compactEmptyGutters(collectExportBoxes([...nodes])),
    edges,
  );
  return { k: plan.k, paperBound: plan.wantedK > plan.roomK + 1e-9 };
}

/**
 * The widest a badge for this route may be drawn, from the two tiles it is
 * actually drawn between.
 *
 * Three sheet-wide statistics were tried here and all three failed, which was
 * the lesson: the property that kept breaking was not which statistic, it was
 * that a statistic over the sheet answers a question about one connector. The
 * minimum let a parked sliver cap every disc. The median failed the parity
 * flip. The median over badged endpoints failed as soon as a sheet carried two
 * numbered chains - numbering six sensors cut an unrelated cloud pipeline's
 * badges by 43%, to 8.8% of the tiles they number, purely by moving the median
 * of a set they had no part in. A badge is drawn on one arrow between two
 * tiles, and those two tiles are the whole of what it can dwarf.
 */
function badgeCeilingIn(
  route: { sourceId: string; targetId: string },
  boxes: Map<string, { w: number }>,
): number {
  const a = boxes.get(route.sourceId)?.w;
  const b = boxes.get(route.targetId)?.w;
  const ends = [a, b].filter((w): w is number => typeof w === 'number' && w > 0);
  if (ends.length === 0) return Infinity;
  // The narrower end, not the average. A disc that is proportionate to one
  // tile and swamps the other is still swamping a tile.
  return (Math.min(...ends) / PX_PER_INCH) * 0.55;
}

/**
 * The diameter a badge for this step will actually be drawn at.
 *
 * One expression, because the layout reserves room for a badge and the
 * renderer draws one, and a reservation that is not an upper bound on the
 * thing reserved is a collision waiting for a denser fan. The reservation
 * used to be `min(natural, ceiling)` while the draw was `max(floor,
 * min(natural, ceiling))`, so on a sheet whose ceiling fell under the floor
 * the disc came out 39.5% wider than the space held for it, and 104% wider
 * with two-digit steps.
 */
function badgeDiameterIn(stepNumber: number, fonts: VisioFonts, maxIn = Infinity): number {
  return Math.max(
    badgeMinDiameterIn(stepNumber, fonts),
    Math.min(STEP_BADGE_IN * fonts.scale, maxIn),
  );
}

function stepBadgeXml(  id: number,
  centre: Point,
  stepNumber: number,
  fonts: VisioFonts = NATURAL_FONTS,
  /** The edge this callout numbers, so the shape carries a name a reader knows. */
  edgeId?: string,
  /**
   * Diameter, when the natural one will not fit. On a dense sheet the only gap
   * between two tiles can be narrower than a badge, and the placement search
   * then has nowhere to put it but out on the next hop's arrow — where the
   * reader reads the number as that hop's. A smaller disc that sits on its own
   * arrow says the right thing; a full-size one parked on a stranger does not.
   */
  diameterIn?: number,
  /** The ceiling from this badge's own two endpoint tiles. */
  ceilingIn = Infinity,
): string {
  // A badge is drawn on the arrows, between the tiles, so it scales with them.
  // Held at its natural size it was 109% of a whole service tile once the
  // sheet was down to a seventh: a callout larger than the thing it calls out.
  const natural = badgeDiameterIn(stepNumber, fonts, ceilingIn);
  const floor = badgeMinDiameterIn(stepNumber, fonts);
  const d = diameterIn !== undefined && diameterIn > 0
    ? Math.min(natural, Math.max(floor, diameterIn))
    : natural;
  // The digits come down with the disc so they do not spill out of it, but
  // never below the floor both exporters promise: an illegible number in the
  // right place is no better than a legible one in the wrong place.
  const size = Math.max(badgeFontFloorIn(fonts), fonts.badge * Math.min(1, d / natural));
  return `    <Shape ID="${id}" NameU="StepBadge.${id}"${edgeId ? ` Name="step-${esc(edgeId)}"` : ''} Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
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
      <Cell N="LineWeight" V="${f(penIn(0.0125, fonts.scale, fonts.penMaxIn))}"/>
      <Cell N="LinePattern" V="1"/>
      <Section N="Character">
        <Row IX="0"><Cell N="Font" V="1"/><Cell N="Color" V="#FFFFFF"/><Cell N="Size" V="${ff(size)}"/><Cell N="Style" V="1"/></Row>
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
function legendTextXml(id: number, x: number, y: number, width: number, text: string, height = 0.18, name = ''): string {
  return `    <Shape ID="${id}" NameU="LegendText.${id}"${name ? ` Name="${esc(name)}"` : ''} Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PinX" V="${f(x + width / 2)}"/>
      <Cell N="PinY" V="${f(y)}"/>
      <Cell N="Width" V="${f(width)}"/>
      <Cell N="Height" V="${f(height)}"/>
      <Cell N="LocPinX" V="${f(width / 2)}"/>
      <Cell N="LocPinY" V="${f(height / 2)}"/>
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
 * How many lines the text actually takes, wrapped the way a renderer wraps it.
 *
 * `ceil(width / column)` is only a lower bound: Latin prose breaks between
 * words, so a run that will not fit on the current line abandons the rest of
 * that line and starts the next one. On the workflow band that lower bound
 * budgeted four lines for prose that drew five, and six for unbreakable
 * resource names that drew eight — every row overrunning its neighbours by
 * roughly a tenth of an inch, invisibly, because the guard used the same ratio.
 */
export function wrappedLinesIn(text: string, widthIn: number, fontSizeIn: number): number {
  if (!text) return 1;
  // Visio wraps in `TxtWidth` and in nothing else. A floor here budgeted the
  // lines for a column the shape does not have: on a 0.1042in tile the count
  // was taken at 0.2in, so the band was sized for 9 lines and the renderer set
  // 14, painting 0.357in of the name out through the bottom of the tile. The
  // band above this had exactly the same floor removed for exactly this reason;
  // the defect survived one level down in the line counter.
  //
  // The guard that is actually needed is against a non-positive column. An
  // overlong run already breaks a glyph at a time in `wrapOneLineIn`, and that
  // loop advances on every glyph, so a column narrower than one glyph charges
  // one line per glyph and terminates.
  const column = Math.max(1e-4, widthIn);
  // Hard breaks first: Visio renders a raw newline in `<Text>` as a paragraph
  // break, and counting it as ordinary whitespace budgeted two lines for rows
  // that drew three.
  return String(text).split(/\r\n|\r|\n/)
    .reduce((total, paragraph) => total + wrapOneLineIn(paragraph, column, fontSizeIn), 0);
}

/**
 * One paragraph's worth of wrapping.
 *
 * Deliberately the same tokenisation as the PowerPoint exporter's
 * `wrappedLineCount`: whitespace stays attached to the run in front of it and
 * every space is charged. The earlier version collapsed a whitespace *run* to a
 * single space and never charged a trailing space at end of line, which made it
 * count fewer lines than the PowerPoint side on 15.7% of mixed strings — up to
 * three fewer. Two exporters drawing the same sentence must agree; it is the
 * audit's copy that is supposed to be independent.
 */
function wrapOneLineIn(text: string, column: number, fontSizeIn: number): number {
  if (!text) return 1;
  const runs = text.split(/(?<=[\s\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6])/);
  let lines = 1;
  let used = 0;
  for (const run of runs) {
    const w = estimateTextWidthIn(run, fontSizeIn);
    // The fit test uses visible ink; the run-final spaces hang past the column
    // the way a renderer hangs them. See the PowerPoint exporter's copy.
    const visible = w - trailingWhitespaceIn(run, 72) * fontSizeIn;
    if (used > 0 && used + visible > column) {
      lines += 1;
      used = 0;
    }
    if (w > column) {
      // A single run wider than the box breaks inside itself, one CHARACTER at
      // a time. `ceil(w / column)` assumes a word packs exactly a columnful per
      // line, which is only true if a break may fall part-way through a glyph.
      // Breaks fall between glyphs, so the ratio is a lower bound and never the
      // count - the same defect the deck's copy carried, in the third copy of
      // this algorithm in the repo.
      let lineUsed = used;
      for (const glyph of run) {
        const gw = estimateTextWidthIn(glyph, fontSizeIn);
        if (lineUsed > 0 && lineUsed + gw > column) {
          lines += 1;
          lineUsed = 0;
        }
        lineUsed += gw;
      }
      used = lineUsed;
      continue;
    }
    used += w;
  }
  return Math.max(1, lines);
}

/**
 * How tall one workflow row has to be to hold the sentence written in it.
 *
 * The band used to draw every row at a fixed pitch and give the sentence a
 * 0.18in text block whatever it said. A description too long for its column
 * wraps, and Visio draws the extra lines outside the block, straight through
 * the row beneath: on an 11in page a 76-character step is three lines in a
 * 0.26in pitch, so every row in the band overran the next, all the way down.
 */
function workflowRowHeightIn(description: string, colW: number): number {
  // The column the row is actually DRAWN in, with no floor under it. The floor
  // was a fourth copy of the `Math.max(0.2, widthIn)` defect: it measured the
  // wrap in a column wider than the one the text is given, so a narrow band
  // would reserve fewer lines than the renderer sets and every row would overrun
  // the next - which is the exact failure the comment above says this function
  // exists to prevent. Unreachable today, because `workflowColumnWidthIn` never
  // returns below 1.2in, but it was one call site away from being live.
  const textW = colW - 0.6;
  const lines = wrappedLinesIn(description, textW, LEGEND_FONT_IN);
  return Math.max(WORKFLOW_ROW_IN, lines * LEGEND_FONT_IN * 1.35 + 0.08);
}

/**
 * The median sentence's wrapped line count, for a given column width.
 *
 * The MEDIAN, not the worst. One long sentence among twelve is the author's
 * sentence and no evidence about the column; when HALF the band wraps past
 * three lines the column itself has stopped being prose and become a ribbon.
 */
function workflowMedianLines(entries: WorkflowListEntry[], colW: number): number {
  const lines = entries
    .map((entry) => wrappedLinesIn(entry.description, colW - 0.6, LEGEND_FONT_IN))
    .sort((a, b) => a - b);
  return lines.length === 0 ? 0 : lines[Math.floor(lines.length / 2)];
}

/**
 * Does any single sentence fold more than `multiple` times as far as one full
 * column would have folded it?
 *
 * A median cannot see the rows the drawing exists to explain. A workflow of 48
 * one-word acknowledgements and 3 real paragraphs holds its median at ONE
 * however narrow the column gets - the terse steps were never going to wrap at
 * any width, so they are not evidence about the column - while the sentences
 * that carry the architecture are folded into 9-line ribbons. The median stays
 * quiet through all of it. So the tail is measured too, per entry and against
 * that same entry unsplit, which keeps a long sentence measured against what it
 * would have been rather than against a constant no page can satisfy.
 */
function workflowTailShredded(
  entries: WorkflowListEntry[],
  colW: number,
  unsplitColW: number,
  multiple: number,
): boolean {
  return entries.some((entry) => {
    const alone = wrappedLinesIn(entry.description, unsplitColW - 0.6, LEGEND_FONT_IN);
    const split = wrappedLinesIn(entry.description, colW - 0.6, LEGEND_FONT_IN);
    return split > Math.max(MAX_WORKFLOW_MEDIAN_LINES, alone * multiple);
  });
}

/** The tallest column, which is the height the band has to reserve. */
function workflowStackIn(entries: WorkflowListEntry[], columns: number, colW: number): number {
  const cols = Math.max(1, columns);
  const perColumn = Math.ceil(entries.length / cols);
  let tallest = 0;
  for (let c = 0; c < cols; c += 1) {
    let stack = 0;
    for (let i = c * perColumn; i < Math.min((c + 1) * perColumn, entries.length); i += 1) {
      stack += workflowRowHeightIn(entries[i].description, colW);
    }
    tallest = Math.max(tallest, stack);
  }
  return tallest;
}

/** How wide the sheet will draw the band, for a page of a given width. */
function workflowPanelWidthIn(pageWidthIn: number, columns: number): number {
  return Math.min(Math.max(pageWidthIn - 0.7, 2.4), 7.5 * Math.max(1, columns));
}

/**
 * Emit the numbered workflow narration beside the drawing.
 *
 * An Azure Architecture Center diagram never shows a numbered callout without
 * the sentence it refers to; on its own the badge is just an unexplained digit.
 * Rows stack downward from the top-left of the panel.
 */
/**
 * The sheet's answer to the deck's "Service names" slide.
 *
 * PowerPoint can afford to shorten a caption because it prints an index behind
 * the drawing: "Names shortened on the drawing, in full." Visio has one page,
 * so a name the tile could not hold used to exist nowhere a reader can see -
 * it survived only in the shape's Name attribute and its shape data, which are
 * handles for automation and are never put on paper. That asymmetry is what
 * made the same defect High on the sheet and merely untidy in the deck, and it
 * is why the two exports of one diagram could name different services.
 *
 * Columns rather than one long list, because a hairline-heavy drawing can
 * shorten dozens of names and a single column would run off the top of the page.
 *
 * ON ITS OWN PAGE, which is what the deck does with it and the only arrangement
 * that cannot go wrong. Drawn on the sheet it has to be reserved for, and the
 * reservation is circular: how many names get shortened depends on the scale,
 * and the scale depends on how much room the panel takes. Reserving the worst
 * case costs the drawing 42% of an 8.5in page on the common run where nothing
 * is shortened at all, and reserving nothing is what buried 20 of 48 tiles
 * under an opaque white box that is emitted last.
 */
function buildServiceNamePanel(
  startId: number,
  entries: Array<{ authored: string; drawn: string }>,
  heading = 'Service names shortened on the drawing, in full',
): { shapes: string[]; nextId: number; widthIn: number; heightIn: number } {
  const shapes: string[] = [];
  let id = startId;
  // MEASURE THE ROW, THEN SIZE THE COLUMN TO IT. This was the one text
  // decision in the file that was a constant: colW was 3.4in whatever the row
  // said, and legendTextXml emits no TxtWidth, so Visio wraps at the shape box
  // and a row too long for 3.16in silently became two lines of 0.135in inside a
  // 0.2in box. Rows are pitched at exactly rowH, so the spill had nowhere to go
  // but through the neighbouring row: two services' names drawn through each
  // other, in the panel whose whole purpose is that a shortened name stays
  // readable. Making the row twice as long for ASK-59-C (stub = full name) put
  // the longest string in the file in the only unmeasured box.
  //
  // A name that was never shortened has no stub to look up, and printing
  // `Azure Front Door  =  Azure Front Door` 700 times turns the one page that
  // is supposed to rescue the sheet into a wall of doubled text. Those rows
  // arrive here when the drawing was reduced so far that its labels stopped
  // being readable rather than because any single tile was too narrow, so the
  // full name on its own is the whole content of the row.
  const rowText = entries.map((entry) => {
    if (entry.drawn === entry.authored) return entry.authored;
    return `${entry.drawn === '' ? '(not drawn)' : entry.drawn}  =  ${entry.authored}`;
  });
  const PAD_IN = 0.24;
  // A cap, because one pathological name should not set the width of every
  // column. Rows past it wrap, and the pitch below grows to carry the wrap.
  const MAX_INDEX_COL_IN = 6.0;
  // ROOM OVER THE MEASUREMENT, not room equal to it. `textW` came out as
  // `widest` exactly, so the longest row in the panel was given a text box the
  // precise width of its own estimated ink and fitted on one line only if the
  // renderer's measurement agreed with this file's to the last fraction. It
  // does not have to: the estimate is a model of one font, Visio measures the
  // font it actually loaded, and the two models already in this repo disagree
  // at the boundary (one tests `used + visible > column`, one `>=`). A Thai row
  // estimated at 2.2581in in a 2.2581in box was one line here and two lines to
  // the audit, inside a 0.2000in single-line pitch, drawn straight through the
  // row above - the exact overprint the measurement above was added to stop.
  // Below 2.4in the floor supplied this slack by accident, which is why only
  // scripts with no spaces to wrap at ever reached the case.
  const INDEX_FIT_SLACK = 1.03;
  const widest = rowText.reduce((w, row) => Math.max(w, estimateTextWidthIn(row, LEGEND_FONT_IN)), 0) * INDEX_FIT_SLACK;
  const colW = Math.min(MAX_INDEX_COL_IN, Math.max(2.4, widest + PAD_IN));
  const textW = colW - PAD_IN;
  const tallest = rowText.reduce((n, row) => Math.max(n, wrappedLinesIn(row, textW, LEGEND_FONT_IN)), 1);
  const rowH = Math.max(0.2, tallest * LEGEND_FONT_IN * 1.35 + 0.06);
  // The page is sized to the index, not the index to the page, so there is no
  // clamp to get wrong and nothing to overflow. Kept near a printable shape by
  // filling a column to about a Letter page before starting the next one.
  const fillsALetterPage = Math.max(1, Math.min(entries.length, Math.floor((10.0 - 1.0) / rowH)));
  // AND THEN GROW THE COLUMN UNTIL THE PAGE FITS.
  //
  // Filling to about a Letter page is a readability preference; staying under
  // Visio's 200in page limit is not optional, and only page 1 was ever held to
  // it. Page 2's width is `colW * ceil(rows / perColumn)`, `colW` caps at 6.0in
  // for names past about 125 characters, `perColumn` falls to 19 when those
  // names wrap to three lines, and NOTHING bounded the row count. That was
  // survivable while the index listed only the names a tile had cut. It stopped
  // being survivable when the index started listing every name on a sheet whose
  // type is too small to read: 900 services with a hundred long names went from
  // 100 rows and 24.70in to 900 rows and 288.70in, and Visio refuses to open the
  // file at all - so the reader loses the DRAWING as well, from the page that
  // exists to rescue the names.
  //
  // Trade the axis Visio has room on for the one it does not: a taller column
  // costs scrolling, a wider page costs the file. The same 900 rows come to 33
  // columns of 28 and a 198.70x14.06in page.
  const maxColumns = Math.max(1, Math.floor((MAX_VISIO_PAGE_IN - 0.7) / colW));
  const perColumn = Math.max(fillsALetterPage, Math.ceil(entries.length / maxColumns));
  const columns = Math.max(1, Math.ceil(entries.length / perColumn));
  const rows = Math.min(entries.length, perColumn);
  const boxW = colW * columns;
  const boxH = rowH * rows + 0.34;
  const widthIn = boxW + 0.7;
  const heightIn = boxH + 0.7;
  const originX = 0.35;
  const originY = 0.35;
  shapes.push(`    <Shape ID="${id++}" NameU="ServiceNames.${startId}" Type="Shape" LineStyle="0" FillStyle="0" TextStyle="0">
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
  shapes.push(legendTextXml(id++, originX + 0.12, originY + boxH - 0.18, boxW - 0.24, heading));
  entries.forEach((_entry, index) => {
    const column = Math.floor(index / perColumn);
    const row = index % perColumn;
    // BOTH HALVES. The stub is what the reader sees on the drawing, so it is
    // the key they look the row up by; printing only the full name leaves them
    // matching 30-character strings against two letters by eye.
    shapes.push(legendTextXml(
      id++,
      originX + 0.12 + column * colW,
      originY + 0.14 + row * rowH,
      textW,
      rowText[index],
      rowH,
      `service-name-${index}`,
    ));
  });
  return { shapes, nextId: id, widthIn, heightIn };
}

function buildWorkflowPanel(
  startId: number,
  entries: WorkflowListEntry[],
  originX: number,
  topY: number,
  width: number,
  columns = 1,
  minHeightIn = 0,
): { shapes: string[]; nextId: number } {
  const shapes: string[] = [];
  let id = startId;
  const cols = Math.max(1, columns);
  const perColumn = Math.ceil(entries.length / cols);
  const colW = width / cols;
  // THE PANEL FILLS WHAT THE PAGE RESERVED FOR IT.
  //
  // The reservation counts every labelled edge as though its wording will be
  // handed to the row, because whether a label is muted is not decided until
  // the arrows are routed - which happens after the page has been sized. The
  // panel appends the wording only for the edges actually muted. So the two are
  // measured over different sentences and the reservation is the larger by
  // construction; the difference is not slack the drawing can have back,
  // because the drawing was already fitted against the reservation.
  //
  // Left alone, that difference came out as a hole in the middle of the sheet.
  // Lengthening ONLY the connector label - same nodes, same steps, same
  // descriptions - took a 40-service sheet from 17.09in to 26.39in for a
  // drawing spanning 9.11in, and 3.51in of it was blank paper between the
  // drawing and the band.
  const naturalH = workflowStackIn(entries, cols, colW) + 0.34;
  const boxH = Math.max(naturalH, minHeightIn);
  // Spread as LEADING BETWEEN ROWS, never as a margin under the last one. A
  // deeper empty footer is the same defect moved indoors, and the audit says so
  // ("the workflow band reserves Xin below its last row"). Charged against the
  // tallest column, which is the one that sets `boxH`, so that column still
  // finishes flush with the bottom of the panel and the shorter columns keep
  // the ragged bottom they already had.
  const rowsInTallest = Math.max(1, Math.min(perColumn, entries.length));
  const leadIn = Math.max(0, boxH - naturalH) / rowsInTallest;
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
  // Each row starts where the one above it ended, so a sentence that wraps
  // pushes the rest of its column down instead of being drawn over it.
  const cursor = new Array(cols).fill(0);
  entries.forEach((entry, index) => {
    const column = Math.min(cols - 1, Math.floor(index / perColumn));
    const rowH = workflowRowHeightIn(entry.description, colW) + leadIn;
    const rowTop = originY + boxH - 0.34 - cursor[column];
    cursor[column] += rowH;
    const colX = originX + column * colW;
    shapes.push(legendTextXml(id++, colX + 0.14, rowTop - 0.09, 0.3, `${entry.step}.`));
    shapes.push(legendTextXml(id++, colX + 0.46, rowTop - rowH / 2, colW - 0.6, entry.description, rowH, `workflow-text-${entry.step}`));
  });
  return { shapes, nextId: id };
}

function pagesXml(
  pageWidthIn: number,
  pageHeightIn: number,
  title: string,
  index?: { widthIn: number; heightIn: number },
): string {
  // The index gets its OWN PAGE, which is what the deck does with it and the
  // only arrangement that cannot go wrong. Put on the drawing it has to be
  // reserved for, and the reservation is circular - how many names get
  // shortened depends on the scale, which depends on how much room the panel
  // takes. Reserving the worst case costs the drawing 42% of an 8.5in page on
  // the common run where nothing is shortened at all. On its own page it
  // reserves nothing, buries nothing, and is sized to its own contents.
  const indexPage = index
    ? `
  <Page ID="1" NameU="Service names" Name="Service names" ViewScale="-1" ViewCenterX="${f(index.widthIn / 2)}" ViewCenterY="${f(index.heightIn / 2)}">
    <PageSheet LineStyle="0" FillStyle="0" TextStyle="0">
      <Cell N="PageWidth" V="${f(index.widthIn)}"/>
      <Cell N="PageHeight" V="${f(index.heightIn)}"/>
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
    </PageSheet>
    <Rel r:id="rId2"/>
  </Page>`
    : '';
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
  </Page>${indexPage}
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
  // Same narration the deck gets: only one hop between a given pair of services
  // is ever given a step number, so the other members of a fan carry a callout
  // that the panel never explains — or, once the fan drops its wording, say
  // nothing at all.
  const narrated = narrateEdgeCallouts(edges);
  const legendEntries = usedConnectionLegend(edges);
  // A numbered workflow gets its own band across the top of the sheet, so the
  // prose never lands on the drawing the way an overlaid panel would. It is
  // measured HERE, before the drawing is fitted, because it is page height the
  // drawing does not get: the fit left exactly 0.5in of headroom under Visio's
  // limit and a twenty-step workflow is 5.7in, so a tall architecture with a
  // workflow was written at 202in and Visio refused to open it at all.
  const workflowEntries = workflowListFromEdges(narrated);
  // Sized against the longest the rows can possibly get, not the shortest.
  //
  // A muted connector hands its wording to its workflow row, in a parenthesis
  // appended when the band is drawn — long after the band's reservation was
  // measured from the authored sentences. Muting is decided by label placement,
  // which needs the page, which needs the band, so the exact text is not
  // knowable here. What *is* knowable is the ceiling: muting only ever hands a
  // step the label of its own connector, so reserving as though every numbered
  // connector were muted can only over-reserve. It used to under-reserve, and
  // the panel is opaque white drawn last: on a 3x3 grid with a 20-arrow fan it
  // grew 1.8in past its reservation and painted out six of the nine services.
  //
  // Restricting this to bundled arrows looks tempting and is wrong: a lone
  // arrow's label is muted too, and five service tiles went under the band the
  // moment the reservation stopped counting them.
  const handableWording = new Map<number, string>();
  for (const edge of narrated) {
    const step = (edge.data as { stepNumber?: number } | undefined)?.stepNumber;
    if (Number.isInteger(step) && edge.label && !handableWording.has(step as number)) {
      handableWording.set(step as number, String(edge.label));
    }
  }
  const reservedEntries = workflowEntries.map((entry) => {
    const handed = handableWording.get(entry.step);
    return handed ? { ...entry, description: `${entry.description}（${handed}）` } : entry;
  });
  // Gutters are closed on BOTH sides of the magnifier, for two independent
  // reasons.
  //
  // Before: `calloutMagnificationPlan` divides the room left on the page by the
  // span it is asked to cover, and that span still contained the voids this
  // very line deletes. Above roughly 5754px of separation the quotient fell
  // under 1 and the magnifier declined a free enlargement, shipping a 14.8in
  // sheet with 45in of the 60in budget unspent. Monotonically wrong: a 700px
  // gap planned k=1.395, 4000px planned 1.282, 7000px planned 1.000.
  //
  // After: `compactEmptyGutters` measures a band in the coordinates it is
  // handed, so a gap that was under the bar when it was measured can be over it
  // once the drawing is scaled. That produced a 21.1in band carrying nothing.
  // No fixture reproduces it today - round 74 moved the magnifier onto the
  // narrowest BADGED service, which lowered k on every sliver drawing far
  // enough that no surviving gutter crosses VOID_GUTTER_PX after scaling, and a
  // purpose-built 250px gutter at k=1.395 measured byte-identical with and
  // without this call. It is kept as a guard, not as a proven line: the defect
  // was real, and what removed its reproduction was a change to how k is
  // chosen, which the next round is free to change back.
  const drawing = compactEmptyGutters(
    magnifiedForCallouts(compactEmptyGutters(collectExportBoxes(nodes)), edges),
  );
  // The band is page furniture, and furniture does not get to evict the thing
  // it describes. `workflowListFromEdges` has no cap, so a fully-meshed
  // architecture produces hundreds of numbered steps and a single-column band
  // taller than the page itself — at 762 steps it took the whole 198in and
  // handed the fit a negative height budget. Lay the rows in as many columns as
  // it takes to stay inside a third of the sheet: every sentence the author
  // wrote is still on the page, which is the one thing that must not be traded,
  // and the drawing keeps the two thirds it was drawn for.
  //
  // "A third of the sheet" is a third of *this* drawing's sheet, not of Visio's
  // 200in ceiling. Measured against the ceiling, a nine-service architecture on
  // an 11in page kept a single-column band 9.4in tall — 70% of the page, with
  // the drawing squeezed into what was left — because 9.4in is nothing next to
  // 66in. The drawing's own natural height is what the band has to be modest
  // beside.
  //
  // What is measured is now the height the rows actually occupy rather than one
  // pitch each, because a sentence too long for its column wraps and a band
  // sized for one line per step is short by however many lines it did not count.
  //
  // A column is narrowest, and so its sentences wrap longest, on the narrowest
  // page the exporter will emit. Measuring there can only ever reserve too much
  // room, never too little, which is the safe direction: the band is drawn at
  // the real page width and comes out no taller than its reservation.
  const stackFor = (cols: number, widthIn: number): number => workflowStackIn(
    reservedEntries,
    cols,
    workflowPanelWidthIn(widthIn, cols) / cols,
  ) + 0.5;
  // Splitting trades column height for column width, and the sentences have to
  // fit the width: past a point the extra wrapping costs more lines than the
  // split saves rows.
  //
  // Always the SHORTEST split, never the first one under the target. "First
  // under target" is not comparable between two page widths: wider columns wrap
  // less, so the wide pass reaches the target at fewer columns, and a band in
  // fewer columns is taller. Both results are under the target and the wide one
  // can still be much taller than the narrow one — which is exactly the
  // reservation this feeds. On a 500-step CJK workflow the narrow pass reserved
  // 58.95in at 3 columns and the wide pass drew 65.50in at 2, and the page came
  // out at 206in: a file Visio will not open at all. The shortest band at a
  // wider page is no taller than the shortest at a narrower one for every
  // column count, so taking the shortest makes the reservation an upper bound
  // by construction.
  const naturalHIn = (() => {
    const b = computeBounds(drawing.values());
    return Math.max(b.maxY - b.minY, 1) / PX_PER_INCH;
  })();
  const bandTargetIn = Math.min(
    MAX_VISIO_PAGE_IN / 3,
    Math.max(MIN_PAGE_H_IN / 3, naturalHIn / 2),
  );
  // HOW MANY COLUMNS THE PROSE WILL STAND, decided ONCE and for every page
  // width. Both bounds below depend on the column's width, so asking them per
  // page made the admissible SET depend on the page - and the reservation pass
  // runs at the narrowest page the exporter emits while the layout pass runs at
  // the real one. A narrow page allowed fewer columns and so reserved a taller
  // band than the wide page went on to draw, and the difference was printed as
  // blank paper between the drawing and the band: 1.30in of it. Deciding the
  // set at the reserve width keeps "the reservation is an upper bound" true by
  // construction, which is the property the two-pass sizing rests on.
  //
  // A COLUMN TOO NARROW TO READ IS NOT A SHORTER BAND, IT IS A LOST SENTENCE.
  // The minimiser scores nothing but stack height, and with brief descriptions
  // the height falls monotonically in the column count, so it took the cap
  // every time: twelve two-syllable steps were set in columns 0.8583in wide,
  // leaving 0.2583in of text column - about two characters. Splitting has to
  // stop while the sentence is still a sentence.
  //
  // And the same bound asked the other way. A wide column can still shred long
  // prose into ribbons, which the width test cannot see and the height score
  // actively rewards.
  //
  // RELATIVE to what one column would have set, not an absolute three. An
  // 800-character sentence takes six lines in the widest column this page has
  // to give, so an absolute bound accuses the split of damage the prose did:
  // it refused to split at all, the band came out 2.98in tall in one column,
  // and 1.30in of the page it reserved was printed as blank paper. What the
  // split owns is the DIFFERENCE it makes, which is what this measures - and
  // it is the same test the audit applies to the emitted sheet.
  // HOW MANY COLUMNS THE PROSE WILL STAND, asked at the width the band is being
  // measured for. This used to be decided ONCE, at the reserve width, and for
  // a good reason: the reservation pass runs at the narrowest page the exporter
  // emits while the layout pass runs at the real one, a narrow page allows
  // fewer columns and so reserves a taller band than the wide page goes on to
  // draw, and that difference was printed as blank paper between the drawing
  // and the band - 1.30in of it.
  //
  // That failure mode no longer exists. The panel is now handed the height the
  // page reserved and fills it (see `workflowBandIn` below), so a reservation
  // larger than the drawn stack is leading between rows, never a hole. Pinning
  // the set to the narrow page therefore bought nothing and cost a great deal:
  // an 11in page allows 2 columns of 5.15in, which the median bound refuses, so
  // a 20.26in sheet that could have carried two 7.5in columns was held to one
  // and its band came out 15.05in tall for a drawing spanning 9.11in.
  //
  // A COLUMN TOO NARROW TO READ IS NOT A SHORTER BAND, IT IS A LOST SENTENCE.
  // The minimiser scores nothing but stack height, and with brief descriptions
  // the height falls monotonically in the column count, so it took the cap
  // every time: twelve two-syllable steps were set in columns 0.8583in wide,
  // leaving 0.2583in of text column - about two characters. Splitting has to
  // stop while the sentence is still a sentence.
  //
  // And the same bound asked the other way. A wide column can still shred long
  // prose into ribbons, which the width test cannot see and the height score
  // actively rewards.
  //
  // RELATIVE to what one column would have set, not an absolute three. An
  // 800-character sentence takes six lines in the widest column this page has
  // to give, so an absolute bound accuses the split of damage the prose did:
  // it refused to split at all, the band came out 2.98in tall in one column,
  // and 1.30in of the page it reserved was printed as blank paper. What the
  // split owns is the DIFFERENCE it makes, which is what this measures - and
  // it is the same test the audit applies to the emitted sheet.
  //
  // The unsplit reference stays at MIN_PAGE_W_IN because it is a READING
  // MEASURE, not a page measure: `workflowPanelWidthIn(w, 1)` is capped at
  // 7.5in for every page at or above 8.2in, so it is the same 7.5in column on
  // both passes and the bound means the same thing on both.
  const admissibleColumnsAt = (widthIn: number): number => {
    const unsplitColW = workflowPanelWidthIn(MIN_PAGE_W_IN, 1);
    const unsplit = workflowMedianLines(reservedEntries, unsplitColW);
    const bound = Math.max(MAX_WORKFLOW_MEDIAN_LINES, unsplit);
    let last = 1;
    for (let cols = 2; cols <= MAX_WORKFLOW_COLUMNS; cols += 1) {
      const colW = workflowPanelWidthIn(widthIn, cols) / cols;
      if (colW < MIN_WORKFLOW_COL_IN) break;
      if (workflowMedianLines(reservedEntries, colW) > bound) break;
      if (workflowTailShredded(reservedEntries, colW, unsplitColW, WORKFLOW_TAIL_MULTIPLE)) break;
      last = cols;
    }
    return last;
  };
  const bandFor = (widthIn: number): { columns: number; height: number } => {
    if (workflowEntries.length === 0) return { columns: 1, height: 0 };
    const admissible = admissibleColumnsAt(widthIn);
    let columns = 1;
    let shortest = stackFor(1, widthIn);
    if (shortest > bandTargetIn) {
      for (let cols = 2; cols <= admissible; cols += 1) {
        const height = stackFor(cols, widthIn);
        if (height < shortest) {
          shortest = height;
          columns = cols;
        }
      }
    }
    return { columns, height: shortest };
  };
  // Sized twice, because the band's height depends on how wide the page turns
  // out to be and the page's width is not known until the drawing is fitted.
  // The first pass reserves at the narrowest page the exporter emits, which is
  // the safe direction — a narrow column wraps longest, so it can only reserve
  // too much. That is what the fit is given. The second pass, once the real
  // width is known, is what the page is actually sized and laid out from: on a
  // 21.5in sheet the first pass over-reserved by 2.5in, and every inch of it
  // was drawn as blank paper between the drawing and the band.
  // The colour key is the same construction as the workflow band — opaque white
  // fill, drawn after every service — but it was pinned to the bottom-left
  // corner and reserved nothing, so on any drawing that reached the bottom of
  // its page it was simply painted over a service tile. Give it a strip of its
  // own, the way the band has one at the top.
  const legendBandIn = legendEntries.length > 0 ? 0.24 * legendEntries.length + 0.79 : 0;
  // Reserved TWICE, for the same reason it is sized twice.
  //
  // The first reservation is taken at the narrowest page the exporter emits,
  // which is the only width known before the drawing is fitted. On a band of a
  // dozen rows that estimate is within a tenth of an inch of the truth. On a
  // band of ninety it is not: at MIN_PAGE_W_IN every row wraps its longest, so
  // the reserve came out 1.48in taller than the band the sheet actually drew,
  // the drawing was shrunk by all of it, and the whole 1.48in was printed as
  // blank paper between the drawing and the band.
  //
  // So the first fit is used only to learn a page width, and the reserve is
  // then retaken at that width. It stays an upper bound by construction: a
  // band's height falls monotonically as its column widens, and this width is
  // itself a lower bound on the final one (a smaller reserve leaves a larger
  // drawing, and a larger drawing only widens the page), so the band measured
  // here can only be taller than the band finally drawn, never shorter.
  const pageWidthForBand = (reserveIn: number): number => {
    const probe = scaleBoxesWithin(
      fitBoxesWithin(drawing, limitPx, limitPx - (reserveIn + legendBandIn) * PX_PER_INCH),
      limitPx,
      limitPx - (reserveIn + legendBandIn) * PX_PER_INCH,
    );
    const b = computeBounds(probe.values());
    return Math.max(Math.max(b.maxX - b.minX, 1) / PX_PER_INCH + PAGE_PADDING_IN * 2, MIN_PAGE_W_IN);
  };
  const reserveBandIn = workflowEntries.length === 0
    ? 0
    : bandFor(pageWidthForBand(bandFor(MIN_PAGE_W_IN).height)).height;
  const fitted = fitBoxesWithin(
    drawing,
    limitPx,
    limitPx - (reserveBandIn + legendBandIn) * PX_PER_INCH,
  );
  // Squeezing the gaps is the first answer because it costs nothing but
  // distance. It has nothing left to give when the shapes alone are over the
  // limit — a single row of more than 127 tiles, or a zone rectangle spanning
  // the drawing, which counts as solid. Scaling is worse: it takes the type
  // down with it. A file Visio will not open is worse still.
  const raw = scaleBoxesWithin(fitted, limitPx, limitPx - (reserveBandIn + legendBandIn) * PX_PER_INCH);
  // Take the type down with the drawing. Holding it fixed does not rescue the
  // words on a tile an eighth of an inch wide, it prints each name over its
  // neighbours and loses the structure as well as the labels.
  const fonts = withDrawingCeilings(
    fontsForScale(boxScaleWithin(fitted, limitPx, limitPx - (reserveBandIn + legendBandIn) * PX_PER_INCH)),
    raw.values(),
  );
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
  const pageWidthIn = f(Math.max(contentWIn + PAGE_PADDING_IN * 2, MIN_PAGE_W_IN));
  // Second pass: the real width is known now, so the band can be laid out for
  // the page it is actually going on instead of for the narrowest one.
  const refined = bandFor(pageWidthIn);
  const workflowColumns = refined.columns;
  // Not `min(refined, reserveBandIn)`: the two are laid out at different widths
  // and so can pick different column counts, which makes the narrow-page
  // estimate the *smaller* of the two. Taking the min then modelled a band
  // 0.03in shorter than the one drawn, and a badge stepped clear of the model
  // still had a sliver under the panel. The band drawn is the band at this
  // width; the reservation only ever has to be an upper bound of it.
  // The reservation is what the page is sized from, so the panel must FILL it.
  //
  // The two are measured over different sentence sets and always will be: the
  // reservation counts every labelled edge as if its wording will be handed to
  // the row, because whether a label is muted is not decided until the arrows
  // are routed, which is after the page is sized. On a 91-row band that made
  // the reservation 1.48in taller than the panel drawn, and because the panel
  // hangs from the top of the sheet the whole 1.48in came out as blank paper
  // between the drawing and the band.
  //
  // Sizing the page down to the drawn panel instead is the wrong direction: it
  // puts the drawing where the reservation used to be, and on 29 scenarios the
  // opaque band was then painted straight over the service tiles. It cannot be
  // made to work, either, because the drawn panel's own height is not known
  // here - `mutedWording` is filled in during label placement, which needs the
  // page height this line is computing.
  //
  // So the panel is handed this height and spreads the difference as LEADING
  // BETWEEN ITS ROWS. Not as a deeper footer: that is the same hole moved
  // indoors and the audit rejects it by name. Holding a 40-service sheet fixed
  // and lengthening only the connector label used to take it from 17.09in to
  // 26.39in for a drawing spanning 9.11in, with 3.51in of blank paper in the
  // middle of it.
  const workflowBandIn = refined.height;
  const pageHeightIn = f(Math.max(contentHIn + PAGE_PADDING_IN * 2 + workflowBandIn + legendBandIn, MIN_PAGE_H_IN));

  // Centre the drawing between the two panels, converting to Visio's
  // bottom-left origin. `topY` counts down from the top of the page, so the
  // band — which is drawn at the top — is what the offset has to clear; the
  // legend sits at the bottom and is cleared by leaving its strip out of the
  // height the drawing is centred in.
  const offsetXIn = (pageWidthIn - contentWIn) / 2;
  const offsetYIn = (pageHeightIn - workflowBandIn - legendBandIn - contentHIn) / 2 + workflowBandIn;
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
      topY = clampIn(topY, drawingTopIn, pageHeightIn - h - Math.max(legendBandIn, PAGE_PADDING_IN / 2));
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

  // 256px, not the 128 the deck uses. A Visio icon is drawn about 0.36in wide,
  // so 128px is ~356 dpi: fine printed, but Visio is a *zooming* tool and at
  // the 400% people actually work at that falls to ~89 dpi, which is visibly
  // soft. 256px doubles it. Measured cost across 25 distinct icons: 210KB ->
  // 502KB of PNG, which the package cannot deflate further -- accepted,
  // because a drawing that goes blurry the moment it is inspected is not
  // usable, and Visio drawings are inspected closely by definition.
  //
  // Not vector, unlike the PowerPoint path: a VSDX carries images as
  // `ForeignData`, and no Visio is available here to confirm what it accepts,
  // so this stays a change of degree in a format already proven to render.
  const icons = presetIcons ?? await rasterizeIcons(services.map((box) => box.iconPath), 256);

  const shapes: string[] = [];
  const connects: string[] = [];
  const media: Array<{ file: string; bytes: Uint8Array }> = [];
  const pageRels: string[] = [];
  const shapeIdByNode = new Map<string, number>();
  let nextId = 1;
  let mediaIndex = 0;
  /**
   * One media part per distinct icon, not per service that draws it.
   *
   * A drawing repeats the same service icon constantly, and a part per shape
   * stored the identical PNG once for every tile: a zip compresses each entry
   * on its own, so twenty App Service tiles really did carry twenty copies of
   * the same bytes. Nothing referenced them separately -- a relationship is
   * just a pointer, and several shapes may share one.
   */
  const iconRelByPath = new Map<string, string>();

  for (let zoneIndex = 0; zoneIndex < groups.length; zoneIndex += 1) {
    const zone = groups[zoneIndex];
    const id = nextId++;
    shapeIdByNode.set(zone.id, id);
    const palette = paletteForZone(zone);
    shapes.push(zoneShapeXml(id, toRect(zone), zone.label, palette, fonts));
  }

  // Names the tiles could not hold in full, for the panel that gives the sheet
  // the recovery route the deck has had all along.
  const shortenedNames: Array<{ authored: string; drawn: string }> = [];
  // Every drawn name, for the case where the sheet is legible nowhere rather
  // than on one tile. See the index decision below.
  const allNames: Array<{ authored: string; drawn: string }> = [];
  const drawnByTile = new Map<string, string>();
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
      const key = service.iconPath ?? '';
      relId = iconRelByPath.get(key) ?? null;
      if (!relId) {
        mediaIndex += 1;
        const file = `image${mediaIndex}.png`;
        relId = `rId${mediaIndex}`;
        media.push({ file, bytes: icon.bytes });
        pageRels.push(
          `  <Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${file}"/>`,
        );
        iconRelByPath.set(key, relId);
      }
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
    // Tags are chips on the canvas tile, and they were reaching neither the
    // drawn text nor the shape data — so a diagram tagged "PCI" exported with
    // no trace of it. Visio has no chip, but shape data is better: a reader can
    // filter and report on it. One row per tag keeps each individually
    // searchable rather than hiding them in one comma-joined string.
    for (const [index, tag] of (service.tags ?? []).entries()) {
      properties.push({
        name: `Tag${index + 1}`,
        label: (service.tags ?? []).length > 1 ? `Tag ${index + 1}` : 'Tag',
        value: tag,
      });
    }

    // Read back from the XML that was just emitted, rather than re-deciding.
    // "Which names did the tile shorten?" is answered by exactly one piece of
    // code - the one that wrote the shape - so the panel cannot drift out of
    // step with the drawing the way a second copy of the fitting rules would.
    let tileXml = serviceGroupXml(
      { group: groupId, rect: rectId, icon: iconId },
      rect,
      service,
      paletteForService(service),
      relId,
      properties,
      meta,
      fonts,
    );
    const authoredName = String(service.label ?? '').trim();
    const readDrawn = (src: string): string => stripMarkup(/<Text>([\s\S]*?)<\/Text>/.exec(src)?.[1] ?? '')
      .split('\n')[0]
      .trim()
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
    let drawnName = readDrawn(tileXml);
    // A STUB IS A LOOKUP KEY, AND A KEY THAT REPEATS IS NOT A KEY. The index is
    // on another page, so the reader cannot see it and the drawing at the same
    // time: the drawing has to be self-consistent alone. Eight services sharing
    // a long prefix on narrow tiles all shortened to "C…", which makes both the
    // tiles and their index rows unmatchable. Lengthening is not always
    // available - a tile with room for 1.79 characters has nothing to lengthen
    // into - so a colliding tile falls back to a numeric key, which is unique,
    // narrower than most letters, and resolves in the index exactly as a stub
    // does. Never blank the tile: a longer key costs nothing, a blank tile
    // costs everything.
    if (authoredName && drawnName !== authoredName
      && (!/[\p{L}\p{N}]/u.test(drawnName)
        || (drawnByTile.has(drawnName) && drawnByTile.get(drawnName) !== authoredName))) {
      const key = `${drawnByTile.size + 1}`;
      tileXml = serviceGroupXml(
        { group: groupId, rect: rectId, icon: iconId },
        rect,
        { ...service, label: key },
        paletteForService(service),
        relId,
        properties,
        meta,
        fonts,
      );
      // The KEY is what gets drawn; the NAME attribute still has to carry the
      // service the shape actually is. Visio's Name is what a reader sees in
      // the shape window and what every cross-format check identifies a tile
      // by, so letting the re-render stamp "2" over it would rename the service
      // rather than merely abbreviate its caption.
      tileXml = tileXml.replace(
        `NameU="Service.${groupId}" Name="${esc(key)}"`,
        `NameU="Service.${groupId}" Name="${esc(authoredName)}"`,
      );
      drawnName = readDrawn(tileXml);
    }
    if (drawnName) drawnByTile.set(drawnName, authoredName);
    if (authoredName) allNames.push({ authored: authoredName, drawn: drawnName });
    if (authoredName && drawnName !== authoredName) {
      shortenedNames.push({ authored: authoredName, drawn: drawnName });
    }
    shapes.push(tileXml);
  }

  // Parallel hops between the same two services all put their text at the same
  // midpoint, so a fan of five wrote five sentences on top of each other and
  // the sheet was unreadable at exactly the place it had the most to say. Fan
  // the labels the way the arrows themselves are fanned, and hang each badge
  // off its own rung.
  const bundleOf = (route: ExportRoute): string => (route.sourceId < route.targetId
    ? `${route.sourceId}|${route.targetId}`
    : `${route.targetId}|${route.sourceId}`);
  // The badge scales with the drawing, so every allowance made for one has
  // to scale too, or the rungs reserve a quarter inch for a mark an eighth of
  // an inch across and the fan is spaced for furniture that is not there.
  // Taken at the highest step on the sheet, because the digits are what a
  // badge cannot give up and a three-digit disc is the widest one that will
  // be drawn - the reservation has to bound every badge, not the median one.
  // Taken as the widest badge any route on the sheet will actually be drawn
  // at, because the reservation has to bound every badge and each one is now
  // capped by its own two endpoint tiles rather than by one sheet-wide figure.
  const highestStep = routes.reduce((hi, r) => Math.max(hi, r.stepNumber ?? 0), 1);
  const badgeIn = (routes as ExportRoute[]).reduce(
    (hi, r) => Math.max(hi, badgeDiameterIn(r.stepNumber ?? highestStep, fonts, badgeCeilingIn(r, boxes))),
    badgeDiameterIn(highestStep, fonts, 0),
  );
  const labelSize = (label: string): { w: number; h: number } => {
    // `natural` is a width and a hard break contributes nothing to it, so the
    // chip is only ever as wide as the longest single line — which is right.
    // The *height* has to be counted, not divided: this was the third copy of
    // `ceil(width / column)` in the file, and because a newline adds no width
    // it sized a four-paragraph label at one line. `h` is not decoration — it
    // is the collision rectangle for seating, the rung pitch of a fan, the walk
    // step when a chip is settled, and the emitted `TxtHeight` — so a fan of
    // four hops kept a frozen 0.490in pitch while the text grew to 0.570in and
    // the chips were written through each other.
    const natural = estimateTextWidthIn(label, fonts.connector) + 0.08;
    const w = Math.min(Math.max(natural, 0.5), 1.7);
    const lines = wrappedLinesIn(label, Math.max(w - 0.08, 0.1), fonts.connector);
    return { w, h: lines * fonts.connector * 1.3 + 0.05 };
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
  // The two panels are opaque white and are drawn last, over everything.
  //
  // The search used to bound itself against the page alone, so the strip the
  // workflow band occupies read as a wide expanse of clear paper holding no
  // service and no label — and the ring search walked whole ladders into it.
  // On `twin-ladders` that hid ten of thirty-three connector labels and ten
  // step badges under the panel, while every rule in the audit passed: the
  // exporter was scoring itself on text it had written, not on text a reader
  // can see. Measured as rectangles rather than as a ceiling on `y`, because
  // the band is 7.5in wide on a 14.85in sheet and the paper beside it is real.
  const furnitureRects: Array<{ x: number; y: number; w: number; h: number }> = [];
  if (workflowEntries.length > 0) {
    const bandW = workflowPanelWidthIn(pageWidthIn, workflowColumns);
    // A LABEL THAT GRAZES THE PANEL IS UNDER IT. The panel used to be drawn at
    // its own natural height and the strip avoided here is the reservation, so
    // there was accidental clearance between the two and a label parked flush
    // under the strip still came out on clear paper. The panel now fills the
    // reservation, so that slack is gone and flush means covered: `outlier`
    // put "Managed identity" exactly on the boundary. The clearance has to be
    // asked for rather than inherited from a sizing bug.
    //
    // A HAIR OF IT, not a line of type. Widening the avoided strip by a full
    // line pitch is not a clearance, it is a smaller page: the search reads the
    // strip as blocked, the ladder is judged stuck, and five sentences that had
    // clear paper beside them were muted into the band instead of being drawn
    // where they belong.
    const clear = 0.05;
    furnitureRects.push({
      x: 0.35,
      y: pageHeightIn - 0.2 - workflowBandIn - clear,
      w: bandW,
      h: workflowBandIn + clear,
    });
  }
  if (legendEntries.length > 0) {
    furnitureRects.push({ x: 0.35, y: 0.35, w: 2.4, h: 0.24 * legendEntries.length + 0.34 });
  }
  // The index panel reserves NOTHING here on purpose: it is on its own page, so
  // there is nothing on this sheet for a connector label to be hidden under.
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
    let colPitch = 0;
    for (const member of members) {
      const size = labelOf(member) ? labelSize(labelOf(member)) : { w: badgeIn, h: 0 };
      rung = Math.max(rung, size.h + (member.stepNumber === undefined ? 0.05 : badgeIn + 0.07));
      colPitch = Math.max(colPitch, size.w + 0.12);
    }
    // Ranked by where the ARROW was fanned to, not by the order the edges were
    // declared, so rung n sits beside arrow n instead of crossing over it.
    const n = normalOf(members[0]);
    const ordered = [...members].sort((l, r) => {
      const lm = midOf(l);
      const rm = midOf(r);
      return (lm.x * n.x + lm.y * n.y) - (rm.x * n.x + rm.y * n.y);
    });
    // A single column of rungs is only a ladder while it is shorter than the
    // paper it stands in. Twenty numbered arrows between one pair of services
    // is 6.2in of rungs beside a 2.6in drawing, and the ladder simply kept
    // climbing — off the top of the sheet before, and into the workflow band
    // once the band was reserved. Fold it instead: the same rungs in as many
    // columns as it takes, which is the one arrangement that keeps every badge
    // both beside its arrow and on visible paper.
    const roomIn = Math.max(rung * 2, pageHeightIn - workflowBandIn - 0.4);
    const perColumn = Math.max(2, Math.min(ordered.length, Math.floor(roomIn / Math.max(rung, 0.01))));
    const columns = Math.ceil(ordered.length / perColumn);
    ordered.forEach((member, index) => {
      const column = Math.floor(index / perColumn);
      const row = index % perColumn;
      const rows = Math.min(perColumn, ordered.length - column * perColumn);
      ladder.set(member.id, {
        drop: (row - (rows - 1) / 2) * rung,
        along: seatOf(member).along + (column - (columns - 1) / 2) * colPitch,
      });
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
  // Every arrow on the sheet as the polyline it is actually drawn as, so a
  // placement can be scored against the hops it is NOT the label of. The
  // begin→end chord is the wrong line to measure against even though Visio pins
  // text to it: on an elbowed hop the chord runs through paper the arrow never
  // touches, so a seat 0.1in from the chord could be 0.7in from anything the
  // reader can see — and the export audit, which reads the drawn geometry, says
  // so. Measuring what is drawn is the only way the exporter and the rule that
  // judges it can agree.
  const polyOf = (route: ExportRoute): Point[] => route.points.map(toPoint);
  const polyCache = new Map<string, Point[]>();
  const polyFor = (route: ExportRoute): Point[] => {
    const seen = polyCache.get(route.id);
    if (seen) return seen;
    const pts = polyOf(route);
    polyCache.set(route.id, pts);
    return pts;
  };
  const gapToSegment = (a: Point, b: Point, at: Point): number => {
    const vx = b.x - a.x;
    const vy = b.y - a.y;
    const len2 = vx * vx + vy * vy;
    const t = len2 > 0
      ? Math.min(1, Math.max(0, ((at.x - a.x) * vx + (at.y - a.y) * vy) / len2))
      : 0;
    return Math.hypot(at.x - (a.x + vx * t), at.y - (a.y + vy * t));
  };
  const gapToPoly = (pts: Point[], at: Point): number => {
    if (pts.length === 0) return Infinity;
    if (pts.length === 1) return Math.hypot(at.x - pts[0].x, at.y - pts[0].y);
    let best = Infinity;
    for (let i = 1; i < pts.length; i += 1) {
      best = Math.min(best, gapToSegment(pts[i - 1], pts[i], at));
      if (best === 0) break;
    }
    return best;
  };
  const arrows = (routes as ExportRoute[]).map((route) => ({
    bundle: bundleOf(route),
    pts: polyFor(route),
  }));
  // Where a step badge sits on its rung: on the far side of its own label, so
  // the number and the words read as one object and the badge never covers the
  // text it belongs to.
  // Shared by the placement search and the emit — the two computing this
  // differently is how a badge ends up nowhere near the seat that was scored.
  // Leading the label instead (number first, then words) was tried and is
  // worse: it moves the badge *along* the hop rather than away from it, and on
  // a grid that carries it past its own arrow's end and onto the perpendicular
  // neighbour, turning 10 stray callouts into 26.
  const badgeSeatIn = (
    drop: number,
    along: number,
    textH: number,
  ): { away: number; along: number } => ({
    away: drop + (textH > 0 ? textH / 2 + badgeIn / 2 + 0.03 : 0),
    along,
  });
  /**
   * Move one fan as a body and report where it came to rest. `blocked` is the
   * collision area alone: the drift terms only break ties, so mixing them into
   * the total would make a clean placement look like a dirty one and the
   * caller could never tell whether the fan actually fits.
   */
  const settle = (members: ExportRoute[]): { shift: number; slide: number; blocked: number; buried: number } => {
    const step = Math.max(
      0.14,
      ...members.map((member) => (labelOf(member) ? labelSize(labelOf(member)).h : badgeIn) + 0.05),
    );
    // The arrows this fan's text could be mistaken for. Fan siblings are not
    // among them: a bundle of parallel edges between one pair of services is a
    // single object to the reader, so a rung nearer sibling 5 than sibling 6
    // misleads nobody. Gathered once per fan and filtered to the arrows that
    // pass anywhere near it, because the alternative is every candidate seat
    // measured against every arrow on a five-hundred hop sheet.
    const ownBundle = bundleOf(members[0]);
    const near = chordOf(members[0]);
    const strangers = arrows.filter((arrow) => arrow.bundle !== ownBundle
      && gapToPoly(arrow.pts, near) < 6);
    const badgeRoom = members.some((member) => member.stepNumber !== undefined) ? badgeIn + 0.06 : 0;
    // How far this ladder's text can drift before it stops belonging to its own
    // arrow. Sliding along a hop keeps the words reading as that hop's;
    // stepping away from it does not, and past a point the label is simply
    // nearer a different arrow — the reader then matches it to the wrong hop
    // and never knows they did. The walk was bounded only by blockage, so on a
    // tight grid a sentence crossed six inches of paper to find clear air and
    // arrived beside somebody else's arrow: 92 labels and 241 callouts across
    // the corpus, one of them 6.84in from its own hop and 0.59in from a
    // stranger's. The deck bounds the same walk with `chipReach`.
    //
    // Priced rather than forbidden. A hard cap on the search radius stopped the
    // wandering but also walled labels in: on a dense sheet the only clear seat
    // is often further than a comfortable reach, and capping the walk left
    // seven labels buried under icons on `estate72` and one on `tight-seam` —
    // trading a label the reader misattributes for one they cannot read at all.
    // Free inside the reach, then growing with the square of the excess, so a
    // ladder escaping a tile pays a little and one escaping an opaque panel can
    // still pay its way across the page.
    const stackIn = members.reduce(
      (sum, member) => sum + (labelOf(member) ? labelSize(labelOf(member)).h : badgeIn) + 0.05,
      0,
    );
    const reachIn = Math.max(
      1.2 * stackIn,
      0.6 * Math.max(...members.map((member) => labelSize(labelOf(member)).w)),
      0.5,
    );
    const blockage = (shift: number, slide: number): number => {
      let cost = 0;
      for (const member of members) {
        const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
        const drop = seat.drop + shift;
        const along = seat.along + slide;
        const text = labelOf(member);
        const box = rectAt(member, drop, along);
        const parts: Array<{ x: number; y: number; w: number; h: number; disc: boolean }> = text
          ? [{ ...box, disc: false }]
          : [];
        if (member.stepNumber !== undefined && badgeRoom > 0) {
          const n = normalOf(member);
          const u = directionOf(member);
          const centre = chordOf(member);
          const away = badgeSeatIn(drop, along, text ? box.h : 0);
          parts.push({
            x: centre.x + n.x * away.away + u.x * away.along - badgeIn / 2,
            y: centre.y + n.y * away.away + u.y * away.along - badgeIn / 2,
            w: badgeIn,
            h: badgeIn,
            disc: true,
          });
        }
        for (const part of parts) {
          if (part.x < 0.05 || part.y < 0.05
            || part.x + part.w > pageWidthIn - 0.05 || part.y + part.h > pageHeightIn - 0.05) {
            // Off the sheet is not a cost to be traded against, it is a deletion.
            cost += 100;
          }
          // A step badge is a filled disc, not text: where a sentence written
          // across a tile is untidy but still legible over mostly-empty fill, a
          // badge paints the icon out entirely. Priced as a fixed penalty
          // rather than by area — a half-inch disc has so little of it that
          // `area * 4` was small change beside the terms below, and the search
          // happily parked badges on icons to buy a tenth of an inch elsewhere.
          if (part.disc) {
            for (const rect of serviceRects) if (hit(part, rect) > 0) cost += 30;
          } else {
            for (const rect of serviceRects) cost += hit(part, rect) * 4;
          }
          // Under the panel is not a position to be traded against either — the
          // panel is drawn over it and the reader never sees the text at all.
          for (const rect of furnitureRects) if (hit(part, rect) > 0) cost += 100;
          for (const other of placedLabels) cost += hit(part, other) * 12;
          // Nearer somebody else's hop than its own is the one failure a reader
          // cannot detect: the words look like a perfectly ordinary label, just
          // of the wrong arrow. Scored as the distance the seat is on the wrong
          // side of that comparison, so a seat that is merely close to a
          // neighbour costs nothing and one that has crossed over costs in
          // proportion to how far.
          //
          // Deleting this term leaves the export audit at zero, because a
          // sentence that lands beside a stranger is judged lost and the muting
          // pass moves it into the workflow band — correct, but no longer on
          // the drawing. What it actually buys is text: `estate-chain` draws 35
          // labels with this term and 10 without. It is doing the work the
          // audit cannot see, and it is not dead code.
          const at = { x: part.x + part.w / 2, y: part.y + part.h / 2 };
          const mine = gapToPoly(polyFor(member), at);
          if (strangers.length > 0) {
            let closest = Infinity;
            for (const arrow of strangers) closest = Math.min(closest, gapToPoly(arrow.pts, at));
            if (closest < mine) cost += (mine - closest) * 6;
          }
          // And even with nobody to be confused with, text an arm's length from
          // its arrow belongs to nothing the reader can see. Free within reach,
          // then squared, so escaping a tile is affordable and drifting across
          // the page is not.
          const over = mine - reachIn;
          if (over > 0) cost += over * over * 3;
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
        // Words that sit nearer another hop than the one they describe are not
        // degraded, they are wrong: the reader attaches them to that hop with
        // complete confidence and never finds out. Counted as lost so the
        // muting pass moves the sentence into the workflow band, where it is
        // printed against its own step number and cannot be misread.
        if (strangers.length > 0) {
          const at = { x: box.x + box.w / 2, y: box.y + box.h / 2 };
          const mine = gapToPoly(polyFor(member), at);
          let closest = Infinity;
          for (const arrow of strangers) closest = Math.min(closest, gapToPoly(arrow.pts, at));
          if (closest < mine) return 1;
        }
        // An opaque panel over the words is a deletion, not a degradation: the
        // muting pass has to be told the trade failed so it puts the wording
        // back into the workflow row, where it can actually be read.
        for (const rect of furnitureRects) {
          if (hit(box, rect) > 0.01) return 1;
        }
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
    // Far enough that a ladder can cross a row of tiles to find clear air, and
    // no further. The bound used to be `reachIn` itself, which walled labels in
    // and left seven buried under icons on `estate72`; lifting it entirely
    // fixed that but made the walk scan every ring on a 200in page — over a
    // thousand of them per fan, and the export test file went from seconds to
    // beyond twenty minutes. Four times the reach is room enough to escape and
    // still a bounded search.
    const rings = Math.max(2, Math.min(
      Math.ceil(Math.max(pageWidthIn, pageHeightIn) / Math.max(step, 0.05)),
      Math.ceil(4 * reachIn / Math.max(step, 0.05)),
    ));
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
    // The travel term is priced on distance to the drawn polyline, but `ring`
    // counts displacement along the chord normal, and those are not the same
    // quantity: the natural seat already sits `drop` off the chord, and on an
    // elbowed hop a shift can walk *toward* the polyline before it walks away.
    // A break that estimated travel as `ring * step` was therefore free to cut
    // off a ring cheaper than its own estimate. Distance to a fixed set moves
    // by at most the distance the point moved, so `ring * step - approachIn` is
    // a sound lower bound on how far out the seat really is, where `approachIn`
    // is the furthest any member's seat starts from its own arrow.
    const approachIn = members.reduce((far, member) => {
      const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
      const box = rectAt(member, seat.drop, seat.along);
      const at = { x: box.x + box.w / 2, y: box.y + box.h / 2 };
      return Math.max(far, gapToPoly(polyFor(member), at));
    }, 0);
    for (let ring = 1; bestBlocked > 0 && ring <= rings; ring += 1) {
      // Travel is priced superlinearly and every other term is non-negative, so
      // once a ring's own travel already costs more than the best seat found,
      // nothing further out can beat it and the rest of the walk is waste.
      const excess = ring * step - approachIn - reachIn;
      if (excess > 0 && excess * excess * 3 >= bestCost) break;
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
  // Where a step badge sits, resolved once. Three copies of this arithmetic
  // used to exist — the placement search's model, the reservation that keeps
  // later fans off it, and the emit — and a badge is only ever as good as the
  // worst of them agreeing.
  //
  // The natural seat is the far side of its own label, so number and words read
  // as one object. That seat is a whole text-height further from the arrow than
  // the text is, though, and on a tight grid that extra 0.58in is the *next*
  // row's arrow: a numbered callout the reader confidently attaches to the
  // wrong hop. The fan-wide search cannot help — moving the ladder moves the
  // label too — so the badge is allowed to look for its own seat: the other
  // side of its arrow, closer in, or slid along it. Judged by the same tests
  // the export audit applies, and by whether it is nearer a stranger's arrow
  // than its own.
  //
  // Leading the label instead (number first, then words) was tried and is
  // worse: it moves the badge *along* the hop rather than away from it, and on
  // a grid that carries it past its own arrow's end onto the perpendicular
  // neighbour — 10 stray callouts became 26.
  const badgeAt = new Map<string, { x: number; y: number; d: number }>();
  // A chord pinned at the page edge, or under the band, has no seat within
  // reach and the search settles for the least bad one — still a callout the
  // reader cannot see. A badge is a small disc: stepping it clear of an opaque
  // panel is always possible and always better than leaving it invisible, even
  // when the only room left is over a tile.
  const clampBadge = (seat: { x: number; y: number; d: number }): {
    x: number; y: number; d: number;
  } => {
    const half = seat.d / 2;
    const ceiling = furnitureRects.reduce(
      (top, rect) => (rect.y + rect.h >= pageHeightIn - 0.5 ? Math.min(top, rect.y) : top),
      pageHeightIn,
    );
    const at = {
      x: clampIn(seat.x, half, pageWidthIn - half),
      y: clampIn(seat.y, half, ceiling - half),
      d: seat.d,
    };
    // The badge is stepped clear of any furniture panel it lands in. One pass
    // is enough: the panels are disjoint in `y` by construction — the page
    // always reserves `contentH + 1.2 + band + legend`, and the legend's
    // reservation runs 0.45in past its drawn box, which buys 1.10in of
    // guaranteed separation — so stepping out of one cannot land inside the
    // other. The export audit asserts that disjointness directly rather than
    // leaving it as an argument about four constants in two files.
    for (const rect of furnitureRects) {
      if (at.x + half <= rect.x || at.x - half >= rect.x + rect.w
        || at.y + half <= rect.y || at.y - half >= rect.y + rect.h) continue;
      const below = rect.y - half - 0.02;
      const above = rect.y + rect.h + half + 0.02;
      at.y = below >= half && (at.y <= rect.y + rect.h / 2 || above > pageHeightIn - half)
        ? below
        : above;
    }
    return at;
  };
  const badgeSeatFor = (
    member: ExportRoute,
    placed: { drop: number; along: number },
  ): { x: number; y: number; d: number } => {
    const n = normalOf(member);
    const u = directionOf(member);
    const centre = chordOf(member);
    const box = rectAt(member, placed.drop, placed.along);
    const textH = labelOf(member) ? box.h : 0;
    const ownPts = polyFor(member);
    const ownBundle = bundleOf(member);
    const strangers = arrows.filter((arrow) => arrow.bundle !== ownBundle
      && gapToPoly(arrow.pts, centre) < 3);
    const run = Math.max(runOf(member), 0);
    const seatAt = (d: number, side: number, away: number, slide: number): Point => {
      // Scored where the badge will actually sit, not where the search would
      // like it to. `clampBadge` pulls a seat back onto the page and out from
      // under the top furniture, and it used to do so *after* the winner was
      // chosen — so every term below was computed at a position the badge would
      // then be moved away from, and its attribution at the position it really
      // occupies was never checked at all. That was harmless only by accident,
      // because the candidate list happens to include a seat on the arrow
      // itself; it is an invariant now rather than a coincidence.
      const at = clampBadge({
        x: centre.x + n.x * side * away + u.x * slide,
        y: centre.y + n.y * side * away + u.y * slide,
        d,
      });
      return { x: at.x, y: at.y };
    };
    // Full size first, then progressively smaller, stopping at the first
    // diameter that finds a seat with nothing wrong with it. Shrinking is a
    // real cost — a tighter badge has less air around its number — so it is
    // only ever paid to buy a clean seat, never a marginally cheaper one. The
    // last step is the smallest disc that still holds the number at the 7pt
    // floor: below that the badge is no longer readable and there is nothing
    // left to trade.
    const natural = badgeIn;
    const floorD = badgeMinDiameterIn(member.stepNumber ?? 1, fonts);
    const diameters = [...new Set([natural, natural * 0.82, natural * 0.68, floorD])]
      .filter((d) => d >= floorD - 1e-9 && d <= natural + 1e-9)
      .sort((l, r) => r - l);
    // On a sheet scaled so far down that even a full-size badge is under the
    // floor, there is nothing to choose between: the drawing has already made
    // that trade everywhere else, and an empty candidate list would leave the
    // badge sitting on the raw chord — which is a tile.
    if (diameters.length === 0) diameters.push(natural);
    // The least-bad seat across every diameter, not the first one tried.
    //
    // When no seat anywhere is clean, this is what the badge gets, and keeping
    // the first meant keeping the largest disc's — even when giving up a little
    // size bought a seat on the callout's own arrow instead of on a stranger's.
    // The scorer already weighs that trade at 6 against inches; it was simply
    // never asked. On a grid pitched tighter than a badge is wide, five
    // callouts in twelve sat on a hop they did not belong to, which tells the
    // reader the architecture does something it does not.
    let fallback: { seat: { x: number; y: number; d: number }; total: number; attributed: boolean } | undefined;
    for (const d of diameters) {
      const half = d / 2;
      const naturalAway = placed.drop + (textH > 0 ? textH / 2 + half + 0.03 : 0);
      const cost = (at: Point, side: number, away: number, slide: number): {
        total: number;
        clean: boolean;
        /** Nearer its own arrow than any other, so it reads as its own hop's. */
        attributed: boolean;
      } => {
        const rect = { x: at.x - half, y: at.y - half, w: d, h: d };
        let total = 0;
        let clean = true;
        if (rect.x < 0.05 || rect.y < 0.05
          || rect.x + rect.w > pageWidthIn - 0.05 || rect.y + rect.h > pageHeightIn - 0.05) {
          total += 100;
          clean = false;
        }
        for (const furniture of furnitureRects) {
          if (hit(rect, furniture) > 0) { total += 100; clean = false; }
        }
        for (const service of serviceRects) {
          if (hit(rect, service) > 0) { total += 30; clean = false; }
        }
        for (const other of placedLabels) {
          const over = hit(rect, other);
          if (over > 0.01) clean = false;
          total += over * 12;
        }
        const mine = gapToPoly(ownPts, at);
        let closest = Infinity;
        for (const arrow of strangers) closest = Math.min(closest, gapToPoly(arrow.pts, at));
        // Deliberately not scaled by the drawing. Every other term here is a
        // preference measured in inches on the sheet, but this one is the seat
        // being nearer a stranger's arrow than its own — a callout that reads as
        // belonging to the wrong hop, which is a misstatement of the
        // architecture rather than an untidy sheet. Scaling it would let a
        // large drawing buy its way out of the misattribution with a little
        // tidiness elsewhere, and the weight of 6 is what keeps it dominant
        // over the `mine * 0.4` proximity term it competes with.
        if (closest < mine) { total += (mine - closest) * 6; clean = false; }
        // Near its own arrow, on the side its label chose, at the seat that
        // reads as part of the sentence: all three are preferences, not
        // requirements, and each is worth less than being read as another
        // hop's step.
        total += mine * 0.4;
        total += side < 0 ? 0.12 : 0;
        total += Math.abs(away - naturalAway) * 0.3;
        total += Math.abs(slide - placed.along) * 0.25;
        return { total, clean, attributed: closest >= mine };
      };
      // Three search axes, all three load-bearing. Measured over the audit
      // corpus, which seats 2052 callouts: the default seat wins 1841 of them,
      // and of the 211 that move, the slide-along-the-arrow offset decides 140,
      // the distance-from-the-arrow 169, and the side 23. None of these is dead
      // weight to be simplified away — dropping `slides` alone would take 140
      // callouts off the seat they were placed on for a reason.
      const aways = [naturalAway, placed.drop, half + 0.04];
      const slides = run > 0
        ? [placed.along, placed.along - run * 0.2, placed.along + run * 0.2,
          placed.along - run * 0.35, placed.along + run * 0.35]
        : [placed.along];
      let best = seatAt(d, 1, naturalAway, placed.along);
      let bestScore = cost(best, 1, naturalAway, placed.along);
      // Attribution first, then everything else. The weight of 6 keeps
      // misattribution ahead of the proximity term it directly competes with,
      // but not ahead of the 30 a seat pays for covering a service — so on a
      // grid with no clear air the search bought a tidy sheet with a callout
      // parked on someone else's arrow. Covering part of a tile is untidy and
      // the reader still sees which hop the number belongs to; sitting on the
      // wrong arrow says the architecture does something it does not. Rank the
      // two, rather than pricing them against each other.
      const beats = (
        candidate: { total: number; attributed: boolean },
        incumbent: { total: number; attributed: boolean },
      ): boolean => (candidate.attributed !== incumbent.attributed
        ? candidate.attributed
        : candidate.total < incumbent.total);
      for (const side of [1, -1]) {
        for (const away of aways) {
          for (const slide of slides) {
            const at = seatAt(d, side, away, slide);
            const score = cost(at, side, away, slide);
            if (beats(score, bestScore)) {
              bestScore = score;
              best = at;
            }
          }
        }
      }
      const seat = { x: best.x, y: best.y, d };
      if (bestScore.clean) return clampBadge(seat);
      if (!fallback || beats(bestScore, fallback)) {
        fallback = { seat, total: bestScore.total, attributed: bestScore.attributed };
      }
    }
    return clampBadge(fallback?.seat ?? { x: centre.x, y: centre.y, d: natural });
  };

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
  // Where a step badge sits, resolved once. Three copies of this arithmetic
  const best = { shift: placement.shift, slide: placement.slide };
    for (const member of members) {
      const seat = ladder.get(member.id) ?? { drop: 0, along: 0 };
      const placed = { drop: seat.drop + best.shift, along: seat.along + best.slide };
      ladder.set(member.id, placed);
      if (labelOf(member)) placedLabels.push(rectAt(member, placed.drop, placed.along));
      if (member.stepNumber !== undefined) {
        const at = badgeSeatFor(member, placed);
        badgeAt.set(member.id, at);
        placedLabels.push({
          x: at.x - at.d / 2,
          y: at.y - at.d / 2,
          w: at.d,
          h: at.d,
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
    // Local +Y is the line's left normal, and the page-space normal the badge
    // seat is resolved in is its opposite whenever the line runs left to right,
    // so the sign has to be carried across or the ladder and the badges fan
    // apart.
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
        fonts,
        route.id,
      ),
    );
    connects.push(connectXml(id, sourceId, targetId));
    if (route.stepNumber !== undefined) {
      // Resolved with the reservation, so the seat that was scored against the
      // panels, the tiles and every other arrow is the seat that gets drawn.
      const at = badgeAt.get(route.id) ?? badgeSeatFor(route, seat);
      shapes.push(stepBadgeXml(
        nextId++, at, route.stepNumber, fonts, route.id, at.d, badgeCeilingIn(route, boxes),
      ));
    }
  }

  // Colour key so the Visio page can't contradict the PNG's connection legend.
  // Resolved before the labels are placed: it is an opaque panel in the
  // bottom-left corner, so the search has to know where it will be.
  if (legendEntries.length > 0) {
    const legend = buildConnectionLegend(nextId, legendEntries, 0.35, 0.35);
    nextId = legend.nextId;
    shapes.push(...legend.shapes);
  }

  // The drawing's index, on its own page. Without it a name the tile could not
  // hold is nowhere a reader can see, which is the one thing a deck never does.
  // Built here, emitted as page 2 below: it takes no room on the drawing, so it
  // cannot be painted over a tile and needs no reservation in the fit.
  let indexPage: { shapes: string[]; widthIn: number; heightIn: number } | null = null;
  //
  // THE INDEX ANSWERS "CAN THE READER GET THIS NAME", NOT "WAS IT CUT".
  //
  // The trigger used to be `shortenedNames.length > 0`, which reads the wrong
  // property. A name is out of reach if the tile could not hold it OR if the
  // sheet prints it too small to read, and only the first of those shortens
  // anything. Visio refuses a page over 200in on a side, so a drawing that
  // does not fit is reduced instead, and past a certain size that reduction
  // takes the type down with it: 700 services came out at 1.22pt, 480 at
  // 1.78pt, 260 at 3.30pt, 150 at 5.72pt. All four printed every name in full,
  // so all four had `shortenedNames.length === 0` and emitted NO INDEX PAGE AT
  // ALL. Seven hundred service names on one sheet at a sixth of a point, and
  // nowhere in the file a reader could recover a single one of them.
  //
  // That is the exact asymmetry this panel was added to close, arriving by the
  // other road. The comment above still describes it correctly - a name the
  // sheet cannot show has to exist somewhere a reader can see - the trigger
  // just tested the one cause that had come up so far.
  //
  // The drawing is NOT stripped of its labels to compensate. A deck slide is a
  // fixed viewing surface, so there the answer is to drop a caption that cannot
  // be read and list it instead; a Visio page is a canvas the reader zooms, so
  // 1.22pt type is recoverable in the application and a deleted label is
  // recoverable nowhere. The sheet keeps its labels and gains the page that
  // makes them readable on paper too.
  const labelPt = fonts.label * 72;
  const tooSmallToRead = labelPt < LEGIBLE_PT - 1e-9;
  const indexEntries = tooSmallToRead ? allNames : shortenedNames;
  if (indexEntries.length > 0) {
    const namePanel = buildServiceNamePanel(
      nextId,
      indexEntries,
      tooSmallToRead
        ? `Service names in full. The drawing is reduced to ${Math.round(fonts.scale * 100)}% to fit a Visio page, so it prints labels at about ${labelPt.toFixed(1)} pt.`
        : 'Service names shortened on the drawing, in full',
    );
    nextId = namePanel.nextId;
    indexPage = { shapes: namePanel.shapes, widthIn: namePanel.widthIn, heightIn: namePanel.heightIn };
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
      // One column is the 7.5in reading measure the panel has always used;
      // more columns need proportionally more of the sheet, bounded by it.
      workflowPanelWidthIn(pageWidthIn, workflowColumns),
      workflowColumns,
      // The page was sized and the drawing was offset from `workflowBandIn`, so
      // that is the height this panel has to occupy - not the height its own
      // rows happen to need.
      workflowBandIn,
    );
    nextId = panel.nextId;
    shapes.push(...panel.shapes);
  }

  const parts: Array<{ path: string; data: string | Uint8Array }> = [
    { path: '[Content_Types].xml', data: contentTypesXml(indexPage ? 2 : 1) },
    { path: '_rels/.rels', data: ROOT_RELS },
    { path: 'docProps/core.xml', data: coreXml(diagramName) },
    { path: 'docProps/app.xml', data: APP_XML },
    { path: 'visio/document.xml', data: DOCUMENT_XML },
    { path: 'visio/_rels/document.xml.rels', data: DOCUMENT_RELS },
    { path: 'visio/windows.xml', data: windowsXml(pageWidthIn, pageHeightIn) },
    {
      path: 'visio/pages/pages.xml',
      data: pagesXml(
        pageWidthIn,
        pageHeightIn,
        diagramName,
        indexPage ? { widthIn: indexPage.widthIn, heightIn: indexPage.heightIn } : undefined,
      ),
    },
    { path: 'visio/pages/_rels/pages.xml.rels', data: indexPage ? PAGES_RELS_WITH_INDEX : PAGES_RELS },
    { path: 'visio/pages/page1.xml', data: pageContentsXml(shapes, connects) },
  ];

  if (indexPage) {
    parts.push({ path: 'visio/pages/page2.xml', data: pageContentsXml(indexPage.shapes, []) });
  }

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



