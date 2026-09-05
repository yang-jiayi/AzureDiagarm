# PowerPoint / Visio Export Quality and Canvas Capture

## Overview

Office exports use the diagram model directly, independently of canvas size,
pan, zoom, selection, and floating UI panels. PNG/SVG capture remains a separate
path, described in section 2.

1. **Native Office exports**: editable PowerPoint and Visio diagrams.
2. **SVG edge rendering fix**: reliable PNG/SVG and validation snapshots.

---

## 1. Export Diagram as PowerPoint Slide

### Feature

The **"Export PPTX Slide"** option generates a `.pptx` with native shapes,
text, connectors, and embedded service icons. Small diagrams use a widescreen
(16:9) slide. Larger diagrams receive readable detail slides or a larger page,
within PowerPoint's 56-inch limit, instead of shrinking every label illegibly.
The customer deck keeps its 16:9 page and uses the same renderer with detail
slides when needed.

### Implementation

**Shared implementation**

- `src/services/diagramExportGeometry.ts` resolves nested/negative positions,
  shares connection semantics and category/zone colors with the canvas, and
  measures text using the actual Arial and East Asian font metrics. Its routing,
  gutter compaction, callout numbering, and label-fitting logic are shared by
  the Office exporters.
- `src/utils/exportIconRaster.ts` loads bundled SVGs through the existing asset
  loader and creates bitmap fallbacks. `diagramExportIcons.ts` is a small,
  strict preloading adapter over that loader; it does not implement a separate
  icon pipeline. Preloaded icons use the `RasterizedIcon` contract
  (`bytes`, `dataUrl`, `sizePx`, optional original `svg`).
- `src/services/pptxNativeShapes.ts` repairs grouping, connection attachments,
  script-specific fonts, and accessibility descriptions without moving shapes.
  `pptxVectorIcons.ts` embeds the SVG originals beside their PNG fallbacks.
  `pptxNativeDiagram.ts` only packages those established transforms; it is not
  a second diagram renderer. It also repairs duplicate table IDs emitted by
  PptxGenJS without changing any connector's attached shape ID.

**PowerPoint** (`src/services/pptxExporter.ts`)

PptxGenJS v4 is used entirely client-side. The full model, not the visible
viewport, supplies the drawing. Service inventories are paginated by measured
row height rather than cut off after 20 rows. Numbered workflows, service-name
indexes, speaker notes, tag chips, documentation links, and accessibility
descriptions remain part of the native output.

Straight routes use preset native PowerPoint connectors, attached where their
endpoints coincide with real connection sites. Bent routes retain their drawn
geometry as editable vector shapes rather than being converted into connectors
that PowerPoint could reroute. Desktop PowerPoint rejects arbitrary custom
geometry inside a connector element, so that invalid combination is never
emitted. A bent route must be repositioned separately when rearranging cards.
An overview can abbreviate wording, while detail, workflow, and index slides
carry the readable interpretation.

The slide layout:

```
┌─────────────────────────────────────────────┐
│ ██ Azure-blue accent bar (0.08" top) ████  │
├─────────────────────────────────────────────┤
│  Header strip (slate-900 / slate-200)       │
│  Diagram title (bold)       Author · Date  │
├─────────────────────────────────────────────┤
│                                             │
│         Native diagram (aspect-fit)          │
│                                             │
├─────────────────────────────────────────────┤
│ Footer text                                 │
└─────────────────────────────────────────────┘
```

**Theme palettes** (automatically matched to the current dark/light canvas mode):

| Token | Dark mode | Light mode |
|-------|-----------|------------|
| `bg` | `1e293b` (slate-800) | `f8fafc` (slate-50) |
| `headerBg` | `0f172a` (slate-900) | `e2e8f0` (slate-200) |
| `accent` | `0078d4` (Azure blue) | `0078d4` (Azure blue) |
| `titleText` | `ffffff` | `0f172a` |
| `metaText` | `94a3b8` | `475569` |
| `footerText` | `94a3b8` | `64748b` |

The graph-based APIs do not call `fitView()` or capture the DOM. Existing
callers that supply an image data URL and `options.diagram` remain supported.

**Visio** (`src/services/visioVsdxExporter.ts`)

The `.vsdx` package contains native service groups (card, icon, text, metadata
and searchable shape data), editable zone backgrounds, and 1-D connectors glued
to their endpoint shapes. The routed geometry, connection legend, numbered
workflow, and full-name index remain intact. Page fitting respects Visio's
200-inch limit.

Zone backgrounds remain independently editable; they are not automatic Visio
containers. Embedded PNG icons use the complete `ForeignData` → page
relationship → media-part chain. Visio retains its print-friendly light palette;
the optional `isDarkMode` argument is accepted for caller compatibility, not
as a sheet theme switch. PowerPoint does honor its dark/light theme option.
Service groups emit child `Shapes` before parent `Text`, as required by the
[Visio ShapeSheet schema](https://learn.microsoft.com/en-us/office/client-developer/visio/shapesheet_type-complextypevisio-xml).

### Honest pricing disclosures

`pricing.estimatedCost: null` means **Price unavailable**, not **Free**.
Only an explicitly known zero can be free. Fabric workload items with a
capacity-covered estimate say **incl. capacity**, not zero dollars; separately
priced Fabric Capacity and OneLake keep their own estimates. Callers may
localize these labels with `priceUnavailableLabel` and `capacityLabel`.
Usage-based positive estimates retain their `~` marker.

Apply `nodesForExport(nodes, showCostBadges)` at the application boundary.
Removed pricing stays removed: the exporter does not recreate a capacity or
unpriced disclosure from the service name after the user has hidden pricing.

`DeckCost.unpricedServices?: string[]` makes a customer deck's total a
**known-cost subtotal (incomplete)**. Every exclusion is listed on additional,
legible pages, even if a name spans pages. Such a deck makes no cheapest-region
or savings claim. The existing incomplete-region notices and price-freshness
fields remain supported separately.

### Key API

```ts
import { exportDiagramAsPptx } from './services/pptxExporter';
import { buildVsdxBlob } from './services/visioVsdxExporter';

const fileName = await exportDiagramAsPptx({ nodes, edges }, {
  diagramName: 'Application architecture',
  author: 'Architecture team',
  date: '2026-09-05',
  isDarkMode: false,
});
// Downloads the PPTX and returns its timestamped/model-suffixed filename.

const drawing = await buildVsdxBlob(nodes, edges, 'Application architecture', {
  capacityLabel: 'Included in shared capacity',
  priceUnavailableLabel: 'Price unavailable',
});
```

`buildDiagramPptxBlob(input, options, icons?)` and
`buildArchitectureDeckBlob(input, options, icons?)` return the same repaired
native package without downloading it. `input` is `{ nodes, edges }` or the
legacy image string. `exportArchitectureDeck` accepts both forms too.
The existing `buildDiagramSlidePptx` / `buildArchitectureDeckPptx` builder APIs
and `buildVsdxPackage(nodes, edges, name, presetIcons?, pricingOptions?)`
remain available for package audits.

### Regression coverage

Run `npm run test:exports` for geometry, native PPTX, and VSDX package checks.
The broader unit suite and `npm run test:export-quality` retain the established
typography, composition, editability, self-description, and golden audits.
The Blob APIs are compared with those audited builders, rather than tested
against a different rendering model. New coverage checks null versus zero,
capacity disclosures, hidden pricing, and complete exclusion pagination.

`npm run test:exports:browser` builds a production-mode fixture using the existing
Vite/Playwright dependencies, loads real bundled icons, checks generated XML, and
exercises both Office formats and all three UI download entry points. Set
`OFFICE_EXPORT_ARTIFACT_DIR` to a project-local directory to retain sample
PPTX/VSDX files. Without it, build output and samples are removed after the run.
Use `OFFICE_EXPORT_BROWSER_CHANNEL=msedge` to use an installed Microsoft Edge.
For an isolated exporter check while application code is unavailable,
`OFFICE_EXPORT_FIXTURES_ONLY=1` omits the app build/UI checks and reports
`uiExports: 0`; it is not a replacement for the full browser test.

On Windows with PowerPoint installed, run
`npm run test:exports:desktop -- -ArtifactDirectory ".\office-export-artifacts"`
against those samples. This opens disposable copies within the artifact
directory, renders PNGs at the presentation's aspect ratio,
checks actual text bounds and confirms that a moved card retains its native
connector. It never saves changes to the originals or closes existing user
presentations. Connector chips and exclusion pages use absolute line spacing;
chip vertical insets match the fitter's 0.06-inch total allowance. Diagram
typography retains its measured Arial/Yu Gothic UI font pairing and independent
quality gates.

---

## 2. SVG Edge Rendering Fix (PNG/SVG capture)

### The Problem

All export formats (PNG via html2canvas, and later PNG/SVG via html-to-image) produced images where **all ReactFlow edge lines were completely invisible** — no strokes, no arrowheads, no dashed patterns.

### Root Cause Analysis

ReactFlow renders edges as `<path>` elements inside an inline `<svg>` block. The `stroke` colour is not set as an SVG presentation attribute on the element; instead it comes from a CSS class rule:

```css
/* reactflow/dist/style.css */
.react-flow__edge-path {
  stroke: #b1b1b7;
  stroke-width: 1;
  fill: none;
}
```

When **html2canvas** was used, it rasterised the HTML but largely ignored SVG subtrees — edges were absent or at best blurred.

After switching to **html-to-image**, the DOM is serialised into an SVG document where the ReactFlow canvas sits inside a `<foreignObject>` element. This is the correct approach and preserves the DOM tree faithfully, but it introduces a different problem:

> **SVG `<foreignObject>` content does not inherit the document's external stylesheets.**

Inside the `<foreignObject>`, the browser's style resolution no longer applies the page CSS. Every `<path>` with `class="react-flow__edge-path"` but no inline `stroke` attribute gets `stroke: none` by default in SVG — completely invisible.

This affects:

- Sync edges (solid dark/light adaptive stroke)
- Async edges (dashed)
- Optional edges (dotted)
- Animated directional-flow edges
- Bidirectional pulse edges

### The Fix

**`src/utils/captureCanvas.ts`** — `prepareEdgesForCapture(wrapper: HTMLElement): () => void`

The strategy: **read computed style, write as presentation attribute, restore after capture**.

```
Before capture:
  for each svg path/line/polyline/circle in wrapper
    for each of 11 SVG attributes (stroke etc.)
      val = window.getComputedStyle(el)[camelCaseAttr]  ← CSS-class value IS present here
      el.setAttribute(attr, val)                         ← inline on element → survives foreignObject

After capture (try / finally):
  restore original attribute state (setAttribute / removeAttribute)
```

#### Why `getComputedStyle` works here but not inside foreignObject

`window.getComputedStyle()` is called on the **live DOM before serialisation**, where all stylesheets are fully applied and cascade resolution has already happened. The computed value for `stroke` on a ReactFlow edge path at this point is the resolved colour string (e.g. `rgb(177, 177, 183)`). Writing it back as a presentation attribute bakes it into the element itself, so it survives the serialisation context change.

#### Attributes inlined

```ts
const SVG_ATTRS_TO_INLINE = [
  'stroke',
  'stroke-width',
  'stroke-dasharray',
  'stroke-dashoffset',
  'stroke-opacity',
  'stroke-linecap',
  'stroke-linejoin',
  'fill',
  'opacity',
  'marker-end',
  'marker-start',
];
```

`fill: rgba(0, 0, 0, 0)` and similar fully transparent values are normalised to `fill: none` (the SVG convention for "no fill") to avoid accidental black fills from `fill: transparent` being misinterpreted.

#### Restore contract

The function returns a restore callback. It is **always called in a `finally` block** so the live canvas SVG is never left in a modified state, even if the capture throws:

```ts
export async function captureDiagramAsPng(element, options) {
  const restore = prepareEdgesForCapture(element);
  try {
    return await toPng(element, resolvedOptions);
  } finally {
    restore();
  }
}
```

### Before / After

| Aspect | Before fix | After fix |
|--------|-----------|-----------|
| Edge lines | Invisible (no stroke) | Correctly coloured |
| Dashed async edges | Invisible | Dashed pattern preserved |
| Arrowheads | Missing | Rendered via `marker-end` |
| Animated edges | Invisible | Base style visible |
| Live canvas | N/A | Unaffected (attributes restored) |
| SVG export type | PNG rasterised inside SVG wrapper | True native vector SVG |
| Bundle size | +190KB (html2canvas) | −190KB (html-to-image) |

### Files Changed

| File | Change |
|------|--------|
| `src/utils/captureCanvas.ts` | **New**: capture utility with edge fix |
| `src/services/pptxExporter.ts` | **New**: PptxGenJS slide builder |
| `src/App.tsx` | Replaced 4× `html2canvas` call sites; added `exportAsPptx` callback and menu item |
| `package.json` | Added `pptxgenjs@^4.0.1`; `html2canvas` still present but no longer imported |

---

## Further Reading

- [html-to-image — foreignObject limitation](https://github.com/bubkoo/html-to-image#faqs)
- [SVG presentation attributes vs CSS properties](https://developer.mozilla.org/en-US/docs/Web/SVG/Attribute/Presentation)
- [PptxGenJS v4 documentation](https://gitbrent.github.io/PptxGenJS/)
- [ReactFlow edge styling](https://reactflow.dev/docs/guides/custom-edges/)
