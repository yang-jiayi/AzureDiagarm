// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Diagram capture utility
 *
 * Wraps html-to-image to reliably capture the ReactFlow canvas — including
 * SVG edge paths (smooth / bezier / orthogonal curves, dashed stroke patterns,
 * etc.) that html2canvas would silently drop.
 *
 * Root-cause of edge invisibility with html-to-image:
 *   html-to-image serialises the DOM into an SVG <foreignObject> block.
 *   Inside that foreign object, the original page's CSS stylesheets are no
 *   longer in scope, so SVG <path> elements that relied on CSS class rules
 *   for their stroke colour (e.g. ReactFlow's sync edges) render as invisible.
 *
 * Fix: `prepareEdgesForCapture()` reads each SVG element's COMPUTED styles
 * (which DO include the CSS-derived values) and copies them as SVG presentation
 * attributes directly onto the elements just before capture.  Presentation
 * attributes survive serialisation regardless of whether stylesheets are
 * present.  After capture, all attributes are restored to their original state.
 */

import { toPng } from 'html-to-image';
import {
  calculateContentCapturePlan,
  type DiagramContentBounds,
} from './exportComposition';

export type ExportBackground = 'plain' | 'dots' | 'grid';

// ─── SVG edge pre-inlining ────────────────────────────────────────────────────

/** Presentation attributes we forcibly inline before capture. */
const SVG_ATTRS_TO_INLINE = [
  'stroke',
  'stroke-width',
  'stroke-dasharray',
  'stroke-dashoffset',
  'stroke-opacity',
  'stroke-linecap',
  'stroke-linejoin',
  'fill',
  'fill-opacity',
  'opacity',
  'marker-end',
  'marker-start',
] as const;

/**
 * Copy computed CSS properties onto SVG path/line/polyline/circle elements as
 * presentation attributes.  Returns a cleanup function that restores the
 * original attribute state.
 */
function prepareEdgesForCapture(wrapper: HTMLElement): () => void {
  const restorers: Array<() => void> = [];

  wrapper
    .querySelectorAll<SVGElement>('svg path, svg line, svg polyline, svg circle')
    .forEach((el) => {
      const cs = window.getComputedStyle(el as unknown as Element);

      SVG_ATTRS_TO_INLINE.forEach((attr) => {
        // getComputedStyle uses camelCase for hyphenated properties
        const camel = attr.replace(/-([a-z])/g, (_, c) => c.toUpperCase()) as
          keyof CSSStyleDeclaration;
        const computed = cs[camel] as string | undefined;
        if (!computed || computed === '') return;

        const prev = el.getAttribute(attr);

        // Normalise transparent fills → 'none' (SVG convention)
        const normalised =
          attr === 'fill' &&
          (computed === 'rgba(0, 0, 0, 0)' || computed === 'transparent')
            ? 'none'
            : computed;

        el.setAttribute(attr, normalised);

        restorers.push(() => {
          if (prev === null) el.removeAttribute(attr);
          else el.setAttribute(attr, prev);
        });
      });
    });

  return () => restorers.forEach((fn) => fn());
}

// ─── Classes that should always be hidden from captured output ────────────────

const UI_CHROME_CLASSES = [
  'react-flow__minimap',
  'react-flow__controls',
  'react-flow__attribution',
  'react-flow__handle',
  'react-flow__resize-control',
  'ungroup-button',
  'fit-to-content-button',
  'color-picker-button',
  'color-picker-panel',
];

/** Extended set used for SVG/PPTX export (hides floating panels too). */
const PANEL_CLASSES = [
  ...UI_CHROME_CLASSES,
  'info-panel',
  'workflow-panel',
  'alignment-toolbar',
  'icon-palette',
  'canvas-nav-hint',
];

// ─── Types ────────────────────────────────────────────────────────────────────

export interface CaptureOptions {
  /** CSS color string for the background (any valid CSS value). */
  backgroundColor: string;
  /**
   * Output pixel ratio (default 2 for @2x / retina quality).
   * Uses the actual device pixel ratio when omitted.
   */
  pixelRatio?: number;
  /**
   * When true, also hide floating info/workflow/palette panels
   * in addition to the standard ReactFlow chrome. Useful for
   * "clean" exports (SVG, PPTX) where UI panels would be noise.
   */
  excludePanels?: boolean;
  /**
   * Background pattern for the exported artifact. The live React Flow canvas
   * remains dotted; this controls only the cloned capture output.
   * Defaults to `plain` for presentation-ready exports.
   */
  exportBackground?: ExportBackground;
  /**
   * Build a presentation-ready export around the React Flow viewport. The
   * viewport is cloned off-screen, tightly framed to these content bounds,
   * and combined with an optional title and relationship legend.
   */
  composition?: {
    bounds: DiagramContentBounds;
    title?: string;
    subtitle?: string;
    legendTitle?: string;
    legendItems?: Array<{
      label: string;
      description: string;
      color: string;
      lineStyle?: 'solid' | 'dashed' | 'dotted';
    }>;
  };
}

// ─── Internal helpers ─────────────────────────────────────────────────────────

function makeFilter(excludeClasses: string[]) {
  return (node: HTMLElement): boolean => {
    if (!node.classList) return true;
    return !excludeClasses.some((cls) => node.classList.contains(cls));
  };
}

function captureStyle(exportBackground: ExportBackground): Partial<CSSStyleDeclaration> {
  if (exportBackground !== 'grid') return {};
  return {
    backgroundImage: [
      'linear-gradient(rgba(96, 165, 250, 0.24) 1px, transparent 1px)',
      'linear-gradient(90deg, rgba(96, 165, 250, 0.24) 1px, transparent 1px)',
    ].join(', '),
    backgroundSize: '20px 20px',
    backgroundPosition: '0 0',
  };
}

function captureClasses(options: CaptureOptions): string[] {
  const classes = options.excludePanels ? [...PANEL_CLASSES] : [...UI_CHROME_CLASSES];
  if ((options.exportBackground ?? 'plain') !== 'dots') {
    classes.push('react-flow__background');
  }
  return classes;
}

  interface PreparedCaptureTarget {
    target: HTMLElement;
    width?: number;
    height?: number;
    cleanup: () => void;
  }

  function designToken(name: string, fallback: string): string {
    const value = window.getComputedStyle(document.body).getPropertyValue(name).trim();
    return value || fallback;
  }

  function applyExportBackground(
    element: HTMLElement,
    exportBackground: ExportBackground,
  ): void {
    if (exportBackground === 'dots') {
      element.style.backgroundImage = 'radial-gradient(circle, rgba(96, 165, 250, 0.32) 1.2px, transparent 1.3px)';
      element.style.backgroundSize = '20px 20px';
      return;
    }
    if (exportBackground === 'grid') {
      element.style.backgroundImage = [
        'linear-gradient(rgba(96, 165, 250, 0.24) 1px, transparent 1px)',
        'linear-gradient(90deg, rgba(96, 165, 250, 0.24) 1px, transparent 1px)',
      ].join(', ');
      element.style.backgroundSize = '20px 20px';
    }
  }

  function appendExportHeader(
    host: HTMLElement,
    height: number,
    title: string,
    subtitle: string | undefined,
  ): void {
    const header = document.createElement('header');
    Object.assign(header.style, {
      height: `${height}px`,
      boxSizing: 'border-box',
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      padding: '14px 28px',
      borderBottom: `1px solid ${designToken('--azd-color-border', '#d7e0ea')}`,
      background: designToken('--azd-color-surface-elevated', '#ffffff'),
    });

    const heading = document.createElement('div');
    heading.textContent = title;
    Object.assign(heading.style, {
      color: designToken('--azd-color-text-strong', '#0f172a'),
      fontSize: '24px',
      fontWeight: '700',
      lineHeight: '1.2',
    });
    header.appendChild(heading);

    if (subtitle) {
      const metadata = document.createElement('div');
      metadata.textContent = subtitle;
      Object.assign(metadata.style, {
        marginTop: '5px',
        color: designToken('--azd-color-text-muted', '#52677b'),
        fontSize: '12px',
        fontWeight: '600',
      });
      header.appendChild(metadata);
    }

    host.appendChild(header);
  }

  function appendExportLegend(
    host: HTMLElement,
    top: number,
    height: number,
    columns: number,
    title: string,
    items: NonNullable<NonNullable<CaptureOptions['composition']>['legendItems']>,
  ): void {
    const legend = document.createElement('section');
    Object.assign(legend.style, {
      position: 'absolute',
      top: `${top}px`,
      left: '0',
      width: '100%',
      height: `${height}px`,
      boxSizing: 'border-box',
      padding: '12px 24px 16px',
      borderTop: `1px solid ${designToken('--azd-color-border', '#d7e0ea')}`,
      background: designToken('--azd-color-surface-panel', '#f8fafc'),
    });

    const heading = document.createElement('div');
    heading.textContent = title;
    Object.assign(heading.style, {
      marginBottom: '10px',
      color: designToken('--azd-color-text-muted', '#52677b'),
      fontSize: '11px',
      fontWeight: '800',
      letterSpacing: '0.06em',
      textTransform: 'uppercase',
    });
    legend.appendChild(heading);

    const grid = document.createElement('div');
    Object.assign(grid.style, {
      display: 'grid',
      gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`,
      gap: '8px 18px',
    });

    for (const item of items) {
      const row = document.createElement('div');
      Object.assign(row.style, {
        minWidth: '0',
        display: 'grid',
        gridTemplateColumns: '48px minmax(0, 1fr)',
        alignItems: 'center',
        gap: '9px',
      });

      const sample = document.createElement('span');
      Object.assign(sample.style, {
        display: 'block',
        width: '44px',
        borderTopWidth: '3px',
        borderTopStyle: item.lineStyle ?? 'solid',
        borderTopColor: item.color,
      });
      row.appendChild(sample);

      const copy = document.createElement('span');
      copy.style.minWidth = '0';
      const label = document.createElement('strong');
      label.textContent = item.label;
      Object.assign(label.style, {
        display: 'block',
        color: designToken('--azd-color-text-strong', '#0f172a'),
        fontSize: '12px',
        lineHeight: '1.2',
      });
      const description = document.createElement('small');
      description.textContent = item.description;
      Object.assign(description.style, {
        display: 'block',
        marginTop: '2px',
        overflow: 'hidden',
        color: designToken('--azd-color-text-muted', '#52677b'),
        fontSize: '10px',
        lineHeight: '1.25',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap',
      });
      copy.append(label, description);
      row.appendChild(copy);
      grid.appendChild(row);
    }

    legend.appendChild(grid);
    host.appendChild(legend);
  }

  function prepareCompositionTarget(
    element: HTMLElement,
    options: CaptureOptions,
  ): PreparedCaptureTarget {
    const composition = options.composition;
    if (!composition) return { target: element, cleanup: () => undefined };

    const sourceViewport = element.querySelector<HTMLElement>('.react-flow__viewport');
    if (!sourceViewport) {
      throw new Error('React Flow viewport was not found for content-aware export.');
    }

    const legendItems = composition.legendItems ?? [];
    const threatOverlay = element.querySelector<HTMLElement>('.threat-model-overlay');
    const threatItemCount = threatOverlay?.querySelectorAll('li').length ?? 0;
    const measuredThreatHeight = threatOverlay
      ? Math.max(
          threatOverlay.scrollHeight,
          threatOverlay.getBoundingClientRect().height,
          112 + threatItemCount * 38,
        )
      : 0;
    const plan = calculateContentCapturePlan(composition.bounds, {
      hasHeader: Boolean(composition.title),
      legendItemCount: legendItems.length,
      minDiagramHeight: measuredThreatHeight > 0
        ? Math.ceil(measuredThreatHeight + 24)
        : undefined,
    });
    const host = document.createElement('div');
    host.className = 'react-flow export-capture-composition';
    host.setAttribute('data-export-composition', 'true');
    Object.assign(host.style, {
      position: 'fixed',
      left: '0',
      top: '0',
      width: `${plan.width}px`,
      height: `${plan.height}px`,
      overflow: 'hidden',
      boxSizing: 'border-box',
      background: options.backgroundColor,
      color: designToken('--azd-color-text', '#1f2937'),
      fontFamily: designToken('--azd-font-family', '"Segoe UI", sans-serif'),
      pointerEvents: 'none',
      zIndex: '-2147483647',
    });

    if (composition.title) {
      appendExportHeader(host, plan.headerHeight, composition.title, composition.subtitle);
    }

    const frame = document.createElement('div');
    Object.assign(frame.style, {
      position: 'absolute',
      top: `${plan.headerHeight}px`,
      left: '0',
      width: `${plan.width}px`,
      height: `${plan.diagramHeight}px`,
      overflow: 'hidden',
      background: options.backgroundColor,
    });
    applyExportBackground(frame, options.exportBackground ?? 'plain');

    const viewportClone = sourceViewport.cloneNode(true) as HTMLElement;
    Object.assign(viewportClone.style, {
      position: 'absolute',
      inset: '0',
      width: `${plan.width}px`,
      height: `${plan.diagramHeight}px`,
      transformOrigin: '0 0',
      transform: `translate(${plan.transformX}px, ${plan.transformY}px) scale(${plan.scale})`,
    });
    viewportClone.querySelectorAll<HTMLElement>('.selected, .is-selected').forEach((node) => {
      node.classList.remove('selected', 'is-selected');
    });
    viewportClone.querySelectorAll<HTMLElement>('.editable-edge-label.is-empty').forEach((label) => {
      label.closest('.editable-edge-label-shell')?.remove();
    });
    frame.appendChild(viewportClone);

    if (threatOverlay) {
      const overlayClone = threatOverlay.cloneNode(true) as HTMLElement;
      overlayClone.querySelectorAll('button').forEach(button => button.remove());
      overlayClone.querySelectorAll<HTMLElement>('li').forEach((item) => {
        item.style.setProperty('display', 'grid', 'important');
      });
      overlayClone.setAttribute('aria-hidden', 'true');
      overlayClone.setAttribute('data-export-threat-overlay', 'true');
      Object.assign(overlayClone.style, {
        pointerEvents: 'none',
        width: '310px',
        maxHeight: 'none',
        overflow: 'visible',
      });
      frame.appendChild(overlayClone);
    }

    host.appendChild(frame);

    if (legendItems.length > 0) {
      appendExportLegend(
        host,
        plan.headerHeight + plan.diagramHeight,
        plan.legendHeight,
        plan.legendColumns,
        composition.legendTitle || 'Connections',
        legendItems,
      );
    }

    document.body.appendChild(host);
    return {
      target: host,
      width: plan.width,
      height: plan.height,
      cleanup: () => host.remove(),
    };
  }

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Capture the element as a PNG and return a `data:image/png` data URL.
 *
 * SVG edge styles are pre-inlined as presentation attributes so they survive
 * the html-to-image foreignObject serialisation.
 */
export async function captureDiagramAsPng(
  element: HTMLElement,
  options: CaptureOptions,
): Promise<string> {
  const exportBackground = options.exportBackground ?? 'plain';
  const excludeClasses = captureClasses(options);
  const prepared = prepareCompositionTarget(element, options);
  const restore = prepareEdgesForCapture(prepared.target);

  try {
    return await toPng(prepared.target, {
      backgroundColor: options.backgroundColor,
      pixelRatio: options.pixelRatio ?? 2,
      width: prepared.width,
      height: prepared.height,
      filter: makeFilter(excludeClasses),
      style: options.composition ? {} : captureStyle(exportBackground),
      imagePlaceholder:
        'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    });
  } finally {
    restore();
    prepared.cleanup();
  }
}

/**
 * There is deliberately no `captureDiagramAsSvg` here any more.
 *
 * html-to-image's `toSvg` produces an SVG whose entire content is one
 * `<foreignObject>` full of XHTML. Browsers render that, so the output looked
 * correct on the only surface anyone checked, and it opened blank in Inkscape,
 * in Illustrator, through librsvg, in Office's Insert > Picture and in macOS
 * Preview. A user who chooses SVG over PNG is nearly always choosing it in
 * order to open the file somewhere else and edit it, so "renders in a browser"
 * is not the bar.
 *
 * `src/services/vectorSvgExporter.ts` builds the same drawing out of the shared
 * export geometry as real `<rect>`, `<path>`, `<text>` and nested `<svg>`
 * elements, which every tool can open and every shape of which is selectable.
 * Reach for that instead.
 */
