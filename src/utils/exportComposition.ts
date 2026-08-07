// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface DiagramContentBounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface DiagramPoint {
  x: number;
  y: number;
}

export interface ScreenContentRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface ContentCapturePlan {
  width: number;
  height: number;
  diagramHeight: number;
  headerHeight: number;
  legendHeight: number;
  legendColumns: number;
  scale: number;
  transformX: number;
  transformY: number;
}

interface ContentCapturePlanOptions {
  hasHeader?: boolean;
  legendItemCount?: number;
  margin?: number;
  minWidth?: number;
  maxWidth?: number;
  maxDiagramHeight?: number;
  minScale?: number;
  maxScale?: number;
}

export function screenRectToDiagramBounds(
  rect: ScreenContentRect,
  flowOrigin: DiagramPoint,
  zoom: number,
): DiagramContentBounds | null {
  if (
    !Number.isFinite(zoom)
    || zoom <= 0
    || ![rect.left, rect.top, rect.width, rect.height, flowOrigin.x, flowOrigin.y]
      .every(Number.isFinite)
    || rect.width <= 0
    || rect.height <= 0
  ) {
    return null;
  }

  return {
    x: (rect.left - flowOrigin.x) / zoom,
    y: (rect.top - flowOrigin.y) / zoom,
    width: rect.width / zoom,
    height: rect.height / zoom,
  };
}

export function expandDiagramContentBounds(
  bounds: DiagramContentBounds,
  additionalBounds: DiagramContentBounds[],
  padding = 0,
): DiagramContentBounds {
  const validAdditionalBounds = additionalBounds.filter((candidate) => (
    [candidate.x, candidate.y, candidate.width, candidate.height].every(Number.isFinite)
    && candidate.width > 0
    && candidate.height > 0
  ));
  if (validAdditionalBounds.length === 0) return bounds;

  const safePadding = Number.isFinite(padding) ? Math.max(0, padding) : 0;
  const minX = Math.min(
    bounds.x,
    ...validAdditionalBounds.map(candidate => candidate.x - safePadding),
  );
  const minY = Math.min(
    bounds.y,
    ...validAdditionalBounds.map(candidate => candidate.y - safePadding),
  );
  const maxX = Math.max(
    bounds.x + bounds.width,
    ...validAdditionalBounds.map(candidate => (
      candidate.x + candidate.width + safePadding
    )),
  );
  const maxY = Math.max(
    bounds.y + bounds.height,
    ...validAdditionalBounds.map(candidate => (
      candidate.y + candidate.height + safePadding
    )),
  );

  return {
    x: minX,
    y: minY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

export function calculateContentCapturePlan(
  bounds: DiagramContentBounds,
  options: ContentCapturePlanOptions = {},
): ContentCapturePlan {
  const margin = options.margin ?? 72;
  const minWidth = options.minWidth ?? 720;
  const maxWidth = options.maxWidth ?? 2200;
  const maxDiagramHeight = options.maxDiagramHeight ?? 1600;
  const minScale = options.minScale ?? 0.65;
  const maxScale = options.maxScale ?? 1.2;
  const contentWidth = Math.max(1, bounds.width);
  const contentHeight = Math.max(1, bounds.height);
  const availableScale = Math.min(
    (maxWidth - 2 * margin) / contentWidth,
    (maxDiagramHeight - 2 * margin) / contentHeight,
  );
  const scale = Math.max(minScale, Math.min(maxScale, availableScale));
  const scaledWidth = contentWidth * scale;
  const scaledHeight = contentHeight * scale;
  const width = Math.max(minWidth, Math.ceil(scaledWidth + 2 * margin));
  const diagramHeight = Math.ceil(scaledHeight + 2 * margin);
  const headerHeight = options.hasHeader ? 84 : 0;
  const legendItemCount = options.legendItemCount ?? 0;
  const legendColumns = width >= 1400 ? 5 : width >= 900 ? 3 : 2;
  const legendRows = legendItemCount > 0
    ? Math.ceil(legendItemCount / legendColumns)
    : 0;
  const legendHeight = legendRows > 0 ? 48 + legendRows * 58 : 0;

  return {
    width,
    height: headerHeight + diagramHeight + legendHeight,
    diagramHeight,
    headerHeight,
    legendHeight,
    legendColumns,
    scale,
    transformX: (width - scaledWidth) / 2 - bounds.x * scale,
    transformY: margin - bounds.y * scale,
  };
}
