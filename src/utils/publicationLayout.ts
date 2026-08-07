// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { ReferenceArchitecture } from '../services/referenceArchitectureAI';
import type { BlueprintArchitecture } from '../services/blueprintArchitectureAI';

const REFERENCE_MIN_WIDTH = 960;
const REFERENCE_MAX_WIDTH = 1920;
const REFERENCE_HORIZONTAL_PADDING = 96;
const REFERENCE_SIDE_COLUMN_WIDTH = 196;
const REFERENCE_COLUMN_GAP = 20;
const REFERENCE_STAGE_GAP = 12;

export const BLUEPRINT_SIDEBAR_WIDTH = 400;
const BLUEPRINT_HORIZONTAL_PADDING = 96;
const BLUEPRINT_BODY_GAP = 20;
const BLUEPRINT_CONTENT_PADDING = 64;
const BLUEPRINT_MIN_WIDTH = 960;
const BLUEPRINT_MIN_HEIGHT = 520;
const BLUEPRINT_NODE_WIDTH = 180;
const BLUEPRINT_NODE_HEIGHT = 120;

function roundUp(value: number, increment: number): number {
  return Math.ceil(value / increment) * increment;
}

function preferredStageWidth(serviceCount: number): number {
  if (serviceCount >= 4) return 268;
  if (serviceCount >= 2) return 190;
  return 160;
}

export function calculateReferenceCanvasWidth(
  data: ReferenceArchitecture,
  requestedWidth?: number,
): number {
  if (requestedWidth !== undefined) {
    return Math.max(REFERENCE_MIN_WIDTH, Math.round(requestedWidth));
  }

  const stageCount = Math.max(1, data.stages.length);
  const maxServices = Math.max(0, ...data.stages.map(stage => stage.services.length));
  const stageAreaWidth =
    stageCount * preferredStageWidth(maxServices)
    + Math.max(0, stageCount - 1) * REFERENCE_STAGE_GAP;
  const sideColumnCount =
    (data.dataSources && data.dataSources.length > 0 ? 1 : 0)
    + (data.actors && data.actors.length > 0 ? 1 : 0);
  const bodyColumnCount = 1 + sideColumnCount;
  const contentWidth =
    stageAreaWidth
    + sideColumnCount * REFERENCE_SIDE_COLUMN_WIDTH
    + Math.max(0, bodyColumnCount - 1) * REFERENCE_COLUMN_GAP
    + REFERENCE_HORIZONTAL_PADDING;

  return Math.min(
    REFERENCE_MAX_WIDTH,
    Math.max(REFERENCE_MIN_WIDTH, roundUp(contentWidth, 8)),
  );
}

export function calculateBlueprintHostWidth(
  canvasWidth: number,
  legendPosition: 'bottom' | 'right',
): number {
  return legendPosition === 'right'
    ? canvasWidth + BLUEPRINT_SIDEBAR_WIDTH + BLUEPRINT_BODY_GAP + BLUEPRINT_HORIZONTAL_PADDING
    : canvasWidth + BLUEPRINT_HORIZONTAL_PADDING;
}

export interface BlueprintContentFrame {
  width: number;
  height: number;
  offsetX: number;
  offsetY: number;
}

function framedDimension(
  contentSize: number,
  authoredSize: number,
  minimumSize: number,
): number {
  if (contentSize > authoredSize) return Math.ceil(contentSize + 2 * BLUEPRINT_CONTENT_PADDING);
  return Math.ceil(Math.max(
    minimumSize,
    Math.min(authoredSize, contentSize + 2 * BLUEPRINT_CONTENT_PADDING),
  ));
}

export function calculateBlueprintContentFrame(
  data: BlueprintArchitecture,
): BlueprintContentFrame {
  const rectangles = [
    ...data.zones.map(zone => ({
      x1: zone.x,
      y1: zone.y,
      x2: zone.x + zone.width,
      y2: zone.y + zone.height,
    })),
    ...data.nodes.map(node => ({
      x1: node.x,
      y1: node.y,
      x2: node.x + BLUEPRINT_NODE_WIDTH,
      y2: node.y + BLUEPRINT_NODE_HEIGHT,
    })),
  ];

  if (rectangles.length === 0) {
    return {
      width: data.canvas.width,
      height: data.canvas.height,
      offsetX: 0,
      offsetY: 0,
    };
  }

  const minX = Math.min(...rectangles.map(rect => rect.x1));
  const minY = Math.min(...rectangles.map(rect => rect.y1));
  const maxX = Math.max(...rectangles.map(rect => rect.x2));
  const maxY = Math.max(...rectangles.map(rect => rect.y2));
  const contentWidth = Math.max(1, maxX - minX);
  const contentHeight = Math.max(1, maxY - minY);
  const width = framedDimension(contentWidth, data.canvas.width, BLUEPRINT_MIN_WIDTH);
  const height = framedDimension(contentHeight, data.canvas.height, BLUEPRINT_MIN_HEIGHT);

  return {
    width,
    height,
    offsetX: (width - contentWidth) / 2 - minX,
    offsetY: (height - contentHeight) / 2 - minY,
  };
}
