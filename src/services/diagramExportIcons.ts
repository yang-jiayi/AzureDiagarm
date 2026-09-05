// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Node } from 'reactflow';
import { collectExportBoxes } from './diagramExportGeometry';
import { rasterizeIconToPng, rasterizeIcons, type RasterizedIcon } from '../utils/exportIconRaster';

export type ExportIcon = RasterizedIcon;
export type ExportIcons = ReadonlyMap<string, RasterizedIcon>;

/** Explicit asset loading for callers that require every icon to be present. */
export async function loadExportIcon(iconPath: string): Promise<ExportIcon> {
  const icon = await rasterizeIconToPng(iconPath, 256);
  if (!icon) throw new Error(`Unable to load export icon: ${iconPath}`);
  return icon;
}

export async function loadDiagramIcons(nodes: readonly Node[]): Promise<ExportIcons> {
  const paths = [...new Set([...collectExportBoxes([...nodes]).values()]
    .flatMap(box => box.iconPath ? [box.iconPath] : []))];
  const icons = await rasterizeIcons(paths, 256);
  for (const path of paths) {
    if (!icons.has(path)) throw new Error(`Unable to load export icon: ${path}`);
  }
  return icons;
}
