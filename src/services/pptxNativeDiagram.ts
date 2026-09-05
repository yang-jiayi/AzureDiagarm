// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import JSZip from 'jszip';
import type PptxGenJS from 'pptxgenjs';
import { GEOMETRY_EA_FONT, GEOMETRY_LATIN_FONT } from './diagramExportGeometry';
import { nativizePackage } from './pptxNativeShapes';
import { embedVectorIcons } from './pptxVectorIcons';

function repairTableShapeIds(xml: string): string {
  const used = new Map<string, number>();
  let highest = 0;
  for (const match of xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)) {
    used.set(match[1], (used.get(match[1]) ?? 0) + 1);
    highest = Math.max(highest, Number(match[1]));
  }
  // PptxGenJS 4 derives table IDs from slide numbers, which can collide with
  // the slide's chrome. Tables have no connector attachments; only their
  // duplicate nonvisual IDs change, leaving every glued shape ID untouched.
  return xml.replace(/<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g, frame =>
    frame.replace(/(<p:cNvPr\b[^>]*\bid=")(\d+)(")/, (tag, head: string, id: string, tail: string) => {
      if ((used.get(id) ?? 0) < 2) return tag;
      used.set(id, used.get(id)! - 1);
      highest += 1;
      return `${head}${highest}${tail}`;
    }));
}

/**
 * The download and Blob APIs share the established, fidelity-preserving
 * shape repair and SVG extension chain. No second diagram renderer lives here.
 */
export async function buildNativePptxBlob(
  pptx: PptxGenJS, vectorIcons: Map<string, string> = new Map(),
): Promise<Blob> {
  const output = await pptx.write({ outputType: 'arraybuffer' });
  if (!(output instanceof ArrayBuffer)) {
    throw new Error('PowerPoint generation did not produce an ArrayBuffer.');
  }
  const zip = await nativizePackage(await JSZip.loadAsync(output), {
    latin: GEOMETRY_LATIN_FONT,
    ea: GEOMETRY_EA_FONT,
  });
  await embedVectorIcons(zip, vectorIcons);
  for (const name of Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))) {
    const xml = await zip.file(name)!.async('string');
    const fixed = repairTableShapeIds(xml);
    if (fixed !== xml) zip.file(name, fixed);
  }
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}
