// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Vector icons for the exported deck.
 *
 * The icons on the canvas are SVG, but a slide could only carry them as a
 * rasterised PNG: pptxgenjs has no way to emit anything else. At the size a
 * tile draws an icon that PNG works out at roughly 213 dpi, so it is soft the
 * moment anyone projects the deck large, prints it, or -- most commonly --
 * pulls one service out of the diagram and scales it up for a title slide.
 *
 * OOXML does not make this a choice between raster and vector. A picture is a
 * raster blip that may carry the vector original in an extension, and each
 * version of PowerPoint renders whichever it understands: 2016 and later draw
 * the SVG at whatever size the shape happens to be, and anything older ignores
 * the extension it does not know and draws the PNG exactly as before. So the
 * deck ships both and nothing regresses.
 *
 * This runs on the finished package, for the same reason the shape conversion
 * does -- pptxgenjs cannot express it, and a deck that failed to gain vector
 * icons is very much better than no deck at all, so every failure here leaves
 * the package untouched.
 */

/** The extension namespace and uri Office itself writes for an SVG picture. */
const SVG_EXT_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

interface ZipLike {
  files: Record<string, unknown>;
  file(path: string): { async(type: 'string'): Promise<string> } | null;
  file(path: string, data: string): unknown;
}

/** `<p:pic>` blocks, with the name and blip each one carries. */
function picBlocks(slideXml: string): Array<{ start: number; end: number; xml: string }> {
  const out: Array<{ start: number; end: number; xml: string }> = [];
  const open = /<p:pic>/g;
  let hit: RegExpExecArray | null;
  while ((hit = open.exec(slideXml)) !== null) {
    const end = slideXml.indexOf('</p:pic>', hit.index);
    if (end === -1) continue;
    const stop = end + '</p:pic>'.length;
    out.push({ start: hit.index, end: stop, xml: slideXml.slice(hit.index, stop) });
    open.lastIndex = stop;
  }
  return out;
}

/**
 * The highest `rId` already spoken for.
 *
 * Numbering a new relationship from the *count* rather than the maximum is the
 * classic way to produce a duplicate id: parts are not always numbered
 * contiguously, and a duplicate silently repoints an existing picture.
 */
function maxRelId(relsXml: string): number {
  let max = 0;
  for (const hit of relsXml.matchAll(/Id="rId(\d+)"/g)) {
    max = Math.max(max, Number(hit[1]));
  }
  return max;
}

function relsPathFor(slidePath: string): string {
  return slidePath.replace(/([^/]+)$/, '_rels/$1.rels');
}

/**
 * Give every icon picture on every slide its vector original.
 *
 * `svgByName` is keyed by the `objectName` the exporter gave the picture, which
 * is the only stable way back from a shape in the finished package to the icon
 * it was drawn from.
 */
export async function embedVectorIcons<T extends ZipLike>(
  zip: T,
  svgByName: Map<string, string>,
): Promise<T> {
  if (svgByName.size === 0) return zip;

  const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  for (const slidePath of slides) {
    const slideEntry = zip.file(slidePath);
    if (!slideEntry || typeof (slideEntry as { async?: unknown }).async !== 'function') continue;
    const slideXml = await (slideEntry as { async(type: 'string'): Promise<string> }).async('string');

    const relsPath = relsPathFor(slidePath);
    const relsEntry = zip.file(relsPath);
    if (!relsEntry || typeof (relsEntry as { async?: unknown }).async !== 'function') continue;
    let relsXml = await (relsEntry as { async(type: 'string'): Promise<string> }).async('string');

    const slideNumber = /slide(\d+)\.xml$/.exec(slidePath)?.[1] ?? '0';
    let nextRel = maxRelId(relsXml);
    // One media part per distinct icon per slide: a diagram repeats the same
    // service icon constantly, and a part per picture would multiply the file
    // size by the repetition for no gain.
    const relForSvg = new Map<string, string>();
    const added: string[] = [];

    let out = '';
    let cursor = 0;
    let changed = false;

    for (const pic of picBlocks(slideXml)) {
      const name = /<p:cNvPr\b[^>]*\bname="([^"]*)"/.exec(pic.xml)?.[1];
      const svg = name ? svgByName.get(name) : undefined;
      if (!svg) continue;

      // The blip is matched *with* whatever it contains, not just when it is
      // empty. pptxgenjs writes an empty one today, but a blip legitimately
      // carries children -- `<a:alphaModFix/>` for transparency, a duotone
      // recolour -- and an empty-only match would silently skip those pictures
      // and leave them raster with no sign that anything had gone wrong.
      const blip = /<a:blip\b[^>]*?\/>|<a:blip\b[^>]*?>[\s\S]*?<\/a:blip>/.exec(pic.xml);
      if (!blip) continue;
      // Never touch a blip that already carries an extension list: it is either
      // already vector or something this code did not write, and appending a
      // second `<a:extLst>` to one blip is invalid OOXML.
      if (blip[0].includes('<a:extLst>')) continue;
      const embed = /r:embed="([^"]+)"/.exec(blip[0])?.[1];
      if (!embed) continue;

      let rel = relForSvg.get(svg);
      if (!rel) {
        nextRel += 1;
        rel = `rId${nextRel}`;
        const target = `../media/vector-${slideNumber}-${relForSvg.size + 1}.svg`;
        zip.file(`ppt/media/vector-${slideNumber}-${relForSvg.size + 1}.svg`, svg);
        added.push(`<Relationship Id="${rel}" Type="${IMAGE_REL_TYPE}" Target="${target}"/>`);
        relForSvg.set(svg, rel);
      }

      // Whatever the blip already contained is kept and the extension appended
      // after it: `<a:extLst>` is last in the schema's sequence, and rebuilding
      // the element from just its `r:embed` would quietly drop a transparency
      // or recolour the diagram had asked for.
      const inner = blip[0].endsWith('/>') ? '' : blip[0].replace(/^<a:blip\b[^>]*?>/, '').replace(/<\/a:blip>$/, '');
      const withSvg =
        `<a:blip r:embed="${embed}">${inner}<a:extLst><a:ext uri="${SVG_EXT_URI}">` +
        `<asvg:svgBlip xmlns:asvg="${SVG_NS}" r:embed="${rel}"/></a:ext></a:extLst></a:blip>`;
      const replaced = pic.xml.slice(0, blip.index) + withSvg
        + pic.xml.slice(blip.index + blip[0].length);

      out += slideXml.slice(cursor, pic.start) + replaced;
      cursor = pic.end;
      changed = true;
    }

    if (!changed) continue;
    zip.file(slidePath, out + slideXml.slice(cursor));
    relsXml = relsXml.replace('</Relationships>', `${added.join('')}</Relationships>`);
    zip.file(relsPath, relsXml);
  }

  return zip;
}
