/**
 * Turn a pptxgenjs deck into shapes PowerPoint recognises as its own.
 *
 * pptxgenjs can only emit `<p:sp>`. That is enough to *draw* an architecture,
 * but not enough to *edit* one: every arrow is a loose line that stays behind
 * when the service it points at is dragged, and every service name is a
 * separate text box that stays behind when its tile is dragged. A recipient
 * therefore has to redraw the deck by hand, which is exactly the complaint
 * this module exists to answer.
 *
 * The library has no API for either fix — `stCxn`, `endCxn`, `cxnSp` and any
 * grouping primitive appear nowhere in pptxgenjs 4.x — so the deck is repaired
 * after it is written, by rewriting the slide XML in the .pptx zip.
 *
 * The transform is deliberately conservative: it never moves anything. A
 * connector is glued only where its endpoint already sits on the connection
 * site PowerPoint would have chosen, so the drawing is byte-identical in
 * appearance and only its editability changes.
 */

/**
 * Give East Asian text the East Asian face.
 *
 * pptxgenjs writes ONE name into all three script slots - `<a:latin>`, `<a:ea>`
 * and `<a:cs>` all get `opts.fontFace` (pptxgen.cjs.js, `genXmlTextRunProps`) -
 * and there is no option to separate them. That is fine while the deck names a
 * face that covers everything, and wrong the moment it names Arial: Arial has
 * no kana and no kanji, so `<a:ea typeface="Arial"/>` sends every Japanese run
 * to the theme's `Jpan` fallback, which is plain `Yu Gothic`. That face states
 * a 1.6021 hhea line height against Yu Gothic UI's 1.3301, i.e. 19% above the
 * 1.35 factor every line-height rule in the exporters reserves against - so a
 * Japanese label would silently spill out of a box the sizer had proved it fit.
 *
 * `<a:cs>` is deliberately left alone. Arial does contain Arabic and Hebrew,
 * and the width tables now measure both from Arial, so pointing complex scripts
 * anywhere else would break the agreement between what is priced and what is
 * drawn.
 *
 * Only the exact Latin name is rewritten, so a run that already carries a
 * deliberate East Asian face, or a theme reference like `+mn-ea`, is untouched.
 */
export function repairEastAsianFont(xml: string, latin: string, ea: string): string {
  const from = new RegExp(`(<a:ea\\b[^>]*\\btypeface=")${latin.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(")`, 'g');
  return xml.replace(from, `$1${ea}$2`);
}

const EMU_PER_IN = 914400;
/** A tenth of a line width. Endpoints are emitted as integers, so this only
 *  has to absorb rounding, not tolerate a genuinely different position. */
const SITE_TOLERANCE_EMU = Math.round(0.02 * EMU_PER_IN);

interface Xfrm {
  x: number;
  y: number;
  w: number;
  h: number;
  flipH: boolean;
  flipV: boolean;
}

interface ShapeXml {
  xml: string;
  kind: 'sp' | 'pic';
  start: number;
  end: number;
  id: number;
  name: string;
  xfrm: Xfrm | null;
  prst: string | null;
}

function readXfrm(xml: string): Xfrm | null {
  const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(xml);
  const ext = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/.exec(xml);
  if (!off || !ext) return null;
  const frame = /<a:xfrm([^>]*)>/.exec(xml)?.[1] ?? '';
  return {
    x: +off[1],
    y: +off[2],
    w: +ext[1],
    h: +ext[2],
    flipH: /flipH="1"/.test(frame),
    flipV: /flipV="1"/.test(frame),
  };
}

function parseShapes(slideXml: string): ShapeXml[] {
  const shapes: ShapeXml[] = [];
  // pptxgenjs never nests these, so a non-greedy scan is exact.
  const re = /<p:(sp|pic)>[\s\S]*?<\/p:\1>/g;
  let match: RegExpExecArray | null;
  while ((match = re.exec(slideXml)) !== null) {
    const xml = match[0];
    const cNvPr = /<p:cNvPr id="(\d+)" name="([^"]*)"/.exec(xml);
    if (!cNvPr) continue;
    shapes.push({
      xml,
      kind: match[1] as 'sp' | 'pic',
      start: match.index,
      end: match.index + xml.length,
      id: +cNvPr[1],
      name: cNvPr[2],
      xfrm: readXfrm(xml),
      prst: /<a:prstGeom prst="([^"]+)"/.exec(xml)?.[1] ?? null,
    });
  }
  return shapes;
}

/** Where a line actually starts and ends, after the flips that encode direction. */
function endpoints(f: Xfrm): { x1: number; y1: number; x2: number; y2: number } {
  const x1 = f.flipH ? f.x + f.w : f.x;
  const x2 = f.flipH ? f.x : f.x + f.w;
  const y1 = f.flipV ? f.y + f.h : f.y;
  const y2 = f.flipV ? f.y : f.y + f.h;
  return { x1, y1, x2, y2 };
}

/**
 * PowerPoint's four connection sites on a rectangle, in the order the format
 * numbers them. A preset rectangle has exactly these and nothing else, which
 * is why glue is only safe where the drawn endpoint already coincides with one.
 */
function sites(f: Xfrm): { idx: number; x: number; y: number }[] {
  return [
    { idx: 0, x: f.x + f.w / 2, y: f.y },
    { idx: 1, x: f.x, y: f.y + f.h / 2 },
    { idx: 2, x: f.x + f.w / 2, y: f.y + f.h },
    { idx: 3, x: f.x + f.w, y: f.y + f.h / 2 },
  ];
}

function glueFor(
  point: { x: number; y: number },
  tiles: ShapeXml[],
): { id: number; idx: number } | null {
  let best: { id: number; idx: number; d: number } | null = null;
  for (const tile of tiles) {
    if (!tile.xfrm) continue;
    for (const site of sites(tile.xfrm)) {
      const d = Math.hypot(site.x - point.x, site.y - point.y);
      if (d > SITE_TOLERANCE_EMU) continue;
      if (!best || d < best.d) best = { id: tile.id, idx: site.idx, d };
    }
  }
  return best ? { id: best.id, idx: best.idx } : null;
}

function isTile(name: string): boolean {
  return (
    name.startsWith('service-') &&
    !name.startsWith('service-label-') &&
    !name.startsWith('service-meta-')
  );
}

/** The container shape each floating caption belongs to, by object-name prefix. */
const CAPTIONS: { label: string; owner: (id: string) => string }[] = [
  { label: 'service-label-', owner: (id) => `service-${id}` },
  { label: 'zone-label-', owner: (id) => `zone-${id}` },
];

/** EMU insets that place `inner` exactly where it is today, inside `outer`. */
function insets(outer: Xfrm, inner: Xfrm): { l: number; t: number; r: number; b: number } {
  return {
    l: Math.max(0, Math.round(inner.x - outer.x)),
    t: Math.max(0, Math.round(inner.y - outer.y)),
    r: Math.max(0, Math.round(outer.x + outer.w - (inner.x + inner.w))),
    b: Math.max(0, Math.round(outer.y + outer.h - (inner.y + inner.h))),
  };
}

/**
 * Move a caption inside the shape it names, so dragging the shape takes the
 * words with it. The caption keeps its exact position: the space it used to
 * float in becomes the container's text insets.
 */
function foldLabels(shapes: ShapeXml[]): Map<number, string> {
  const byName = new Map<string, ShapeXml>();
  for (const s of shapes) if (s.xfrm) byName.set(s.name, s);

  const folded = new Map<number, string>();
  for (const label of shapes) {
    const kind = CAPTIONS.find((c) => label.name.startsWith(c.label));
    if (!kind || !label.xfrm) continue;
    const owner = byName.get(kind.owner(label.name.slice(kind.label.length)));
    if (!owner?.xfrm || /<p:txBody>/.test(owner.xml)) continue;
    const body = /<p:txBody>[\s\S]*<\/p:txBody>/.exec(label.xml)?.[0];
    if (!body) continue;
    // The caption must already lie inside its container, or folding it in
    // would move it. A zone title that sits above its box stays where it is.
    if (
      label.xfrm.x < owner.xfrm.x - 1 ||
      label.xfrm.y < owner.xfrm.y - 1 ||
      label.xfrm.x + label.xfrm.w > owner.xfrm.x + owner.xfrm.w + 1 ||
      label.xfrm.y + label.xfrm.h > owner.xfrm.y + owner.xfrm.h + 1
    ) {
      continue;
    }
    const pad = insets(owner.xfrm, label.xfrm);
    const anchored = body.replace(
      /<a:bodyPr([^>]*)>/,
      (_all, attrs: string) =>
        `<a:bodyPr${attrs
          .replace(/\s(lIns|tIns|rIns|bIns)="[^"]*"/g, '')
        } lIns="${pad.l}" tIns="${pad.t}" rIns="${pad.r}" bIns="${pad.b}">`,
    );
    folded.set(label.id, '');
    // Build on any caption already folded into this owner, not on the original
    // shape: a tile that owns two captions would otherwise keep only the last,
    // having already deleted the first.
    const base = folded.get(owner.id) || owner.xml;
    folded.set(owner.id, base.replace(/<\/p:sp>$/, `${anchored}</p:sp>`));
  }
  return folded;
}

/**
 * Rewrite a drawn straight line as a real PowerPoint connector, glued to the
 * services it joins wherever it already meets them at a connection site.
 *
 * Nothing moves; the arrow simply stops being an orphan when one of its
 * endpoints is dragged. A bent hop is deliberately left alone -- see below.
 */
function connectorXml(shape: ShapeXml, tiles: ShapeXml[]): string | null {
  if (!shape.name.startsWith('connector-') || !shape.xfrm) return null;
  const bent = /<a:custGeom>/.test(shape.xml);
  if (shape.prst !== 'line' && !bent) return null;
  // A connector shape may not carry custom geometry. PowerPoint does not merely
  // ignore an <a:custGeom> inside <p:cxnSp>: it refuses to open the package at
  // all, reporting "the file is corrupted and unreadable", so a single bent hop
  // cost the user the entire deck. Verified against PowerPoint 16.0 -- the same
  // deck opens once the geometry is preset, and opens once the hop is left as a
  // plain shape, with or without the glue.
  //
  // A bent hop therefore keeps its drawn route and stays an ordinary shape. The
  // alternative, a preset bentConnector3, would open but would let PowerPoint
  // re-route the hop, which is exactly the fidelity this exporter is for: what
  // is on the canvas is what has to come out. Editability is worth having, but
  // never at the price of the route or of the file.
  if (bent) return null;
  const ends = endpoints(shape.xfrm);
  const from = glueFor({ x: ends.x1, y: ends.y1 }, tiles);
  const to = glueFor({ x: ends.x2, y: ends.y2 }, tiles);
  // Tiles that touch exactly share an edge, so both ends of the hop between
  // them land on the same site of the same shape. Gluing that tells PowerPoint
  // the arrow starts and finishes in one place, which it has no sane way to
  // reroute, so leave the hop unglued and let it stay a plain line.
  if (from && to && from.id === to.id && from.idx === to.idx) return null;

  const spPr = /<p:spPr>[\s\S]*<\/p:spPr>/.exec(shape.xml)?.[0];
  if (!spPr) return null;
  const geom = spPr.replace('<a:prstGeom prst="line">', '<a:prstGeom prst="straightConnector1">');
  const style = /<p:style>[\s\S]*?<\/p:style>/.exec(shape.xml)?.[0] ?? '';
  const glue = `${from ? `<a:stCxn id="${from.id}" idx="${from.idx}"/>` : ''}${to ? `<a:endCxn id="${to.id}" idx="${to.idx}"/>` : ''}`;

  return (
    `<p:cxnSp><p:nvCxnSpPr>` +
    `<p:cNvPr id="${shape.id}" name="${shape.name}"/>` +
    `<p:cNvCxnSpPr>${glue}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
    `${geom}${style}</p:cxnSp>`
  );
}

/** A node id can contain any character a user can type, including regex ones. */
function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function decodeXml(text: string): string {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_all, code: string) => String.fromCodePoint(+code))
    .replace(/&amp;/g, '&');
}

function encodeAttr(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Every run of text in a shape, in reading order. */
function textOf(xml: string): string {
  return [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
    .map((match) => decodeXml(match[1]))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** PowerPoint shows alt text in a small pane; a paragraph in it is unreadable. */
const ALT_TEXT_MAX = 300;

function clip(text: string): string {
  return text.length <= ALT_TEXT_MAX ? text : `${text.slice(0, ALT_TEXT_MAX - 1).trimEnd()}…`;
}

/**
 * What each shape is, in words, for the accessibility checker and for anyone
 * reading the deck with a screen reader.
 *
 * A drawing made of rectangles and lines is silent: PowerPoint reads out the
 * text in a shape, so a tile announces its own name, but an arrow has no text
 * at all and a zone's title is a separate caption. The reader hears a list of
 * service names in z-order and nothing about how they connect — which is the
 * entire content of an architecture diagram. Every deck this tool exports also
 * fails PowerPoint's built-in accessibility check for the same reason.
 *
 * Only what can be stated from the drawing is stated. An arrow's endpoints are
 * named only when it is a straight hop whose ends sit on a connection site,
 * because that is the only case where the tile it touches is known; a bent hop
 * carries a bounding box whose corners can land on an unrelated tile, and
 * naming that tile would be an invented fact, which is worse than silence.
 */
function altTextByName(shapes: ShapeXml[], tiles: ShapeXml[]): Map<string, string> {
  const byName = new Map<string, ShapeXml>();
  for (const shape of shapes) byName.set(shape.name, shape);
  const tileName = new Map<number, string>();
  const alt = new Map<string, string>();

  for (const tile of tiles) {
    const key = tile.name.slice('service-'.length);
    const chip = new RegExp(`^tagtext-${escapeRegExp(key)}-\\d+$`);
    const name = textOf(byName.get(`service-label-${key}`)?.xml ?? '') || textOf(tile.xml);
    const meta = textOf(byName.get(`service-meta-${key}`)?.xml ?? '');
    const tags = shapes.filter((s) => chip.test(s.name)).map((s) => textOf(s.xml)).filter(Boolean);
    tileName.set(tile.id, name);
    if (!name) continue;
    const text = clip([`Service: ${name}`, meta, tags.length > 0 ? `Tags: ${tags.join(', ')}` : '']
      .filter(Boolean)
      .join('. '));
    alt.set(tile.name, text);
    // The group is what a screen reader reaches first once the tile, its icon
    // and its chips have been folded into one object.
    alt.set(`node-${key}`, text);
  }

  for (const zone of shapes) {
    if (!zone.name.startsWith('zone-') || zone.name.startsWith('zone-label-')) continue;
    const key = zone.name.slice('zone-'.length);
    const label = textOf(byName.get(`zone-label-${key}`)?.xml ?? '') || textOf(zone.xml);
    if (label) alt.set(zone.name, clip(`Zone: ${label}`));
  }

  for (const line of shapes) {
    if (!line.name.startsWith('connector-')) continue;
    if (line.name.startsWith('connector-label-') || line.name.startsWith('connector-step-')) continue;
    const key = line.name.slice('connector-'.length);
    const straight = !!line.xfrm && line.prst === 'line' && !/<a:custGeom>/.test(line.xml);
    const ends = straight && line.xfrm ? endpoints(line.xfrm) : null;
    const from = ends ? glueFor({ x: ends.x1, y: ends.y1 }, tiles) : null;
    const to = ends ? glueFor({ x: ends.x2, y: ends.y2 }, tiles) : null;
    const a = from ? tileName.get(from.id) : undefined;
    const b = to ? tileName.get(to.id) : undefined;
    const step = textOf(byName.get(`connector-step-${key}`)?.xml ?? '');
    const label = textOf(byName.get(`connector-label-${key}`)?.xml ?? '');
    const text = [
      step ? `Step ${step}.` : '',
      a && b && a !== b ? `Connection from ${a} to ${b}` : 'Connection',
      label ? `: ${label}` : '',
    ].filter(Boolean).join(' ').replace(' :', ':');
    alt.set(line.name, clip(text));
  }

  return alt;
}

/** Write `descr` on the shapes named in `alt`, leaving any existing one alone. */
function applyAltText(slideXml: string, alt: Map<string, string>): string {
  if (alt.size === 0) return slideXml;
  return slideXml.replace(
    /<p:cNvPr id="(\d+)" name="([^"]*)"([^>]*?)(\/?)>/g,
    (all, id: string, name: string, rest: string, close: string) => {
      if (/\bdescr="/.test(rest)) return all;
      const text = alt.get(name);
      return text ? `<p:cNvPr id="${id}" name="${name}" descr="${encodeAttr(text)}"${rest}${close}>` : all;
    },
  );
}

/**
 * Repair one slide. Returns the same XML when there is nothing to convert, so
 * callers can cheaply skip untouched parts.
 */
export function nativizeSlideXml(slideXml: string): string {
  const shapes = parseShapes(slideXml);
  if (shapes.length === 0) return slideXml;
  const tiles = shapes.filter((s) => s.kind === 'sp' && isTile(s.name) && s.xfrm);

  const replacements = new Map<number, string>(foldLabels(shapes.filter((s) => s.kind === 'sp')));
  for (const shape of shapes) {
    if (shape.kind !== 'sp') continue;
    const converted = connectorXml(shape, tiles);
    if (converted !== null) replacements.set(shape.id, converted);
  }

  // A tile and the icon drawn on it are one thing to the reader, so make them
  // one thing to PowerPoint. Without this, dragging a service leaves its icon
  // behind — which is most of what makes an exported deck unusable. The
  // category stripe is a separate shape for the same reason the icon is (a
  // rounded rect has no per-side border), so it has to be folded in too or
  // moving a tile leaves a bare coloured bar sitting on the slide.
  //
  // Anything drawn ON a tile has to be listed here or it strays, which is why
  // the tag chips are matched too — and matched by an anchored pattern rather
  // than a prefix, because there is no separator that cannot occur inside an
  // id: a bare `startsWith('tag-a-')` also swallows the chips of a node called
  // `a-1`, silently folding one service's tags into a different service's tile.
  let nextId = Math.max(0, ...shapes.map((s) => s.id)) + 1;
  for (const tile of tiles) {
    const key = tile.name.slice('service-'.length);
    const chip = new RegExp(`^tag(?:text)?-${escapeRegExp(key)}-\\d+$`);
    const parts = shapes.filter(
      (s) =>
        (s.name === `icon-${key}` ||
          s.name === `service-meta-${key}` ||
          s.name === `accent-${key}` ||
          chip.test(s.name)) &&
        s.xfrm,
    );
    if (parts.length === 0 || !tile.xfrm) continue;
    const body = replacements.get(tile.id) ?? tile.xml;
    const f = tile.xfrm;
    const frame =
      `<a:xfrm><a:off x="${f.x}" y="${f.y}"/><a:ext cx="${f.w}" cy="${f.h}"/>` +
      `<a:chOff x="${f.x}" y="${f.y}"/><a:chExt cx="${f.w}" cy="${f.h}"/></a:xfrm>`;
    replacements.set(
      tile.id,
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${nextId}" name="node-${key}"/>` +
        `<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${frame}</p:grpSpPr>` +
        `${body}${parts.map((p) => p.xml).join('')}</p:grpSp>`,
    );
    for (const part of parts) replacements.set(part.id, '');
    nextId += 1;
  }
  const alt = altTextByName(shapes, tiles);
  if (replacements.size === 0) return applyAltText(slideXml, alt);

  let out = '';
  let cursor = 0;
  for (const shape of shapes) {
    const replacement = replacements.get(shape.id);
    if (replacement === undefined) continue;
    out += slideXml.slice(cursor, shape.start) + replacement;
    cursor = shape.end;
  }
  return applyAltText(out + slideXml.slice(cursor), alt);
}

/**
 * Repair every slide in a written .pptx. `zip` is a loaded JSZip of the deck;
 * the same instance is returned so the caller can write it out in whatever
 * form its platform wants.
 */
export async function nativizePackage<T extends {
  files: Record<string, unknown>;
  file(path: string): { async(type: 'string'): Promise<string> } | null;
  file(path: string, data: string): unknown;
}>(zip: T, fonts?: { latin: string; ea: string }): Promise<T> {
  const slides = Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  for (const name of slides) {
    const entry = zip.file(name);
    if (!entry || typeof (entry as { async?: unknown }).async !== 'function') continue;
    const xml = await (entry as { async(type: 'string'): Promise<string> }).async('string');
    let fixed = nativizeSlideXml(xml);
    if (fonts) fixed = repairEastAsianFont(fixed, fonts.latin, fonts.ea);
    if (fixed !== xml) zip.file(name, fixed);
  }
  return zip;
}
