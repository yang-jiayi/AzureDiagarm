/**
 * Per-character advance widths, read out of the font file itself.
 *
 * The tables in `diagramExportGeometry.ts` and `exportQualityAudit.ts` decide
 * whether a name fits its tile, so a wrong number there is a label painted
 * outside its box -- the exact defect that made exports need hand repair. They
 * were originally measured through GDI+ `MeasureString`, which is a rasteriser:
 * it answers in whole device pixels, so the number it gives depends on the
 * point size and the DPI it was asked at, and recovering the design advance
 * means dividing out a fitting error. That is why the old comments talk about
 * "20 repeats at 100pt".
 *
 * This reads `hmtx` instead, which is where the advance the rasteriser is
 * quantising actually comes from. It is exact, it needs no display, and it runs
 * the same on any machine -- so the numbers can be regenerated and diffed
 * rather than trusted.
 *
 * Deliberately self-contained. A font parser is a few hundred lines of table
 * offsets and this repo would otherwise take a dependency that runs on
 * arbitrary binary input for a script that is run by hand a few times a year.
 *
 * Usage:
 *   npx tsx scripts/measureFontAdvances.ts "Arial"
 *   npx tsx scripts/measureFontAdvances.ts "Yu Gothic UI" --verify
 */
import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';

const FONT_DIRS = [
  path.join(process.env.WINDIR ?? 'C:/Windows', 'Fonts'),
  path.join(process.env.LOCALAPPDATA ?? '', 'Microsoft/Windows/Fonts'),
].filter(Boolean);

interface Sfnt {
  buf: Buffer;
  tables: Map<string, { offset: number; length: number }>;
}

/** The sfnt table directory at `base`, which is 0 for a plain font. */
function readSfnt(buf: Buffer, base: number): Sfnt {
  const numTables = buf.readUInt16BE(base + 4);
  const tables = new Map<string, { offset: number; length: number }>();
  for (let i = 0; i < numTables; i += 1) {
    const rec = base + 12 + i * 16;
    tables.set(buf.toString('latin1', rec, rec + 4), {
      offset: buf.readUInt32BE(rec + 8),
      length: buf.readUInt32BE(rec + 12),
    });
  }
  return { buf, tables };
}

/** Every font in a file: one for a .ttf/.otf, several for a .ttc collection. */
function readFonts(file: string): Sfnt[] {
  const buf = readFileSync(file);
  if (buf.length < 12) return [];
  if (buf.toString('latin1', 0, 4) === 'ttcf') {
    const count = buf.readUInt32BE(8);
    const out: Sfnt[] = [];
    for (let i = 0; i < count; i += 1) out.push(readSfnt(buf, buf.readUInt32BE(12 + i * 4)));
    return out;
  }
  return [readSfnt(buf, 0)];
}

/**
 * The family and style a document would name.
 *
 * Prefers the typographic names (16/17) over the legacy pair (1/2), because
 * the legacy pair lies about exactly the case that matters here: the semilight
 * face of Yu Gothic UI calls itself family "Yu Gothic UI Semilight", style
 * "Regular", so matching on 1/2 accepts it as the regular weight and every
 * advance comes out too narrow.
 *
 * "Yu Gothic UI" and "Yu Gothic" also live in ONE .ttc, so the file name cannot
 * be used to tell them apart either.
 */
function faceIdentity(font: Sfnt): { family: string; style: string } {
  const t = font.tables.get('name');
  if (!t) return { family: '', style: '' };
  const { buf } = font;
  const base = t.offset;
  const count = buf.readUInt16BE(base + 2);
  const storage = base + buf.readUInt16BE(base + 4);
  const byId = new Map<number, string>();
  for (let i = 0; i < count; i += 1) {
    const rec = base + 6 + i * 12;
    const platform = buf.readUInt16BE(rec);
    const language = buf.readUInt16BE(rec + 4);
    const nameId = buf.readUInt16BE(rec + 6);
    if (nameId !== 1 && nameId !== 2 && nameId !== 16 && nameId !== 17) continue;
    // Windows/Unicode records are UTF-16BE; language 0x409 is US English, which
    // is the name a document written on an English or a Japanese Windows both
    // resolve. The Japanese record says 游ゴシック, which no .pptx here names.
    if (platform !== 3 || language !== 0x409) continue;
    const len = buf.readUInt16BE(rec + 8);
    const off = storage + buf.readUInt16BE(rec + 10);
    let name = '';
    for (let j = 0; j + 1 < len; j += 2) name += String.fromCharCode(buf.readUInt16BE(off + j));
    if (!byId.has(nameId)) byId.set(nameId, name);
  }
  return { family: byId.get(16) ?? byId.get(1) ?? '', style: byId.get(17) ?? byId.get(2) ?? '' };
}

/** Advance widths in font units, indexed by glyph id. */
function advances(font: Sfnt): number[] {
  const { buf } = font;
  const hhea = font.tables.get('hhea')!;
  const maxp = font.tables.get('maxp')!;
  const hmtx = font.tables.get('hmtx')!;
  const numMetrics = buf.readUInt16BE(hhea.offset + 34);
  const numGlyphs = buf.readUInt16BE(maxp.offset + 4);
  const out = new Array<number>(numGlyphs);
  let last = 0;
  for (let g = 0; g < numGlyphs; g += 1) {
    if (g < numMetrics) last = buf.readUInt16BE(hmtx.offset + g * 4);
    out[g] = last;
  }
  return out;
}

/** Code point to glyph id, from the best Unicode cmap subtable present. */
function charMap(font: Sfnt): Map<number, number> {
  const { buf } = font;
  const map = new Map<number, number>();
  const cmapBase = font.tables.get('cmap')!.offset;
  const n = buf.readUInt16BE(cmapBase + 2);
  let best = -1;
  let bestScore = -1;
  for (let i = 0; i < n; i += 1) {
    const rec = cmapBase + 4 + i * 8;
    const platform = buf.readUInt16BE(rec);
    const encoding = buf.readUInt16BE(rec + 2);
    const offset = cmapBase + buf.readUInt32BE(rec + 4);
    if (offset + 2 > buf.length) continue;
    const format = buf.readUInt16BE(offset);
    // Score on the SUBTABLE FORMAT, not on platform/encoding. Yu Gothic
    // publishes (0,5) as a format 14 variation-selector table, which carries no
    // code-point-to-glyph mapping at all; picking it by encoding number yielded
    // an empty map and a verification pass with nothing verified.
    const unicode = platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10));
    if (!unicode) continue;
    const score = format === 12 ? 2 : format === 4 ? 1 : -1;
    if (score > bestScore) { bestScore = score; best = offset; }
  }
  if (best < 0) return map;
  const format = buf.readUInt16BE(best);
  if (format === 4) {
    const segX2 = buf.readUInt16BE(best + 6);
    const ends = best + 14;
    const starts = ends + segX2 + 2;
    const deltas = starts + segX2;
    const ranges = deltas + segX2;
    for (let s = 0; s < segX2 / 2; s += 1) {
      const end = buf.readUInt16BE(ends + s * 2);
      const start = buf.readUInt16BE(starts + s * 2);
      const delta = buf.readInt16BE(deltas + s * 2);
      const rangeOff = buf.readUInt16BE(ranges + s * 2);
      if (start === 0xffff) continue;
      for (let c = start; c <= end && c !== 0x10000; c += 1) {
        let g: number;
        if (rangeOff === 0) g = (c + delta) & 0xffff;
        else {
          const at = ranges + s * 2 + rangeOff + (c - start) * 2;
          if (at + 1 >= buf.length) continue;
          g = buf.readUInt16BE(at);
          if (g !== 0) g = (g + delta) & 0xffff;
        }
        if (g !== 0) map.set(c, g);
      }
    }
  } else if (format === 12) {
    const groups = buf.readUInt32BE(best + 12);
    for (let i = 0; i < groups; i += 1) {
      const rec = best + 16 + i * 12;
      const start = buf.readUInt32BE(rec);
      const end = buf.readUInt32BE(rec + 4);
      const startGlyph = buf.readUInt32BE(rec + 8);
      // Guard a hostile or novel font rather than spinning on a huge range.
      if (end - start > 0x20000) continue;
      for (let c = start; c <= end; c += 1) map.set(c, startGlyph + (c - start));
    }
  }
  return map;
}

export interface MeasuredFont {
  family: string;
  file: string;
  unitsPerEm: number;
  /** Advance in em, or undefined when the font has no glyph for it. */
  em(codePoint: number): number | undefined;
  /**
   * The three line-height answers a font gives, all in em.
   *
   * A font states its own line spacing twice and the two rarely agree, so a
   * renderer's default leading depends on which one it believes. Word and
   * PowerPoint use the OS/2 win metrics; a browser uses hhea. Both are
   * reported rather than picked, because the exporters pin their line-height
   * rules to these numbers and a rule justified by the wrong one is a rule
   * with no evidence behind it.
   */
  vmetrics: { hhea: number; win: number; typo: number };
  /**
   * The OS/2 identity fields a Visio `<FaceName>` element has to carry.
   *
   * Visio matches a face by name first, but falls back to these when the name
   * is absent - so wrong values here are a wrong font on any machine that does
   * not have the named one, which is exactly the machine the fallback exists
   * for. Copied from the font file rather than typed out.
   */
  panose: string;
  unicodeRanges: string;
  codePages: string;
}

/** Load a family by the name a document would use, searching the font dirs. */
export function loadFamily(family: string): MeasuredFont {
  const wanted = family.toLowerCase();
  for (const dir of FONT_DIRS) {
    let entries: string[];
    try { entries = readdirSync(dir); } catch { continue; }
    for (const entry of entries) {
      if (!/\.(ttf|otf|ttc|otc)$/i.test(entry)) continue;
      const file = path.join(dir, entry);
      let fonts: Sfnt[];
      try { fonts = readFonts(file); } catch { continue; }
      for (const font of fonts) {
        const id = faceIdentity(font);
        if (id.family.toLowerCase() !== wanted) continue;
        // Regular only. The tables price unstyled text, and a semilight or
        // medium face out of the same collection would silently shift every
        // number without any signal that the wrong file was read.
        if (id.style.toLowerCase() !== 'regular') continue;
        if (!font.tables.has('head') || !font.tables.has('hmtx')
          || !font.tables.has('hhea') || !font.tables.has('maxp')
          || !font.tables.has('cmap')) continue;
        const unitsPerEm = font.buf.readUInt16BE(font.tables.get('head')!.offset + 18);
        const adv = advances(font);
        const map = charMap(font);
        const hheaOff = font.tables.get('hhea')!.offset;
        const os2 = font.tables.get('OS/2');
        const hheaLine = (font.buf.readInt16BE(hheaOff + 4) - font.buf.readInt16BE(hheaOff + 6)
          + font.buf.readInt16BE(hheaOff + 8)) / unitsPerEm;
        const win = os2
          ? (font.buf.readUInt16BE(os2.offset + 74) + font.buf.readUInt16BE(os2.offset + 76)) / unitsPerEm
          : hheaLine;
        const typo = os2
          ? (font.buf.readInt16BE(os2.offset + 68) - font.buf.readInt16BE(os2.offset + 70)
            + font.buf.readInt16BE(os2.offset + 72)) / unitsPerEm
          : hheaLine;
        return {
          family,
          file,
          unitsPerEm,
          vmetrics: { hhea: hheaLine, win, typo },
          panose: os2
            ? [...font.buf.subarray(os2.offset + 32, os2.offset + 42)].join(' ')
            : '0 0 0 0 0 0 0 0 0 0',
          unicodeRanges: os2
            ? [42, 46, 50, 54].map((d) => font.buf.readInt32BE(os2.offset + d)).join(' ')
            : '0 0 0 0',
          codePages: os2 && font.buf.readUInt16BE(os2.offset) >= 1
            ? [78, 82].map((d) => font.buf.readInt32BE(os2.offset + d)).join(' ')
            : '0 0',
          em: (cp: number) => {
            const g = map.get(cp);
            return g === undefined ? undefined : adv[g] / unitsPerEm;
          },
        };
      }
    }
  }
  throw new Error(`no regular face found for family "${family}" in ${FONT_DIRS.join(', ')}`);
}

/** The committed Yu Gothic UI table, so this tool can be checked against it. */
const COMMITTED_ASCII = [
  0.284, 0.392, 0.591, 0.539, 0.818, 0.800, 0.230, 0.302, 0.302, 0.417,
  0.684, 0.217, 0.400, 0.217, 0.390, 0.539, 0.539, 0.539, 0.539, 0.539,
  0.539, 0.539, 0.539, 0.539, 0.539, 0.217, 0.217, 0.684, 0.684, 0.684,
  0.448, 0.955, 0.645, 0.573, 0.619, 0.701, 0.506, 0.488, 0.686, 0.710,
  0.266, 0.357, 0.580, 0.471, 0.898, 0.748, 0.754, 0.560, 0.754, 0.598,
  0.531, 0.524, 0.687, 0.621, 0.934, 0.590, 0.553, 0.570, 0.302, 0.539,
  0.302, 0.684, 0.415, 0.268, 0.509, 0.588, 0.462, 0.589, 0.523, 0.313,
  0.589, 0.566, 0.242, 0.242, 0.497, 0.242, 0.861, 0.566, 0.586, 0.588,
  0.589, 0.348, 0.424, 0.339, 0.566, 0.479, 0.723, 0.459, 0.484, 0.452,
  0.302, 0.239, 0.302, 0.684,
];

function round(n: number): number { return Math.round(n * 1000) / 1000; }

/**
 * What a Windows renderer reaches for when the label font has no glyph.
 *
 * Ordered the way the shell's font linking is: the UI font first, then the
 * script-specific UI faces, then the Japanese face, then the symbol faces. The
 * previous tables recorded a substitute advance without recording WHICH font
 * supplied it, which is why their comment can only say the number is "weaker
 * evidence"; naming the donor makes the same number auditable.
 */
const SUBSTITUTE_CHAIN = [
  'Segoe UI',
  'Leelawadee UI',
  'Nirmala UI',
  'Yu Gothic UI',
  'Segoe UI Symbol',
  'Segoe UI Historic',
];

interface Donor { font: MeasuredFont; em: number }

function resolveAdvance(primary: MeasuredFont, chain: MeasuredFont[], cp: number): Donor | undefined {
  const own = primary.em(cp);
  if (own !== undefined) return { font: primary, em: own };
  for (const font of chain) {
    const em = font.em(cp);
    if (em !== undefined) return { font, em };
  }
  return undefined;
}

/**
 * The rules that price a character WITHOUT consulting a table, mirrored here.
 *
 * CJK is charged a full em by construction and every Japanese Gothic face puts
 * ideographs on a full-em grid, so those code points are font-independent and a
 * font migration must not touch them. Counting them as "coverage" drowns the
 * comparison: the first run of this analysis reported 41,230 covered code
 * points, of which some 40,000 were ideographs that no table has ever held.
 */
const RULE_PRICED_RE = /[\u2e80-\u9fff\uac00-\ud7af\uff00-\uff60\uffe0-\uffe6]|[\u200b-\u200f\u2060\ufe00-\ufe0f\ufeff]|[\p{Mn}\p{Me}]|\s/u;

/** Never draws, so never advances, whatever the font file's `hmtx` claims. */
const DEFAULT_IGNORABLE_RE = /\p{Default_Ignorable_Code_Point}/u;

async function coverage(font: MeasuredFont): Promise<void> {
  // Ask the SHIPPING model which code points it prices, rather than re-typing
  // the ranges out of its source. `advanceTier` already distinguishes what the
  // label font draws from what a substitute draws, so the new tables can cover
  // exactly what the old ones covered and the diff is a real comparison.
  const geometry = await import('../src/services/diagramExportGeometry.js');
  const chain = SUBSTITUTE_CHAIN.filter((n) => n !== font.family).map((n) => {
    try { return loadFamily(n); } catch { return undefined; }
  }).filter((f): f is MeasuredFont => Boolean(f));
  process.stderr.write(`substitutes: ${chain.map((c) => c.family).join(', ')}\n`);

  let label = 0;
  let substitute = 0;
  const nowNative: number[] = [];
  const nowSubstituted: number[] = [];
  const lost: number[] = [];
  for (let cp = 0x21; cp <= 0xffff; cp += 1) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const ch = String.fromCodePoint(cp);
    if (RULE_PRICED_RE.test(ch)) continue;
    const tier = geometry.advanceTier(ch);
    if (tier === 'none') continue;
    if (tier === 'label') label += 1; else substitute += 1;
    const own = font.em(cp) !== undefined;
    if (own && tier === 'substitute') nowNative.push(cp);
    if (!own && tier === 'label') nowSubstituted.push(cp);
    if (!own && !resolveAdvance(font, chain, cp)) lost.push(cp);
  }
  const hex = (list: number[]) => list.slice(0, 24).map((c) => `U+${c.toString(16).toUpperCase().padStart(4, '0')}`).join(' ');
  process.stdout.write(
    `tabled today: ${label} by the label font, ${substitute} by a substitute\n`
    + `${font.family} draws natively what used to be substituted: ${nowNative.length}  ${hex(nowNative)}\n`
    + `${font.family} needs a substitute where the old font drew it: ${nowSubstituted.length}  ${hex(nowSubstituted)}\n`
    + `no font in the chain has: ${lost.length}  ${hex(lost)}\n`,
  );
}

/**
 * The same advances, measured by a browser instead of by reading `hmtx`.
 *
 * The export audit deliberately keeps its OWN copy of every table so it can
 * detect the exporter's tables drifting; a copy generated by the same code path
 * would make that check vacuous. So the audit's numbers come from here: Chromium
 * lays the text out through DirectWrite and answers in fractional pixels, which
 * is a different engine and a different API from the font file this script
 * otherwise reads.
 *
 * It is also the more relevant oracle for the product's actual claim. The
 * diagram on screen is drawn by a browser, so a number the browser agrees with
 * is a number the exported file and the screen will both honour.
 *
 * Measured one character at a time with kerning off, because the width model
 * sums per-character advances. Asking for 20 repeats - the trick the GDI+
 * measurement needed to divide out whole-pixel rounding, which canvas does not
 * suffer from - measures something the model never charges: `"f".repeat(20)`
 * becomes ten `ﬀ` ligatures and answers 0.261 em against the font's 0.278, and
 * `"1".repeat(20)` kerns to 0.486 against 0.556. Both are correct answers to a
 * question about a RUN of that character, and both are the wrong number to put
 * in a table that is added up one character at a time.
 *
 * `donors` names the face to lay a SUBSTITUTED character out in, and without it
 * this oracle is not independent, it is unrelated. For a character the label
 * font lacks there is no such thing as its advance - there is only the advance
 * of whichever face gets asked - and left to itself Chromium picks a different
 * face from the one the file reader picks. The audit then priced Thai 6.5%
 * above the exporter and reported the exporter's own correct arithmetic as a
 * defect, on 8 strings, in a corpus whose entire purpose is to catch the
 * exporter mispricing text. Naming the donor puts BOTH methods on the same font
 * file, so what is left between them is the thing worth measuring: two engines
 * reading one file and having to agree.
 */
async function measureInBrowser(
  family: string,
  codePoints: number[],
  donors?: Map<number, string>,
): Promise<Map<number, number>> {
  const { chromium } = await import('playwright');
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const measured = await page.evaluate(({ font, points, size, subs }) => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('no 2d context');
      ctx.fontKerning = 'none';
      const donorFor = new Map(subs);
      const out: Array<[number, number]> = [];
      for (const cp of points) {
        const donor = donorFor.get(cp);
        ctx.font = donor ? `${size}px "${donor}", "${font}"` : `${size}px "${font}"`;
        out.push([cp, ctx.measureText(String.fromCodePoint(cp)).width / size]);
      }
      return out;
    }, {
      font: family,
      points: codePoints,
      size: 512,
      subs: [...(donors ?? new Map<number, string>())],
    });
    return new Map(measured);
  } finally {
    await browser.close();
  }
}

/** Consecutive code points grouped into the run form the tables use. */
function runs(values: Map<number, number>): Array<[number, number[]]> {
  const codes = [...values.keys()].sort((a, b) => a - b);
  const out: Array<[number, number[]]> = [];
  for (const code of codes) {
    const last = out[out.length - 1];
    if (last && code === last[0] + last[1].length) last[1].push(values.get(code)!);
    else out.push([code, [values.get(code)!]]);
  }
  return out;
}

function num(n: number): string { return String(round(n)); }

/** Wrap a flat list of numbers at 10 per line, the way the tables are written. */
function block(list: number[]): string {
  const lines: string[] = [];
  for (let i = 0; i < list.length; i += 10) {
    lines.push(`  ${list.slice(i, i + 10).map(num).join(', ')},`);
  }
  return lines.join('\n');
}

function runBlock(list: Array<[number, number[]]>, donors?: Map<number, string>): string {
  return list.map(([start, values]) => {
    const note = donors?.get(start);
    return `${note ? `  // ${note}\n` : ''}  [0x${start.toString(16)}, [${values.map(num).join(', ')}]],`;
  }).join('\n');
}

async function tables(font: MeasuredFont): Promise<void> {
  const geometry = await import('../src/services/diagramExportGeometry.js');
  const chain = SUBSTITUTE_CHAIN.filter((n) => n !== font.family).map((n) => {
    try { return loadFamily(n); } catch { return undefined; }
  }).filter((f): f is MeasuredFont => Boolean(f));

  const need = (cp: number): Donor => {
    const found = resolveAdvance(font, chain, cp);
    if (!found) throw new Error(`no font in the chain draws U+${cp.toString(16)}`);
    // A default-ignorable code point advances nothing: a soft hyphen draws only
    // where a renderer chooses to break, and a format character never draws at
    // all. `hmtx` still carries a width for them - Arial's soft hyphen says
    // 0.333 em where Yu Gothic UI's said 0 - so reading the file without this
    // rule would have charged a third of an em for ink that is never laid down,
    // purely as a side effect of changing font. Chromium answers 0 for all of
    // them, which is the behaviour being modelled.
    if (DEFAULT_IGNORABLE_RE.test(String.fromCodePoint(cp))) return { font: found.font, em: 0 };
    return found;
  };

  const ascii: number[] = [];
  for (let cp = 33; cp <= 126; cp += 1) ascii.push(need(cp).em);

  const latin: number[] = [];
  for (let cp = 0xa1; cp <= 0x17f; cp += 1) latin.push(need(cp).em);

  // Above U+017F, keep exactly the code points the shipping model already
  // prices, so a font change can never quietly narrow coverage. Split them by
  // WHO measured them rather than by a hand-written list of scripts: with Arial
  // the split moves - Hebrew and Arabic become the label font's own glyphs, and
  // the subscripts and a few maths symbols become the substitute's.
  const wide = new Map<number, number>();
  const fallback = new Map<number, number>();
  const donorOf = new Map<number, string>();
  for (let cp = 0x180; cp <= 0xffff; cp += 1) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const ch = String.fromCodePoint(cp);
    if (RULE_PRICED_RE.test(ch)) continue;
    if (geometry.advanceTier(ch) === 'none') continue;
    const own = font.em(cp);
    if (own !== undefined) wide.set(cp, need(cp).em);
    else { const d = need(cp); fallback.set(cp, d.em); donorOf.set(cp, d.font.family); }
  }

  // The extras map is a fixed, hand-chosen key set: the punctuation and arrows
  const extraKeys = [
    '\u00a0', '\u00b7', '\u00d7', '\u2013', '\u2014', '\u2018', '\u2019',
    '\u201c', '\u201d', '\u2026', '\u2190', '\u2192', '\u2194', '\u21d2',
    '\u2212', '\u2022',
  ];
  const extras = new Map<string, number>();
  for (const ch of extraKeys) extras.set(ch, need(ch.codePointAt(0)!).em);
  let spaceEm = need(0x20).em;

  // The audit's copy has to come from a different measurement path than the
  // exporter's, or the drift check between the two models compares a number
  // with itself. Everything above is read out of the font file; everything
  // below is laid out by Chromium.
  if (process.argv.includes('--browser')) {
    const points = [
      0x20, ...extraKeys.map((c) => c.codePointAt(0)!),
      ...Array.from({ length: 94 }, (_, i) => 33 + i),
      ...Array.from({ length: 0x17f - 0xa1 + 1 }, (_, i) => 0xa1 + i),
      ...wide.keys(), ...fallback.keys(),
    ];
    const seen = await measureInBrowser(font.family, [...new Set(points)], donorOf);
    let worst = 0;
    let worstAt = 0;
    const disagreed: Array<[number, number, number]> = [];
    const take = (cp: number, fileValue: number): number => {
      const browserValue = seen.get(cp);
      if (browserValue === undefined) return fileValue;
      const diff = Math.abs(round(browserValue) - round(fileValue));
      if (diff > 0.002) {
        disagreed.push([cp, fileValue, browserValue]);
        if (diff > worst) { worst = diff; worstAt = cp; }
      }
      return browserValue;
    };
    spaceEm = take(0x20, spaceEm);
    for (const [ch, em] of extras) extras.set(ch, take(ch.codePointAt(0)!, em));
    for (let i = 0; i < ascii.length; i += 1) ascii[i] = take(33 + i, ascii[i]);
    for (let i = 0; i < latin.length; i += 1) latin[i] = take(0xa1 + i, latin[i]);
    for (const [cp, em] of wide) wide.set(cp, take(cp, em));
    for (const [cp, em] of fallback) fallback.set(cp, take(cp, em));
    process.stderr.write(
      `browser vs font file: ${disagreed.length} of ${seen.size} disagree by more than 0.002 em; `
      + `worst ${round(worst)} em at U+${worstAt.toString(16).toUpperCase()}\n`,
    );
    // Split the report by tier. A disagreement about a glyph the label font
    // OWNS is a measurement problem and has to be explained; a disagreement
    // about a glyph it lacks is the two methods choosing different substitute
    // faces, which is the uncertainty this model has always carried and states
    // openly rather than the two methods measuring the same thing differently.
    const own = disagreed.filter(([cp]) => !fallback.has(cp));
    process.stderr.write(
      `  of those, ${own.length} are glyphs ${font.family} itself contains, `
      + `${disagreed.length - own.length} are characters it lacks and the two methods substitute differently\n`,
    );
    for (const [cp, fileValue, browserValue] of disagreed) {
      process.stderr.write(
        `  U+${cp.toString(16).toUpperCase().padStart(4, '0')} file ${num(fileValue)} browser ${num(browserValue)}`
        + `${fallback.has(cp) ? ' (substituted)' : ''}\n`,
      );
    }
  }

  const fallbackRuns = runs(fallback);
  const donorNotes = new Map<number, string>();
  for (const [start, values] of fallbackRuns) {
    const names = new Set<string>();
    for (let i = 0; i < values.length; i += 1) names.add(donorOf.get(start + i)!);
    donorNotes.set(start, `drawn by ${[...names].join(' / ')}`);
  }

  process.stdout.write(
    `// measured from ${process.argv.includes('--browser') ? `Chromium laying out "${font.family}"` : `${font.file} (unitsPerEm ${font.unitsPerEm})`}\n`
    + `const ADVANCE_EM = [\n${block(ascii)}\n];\n\n`
    + `const SPACE_EM = ${num(spaceEm)};\n\n`
    + `const EXTRA_EM: Record<string, number> = {\n`
    + [...extras].map(([ch, em]) => `  '\\u${ch.codePointAt(0)!.toString(16).padStart(4, '0')}': ${num(em)},`).join('\n')
    + `\n};\n\n`
    + `const LATIN_EM = [\n${block(latin)}\n];\n\n`
    + `const WIDE_EM = [\n${runBlock(runs(wide))}\n];\n\n`
    + `const FALLBACK_EM = [\n${runBlock(fallbackRuns, donorNotes)}\n];\n`,
  );
}

async function main(): Promise<void> {
  const family = process.argv[2] ?? 'Arial';
  const font = loadFamily(family);
  process.stderr.write(`${font.family}  ${font.file}  unitsPerEm=${font.unitsPerEm}\n`);

  if (process.argv.includes('--verify')) {
    // Reproducing the committed numbers is the whole warrant for this tool: if
    // reading `hmtx` agrees with what GDI+ was asked years ago, the two methods
    // are measuring the same thing and a new font can be trusted to this one.
    let worst = 0;
    let worstAt = '';
    let missing = 0;
    for (let cp = 33; cp <= 126; cp += 1) {
      const got = font.em(cp);
      // A missing glyph must FAIL, not be skipped. Skipping is how the first
      // run of this script reported "worst disagreement: 0" after checking
      // nothing at all, because a cmap bug had left the map empty.
      if (got === undefined) { missing += 1; process.stderr.write(`MISSING U+${cp.toString(16)}\n`); continue; }
      const diff = Math.abs(round(got) - COMMITTED_ASCII[cp - 33]);
      if (diff > worst) { worst = diff; worstAt = String.fromCodePoint(cp); }
    }
    process.stderr.write(`checked ${94 - missing}/94, missing ${missing}, worst ${round(worst)} em at "${worstAt}"\n`);
    process.exit(missing > 0 || worst > 0.002 ? 1 : 0);
  }

  if (process.argv.includes('--coverage')) { await coverage(font); return; }
  if (process.argv.includes('--tables')) { await tables(font); return; }
  if (process.argv.includes('--facename')) {
    process.stdout.write(
      `<FaceName ID="?" NameU="${font.family}" UnicodeRanges="${font.unicodeRanges}"`
      + ` CharSets="${font.codePages}" Panos="${font.panose}" Flags="325"/>\n`,
    );
    return;
  }
  if (process.argv.includes('--vmetrics')) {
    const v = font.vmetrics;
    process.stdout.write(
      `hhea ${v.hhea.toFixed(4)}  win ${v.win.toFixed(4)}  typo ${v.typo.toFixed(4)}\n`,
    );
    return;
  }

  const rows: string[] = [];
  for (let cp = 33; cp <= 126; cp += 8) {
    const line: string[] = [];
    for (let c = cp; c < Math.min(cp + 8, 127); c += 1) line.push(round(font.em(c) ?? 0).toFixed(3));
    rows.push(`  ${line.join(', ')},`);
  }
  process.stdout.write(`${rows.join('\n')}\n`);
}

if (process.argv[1] && path.basename(process.argv[1]) === 'measureFontAdvances.ts') {
  main().catch((error) => { process.stderr.write(`${String(error)}\n`); process.exit(1); });
}
