// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Serpentine wrapping for flow layouts.
 *
 * Dagre gives a linear flow one rank per service, so a twelve-step
 * architecture comes back as a 4360x100 strip — 43:1. Dropped onto a 16:9
 * slide that leaves the drawing 0.3in tall and the page 97% empty, which is
 * why the exports read as unusable no matter how carefully they are rendered.
 *
 * The Azure Architecture Center never publishes a diagram shaped like that:
 * flows wrap, and the finished drawing stays close to the shape of the page it
 * has to sit on. This module folds an over-wide (or, for top-down flows,
 * over-tall) layout into balanced bands, alternating direction so the
 * connector joining two bands is a short hop instead of a sweep back across
 * the whole drawing.
 *
 * It works on opaque boxes, so the caller decides what a box is: an ungrouped
 * service, or a zone container that must move together with everything inside
 * it.
 */

export interface WrapBox {
  id: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface WrapOffset {
  dx: number;
  dy: number;
}

export interface WrapOptions {
  /** 'LR'/'RL' wrap a too-wide drawing into rows, 'TB'/'BT' a too-tall one into columns. */
  direction?: 'LR' | 'RL' | 'TB' | 'BT';
  /** Distance between two bands, in pixels. */
  bandGap?: number;
}

/** Roughly the shape of the documents these diagrams get embedded in. */
export const LEARN_TARGET_ASPECT = 16 / 9;

/**
 * Leave anything at or below this ratio alone. A 2.6:1 drawing still fits a
 * slide comfortably; wrapping it would churn layouts for no gain.
 */
export const WRAP_TRIGGER_RATIO = 2.6;

/** More bands than this stops reading as one flow. */
export const MAX_WRAP_BANDS = 6;

const DEFAULT_BAND_GAP = 200;

interface Rank {
  ids: string[];
  start: number;
  end: number;
  minorStart: number;
  minorEnd: number;
}

/**
 * Group boxes that share a position along the flow axis. Dagre gives every
 * node in a rank the same major coordinate, so overlap along that axis is a
 * faithful reconstruction of its ranking — and it degrades gracefully for
 * hand-moved diagrams, where "things stacked above each other" is exactly the
 * unit we want to keep together.
 */
function buildRanks(
  boxes: WrapBox[],
  major: (b: WrapBox) => number,
  majorSize: (b: WrapBox) => number,
  minor: (b: WrapBox) => number,
  minorSize: (b: WrapBox) => number,
): Rank[] {
  const sorted = [...boxes].sort((a, b) => major(a) - major(b) || a.id.localeCompare(b.id));
  const ranks: Rank[] = [];
  for (const box of sorted) {
    const start = major(box);
    const end = start + majorSize(box);
    const current = ranks[ranks.length - 1];
    if (current && start < current.end) {
      current.ids.push(box.id);
      current.end = Math.max(current.end, end);
      current.minorStart = Math.min(current.minorStart, minor(box));
      current.minorEnd = Math.max(current.minorEnd, minor(box) + minorSize(box));
      continue;
    }
    ranks.push({
      ids: [box.id],
      start,
      end,
      minorStart: minor(box),
      minorEnd: minor(box) + minorSize(box),
    });
  }
  return ranks;
}

/**
 * Split ranks into at most `bandCount` contiguous bands, keeping the widest
 * band as narrow as possible. Contiguity is what lets a band be reversed
 * later: the ranks either side of a seam were adjacent in the source layout,
 * so the gap between them is still meaningful.
 */
function packBands(ranks: Rank[], gapAfter: number[], bandCount: number): Rank[][] {
  const extents = ranks.map((r) => r.end - r.start);
  const total = extents.reduce((sum, e, i) => sum + e + (i < gapAfter.length ? gapAfter[i] : 0), 0);

  const fits = (budget: number): Rank[][] | null => {
    const bands: Rank[][] = [];
    let current: Rank[] = [];
    let extent = 0;
    for (let i = 0; i < ranks.length; i += 1) {
      const gap = current.length > 0 ? gapAfter[i - 1] ?? 0 : 0;
      const next = current.length === 0 ? extents[i] : extent + gap + extents[i];
      if (current.length > 0 && next > budget) {
        bands.push(current);
        current = [ranks[i]];
        extent = extents[i];
      } else {
        current.push(ranks[i]);
        extent = next;
      }
    }
    if (current.length > 0) bands.push(current);
    return bands.length <= bandCount ? bands : null;
  };

  let low = Math.max(...extents);
  let high = Math.max(low, total);
  let best = fits(high) ?? ranks.map((r) => [r]);
  for (let i = 0; i < 40 && high - low > 1; i += 1) {
    const mid = (low + high) / 2;
    const packed = fits(mid);
    if (packed) {
      best = packed;
      high = mid;
    } else {
      low = mid;
    }
  }
  return best;
}

function bandExtent(band: Rank[], gapAfter: number[], indexOf: Map<Rank, number>): number {
  let extent = 0;
  for (let i = 0; i < band.length; i += 1) {
    extent += band[i].end - band[i].start;
    if (i > 0) extent += gapAfter[indexOf.get(band[i - 1])!] ?? 0;
  }
  return extent;
}

/**
 * Work out how far each box has to move so the drawing stops being a strip.
 *
 * Returns an empty map when the layout is already a reasonable shape, so the
 * common case costs one bounding-box measurement and changes nothing.
 */
export function planSerpentineWrap(
  boxes: WrapBox[],
  options: WrapOptions = {},
): Map<string, WrapOffset> {
  const none = new Map<string, WrapOffset>();
  if (boxes.length < 3) return none;

  const direction = options.direction ?? 'LR';
  const horizontal = direction === 'LR' || direction === 'RL';
  const bandGap = options.bandGap && options.bandGap > 0 ? options.bandGap : DEFAULT_BAND_GAP;

  const major = (b: WrapBox) => (horizontal ? b.x : b.y);
  const majorSize = (b: WrapBox) => (horizontal ? b.width : b.height);
  const minor = (b: WrapBox) => (horizontal ? b.y : b.x);
  const minorSize = (b: WrapBox) => (horizontal ? b.height : b.width);

  const majorMin = Math.min(...boxes.map(major));
  const majorMax = Math.max(...boxes.map((b) => major(b) + majorSize(b)));
  const minorMin = Math.min(...boxes.map(minor));
  const minorMax = Math.max(...boxes.map((b) => minor(b) + minorSize(b)));
  const majorSpan = majorMax - majorMin;
  const minorSpan = minorMax - minorMin;
  if (majorSpan <= 0 || minorSpan <= 0) return none;
  if (majorSpan / minorSpan <= WRAP_TRIGGER_RATIO) return none;

  const ranks = buildRanks(boxes, major, majorSize, minor, minorSize);
  if (ranks.length < 2) return none;

  const gapAfter = ranks.slice(0, -1).map((rank, i) => Math.max(0, ranks[i + 1].start - rank.end));
  const indexOf = new Map<Rank, number>(ranks.map((rank, i) => [rank, i]));

  // Try every band count and keep whichever lands closest to the target shape.
  // Solving for the count in closed form ignores the band gap, which for short
  // flows is most of the height.
  let chosen: Rank[][] | null = null;
  let chosenScore = Infinity;
  const maxBands = Math.min(MAX_WRAP_BANDS, ranks.length);
  for (let count = 1; count <= maxBands; count += 1) {
    const bands = packBands(ranks, gapAfter, count);
    const widest = Math.max(...bands.map((band) => bandExtent(band, gapAfter, indexOf)));
    const stacked = bands.reduce((sum, band) => {
      const top = Math.min(...band.map((r) => r.minorStart));
      const bottom = Math.max(...band.map((r) => r.minorEnd));
      return sum + (bottom - top);
    }, 0) + bandGap * (bands.length - 1);
    if (widest <= 0 || stacked <= 0) continue;
    const score = Math.abs(Math.log((widest / stacked) / LEARN_TARGET_ASPECT));
    if (score < chosenScore) {
      chosenScore = score;
      chosen = bands;
    }
  }
  if (!chosen || chosen.length < 2) return none;

  const widest = Math.max(...chosen.map((band) => bandExtent(band, gapAfter, indexOf)));
  const offsets = new Map<string, WrapOffset>();
  let minorCursor = minorMin;

  chosen.forEach((band, bandIndex) => {
    const bandTop = Math.min(...band.map((r) => r.minorStart));
    const bandBottom = Math.max(...band.map((r) => r.minorEnd));
    const dMinor = minorCursor - bandTop;

    // Reverse every other band so the seam between two bands is a short hop
    // rather than a line sweeping back across the entire drawing.
    const placement = bandIndex % 2 === 0 ? band : [...band].reverse();
    const extent = bandExtent(band, gapAfter, indexOf);
    let majorCursor = majorMin + (widest - extent) / 2;

    placement.forEach((rank, i) => {
      const dMajor = majorCursor - rank.start;
      for (const id of rank.ids) {
        offsets.set(id, horizontal ? { dx: dMajor, dy: dMinor } : { dx: dMinor, dy: dMajor });
      }
      majorCursor += rank.end - rank.start;
      const next = placement[i + 1];
      if (next) {
        // Bands are contiguous, so two ranks adjacent in the placement order
        // were adjacent in the source layout too — in either direction.
        const earlier = Math.min(indexOf.get(rank)!, indexOf.get(next)!);
        majorCursor += gapAfter[earlier] ?? 0;
      }
    });

    minorCursor += (bandBottom - bandTop) + bandGap;
  });

  return offsets;
}

interface PositionedLike {
  id: string;
  position: { x: number; y: number };
  width?: number;
  height?: number;
}

/**
 * Apply the wrap to a finished layout.
 *
 * Both layout engines hand back zone containers in absolute coordinates and
 * their members relative to the container, so moving a zone carries everything
 * inside it for free. Only the containers and the services that live outside
 * any container need an offset.
 */
export function wrapPositionedLayout<
  S extends PositionedLike & { groupId?: string },
  G extends PositionedLike & { width: number; height: number },
>(
  services: S[],
  groups: G[],
  options: WrapOptions & { nodeWidth: number; nodeHeight: number },
): { services: S[]; groups: G[]; bands: number } {
  const groupIds = new Set(groups.map((group) => group.id));
  const topLevel = services.filter((service) => !service.groupId || !groupIds.has(service.groupId));

  const boxes: WrapBox[] = [
    ...groups.map((group) => ({
      id: group.id,
      x: group.position.x,
      y: group.position.y,
      width: group.width,
      height: group.height,
    })),
    ...topLevel.map((service) => ({
      id: service.id,
      x: service.position.x,
      y: service.position.y,
      width: service.width ?? options.nodeWidth,
      height: service.height ?? options.nodeHeight,
    })),
  ];

  const offsets = planSerpentineWrap(boxes, options);
  if (offsets.size === 0) return { services, groups, bands: 1 };

  const shift = <T extends PositionedLike>(item: T): T => {
    const offset = offsets.get(item.id);
    if (!offset) return item;
    return {
      ...item,
      position: { x: item.position.x + offset.dx, y: item.position.y + offset.dy },
    };
  };

  const topLevelIds = new Set(topLevel.map((service) => service.id));
  const bands = new Set(
    [...offsets.values()].map((offset) => (
      options.direction === 'TB' || options.direction === 'BT' ? offset.dx : offset.dy
    )),
  ).size;

  return {
    services: services.map((service) => (topLevelIds.has(service.id) ? shift(service) : service)),
    groups: groups.map(shift),
    bands,
  };
}
