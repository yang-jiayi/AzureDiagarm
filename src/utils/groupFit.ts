/**
 * Shrink-wraps a zone container around the services inside it.
 *
 * Both layout engines size a group from the layout library's own compound-node
 * box plus `groupPadding`, which the presets set to 60-80px. Around a 150px
 * tile that is a border wider than the tile itself: a four-zone diagram of ten
 * services came out on a 30.7in x 20.2in page that was 98.9% empty. The
 * Architecture Center draws boundaries that hug their contents, so this refits
 * every container to the services it actually holds.
 *
 * The library still decides where the services go — only the container and the
 * offset of its contents change.
 */

// The editor already measured these: the visible "Fit to content" button
// reserves GROUP_PADDING + GROUP_HEADER_HEIGHT above the first tile, and the
// zone header wraps to two lines for a long name, so anything less overlaps
// the first row. Reusing them keeps an automatic layout and a manual "fit to
// content" agreeing on what a zone should look like.
import { GROUP_PADDING, GROUP_HEADER_HEIGHT } from './groupUtils';

/** Breathing room on the left, right and bottom of a zone. */
export const GROUP_INNER_PAD_PX = GROUP_PADDING;
/** Top inset: the same breathing room plus the zone's title bar. */
export const GROUP_HEADER_PAD_PX = GROUP_PADDING + GROUP_HEADER_HEIGHT;

const MIN_GROUP_W = 220;
const MIN_GROUP_H = 150;

interface FitGroup {
  id: string;
  position: { x: number; y: number };
  width: number;
  height: number;
}

interface FitService {
  id: string;
  groupId?: string | null;
  position: { x: number; y: number };
  width?: number;
  height?: number;
}

export interface FitOptions {
  nodeWidth: number;
  nodeHeight: number;
}

/**
 * `services` must carry positions relative to their parent group, which is the
 * form both engines produce just before overlap resolution. Ungrouped services
 * are returned untouched.
 */
export function fitGroupsToMembers<G extends FitGroup, S extends FitService>(
  groups: G[],
  services: S[],
  opts: FitOptions,
): { groups: G[]; services: S[] } {
  if (groups.length === 0) return { groups, services };

  const moved = new Map<string, { x: number; y: number }>();
  const resized = new Map<string, { w: number; h: number }>();

  for (const group of groups) {
    const members = services.filter((s) => s.groupId === group.id);
    if (members.length === 0) continue;

    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;
    for (const member of members) {
      const w = member.width ?? opts.nodeWidth;
      const h = member.height ?? opts.nodeHeight;
      minX = Math.min(minX, member.position.x);
      minY = Math.min(minY, member.position.y);
      maxX = Math.max(maxX, member.position.x + w);
      maxY = Math.max(maxY, member.position.y + h);
    }
    if (!Number.isFinite(minX) || !Number.isFinite(minY)) continue;

    const contentW = maxX - minX;
    const contentH = maxY - minY;
    const width = Math.max(MIN_GROUP_W, contentW + GROUP_INNER_PAD_PX * 2);
    const height = Math.max(MIN_GROUP_H, contentH + GROUP_HEADER_PAD_PX + GROUP_INNER_PAD_PX);

    // Re-seat the contents against the new inset. Centre them in whichever
    // axis the minimum size made larger than the contents need, so a one-tile
    // zone does not sit in a corner of its own box.
    const slackX = (width - GROUP_INNER_PAD_PX * 2 - contentW) / 2;
    const slackY = (height - GROUP_HEADER_PAD_PX - GROUP_INNER_PAD_PX - contentH) / 2;
    const shiftX = GROUP_INNER_PAD_PX + slackX - minX;
    const shiftY = GROUP_HEADER_PAD_PX + slackY - minY;

    resized.set(group.id, { w: width, h: height });
    if (shiftX !== 0 || shiftY !== 0) moved.set(group.id, { x: shiftX, y: shiftY });
  }

  if (resized.size === 0) return { groups, services };

  return {
    groups: groups.map((group) => {
      const size = resized.get(group.id);
      return size ? { ...group, width: size.w, height: size.h } : group;
    }),
    services: services.map((service) => {
      if (!service.groupId) return service;
      const shift = moved.get(service.groupId);
      if (!shift) return service;
      return {
        ...service,
        position: { x: service.position.x + shift.x, y: service.position.y + shift.y },
      };
    }),
  };
}
