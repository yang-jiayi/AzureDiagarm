// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * The colours the canvas draws with, in one place.
 *
 * The exporters used to carry their own category map. It had drifted a long
 * way from the canvas: networking was cyan on screen and orange in the file,
 * identity was pink on screen and amber in the file, and -- worse than any
 * single wrong hue -- the export collapsed categories the canvas keeps apart,
 * so integration and security both came out the same red and the colour key a
 * reader relies on stopped meaning anything.
 *
 * Since the whole promise of the export is that the file matches the screen,
 * the screen has to be the source of truth. Both the canvas components and the
 * exporters now read from here, so a hue can only change in one place.
 */

/** Accent hue per service category, keyed by the normalised icon-folder name. */
export const CATEGORY_ACCENTS: Record<string, string> = {
  compute: '#0078d4',
  containers: '#0078d4',
  databases: '#10b981',
  storage: '#10b981',
  'data layer': '#10b981',
  'ai + machine learning': '#f59e0b',
  analytics: '#8b5cf6',
  networking: '#06b6d4',
  identity: '#ec4899',
  security: '#ef4444',
  monitor: '#6366f1',
  integration: '#14b8a6',
  iot: '#f97316',
  'app services': '#3b82f6',
  web: '#3b82f6',
  devops: '#8b5cf6',
  // The canvas map had no entry for this, so governance services fell back to
  // grey on screen while the export gave them their own olive. Giving both the
  // same distinct hue keeps the two in step and stops a real category reading
  // as "uncategorised".
  'management + governance': '#84cc16',
};

export const DEFAULT_ACCENT = '#6b7280';

/** The accent the canvas tints a service with. */
export function categoryAccent(category?: string): string {
  const key = typeof category === 'string' ? category.trim().toLowerCase() : '';
  return CATEGORY_ACCENTS[key] ?? DEFAULT_ACCENT;
}

/**
 * Zone accents, chosen from the label the way the canvas chooses them.
 *
 * The order is the canvas's order, verbatim. It is not the order the canvas
 * comment claims ("check before compute to prioritize AI keywords" sits above
 * a branch that is actually below compute), and it is not the order anyone
 * would design. It is reproduced exactly, because the job here is to match
 * what the user is looking at, not to improve on it. Changing this changes the
 * canvas too, which is the point: they cannot disagree any more.
 */
const ZONE_KEYWORDS: Array<{ accent: string; match: string[] }> = [
  { accent: '#6b7280', match: ['web', 'frontend', 'ingress', 'edge'] },
  { accent: '#0078d4', match: ['compute', 'processing', 'microservices', 'api'] },
  { accent: '#10b981', match: ['data', 'storage', 'database', 'persistence'] },
  { accent: '#f59e0b', match: ['ai', 'intelligence', 'analytics', 'ml', 'cognitive'] },
  { accent: '#f97316', match: ['iot', 'device', 'telemetry'] },
  { accent: '#ef4444', match: ['security', 'auth', 'identity', 'vault'] },
  { accent: '#8b5cf6', match: ['monitor', 'ops', 'observability', 'logging'] },
  { accent: '#06b6d4', match: ['network', 'integration', 'messaging', 'event', 'ingestion'] },
  { accent: '#0078d4', match: ['container', 'registry', 'runtime'] },
];

/** The zone accent for this label, or null when no keyword matches. */
export function matchZoneAccent(label?: string): string | null {
  const text = (label ?? '').toLowerCase();
  const hit = ZONE_KEYWORDS.find(({ match }) => match.some((word) => text.includes(word)));
  return hit?.accent ?? null;
}

/** The accent the canvas gives a zone with this label. */
export function zoneAccent(label?: string): string {
  return matchZoneAccent(label) ?? DEFAULT_ACCENT;
}

function channels(hex: string): [number, number, number] {
  const value = hex.replace('#', '');
  const full = value.length === 3 ? value.split('').map((c) => c + c).join('') : value;
  return [
    parseInt(full.slice(0, 2), 16),
    parseInt(full.slice(2, 4), 16),
    parseInt(full.slice(4, 6), 16),
  ];
}

function toHex([r, g, b]: [number, number, number]): string {
  const clamp = (n: number) => Math.max(0, Math.min(255, Math.round(n)));
  return `#${[r, g, b].map((n) => clamp(n).toString(16).padStart(2, '0')).join('')}`;
}

/** Mix a colour toward white. `amount` 0 keeps it, 1 makes it white. */
export function tint(hex: string, amount: number): string {
  const [r, g, b] = channels(hex);
  return toHex([r + (255 - r) * amount, g + (255 - g) * amount, b + (255 - b) * amount]);
}

/** Mix a colour toward black. `amount` 0 keeps it, 1 makes it black. */
export function shade(hex: string, amount: number): string {
  const [r, g, b] = channels(hex);
  return toHex([r * (1 - amount), g * (1 - amount), b * (1 - amount)]);
}

/** `rgba()` form of a hex colour, for the canvas styles that want one. */
export function rgba(hex: string, alpha: number): string {
  const [r, g, b] = channels(hex);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
