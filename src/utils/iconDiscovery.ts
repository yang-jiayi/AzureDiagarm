// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type PaletteIconSource =
  | 'official-azure'
  | 'fabric'
  | 'power-platform'
  | 'dynamics-365'
  | 'microsoft-365'
  | 'supplemental';

interface PaletteIconCandidate {
  id: string;
  name: string;
  category: string;
  paletteCategory: string;
  path: string;
  source: PaletteIconSource;
  /**
   * Reusable concept shape rather than a named product. The Microsoft 365
   * package ships these under deliberately generic names such as "Search",
   * "Code" and "Apps", which collide with real Azure and Fabric assets.
   */
  generic?: boolean;
}

export interface HighlightSegment {
  text: string;
  matched: boolean;
}

const SOURCE_HINTS: Record<string, string[]> = {
  ai: ['ai machine learning'],
  'app-web': ['app services', 'web'],
  compute: ['compute'],
  containers: ['containers'],
  'data-analytics': ['analytics'],
  databases: ['databases'],
  'developer-devops': ['devops'],
  'hybrid-edge': ['hybrid multicloud'],
  identity: ['identity'],
  integration: ['integration'],
  iot: ['iot'],
  devices: ['intune'],
  management: ['management governance'],
  monitoring: ['monitor'],
  networking: ['networking'],
  security: ['security'],
  storage: ['storage'],
  migration: ['migration'],
  fabric: ['fabric'],
  'microsoft-copilot': ['power platform', 'microsoft 365', 'fabric', 'ai machine learning'],
  'power-platform': ['power platform'],
  'dynamics-365': ['dynamics 365'],
  'microsoft-365': ['microsoft 365'],
};

export function normalizeIconDiscoveryText(value: string): string {
  return value
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/[+&/_.(),:;()[\]{}-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function candidateScore(icon: PaletteIconCandidate): number {
  const category = normalizeIconDiscoveryText(icon.category);
  const name = normalizeIconDiscoveryText(icon.name);
  const hints = SOURCE_HINTS[icon.paletteCategory] || [];
  let score = icon.source === 'official-azure'
    ? 30
    : icon.source === 'power-platform' || icon.source === 'dynamics-365'
      ? 25
      : icon.source === 'fabric'
        ? 20
        : icon.source === 'microsoft-365'
          // Concept symbols are reusable shapes, so several of them match any
          // given search. They must never outrank a product logo.
          ? 15
          : 10;

  if (category === name) score += 120;
  if (hints.includes(category)) score += 80;
  if (name.includes(category) || category.includes(name)) score += 30;
  if (['general', 'new icons', 'other'].includes(category)) score -= 40;
  return score;
}

export function deduplicatePaletteIcons<T extends PaletteIconCandidate>(
  icons: T[],
): { icons: T[]; canonicalIdById: Map<string, string> } {
  const groups = new Map<string, T[]>();
  for (const icon of icons) {
    const key = normalizeIconDiscoveryText(icon.name) || icon.id;
    const group = groups.get(key) || [];
    group.push(icon);
    groups.set(key, group);
  }

  const canonicalIdById = new Map<string, string>();
  const uniqueIcons = [...groups.values()].map((group) => {
    const canonical = group.reduce((best, candidate) => {
      // A generic concept shape never represents a name that a named product or
      // an official asset already owns, however well it scores on category hints.
      const genericDifference = Number(best.generic ?? false) - Number(candidate.generic ?? false);
      if (genericDifference !== 0) return genericDifference > 0 ? candidate : best;
      const scoreDifference = candidateScore(candidate) - candidateScore(best);
      if (scoreDifference > 0) return candidate;
      if (scoreDifference < 0) return best;
      return candidate.path.localeCompare(best.path) < 0 ? candidate : best;
    });
    group.forEach(icon => canonicalIdById.set(icon.id, canonical.id));
    return canonical;
  });

  uniqueIcons.sort((left, right) => left.name.localeCompare(right.name));
  return { icons: uniqueIcons, canonicalIdById };
}

export function splitIconSearchHighlight(text: string, query: string): HighlightSegment[] {
  const terms = [...new Set(
    query
      .trim()
      .split(/\s+/)
      .map(term => term.trim())
      .filter(Boolean),
  )].sort((left, right) => right.length - left.length);
  if (terms.length === 0) return [{ text, matched: false }];

  const escapedTerms = terms.map(term => term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  const matcher = new RegExp(`(${escapedTerms.join('|')})`, 'giu');
  return text
    .split(matcher)
    .filter(Boolean)
    .map(segment => ({
      text: segment,
      matched: terms.some(term => segment.localeCompare(term, undefined, {
        sensitivity: 'accent',
      }) === 0),
    }));
}
