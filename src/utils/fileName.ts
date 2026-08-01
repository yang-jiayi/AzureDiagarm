// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const MAX_SEGMENT_LENGTH = 64;

function stableHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

export function toFileNameSegment(value: string): string {
  const slug = value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'unknown-region';
  if (slug.length <= MAX_SEGMENT_LENGTH) return slug;

  const hash = stableHash(value);
  const prefix = /^mixed\s*\(/i.test(value)
    ? 'mixed-regions'
    : slug.slice(0, MAX_SEGMENT_LENGTH - hash.length - 1).replace(/-+$/g, '');
  return `${prefix || 'value'}-${hash}`;
}
