// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export const BREAKPOINTS = {
  micro: 480,
  compact: 640,
  narrow: 900,
  workspace: 1180,
  wide: 1440,
  lowHeight: 480,
  shortHeight: 600,
} as const;

// CSS custom properties cannot be used in media-query conditions. Stylesheets
// use these same literal boundaries, enforced by tests/breakpoints.test.ts.
const maxWidth = (value: number) => `(max-width: ${value}px)`;
const maxHeight = (value: number) => `(max-height: ${value}px)`;

export const MEDIA_QUERIES = {
  micro: maxWidth(BREAKPOINTS.micro),
  compact: maxWidth(BREAKPOINTS.compact),
  narrow: maxWidth(BREAKPOINTS.narrow),
  workspace: maxWidth(BREAKPOINTS.workspace),
  wide: maxWidth(BREAKPOINTS.wide),
  lowHeight: maxHeight(BREAKPOINTS.lowHeight),
  shortHeight: maxHeight(BREAKPOINTS.shortHeight),
  compactOrLowHeight: `${maxWidth(BREAKPOINTS.compact)}, ${maxHeight(BREAKPOINTS.lowHeight)}`,
  compactOrShortWorkspace:
    `${maxWidth(BREAKPOINTS.compact)}, `
    + `${maxWidth(BREAKPOINTS.workspace)} and ${maxHeight(BREAKPOINTS.shortHeight)}`,
} as const;
