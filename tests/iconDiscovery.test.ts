// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  deduplicatePaletteIcons,
  splitIconSearchHighlight,
} from '../src/utils/iconDiscovery';

test('deduplicatePaletteIcons keeps the semantically canonical service icon', () => {
  const shared = {
    name: 'App Services',
    paletteCategory: 'app-web',
    source: 'official-azure' as const,
  };
  const result = deduplicatePaletteIcons([
    {
      ...shared,
      id: 'compute/app-services',
      category: 'compute',
      path: '/compute/app-services.svg',
    },
    {
      ...shared,
      id: 'app-services/app-services',
      category: 'app services',
      path: '/app services/app-services.svg',
    },
    {
      ...shared,
      id: 'web/app-services',
      category: 'web',
      path: '/web/app-services.svg',
    },
  ]);

  assert.deepEqual(result.icons.map(icon => icon.id), ['app-services/app-services']);
  assert.equal(
    result.canonicalIdById.get('compute/app-services'),
    'app-services/app-services',
  );
});

test('deduplicatePaletteIcons prefers the category matching the curated palette', () => {
  const result = deduplicatePaletteIcons([
    {
      id: 'devops/application-insights',
      name: 'Application Insights',
      category: 'devops',
      paletteCategory: 'monitoring',
      path: '/devops/application-insights.svg',
      source: 'official-azure',
    },
    {
      id: 'monitor/application-insights',
      name: 'Application Insights',
      category: 'monitor',
      paletteCategory: 'monitoring',
      path: '/monitor/application-insights.svg',
      source: 'official-azure',
    },
  ]);

  assert.equal(result.icons[0]?.id, 'monitor/application-insights');
});

test('a generic concept shape never replaces the official icon of the same name', () => {
  // The Microsoft 365 package ships reusable Fluent shapes under names such as
  // "Search" and "Apps". Those names belong to real Azure and Fabric assets.
  const result = deduplicatePaletteIcons([
    {
      id: 'microsoft 365/m365-search',
      name: 'Search',
      category: 'microsoft 365',
      paletteCategory: 'microsoft-365',
      path: '/microsoft 365/m365-search.svg',
      source: 'microsoft-365',
      generic: true,
    },
    {
      id: 'general/10834-icon-service-Search',
      name: 'Search',
      category: 'general',
      paletteCategory: 'azure-general',
      path: '/general/10834-icon-service-Search.svg',
      source: 'official-azure',
    },
  ]);

  assert.deepEqual(result.icons.map(icon => icon.id), ['general/10834-icon-service-Search']);
  assert.equal(
    result.canonicalIdById.get('microsoft 365/m365-search'),
    'general/10834-icon-service-Search',
  );
});

test('a named Microsoft 365 workload still wins against a generic shape', () => {
  const result = deduplicatePaletteIcons([
    {
      id: 'general/10800-icon-service-Teams',
      name: 'Microsoft Teams',
      category: 'general',
      paletteCategory: 'azure-general',
      path: '/general/10800-icon-service-Teams.svg',
      source: 'official-azure',
      generic: true,
    },
    {
      id: 'microsoft 365/m365-app-teams',
      name: 'Microsoft Teams',
      category: 'microsoft 365',
      paletteCategory: 'microsoft-365',
      path: '/microsoft 365/m365-app-teams.svg',
      source: 'microsoft-365',
    },
  ]);

  assert.deepEqual(result.icons.map(icon => icon.id), ['microsoft 365/m365-app-teams']);
});

test('splitIconSearchHighlight marks every matching search token without changing text', () => {
  const segments = splitIconSearchHighlight('Front Door and CDN Profiles', 'front CDN');
  assert.equal(segments.map(segment => segment.text).join(''), 'Front Door and CDN Profiles');
  assert.deepEqual(
    segments.filter(segment => segment.matched).map(segment => segment.text),
    ['Front', 'CDN'],
  );
});
