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

test('splitIconSearchHighlight marks every matching search token without changing text', () => {
  const segments = splitIconSearchHighlight('Front Door and CDN Profiles', 'front CDN');
  assert.equal(segments.map(segment => segment.text).join(''), 'Front Door and CDN Profiles');
  assert.deepEqual(
    segments.filter(segment => segment.matched).map(segment => segment.text),
    ['Front', 'CDN'],
  );
});
