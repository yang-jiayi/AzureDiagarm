import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { SERVICE_ICON_MAP } from '../src/data/serviceIconMapping.ts';

const ICON_ROOT = path.join(process.cwd(), 'Azure_Public_Service_Icons', 'Icons');

/**
 * Every mapped service must point at a file that really ships. A stale
 * `iconFile` silently produces an icon-less tile on the canvas *and* an
 * icon-less shape in the PowerPoint / Visio exports, because `loadIcon` does an
 * exact key lookup and returns '' for a miss.
 */
test('every mapped service icon file exists on disk', () => {
  const missing: string[] = [];
  for (const [serviceName, mapping] of Object.entries(SERVICE_ICON_MAP)) {
    const file = path.join(ICON_ROOT, mapping.category, `${mapping.iconFile}.svg`);
    if (!existsSync(file)) {
      missing.push(`${serviceName} -> ${mapping.category}/${mapping.iconFile}.svg`);
    }
  }
  assert.deepEqual(
    missing,
    [],
    `${missing.length} service mappings reference a missing icon file:\n${missing.join('\n')}`,
  );
});

test('mapped categories are real icon folders', () => {
  const folders = new Set(
    readdirSync(ICON_ROOT, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name),
  );
  const unknown = [
    ...new Set(
      Object.values(SERVICE_ICON_MAP)
        .map((mapping) => mapping.category)
        .filter((category) => !folders.has(category)),
    ),
  ];
  assert.deepEqual(unknown, [], `unknown icon folders: ${unknown.join(', ')}`);
});
