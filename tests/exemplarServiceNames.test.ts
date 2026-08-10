// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import test from 'node:test';
import { resolveServiceIconLoose } from '../src/utils/serviceIconFuzzy';

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '..');

interface ExemplarService {
  name: string;
  kind?: string;
}

// Extract the `{ "id": ..., "name": ..., "category": ... }` service/node
// objects embedded in the few-shot exemplar template literals. Connections
// start with "from", actors with "id"+"label" (no "category"), so keying on the
// id→name→category ordering isolates real service tiles.
function extractExemplarServices(relPath: string): ExemplarService[] {
  const source = readFileSync(path.join(repoRoot, relPath), 'utf8');
  const pattern = /\{\s*"id":\s*"[^"]+",\s*"name":\s*"([^"]+)",\s*"category":\s*"[^"]+"(?:\s*,\s*"kind":\s*"([^"]+)")?/g;
  const services: ExemplarService[] = [];
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(source)) !== null) {
    services.push({ name: match[1], kind: match[2] });
  }
  return services;
}

// Blueprint exemplars intentionally include generic, user-supplied nodes
// (a person, a caller's own app, an unspecified on-prem system). The blueprint
// canvas renders these via its initials `FallbackGlyph` by design, and the
// quality checker deliberately flags them as `unrecognized-service` so a user
// can rename them — so they are not expected to map to a specific Azure icon.
const NON_SERVICE_KINDS = new Set(['persona', 'cloud', 'device']);
const BLUEPRINT_GENERIC_PLACEHOLDERS = new Set([
  'Client Application',
  'On-Prem Service',
]);

function iconPathFor(name: string): string | null {
  const mapping = resolveServiceIconLoose(name);
  if (!mapping || !mapping.iconFile || !mapping.category) return null;
  return `/Azure_Public_Service_Icons/Icons/${mapping.category}/${mapping.iconFile}.svg`;
}

test('every reference exemplar service name resolves to a non-empty icon path', () => {
  const services = extractExemplarServices('src/services/referenceArchitectureAI.ts');
  // Guard against the extraction silently matching nothing.
  assert.ok(services.length >= 30, `expected many reference services, got ${services.length}`);

  const unresolved = services
    .map(service => service.name)
    .filter(name => !iconPathFor(name));

  assert.deepEqual(
    unresolved,
    [],
    `reference exemplar service names must resolve to icons: ${unresolved.join(', ')}`,
  );
});

test('every blueprint exemplar service name resolves to a non-empty icon path', () => {
  const nodes = extractExemplarServices('src/services/blueprintArchitectureAI.ts');
  assert.ok(nodes.length >= 8, `expected several blueprint nodes, got ${nodes.length}`);

  const unresolved = nodes
    .filter(node => !NON_SERVICE_KINDS.has(node.kind ?? ''))
    .filter(node => !BLUEPRINT_GENERIC_PLACEHOLDERS.has(node.name))
    .map(node => node.name)
    .filter(name => !iconPathFor(name));

  assert.deepEqual(
    unresolved,
    [],
    `blueprint exemplar service names must resolve to icons: ${unresolved.join(', ')}`,
  );
});

test('the Fabric shorthand names flagged in the audit all resolve', () => {
  // These are the exact names HIGH 2 called out as un-renderable in the
  // reference exemplar; assert each maps to a Fabric icon explicitly.
  const auditNames = [
    'OneLake Shortcuts',
    'Copy Jobs',
    'Dataflow Gen2',
    'Spark Notebook',
    'ML Experiment',
    'ML Model',
    'Mirroring',
    'Semantic Model (Direct Lake)',
    'Interactive Report',
    'Paginated Report',
    'Real-Time Dashboard',
    'Activator',
  ];
  for (const name of auditNames) {
    const mapping = resolveServiceIconLoose(name);
    assert.ok(mapping, `expected "${name}" to resolve`);
    assert.ok(mapping!.iconFile.length > 0, `expected "${name}" to have an icon file`);
    assert.equal(mapping!.category, 'fabric', `expected "${name}" to route to a Fabric icon`);
  }
});
