// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mapWorkflowStepsToEdges, normalizeWorkflowSteps } from '../src/utils/workflowStepMapping';

const edge = (id: string, source: string, target: string) => ({ id, source, target });

test('numbers each edge with the workflow step that describes it', () => {
  const edges = [
    edge('e1', 'fd', 'apim'),
    edge('e2', 'apim', 'func'),
    edge('e3', 'func', 'sql'),
  ];
  const workflow = [
    { step: 1, services: ['fd', 'apim'] },
    { step: 2, services: ['apim', 'func'] },
    { step: 3, services: ['func', 'sql'] },
  ];

  const mapped = mapWorkflowStepsToEdges(edges, workflow);

  assert.deepEqual([...mapped.entries()].sort(), [['e1', 1], ['e2', 2], ['e3', 3]]);
});

test('uses each step at most once even when several edges qualify', () => {
  const edges = [edge('a', 'x', 'y'), edge('b', 'y', 'z'), edge('c', 'x', 'z')];
  const workflow = [{ step: 1, services: ['x', 'y', 'z'] }];

  const mapped = mapWorkflowStepsToEdges(edges, workflow);

  assert.equal(mapped.size, 1);
});

test('prefers the edge whose target is named last, because services are in flow order', () => {
  const edges = [edge('a', 'x', 'y'), edge('b', 'y', 'z')];
  const workflow = [{ step: 4, services: ['x', 'y', 'z'] }];

  const mapped = mapWorkflowStepsToEdges(edges, workflow);

  assert.equal(mapped.get('b'), 4);
  assert.equal(mapped.has('a'), false);
});

test('keeps the author-declared step number rather than renumbering from one', () => {
  const edges = [edge('a', 'x', 'y')];
  const workflow = [{ step: 7, services: ['x', 'y'] }];

  assert.equal(mapWorkflowStepsToEdges(edges, workflow).get('a'), 7);
});

test('never assigns two steps to the same edge', () => {
  const edges = [edge('only', 'x', 'y')];
  const workflow = [
    { step: 1, services: ['x', 'y'] },
    { step: 2, services: ['x', 'y'] },
  ];

  const mapped = mapWorkflowStepsToEdges(edges, workflow);

  assert.deepEqual([...mapped.entries()], [['only', 1]]);
});

test('ignores steps whose services do not form an edge', () => {
  const edges = [edge('a', 'x', 'y')];
  const workflow = [{ step: 1, services: ['p', 'q'] }];

  assert.equal(mapWorkflowStepsToEdges(edges, workflow).size, 0);
});

test('is deterministic when candidates tie on flow order', () => {
  const edges = [edge('zz', 'x', 'z'), edge('aa', 'y', 'z')];
  const workflow = [{ step: 1, services: ['x', 'y', 'z'] }];

  const first = mapWorkflowStepsToEdges(edges, workflow);
  const reversed = mapWorkflowStepsToEdges([...edges].reverse(), workflow);

  assert.deepEqual([...first.entries()], [['aa', 1]]);
  assert.deepEqual([...reversed.entries()], [['aa', 1]]);
});

test('returns an empty map for a diagram with no workflow', () => {
  const edges = [edge('a', 'x', 'y')];

  assert.equal(mapWorkflowStepsToEdges(edges, undefined).size, 0);
  assert.equal(mapWorkflowStepsToEdges(edges, []).size, 0);
});

test('sorts out-of-order workflow arrays before assigning', () => {
  const edges = [edge('e1', 'a', 'b'), edge('e2', 'b', 'c')];
  const workflow = [
    { step: 2, services: ['a', 'b', 'c'] },
    { step: 1, services: ['a', 'b'] },
  ];

  const mapped = mapWorkflowStepsToEdges(edges, workflow);

  assert.equal(mapped.get('e1'), 1);
  assert.equal(mapped.get('e2'), 2);
});

// ── normalizeWorkflowSteps ───────────────────────────────────────────────────
//
// The AI names services inconsistently and numbers steps carelessly. Every one
// of these cases was observed in real model output; each would otherwise yield
// a diagram with no numbers at all, which looks identical to "this feature is
// not implemented".

/** Mirrors the production alias table: id, display name and type all resolve. */
const ALIASES: Record<string, string> = {
  'front-door': 'front-door',
  'azure-front-door': 'front-door',
  aks: 'aks',
  'azure-kubernetes-service': 'aks',
  sql: 'sql',
  'azure-sql-database': 'sql',
};
const resolveRef = (ref: unknown): string | null => {
  const key = String(ref ?? '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-');
  return ALIASES[key] ?? null;
};

test('workflow service references are repaired from display names to ids', () => {
  const result = normalizeWorkflowSteps(
    [
      { step: 1, description: 'Ingress', services: ['Azure Front Door', 'Azure Kubernetes Service'] },
      { step: 2, description: 'Query', services: ['aks', 'sql'] },
    ],
    resolveRef,
  );
  assert.deepEqual(result.steps.map((s) => s.services), [['front-door', 'aks'], ['aks', 'sql']]);
  assert.equal(result.repairedRefs, 2, 'only the two name references counted as repairs');
  assert.equal(result.droppedSteps, 0);
});

test('steps that cannot describe a hop are dropped and the rest renumbered', () => {
  const result = normalizeWorkflowSteps(
    [
      { step: 1, description: 'Unknown service', services: ['Contoso Widget Service', 'aks'] },
      { step: 2, description: 'Ingress', services: ['front-door', 'aks'] },
      { step: 3, description: '   ', services: ['aks', 'sql'] },
      { step: 4, description: 'Only one endpoint', services: ['sql'] },
      { step: 5, description: 'Query', services: ['aks', 'sql'] },
    ],
    resolveRef,
  );
  assert.deepEqual(result.steps.map((s) => s.description), ['Ingress', 'Query']);
  assert.deepEqual(result.steps.map((s) => s.step), [1, 2], 'numbering is contiguous from 1');
  assert.equal(result.droppedSteps, 3);
});

test('gaps, duplicates and repeated services are all repaired', () => {
  const result = normalizeWorkflowSteps(
    [
      { step: 7, description: 'First', services: ['front-door', 'aks'] },
      { step: 7, description: 'Second', services: ['aks', 'aks', 'sql'] },
    ],
    resolveRef,
  );
  assert.deepEqual(result.steps.map((s) => s.step), [1, 2], 'duplicate step numbers are resolved');
  assert.deepEqual(result.steps[1].services, ['aks', 'sql'], 'a repeat cannot stand in for a second service');
});

test('a malformed workflow degrades to an empty list instead of throwing', () => {
  for (const bad of [undefined, null, 'workflow', 42, {}]) {
    const result = normalizeWorkflowSteps(bad, resolveRef);
    assert.deepEqual(result.steps, [], `${JSON.stringify(bad)} yields no steps`);
    assert.equal(result.droppedSteps, 0);
  }
  const withJunk = normalizeWorkflowSteps([null, 'x', { description: 'no services' }], resolveRef);
  assert.deepEqual(withJunk.steps, []);
  assert.equal(withJunk.droppedSteps, 1, 'only the object entry counts as a dropped step');
});

test('normalized steps map straight onto edges', () => {
  const normalized = normalizeWorkflowSteps(
    [
      { step: 4, description: 'Ingress', services: ['Azure Front Door', 'Azure Kubernetes Service'] },
      { step: 9, description: 'Query', services: ['Azure Kubernetes Service', 'Azure SQL Database'] },
    ],
    resolveRef,
  );
  const mapped = mapWorkflowStepsToEdges(
    [
      { id: 'e1', source: 'front-door', target: 'aks' },
      { id: 'e2', source: 'aks', target: 'sql' },
    ],
    normalized.steps,
  );
  assert.deepEqual([...mapped.entries()].sort(), [['e1', 1], ['e2', 2]]);
});
