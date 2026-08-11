// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import test from 'node:test';
import { mapWorkflowStepsToEdges } from '../src/utils/workflowStepMapping';

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
