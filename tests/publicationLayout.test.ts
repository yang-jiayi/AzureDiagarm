import assert from 'node:assert/strict';
import test from 'node:test';

import type { ReferenceArchitecture } from '../src/services/referenceArchitectureAI';
import {
  calculateBlueprintContentFrame,
  calculateBlueprintHostWidth,
  calculateReferenceCanvasWidth,
} from '../src/utils/publicationLayout';
import type { BlueprintArchitecture } from '../src/services/blueprintArchitectureAI';

function reference(overrides: Partial<ReferenceArchitecture> = {}): ReferenceArchitecture {
  return {
    title: 'Reference',
    stages: [
      { id: 'ingest', label: 'Ingest', services: [{ id: 'a', name: 'A', category: 'integration' }] },
      { id: 'serve', label: 'Serve', services: [{ id: 'b', name: 'B', category: 'web' }] },
    ],
    connections: [],
    ...overrides,
  };
}

test('reference exports shrink below the legacy fixed width for compact diagrams', () => {
  assert.equal(calculateReferenceCanvasWidth(reference()), 960);
});

test('reference exports expand for dense stages and optional side columns', () => {
  const services = Array.from({ length: 5 }, (_, index) => ({
    id: `service-${index}`,
    name: `Service ${index}`,
    category: 'analytics',
  }));
  const width = calculateReferenceCanvasWidth(reference({
    stages: Array.from({ length: 4 }, (_, index) => ({
      id: `stage-${index}`,
      label: `Stage ${index}`,
      services,
    })),
    dataSources: [{ category: 'Sources', items: ['Source'] }],
    actors: [{ id: 'users', label: 'Users' }],
  }));

  assert.equal(width, 1640);
});

test('explicit reference width remains supported without allowing unreadably narrow output', () => {
  assert.equal(calculateReferenceCanvasWidth(reference(), 1440), 1440);
  assert.equal(calculateReferenceCanvasWidth(reference(), 600), 960);
});

test('blueprint right-side support panel is included inside the host width', () => {
  assert.equal(calculateBlueprintHostWidth(1600, 'bottom'), 1696);
  assert.equal(calculateBlueprintHostWidth(1600, 'right'), 2116);
});

test('blueprint framing removes unused authored margins without clipping content', () => {
  const blueprint: BlueprintArchitecture = {
    title: 'Blueprint',
    canvas: { width: 1600, height: 1000 },
    zones: [{
      id: 'zone',
      label: 'Zone',
      x: 100,
      y: 360,
      width: 1360,
      height: 520,
    }],
    nodes: [{
      id: 'node',
      name: 'App Service',
      category: 'app services',
      x: 180,
      y: 420,
    }],
    edges: [],
  };

  assert.deepEqual(calculateBlueprintContentFrame(blueprint), {
    width: 1488,
    height: 648,
    offsetX: -36,
    offsetY: -296,
  });
});
