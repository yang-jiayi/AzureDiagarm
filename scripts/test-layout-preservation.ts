import assert from 'node:assert/strict';
import {
  buildAbsolutePositionMap,
  preserveManualLayout,
  selectHorizontalConnectionHandles,
  type LayoutNode,
} from '../src/utils/preserveManualLayout.ts';

const previous: LayoutNode[] = [
  {
    id: 'old-group',
    type: 'groupNode',
    position: { x: 100, y: 50 },
    data: { label: 'Application Tier' },
    style: { width: 500, height: 300, backgroundColor: '#eef6ff' },
    width: 500,
    height: 300,
  },
  {
    id: 'old-app',
    type: 'azureNode',
    position: { x: 40, y: 60 },
    parentNode: 'old-group',
    data: { label: 'App Service', pricing: { estimatedCost: 125 }, stylePreset: 'presentation' },
    style: { opacity: 0.9 },
  },
  {
    id: 'removed-cache',
    type: 'azureNode',
    position: { x: 720, y: 100 },
    data: { label: 'Azure Cache for Redis' },
  },
];

const generated: LayoutNode[] = [
  {
    id: 'new-group-id',
    type: 'groupNode',
    position: { x: 600, y: 400 },
    data: { label: 'Application Tier' },
    style: { width: 700, height: 500 },
    width: 700,
    height: 500,
  },
  {
    id: 'new-monitor',
    type: 'azureNode',
    position: { x: 40, y: 60 },
    parentNode: 'new-group-id',
    data: { label: 'Application Insights' },
  },
  {
    id: 'regenerated-app-id',
    type: 'azureNode',
    position: { x: 80, y: 90 },
    parentNode: 'new-group-id',
    data: { label: 'app-service', iconPath: '/new-app-service.svg' },
  },
];

const result = preserveManualLayout(previous, generated);
const group = result.find((node) => node.id === 'new-group-id');
const app = result.find((node) => node.id === 'regenerated-app-id');
const monitor = result.find((node) => node.id === 'new-monitor');

assert.deepEqual(group?.position, { x: 100, y: 50 }, 'group position should survive regenerated IDs');
assert.equal((group?.style as Record<string, unknown>).backgroundColor, '#eef6ff', 'manual group style should be retained');

assert.deepEqual(app?.position, { x: 40, y: 60 }, 'service absolute position should survive parent ID changes');
assert.equal(app?.parentNode, 'new-group-id', 'generated topology should own the current parent ID');
assert.equal((app?.data as Record<string, unknown>).iconPath, '/new-app-service.svg', 'generated icon should be accepted');
assert.equal((app?.data as Record<string, unknown>).stylePreset, 'presentation', 'editor display data should survive');

assert.deepEqual(monitor?.position, { x: 300, y: 60 }, 'new services should move only when their generated position overlaps');
assert.equal((group?.style as Record<string, unknown>).width, 528, 'parent groups should expand only enough to contain additions');
assert.equal(result.some((node) => node.id === 'removed-cache'), false, 'removed services should stay removed');

const stableIdResult = preserveManualLayout(
  [{ id: 'sql', type: 'azureNode', position: { x: 20, y: 40 }, data: { label: 'SQL Database' } }],
  [{ id: 'sql', type: 'azureNode', position: { x: 900, y: 800 }, data: { label: 'Renamed SQL Database' } }],
);
assert.deepEqual(stableIdResult[0].position, { x: 20, y: 40 }, 'stable IDs should take precedence over labels');

const unchangedGroupResult = preserveManualLayout(
  previous.slice(0, 2),
  generated.slice(0, 2),
);
assert.equal(
  (unchangedGroupResult[0].style as Record<string, unknown>).width,
  500,
  'manual group dimensions should remain unchanged when no addition needs more room',
);

const localizedLabelResult = preserveManualLayout(
  [{ id: 'old-ja', type: 'azureNode', position: { x: 25, y: 45 }, data: { label: 'データベース' } }],
  [{ id: 'new-ja', type: 'azureNode', position: { x: 700, y: 500 }, data: { label: 'データベース' } }],
);
assert.deepEqual(
  localizedLabelResult[0].position,
  { x: 25, y: 45 },
  'localized labels should preserve layout when generated IDs change',
);

const regroupedResult = preserveManualLayout(
  [
    { id: 'old-group-a', type: 'groupNode', position: { x: 100, y: 100 }, data: { label: 'Group A' } },
    {
      id: 'regrouped-service',
      type: 'azureNode',
      position: { x: 40, y: 60 },
      parentNode: 'old-group-a',
      data: { label: 'App Service' },
    },
  ],
  [
    { id: 'new-group-b', type: 'groupNode', position: { x: 900, y: 400 }, data: { label: 'Group B' } },
    {
      id: 'regrouped-service',
      type: 'azureNode',
      position: { x: 70, y: 80 },
      parentNode: 'new-group-b',
      data: { label: 'App Service' },
    },
  ],
);
const regroupedService = regroupedResult.find((node) => node.id === 'regrouped-service');
assert.deepEqual(
  regroupedService?.position,
  { x: 70, y: 80 },
  'services moved to a different group should use the generated in-parent position',
);

const duplicateLabelResult = preserveManualLayout(
  [
    { id: 'old-app-1', type: 'azureNode', position: { x: 10, y: 20 }, data: { label: 'App Service' } },
    { id: 'old-app-2', type: 'azureNode', position: { x: 300, y: 20 }, data: { label: 'App Service' } },
  ],
  [
    { id: 'new-app-1', type: 'azureNode', position: { x: 50, y: 100 }, data: { label: 'App Service' } },
    { id: 'new-app-2', type: 'azureNode', position: { x: 400, y: 100 }, data: { label: 'App Service' } },
  ],
);
assert.deepEqual(
  duplicateLabelResult.map((node) => node.position),
  [{ x: 50, y: 100 }, { x: 400, y: 100 }],
  'ambiguous duplicate labels should keep generated positions instead of swapping manual positions',
);

const generatedDuplicateResult = preserveManualLayout(
  [
    { id: 'old-group-a', type: 'groupNode', position: { x: 100, y: 100 }, data: { label: 'Group A' } },
    {
      id: 'old-context-app',
      type: 'azureNode',
      position: { x: 30, y: 40 },
      parentNode: 'old-group-a',
      data: { label: 'App Service' },
    },
  ],
  [
    { id: 'new-group-b', type: 'groupNode', position: { x: 800, y: 100 }, data: { label: 'Group B' } },
    { id: 'new-group-a', type: 'groupNode', position: { x: 600, y: 100 }, data: { label: 'Group A' } },
    {
      id: 'new-context-app-b',
      type: 'azureNode',
      position: { x: 40, y: 50 },
      parentNode: 'new-group-b',
      data: { label: 'App Service' },
    },
    {
      id: 'new-context-app-a',
      type: 'azureNode',
      position: { x: 70, y: 80 },
      parentNode: 'new-group-a',
      data: { label: 'App Service' },
    },
  ],
);
assert.deepEqual(
  generatedDuplicateResult.find((node) => node.id === 'new-context-app-b')?.position,
  { x: 40, y: 50 },
  'a duplicate in a different group must not consume the previous node',
);
assert.deepEqual(
  generatedDuplicateResult.find((node) => node.id === 'new-context-app-a')?.position,
  { x: 30, y: 40 },
  'parent context should disambiguate a generated duplicate label',
);

const groupCollisionResult = preserveManualLayout(
  [
    {
      id: 'old-manual-group',
      type: 'groupNode',
      position: { x: 500, y: 100 },
      style: { width: 400, height: 300 },
      data: { label: 'Manual Group' },
    },
  ],
  [
    {
      id: 'new-manual-group-id',
      type: 'groupNode',
      position: { x: 0, y: 100 },
      style: { width: 400, height: 300 },
      data: { label: 'Manual Group' },
    },
    {
      id: 'new-sibling-group',
      type: 'groupNode',
      position: { x: 450, y: 100 },
      style: { width: 300, height: 300 },
      data: { label: 'New Group' },
    },
  ],
);
assert.deepEqual(
  groupCollisionResult.find((node) => node.id === 'new-sibling-group')?.position,
  { x: 980, y: 100 },
  'new top-level groups should move clear of manually positioned groups',
);

const routedNodes = preserveManualLayout(
  [
    { id: 'source', type: 'azureNode', position: { x: 500, y: 40 }, data: { label: 'Source' } },
    { id: 'target', type: 'azureNode', position: { x: 100, y: 40 }, data: { label: 'Target' } },
  ],
  [
    { id: 'source', type: 'azureNode', position: { x: 0, y: 40 }, data: { label: 'Source' } },
    { id: 'target', type: 'azureNode', position: { x: 400, y: 40 }, data: { label: 'Target' } },
  ],
);
assert.deepEqual(
  selectHorizontalConnectionHandles(buildAbsolutePositionMap(routedNodes), 'source', 'target'),
  { sourceHandle: 'left-source', targetHandle: 'right-target' },
  'edge handles should follow preserved positions rather than generated positions',
);

console.log('Layout preservation tests passed.');