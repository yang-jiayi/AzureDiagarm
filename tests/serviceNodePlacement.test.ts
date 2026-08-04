import assert from 'node:assert/strict';
import test from 'node:test';
import type { Node } from 'reactflow';
import { findAvailableServicePosition } from '../src/utils/serviceNodePlacement';

function node(id: string, x: number, y: number, width = 180, height = 132): Node {
  return {
    id,
    type: 'azureNode',
    position: { x, y },
    width,
    height,
    data: { label: id },
  };
}

test('service placement keeps the requested position when it is free', () => {
  assert.deepEqual(findAvailableServicePosition({ x: 100, y: 100 }, []), {
    x: 100,
    y: 100,
  });
});

test('service placement finds a nearby grid slot without overlapping siblings', () => {
  const existing = [
    node('center', 100, 100),
    node('left', -120, 100),
    node('right', 320, 100),
  ];
  const position = findAvailableServicePosition({ x: 100, y: 100 }, existing);

  assert.deepEqual(position, { x: -100, y: -60 });
  assert.ok(Math.abs(position.x % 20) === 0);
  assert.ok(Math.abs(position.y % 20) === 0);
});

test('service placement uses measured node dimensions when avoiding collisions', () => {
  const existing = [node('wide', 0, 0, 420, 160)];
  const position = findAvailableServicePosition({ x: 200, y: 20 }, existing);

  assert.notDeepEqual(position, { x: 200, y: 20 });
  assert.ok(position.x >= 448 || position.x + 180 <= -28 || Math.abs(position.y) >= 188);
});
