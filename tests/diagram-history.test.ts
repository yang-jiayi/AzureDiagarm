import assert from 'node:assert/strict';
import test from 'node:test';
import { DiagramHistory } from '../src/services/diagramHistoryService';

test('diagram history supports bounded undo and redo', () => {
  const history = new DiagramHistory({ value: 0 }, 3);
  history.record({ value: 1 });
  history.record({ value: 2 });
  history.record({ value: 3 });

  assert.deepEqual(history.undo(), { value: 2 });
  assert.deepEqual(history.undo(), { value: 1 });
  assert.equal(history.canUndo(), false);
  assert.deepEqual(history.redo(), { value: 2 });
});

test('recording after undo clears redo entries', () => {
  const history = new DiagramHistory({ value: 0 });
  history.record({ value: 1 });
  history.record({ value: 2 });
  assert.deepEqual(history.undo(), { value: 1 });

  history.record({ value: 9 });
  assert.equal(history.canRedo(), false);
  assert.deepEqual(history.undo(), { value: 1 });
});

test('duplicate snapshots do not create history entries', () => {
  const history = new DiagramHistory({ value: 0 });
  assert.equal(history.record({ value: 0 }), false);
  assert.equal(history.canUndo(), false);
});
