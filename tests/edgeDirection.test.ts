import test from 'node:test';
import assert from 'node:assert/strict';
import { classifyEdgeDirection } from '../src/utils/edgeDirection.ts';

/**
 * The direction is not cosmetic: EditableEdge draws the arrowhead from it,
 * layoutPresets orients the primary chain from it, and the IaC exporter
 * derives deployment ordering from it. A misread label points the arrow the
 * wrong way on the canvas AND inverts the generated dependsOn.
 */
test('direction keywords match whole words, not substrings', () => {
  // Every one of these was classified 'reverse' because it contains "ack".
  for (const label of [
    'Backup',
    'Daily backup',
    'Feedback loop',
    'Track events',
    'Package upload',
    'Rollback',
    'Backup to vault',
  ]) {
    assert.equal(classifyEdgeDirection(label), 'forward', `"${label}" must not read as reverse`);
  }

  // "Async" contains "sync" and was classified bidirectional.
  for (const label of ['Async replication', 'Asynchronous write', 'Async queue']) {
    assert.equal(classifyEdgeDirection(label), 'forward', `"${label}" must not read as bidirectional`);
  }
});

test('direction keywords still match when genuinely present', () => {
  for (const label of ['Response', 'HTTP response', 'Callback', 'Returns data', 'ACK', 'Reply message']) {
    assert.equal(classifyEdgeDirection(label), 'reverse', `"${label}" should be reverse`);
  }
  for (const label of ['Sync', 'Data sync', 'Bidirectional link', 'Two-way exchange', 'Communicate']) {
    assert.equal(classifyEdgeDirection(label), 'bidirectional', `"${label}" should be bidirectional`);
  }
});

test('bidirectional wins over reverse, and unknown labels stay forward', () => {
  assert.equal(classifyEdgeDirection('Sync response'), 'bidirectional');
  assert.equal(classifyEdgeDirection('Sends telemetry'), 'forward');
  assert.equal(classifyEdgeDirection(''), 'forward');
  assert.equal(classifyEdgeDirection(undefined as unknown as string), 'forward');
});

test('hyphenated and multi-word keywords are matched on the raw label', () => {
  assert.equal(classifyEdgeDirection('Two-way'), 'bidirectional');
  assert.equal(classifyEdgeDirection('two-way sync channel'), 'bidirectional');
  // A hyphen is a word separator, so the single-word keywords also survive it.
  assert.equal(classifyEdgeDirection('request-response'), 'reverse');
});
