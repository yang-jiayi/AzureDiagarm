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

/**
 * The label is free text from a model, so the inflected form is at least as
 * likely as the bare verb. Matching exact words only sent all of these back
 * to 'forward', which is the same wrong-way arrow reached from the other side.
 */
test('inflected forms of the keywords are still recognised', () => {
  for (const label of [
    'Responses',
    'HTTP responses',
    'Callbacks',
    'Acknowledges',
    'Acknowledgement',
    'Acknowledgment sent',
    'Returned payload',
    'Returning results',
    'Replies',
  ]) {
    assert.equal(classifyEdgeDirection(label), 'reverse', `"${label}" should be reverse`);
  }

  for (const label of [
    'Synchronize',
    'Synchronizes data',
    'Synchronization',
    'Synchronisation',
    'Syncs config',
    'Exchanges tokens',
    'Communicates with',
    'Communication channel',
  ]) {
    assert.equal(classifyEdgeDirection(label), 'bidirectional', `"${label}" should be bidirectional`);
  }
});

test('camelCase labels are split before matching', () => {
  assert.equal(classifyEdgeDirection('sendResponse'), 'reverse');
  assert.equal(classifyEdgeDirection('onCallback'), 'reverse');
  assert.equal(classifyEdgeDirection('syncState'), 'bidirectional');
  // Splitting must not manufacture a match that is not there.
  assert.equal(classifyEdgeDirection('backupBlob'), 'forward');
});

/**
 * Stemming must not become greedy: these are ordinary Azure vocabulary that
 * happens to start with a keyword stem and means something entirely different.
 */
test('stemming does not swallow unrelated Azure vocabulary', () => {
  for (const label of [
    'Replicates data',
    'Geo-replication',
    'Replica set',
    'Synchronous call',
    'Synchronously invokes',
    'Responsible service',
  ]) {
    assert.equal(classifyEdgeDirection(label), 'forward', `"${label}" should stay forward`);
  }
});

test('bidirectional wins over reverse, and unknown labels stay forward', () => {
  assert.equal(classifyEdgeDirection('Sync response'), 'bidirectional');
  assert.equal(classifyEdgeDirection('Sends telemetry'), 'forward');
  assert.equal(classifyEdgeDirection(''), 'forward');
  assert.equal(classifyEdgeDirection(undefined as unknown as string), 'forward');
});

test('punctuation and phrase keywords are handled', () => {
  assert.equal(classifyEdgeDirection('request/response'), 'reverse');
  assert.equal(classifyEdgeDirection('Two-way'), 'bidirectional');
  assert.equal(classifyEdgeDirection('two way channel'), 'bidirectional');
  assert.equal(classifyEdgeDirection('2-way link'), 'bidirectional');
  assert.equal(classifyEdgeDirection('2 way link'), 'bidirectional');
  assert.equal(classifyEdgeDirection('sync-config'), 'bidirectional');
  // A non-English label has no keywords and must not be forced either way.
  assert.equal(classifyEdgeDirection('データ送信'), 'forward');
});
