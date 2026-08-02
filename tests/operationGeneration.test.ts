import assert from 'node:assert/strict';
import test from 'node:test';
import { OperationGeneration } from '../src/utils/operationGeneration';

test('OperationGeneration invalidates older asynchronous work', () => {
  const generations = new OperationGeneration();
  const first = generations.advance();
  const second = generations.advance();

  assert.equal(generations.isCurrent(first), false);
  assert.equal(generations.isCurrent(second), true);
  assert.equal(generations.current(), second);
});

test('capturing the current generation remains valid until a new intent starts', () => {
  const generations = new OperationGeneration();
  const initial = generations.current();

  assert.equal(generations.isCurrent(initial), true);
  generations.advance();
  assert.equal(generations.isCurrent(initial), false);
});

test('activating a fenced transition invalidates work started during the transition', () => {
  const generations = new OperationGeneration();
  const transition = generations.advance();
  const interimSave = generations.current();

  assert.equal(generations.isCurrent(transition), true);
  assert.equal(generations.isCurrent(interimSave), true);

  const activatedDocument = generations.advance();
  assert.equal(generations.isCurrent(interimSave), false);
  assert.equal(generations.isCurrent(activatedDocument), true);
});
