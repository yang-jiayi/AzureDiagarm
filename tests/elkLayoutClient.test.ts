import assert from 'node:assert/strict';
import test from 'node:test';
import type { ElkNode } from 'elkjs/lib/elk-api.js';
import { ElkLayoutClient, type ElkLayoutBackend } from '../src/utils/elkLayoutClient';

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason: unknown) => void;
  const promise = new Promise<T>((onResolve, onReject) => {
    resolve = onResolve;
    reject = onReject;
  });
  return { promise, resolve, reject };
}

class TestWorker extends EventTarget {
  terminations = 0;

  terminate() {
    this.terminations += 1;
  }
}

function backend(layout: ElkLayoutBackend['layout']): ElkLayoutBackend {
  return { layout, knownLayoutAlgorithms: async () => [] };
}

test('ELK worker client preserves out-of-order results and reuses its worker', async () => {
  const worker = new TestWorker();
  const first = deferred<ElkNode>();
  const second = deferred<ElkNode>();
  const client = new ElkLayoutClient(
    backend(graph => graph.id === 'first' ? first.promise : second.promise),
    worker,
  );
  await client.initializeWorker();
  const firstResult = client.layout({ id: 'first' });
  const secondResult = client.layout({ id: 'second' });
  second.resolve({ id: 'second', x: 200 });
  assert.deepEqual(await secondResult, { id: 'second', x: 200 });
  first.resolve({ id: 'first', x: 100 });
  assert.deepEqual(await firstResult, { id: 'first', x: 100 });
  assert.equal(worker.terminations, 0);
  client.dispose();
  assert.equal(worker.terminations, 1);
});

test('worker load errors reject initialization and do not run layout on the main thread', async () => {
  const worker = new TestWorker();
  const registration = deferred<[]>();
  let layoutCalls = 0;
  const client = new ElkLayoutClient({
    knownLayoutAlgorithms: () => registration.promise,
    layout: async graph => {
      layoutCalls += 1;
      return graph;
    },
  }, worker);
  const rejected = assert.rejects(client.initializeWorker(), /ELK layout worker failed/);
  worker.dispatchEvent(new Event('error'));
  await rejected;
  await assert.rejects(client.layout({ id: 'graph' }), /ELK layout worker failed/);
  registration.resolve([]);
  assert.equal(layoutCalls, 0);
  assert.equal(worker.terminations, 1);
});

test('message errors reject all pending layouts and ignore late replies', async () => {
  const worker = new TestWorker();
  const pending = deferred<ElkNode>();
  const client = new ElkLayoutClient(backend(() => pending.promise), worker);
  await client.initializeWorker();
  const first = client.layout({ id: 'first' });
  const second = client.layout({ id: 'second' });
  const rejected = [
    assert.rejects(first, /could not be deserialized/),
    assert.rejects(second, /could not be deserialized/),
  ];
  worker.dispatchEvent(new Event('messageerror'));
  await Promise.all(rejected);
  pending.resolve({ id: 'late' });
  await assert.rejects(client.layout({ id: 'third' }), /could not be deserialized/);
  assert.equal(worker.terminations, 1);
});

test('disposal rejects pending and queued requests and terminates only once', async () => {
  const worker = new TestWorker();
  let layoutCalls = 0;
  const client = new ElkLayoutClient(backend(async graph => {
    layoutCalls += 1;
    return graph;
  }), worker);
  await client.initializeWorker();
  const rejected = assert.rejects(client.layout({ id: 'queued' }), /disposed/);
  client.dispose();
  client.dispose();
  await rejected;
  assert.equal(layoutCalls, 0);
  assert.equal(worker.terminations, 1);
});

test('initialization timeout rejects explicitly without limiting layout execution time', async () => {
  const worker = new TestWorker();
  const registration = deferred<[]>();
  const client = new ElkLayoutClient({
    knownLayoutAlgorithms: () => registration.promise,
    layout: async graph => graph,
  }, worker);
  await assert.rejects(client.initializeWorker(5), /did not initialize/);
  assert.equal(client.failed, true);
  assert.equal(worker.terminations, 1);
  registration.resolve([]);

  const readyWorker = new TestWorker();
  const slowLayout = deferred<ElkNode>();
  const ready = new ElkLayoutClient(backend(() => slowLayout.promise), readyWorker);
  await ready.initializeWorker(5);
  const result = ready.layout({ id: 'slow-layout' });
  await new Promise(resolve => setTimeout(resolve, 15));
  assert.equal(ready.failed, false);
  slowLayout.resolve({ id: 'slow-layout', x: 20 });
  assert.deepEqual(await result, { id: 'slow-layout', x: 20 });
  ready.dispose();
});

test('an algorithm rejection does not poison the worker or silently retry the graph', async () => {
  const worker = new TestWorker();
  let layoutCalls = 0;
  const client = new ElkLayoutClient(backend(async graph => {
    layoutCalls += 1;
    if (graph.id === 'invalid') throw new Error('Invalid graph');
    return { ...graph, x: 40 };
  }), worker);
  await client.initializeWorker();
  await assert.rejects(client.layout({ id: 'invalid' }), /Invalid graph/);
  assert.deepEqual(await client.layout({ id: 'valid' }), { id: 'valid', x: 40 });
  assert.equal(layoutCalls, 2);
  assert.equal(client.failed, false);
  assert.equal(worker.terminations, 0);
  client.dispose();
});

test('synchronous backend errors reject their request without leaking pending work', async () => {
  const worker = new TestWorker();
  const client = new ElkLayoutClient(backend(graph => {
    if (graph.id === 'invalid') throw new Error('Synchronous failure');
    return Promise.resolve(graph);
  }), worker);
  await client.initializeWorker();
  await assert.rejects(client.layout({ id: 'invalid' }), /Synchronous failure/);
  assert.deepEqual(await client.layout({ id: 'valid' }), { id: 'valid' });
  client.dispose();
});

test('Node transport needs no native worker and still rejects disposed requests', async () => {
  const client = new ElkLayoutClient(backend(async graph => ({ ...graph, x: 10 })));
  await client.initializeWorker();
  assert.deepEqual(await client.layout({ id: 'node' }), { id: 'node', x: 10 });
  client.dispose();
  await assert.rejects(client.layout({ id: 'late' }), /disposed/);
});
