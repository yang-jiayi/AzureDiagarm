// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { ElkNode } from 'elkjs/lib/elk-api.js';
import { ElkLayoutClient } from './elkLayoutClient';

let generation = 0;
let clientPromise: Promise<ElkLayoutClient> | undefined;
let activeClient: ElkLayoutClient | undefined;

async function createClient(sourceGeneration: number): Promise<ElkLayoutClient> {
  let client: ElkLayoutClient;
  if (typeof window === 'undefined') {
    // Node-based tools retain ELK's bundled transport; browser failures never
    // enter this branch or fall back to main-thread layout.
    const { default: Elk } = await import('elkjs/lib/elk.bundled.js');
    if (sourceGeneration !== generation) throw new Error('The ELK layout runtime was disposed.');
    client = new ElkLayoutClient(new Elk());
  } else {
    if (typeof Worker === 'undefined') {
      throw new Error('This browser does not support Web Workers for ELK layout.');
    }
    const [{ default: Elk }, { default: workerAsset }] = await Promise.all([
      import('elkjs/lib/elk-api.js'),
      import('elkjs/lib/elk-worker.min.js?url'),
    ]);
    if (sourceGeneration !== generation) throw new Error('The ELK layout runtime was disposed.');
    const workerUrl = new URL(workerAsset, window.location.href);
    if (workerUrl.origin !== window.location.origin) {
      throw new Error('The ELK layout worker must be served from the application origin.');
    }
    const worker = new Worker(workerUrl, { name: 'azure-diagram-elk' });
    try {
      client = new ElkLayoutClient(new Elk({ workerFactory: () => worker }), worker);
    } catch (error) {
      worker.terminate();
      throw error;
    }
  }
  activeClient = client;
  await client.initializeWorker();
  if (sourceGeneration !== generation) {
    client.dispose();
    throw new Error('The ELK layout runtime was disposed.');
  }
  return client;
}

export async function layoutWithElk(graph: ElkNode): Promise<ElkNode> {
  const pending = clientPromise ??= createClient(generation);
  let client: ElkLayoutClient | undefined;
  try {
    client = await pending;
    return await client.layout(graph);
  } catch (error) {
    // A later explicit request may create a new worker, but this request and
    // every request using a failed worker reject rather than silently retry.
    if ((!client || client.failed) && clientPromise === pending) {
      clientPromise = undefined;
      activeClient = undefined;
    }
    throw error;
  }
}

export function disposeElkLayout(): void {
  generation += 1;
  activeClient?.dispose();
  activeClient = undefined;
  clientPromise = undefined;
}

if (typeof window !== 'undefined') {
  window.addEventListener('pagehide', disposeElkLayout);
  import.meta.hot?.dispose(() => {
    window.removeEventListener('pagehide', disposeElkLayout);
    disposeElkLayout();
  });
}
