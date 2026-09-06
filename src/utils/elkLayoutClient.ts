// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { ElkLayoutAlgorithmDescription, ElkNode } from 'elkjs/lib/elk-api.js';

export interface ElkLayoutBackend {
  layout(graph: ElkNode): Promise<ElkNode>;
  knownLayoutAlgorithms(): Promise<ElkLayoutAlgorithmDescription[]>;
}

type WorkerLifecycle = EventTarget & { terminate(): void };

// ELK's promise adapter handles replies, but not native worker load/crash errors.
export class ElkLayoutClient {
  private failure: Error | null = null;
  private readonly pending = new Set<(reason: unknown) => void>();

  constructor(
    private readonly backend: ElkLayoutBackend,
    private readonly worker?: WorkerLifecycle,
  ) {
    worker?.addEventListener('error', this.onWorkerError);
    worker?.addEventListener('messageerror', this.onMessageError);
  }

  get failed(): boolean {
    return this.failure !== null;
  }

  async initializeWorker(timeoutMs = 30_000): Promise<void> {
    if (!this.worker) return;
    const timer = setTimeout(() => {
      this.fail(new Error(`The ELK layout worker did not initialize within ${timeoutMs} ms.`));
    }, timeoutMs);
    try {
      await this.request(() => this.backend.knownLayoutAlgorithms());
    } catch (error) {
      this.dispose();
      throw error;
    } finally {
      clearTimeout(timer);
    }
  }

  layout(graph: ElkNode): Promise<ElkNode> {
    return this.request(() => this.backend.layout(graph));
  }

  dispose(): void {
    this.fail(new Error('The ELK layout runtime was disposed.'));
  }

  private request<T>(operation: () => Promise<T>): Promise<T> {
    if (this.failure) return Promise.reject(this.failure);
    return new Promise<T>((resolve, reject) => {
      this.pending.add(reject);
      Promise.resolve().then(() => {
        if (this.failure) throw this.failure;
        return operation();
      }).then(
        value => {
          if (this.pending.delete(reject)) resolve(value);
        },
        (error: unknown) => {
          if (this.pending.delete(reject)) reject(error);
        },
      );
    });
  }

  private readonly onWorkerError = (event: Event): void => {
    const detail = typeof ErrorEvent !== 'undefined' && event instanceof ErrorEvent
      ? event.message
      : '';
    this.fail(new Error(`The ELK layout worker failed${detail ? `: ${detail}` : '.'}`));
  };

  private readonly onMessageError = (): void => {
    this.fail(new Error('An ELK layout worker response could not be deserialized.'));
  };

  private fail(error: Error): void {
    if (this.failure) return;
    this.failure = error;
    this.worker?.removeEventListener('error', this.onWorkerError);
    this.worker?.removeEventListener('messageerror', this.onMessageError);
    this.worker?.terminate();
    for (const reject of this.pending) reject(error);
    this.pending.clear();
  }
}
