// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface DiagramHistoryEntry<T> {
  fingerprint: string;
  state: T;
}

const cloneState = <T>(state: T): T => JSON.parse(JSON.stringify(state)) as T;

export class DiagramHistory<T> {
  private entries: DiagramHistoryEntry<T>[] = [];
  private index = -1;

  constructor(
    initialState: T,
    private readonly limit = 50,
  ) {
    this.reset(initialState);
  }

  record(state: T): boolean {
    const next = cloneState(state);
    const fingerprint = JSON.stringify(next);
    if (this.entries[this.index]?.fingerprint === fingerprint) return false;

    this.entries = this.entries.slice(0, this.index + 1);
    this.entries.push({ fingerprint, state: next });
    if (this.entries.length > this.limit) {
      this.entries = this.entries.slice(this.entries.length - this.limit);
    }
    this.index = this.entries.length - 1;
    return true;
  }

  reset(state: T): void {
    const next = cloneState(state);
    this.entries = [{ fingerprint: JSON.stringify(next), state: next }];
    this.index = 0;
  }

  canUndo(): boolean {
    return this.index > 0;
  }

  canRedo(): boolean {
    return this.index >= 0 && this.index < this.entries.length - 1;
  }

  undo(): T | null {
    if (!this.canUndo()) return null;
    this.index -= 1;
    return cloneState(this.entries[this.index].state);
  }

  redo(): T | null {
    if (!this.canRedo()) return null;
    this.index += 1;
    return cloneState(this.entries[this.index].state);
  }
}
