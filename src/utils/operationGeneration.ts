// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Tracks the latest asynchronous operation intent.
 *
 * Call advance() whenever a newer operation supersedes older work, then check
 * isCurrent() before applying an asynchronous result.
 */
export class OperationGeneration {
  private value = 0;

  advance(): number {
    this.value += 1;
    return this.value;
  }

  current(): number {
    return this.value;
  }

  isCurrent(generation: number): boolean {
    return generation === this.value;
  }
}
