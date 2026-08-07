// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type GenerationMode = 'topology' | 'reference' | 'blueprint' | 'both';

export function generationProducesCanvas(mode: GenerationMode): boolean {
  return mode === 'topology' || mode === 'both';
}
