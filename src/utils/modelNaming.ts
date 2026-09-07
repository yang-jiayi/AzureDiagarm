// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/** Public metadata captured on an artifact, not the currently selected connection. */
export interface AIArtifactProvenance {
  source?: string;
  model?: string;
  deployment?: string;
  reasoningEffort?: string;
  isReasoning?: boolean;
}

function slug(value: unknown, maxLength: number): string {
  return typeof value === 'string'
    ? value.trim().toLowerCase().replace(/[^a-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, maxLength) : '';
}

export function getModelSuffix(provenance?: AIArtifactProvenance): string {
  if (!provenance) return '';
  const isBYO = provenance.source === 'bring-your-own';
  const identifier = isBYO ? provenance.deployment || provenance.model : provenance.model || provenance.deployment;
  const model = typeof identifier === 'string' && /^gpt[-\s]?6[-\s]?astra$/i.test(identifier.trim())
    ? 'gpt6astra' : slug(identifier, 80);
  if (!model) return '';
  const effort = provenance.isReasoning === false ? 'none' : slug(provenance.reasoningEffort, 24);
  return `${isBYO ? 'byo-' : ''}${model}${effort ? `-${effort}` : ''}`;
}

export function generateModelFilename(
  prefix: string,
  extension: string,
  timestamp?: number,
  provenance?: AIArtifactProvenance,
): string {
  const ts = timestamp ?? Date.now();
  const suffix = getModelSuffix(provenance);
  return `${prefix}-${ts}${suffix ? `-${suffix}` : ''}.${extension}`;
}
