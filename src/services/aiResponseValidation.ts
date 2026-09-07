// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export class AIResponseValidationError extends Error {
  readonly source = 'client';
  readonly code = 'invalid_model_response';
  readonly retryable = false;
  readonly detail: string;

  constructor(detail: string) {
    super('The AI model returned an invalid architecture. Try again or revise the request.');
    this.name = 'AIResponseValidationError';
    this.detail = detail;
  }
}

export function isResponseObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

export function isResponseText(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

export function assertUniqueResponseIds(
  value: unknown,
  label: string,
): asserts value is Array<Record<string, unknown> & { id: string }> {
  if (!Array.isArray(value)) {
    throw new AIResponseValidationError(`${label} must be an array.`);
  }
  const ids = new Set<string>();
  for (const entry of value) {
    if (!isResponseObject(entry) || !isResponseText(entry.id)) {
      throw new AIResponseValidationError(`${label} must contain objects with non-empty string IDs.`);
    }
    if (ids.has(entry.id)) {
      throw new AIResponseValidationError(`${label} must have unique IDs; duplicate references are ambiguous.`);
    }
    ids.add(entry.id);
  }
}
