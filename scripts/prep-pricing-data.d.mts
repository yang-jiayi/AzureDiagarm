/**
 * Types for the JS data script `prep-pricing-data.mjs`.
 *
 * Kept `.mjs` so the pricing refresh can run under bare node, but
 * `tests/pricing-prep.test.ts` imports it and `tsconfig.scripts.json` has to be
 * able to check that. The payloads really are free-form Retail-Prices JSON, so
 * `unknown` in and `Record<string, unknown>` out is the honest declaration
 * rather than an invented schema.
 */

export function compactPricingData(
  data: unknown,
  options?: { keepProductName?: boolean },
): Record<string, unknown>;
export function expandPricingData(compact: unknown): Record<string, unknown>;
