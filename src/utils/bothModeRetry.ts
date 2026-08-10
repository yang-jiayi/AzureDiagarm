// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Retry planning for the "topology + blueprint" generation mode.
 *
 * When one of the two deliverables fails, the other one has already been
 * applied (the topology is on the canvas, or the blueprint is stashed for
 * export). Re-running the whole pipeline would overwrite that result, rebuild
 * the brief into a "MODIFY EXISTING ARCHITECTURE" instruction against the
 * freshly applied canvas, and bill the successful half a second time. The
 * planner below keeps that decision pure so it can be tested directly.
 */

import type { ComponentManifest } from '../services/componentManifestAI';

export type BothDeliverable = 'topology' | 'blueprint';

/**
 * Snapshot of a partially failed run. Captured *before* the successful half is
 * applied, so a retry sees exactly the inputs the first attempt saw.
 */
export interface PendingRetry {
  missing: BothDeliverable;
  /** The brief as typed. A retry is only valid while the brief is unchanged. */
  brief: string;
  /** Context-enriched prompt from the original attempt. */
  prompt: string;
  /** Manifest from the original pre-pass, so it is not paid for twice. */
  manifest?: ComponentManifest;
  /** Whether the topology reached the canvas during the original attempt. */
  canvasApplied: boolean;
}

export interface BothRunPlan {
  /** The retry being replayed, or null for a normal full run. */
  retry: PendingRetry | null;
  runTopology: boolean;
  runBlueprint: boolean;
  /** True when the manifest pre-pass should be skipped and reused. */
  reuseManifest: boolean;
}

/**
 * Decide what a press of the generate button should actually run.
 *
 * A pending retry is honoured only while the brief is byte-identical to the
 * one that failed: an edited brief is a new request and must regenerate both
 * deliverables from scratch.
 */
export function planBothRun(
  pendingRetry: PendingRetry | null | undefined,
  description: string,
): BothRunPlan {
  const retry = pendingRetry && pendingRetry.brief === description ? pendingRetry : null;
  return {
    retry,
    runTopology: !retry || retry.missing === 'topology',
    runBlueprint: !retry || retry.missing === 'blueprint',
    reuseManifest: Boolean(retry),
  };
}
