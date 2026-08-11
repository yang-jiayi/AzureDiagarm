// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Maps `workflow[]` steps onto diagram edges so an arrow can carry the same
 * number as the step that describes it.
 *
 * This is the convention every reference architecture on the Azure
 * Architecture Center follows: the diagram carries numbered callouts and the
 * prose underneath is a numbered list that matches. Without the mapping the
 * two halves of an export are unrelated and the reader has to guess which
 * arrow a sentence is talking about.
 *
 * The AI returns each step's `services` in flow order but does not number the
 * connections, so the mapping is recovered here and shared by every consumer
 * (the animated SVG, the static canvas, and all four export formats) to stop
 * them numbering the same diagram differently.
 */

export interface WorkflowStepInput {
  step: number;
  description?: string;
  services?: string[];
}

export interface StepMappableEdge {
  id: string;
  source: string;
  target: string;
}

/**
 * Assign at most one step number to each edge, and use each step at most once.
 *
 * A step matches an edge only when *both* endpoints appear in that step's
 * `services`. When several edges qualify, the one whose target appears latest
 * in `services` wins, because the list is authored in flow order and the last
 * service named is the step's destination. Ties fall back to edge id so the
 * numbering is stable across runs and identical in every export.
 */
export function mapWorkflowStepsToEdges(
  edges: readonly StepMappableEdge[],
  workflow: readonly WorkflowStepInput[] | undefined | null,
): Map<string, number> {
  const assigned = new Map<string, number>();
  if (!Array.isArray(workflow) || workflow.length === 0) return assigned;

  const ordered = [...workflow]
    .filter((entry): entry is WorkflowStepInput => !!entry && Number.isFinite(entry.step))
    .sort((a, b) => a.step - b.step);

  for (const entry of ordered) {
    const services = Array.isArray(entry.services) ? entry.services : [];
    if (services.length === 0) continue;
    const membership = new Set(services);

    const candidates = edges.filter(
      (edge) => !assigned.has(edge.id) && membership.has(edge.source) && membership.has(edge.target),
    );
    if (candidates.length === 0) continue;

    const best = candidates.reduce((winner, edge) => {
      const rank = services.indexOf(edge.target);
      const winnerRank = services.indexOf(winner.target);
      if (rank !== winnerRank) return rank > winnerRank ? edge : winner;
      return String(edge.id).localeCompare(String(winner.id)) < 0 ? edge : winner;
    });

    assigned.set(best.id, entry.step);
  }

  return assigned;
}
