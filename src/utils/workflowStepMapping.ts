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

export interface NormalizedWorkflow {
  steps: Array<WorkflowStepInput & { step: number; description: string; services: string[] }>;
  /** Service references rewritten from a name/alias to a real service id. */
  repairedRefs: number;
  /** Steps discarded because they could never describe a hop. */
  droppedSteps: number;
}

/**
 * Repair an AI-authored workflow so its steps can actually be numbered.
 *
 * Models emit display names where ids are required just as often in `services`
 * as they do on connection endpoints, and they skip or repeat step numbers. An
 * unresolvable step silently produces an unnumbered diagram, which is worse
 * than an obviously wrong one because nothing signals that anything is missing.
 *
 * `resolveRef` is the caller's alias table (id, name and type all map to an
 * id); it returns null when a reference cannot be resolved at all.
 */
export function normalizeWorkflowSteps(
  workflow: unknown,
  resolveRef: (ref: unknown) => string | null,
): NormalizedWorkflow {
  let repairedRefs = 0;
  let droppedSteps = 0;
  const source = Array.isArray(workflow) ? workflow : [];

  const steps = source
    .filter((entry): entry is Record<string, unknown> => !!entry && typeof entry === 'object')
    .map((entry) => {
      const refs = Array.isArray(entry.services) ? entry.services : [];
      const services: string[] = [];
      for (const ref of refs) {
        const id = resolveRef(ref);
        if (!id) continue;
        if (id !== String(ref ?? '')) repairedRefs += 1;
        // Keep flow order, but a repeat must not stand in for a second service.
        if (!services.includes(id)) services.push(id);
      }
      return { ...entry, services, description: String(entry.description ?? '').trim() };
    })
    .filter((entry) => {
      // A step needs both ends of a hop, and prose, or it numbers nothing.
      if (entry.services.length >= 2 && entry.description) return true;
      droppedSteps += 1;
      return false;
    })
    // Renumber contiguously from 1: a gap or a duplicate makes the badge on the
    // arrow disagree with the numbered list printed beside it.
    .map((entry, index) => ({ ...entry, step: index + 1 })) as NormalizedWorkflow['steps'];

  return { steps, repairedRefs, droppedSteps };
}
