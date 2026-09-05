import type { Edge, Node } from 'reactflow';
import type { CloudDiagramPayload } from './cloudDiagramService';
import type { EditorDocument } from './editorHistory';

function stripTransientNodeState(node: Node): Node {
  const snapshot = { ...node } as Node & Record<string, unknown>;
  delete snapshot.selected;
  delete snapshot.dragging;
  delete snapshot.width;
  delete snapshot.height;
  delete snapshot.positionAbsolute;
  delete snapshot.measured;
  delete snapshot.resizing;
  return snapshot;
}

function stripTransientEdgeState(edge: Edge): Edge {
  const snapshot = { ...edge };
  delete snapshot.selected;
  return snapshot;
}

export function toCloudDiagramPayload(document: EditorDocument): CloudDiagramPayload {
  return {
    nodes: document.nodes.map(stripTransientNodeState),
    edges: document.edges.map(stripTransientEdgeState),
    architecturePrompt: document.architecturePrompt,
    originalPrompt: document.originalPrompt || document.architecturePrompt || undefined,
    validationScore: document.validationScore,
    titleBlockData: document.titleBlockData,
    workflow: document.workflow,
    pricingScenarios: document.pricingScenarios,
    iacBaseline: document.iacBaseline ?? null,
    settings: document.settings,
    reviewHistory: document.reviewHistory ?? [],
    validationSourceFingerprint: document.validationSourceFingerprint ?? null,
  };
}
