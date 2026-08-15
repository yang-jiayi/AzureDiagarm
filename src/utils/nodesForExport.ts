// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Node } from 'reactflow';

/**
 * Hide from the exports whatever the canvas is hiding.
 *
 * Cost badges are suppressed on screen by two independent switches: the
 * `presentation` style preset, and a standalone preference that exists because
 * users asked for it in as many words — "cost as a fixed value is not
 * acceptable, i would rather hide it". The exporters read `data.pricing`
 * straight off the node and knew about neither, so the one number a user had
 * deliberately taken off the screen before showing the diagram to a customer
 * was printed on every tile of the deck, the Visio sheet and the HTML.
 *
 * Applied at the export boundary rather than inside each exporter: there are
 * six of them, the rule is the same for all six, and a seventh added later
 * inherits it for free.
 */
export function nodesForExport(nodes: Node[], showCostBadges: boolean): Node[] {
  return nodes.map((node) => {
    const data = (node.data ?? {}) as Record<string, unknown>;
    if (!data.pricing) return node;
    const hidden = !showCostBadges || data.stylePreset === 'presentation';
    if (!hidden) return node;
    const { pricing: _dropped, ...rest } = data;
    return { ...node, data: rest };
  });
}
