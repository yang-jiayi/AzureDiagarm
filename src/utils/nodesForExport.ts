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
 *
 * The whole `pricing` object goes, not just the money in it, and that is
 * deliberate on two counts. The SKU and region the exports print in the meta
 * subline are read off `pricing`, but the canvas only ever shows them in a
 * tooltip — so dropping them moves the file *closer* to the screen, not
 * further from it. And several cost sections (the workflow narrative's, for
 * one) are gated on "does any node still carry pricing", so leaving a hollow
 * `pricing` object behind would switch those sections back on with the total
 * the user just hid.
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
