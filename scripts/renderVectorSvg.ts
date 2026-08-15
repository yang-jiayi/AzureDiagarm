// Render a realistic diagram through the native SVG exporter, with real Azure
// icon artwork read off disk, so the file can be opened in a non-browser tool.
// Usage: npx tsx scripts/renderVectorSvg.ts <out.svg> [dark]

import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import type { Node, Edge } from 'reactflow';

import { exportToSvg } from '../src/services/vectorSvgExporter';

const ROOT = join(process.cwd(), 'Azure_Public_Service_Icons', 'Icons');

function findIcon(needle: string): { path: string; svg: string } | null {
  const walk = (dir: string): string[] =>
    readdirSync(dir).flatMap((entry) => {
      const full = join(dir, entry);
      return statSync(full).isDirectory() ? walk(full) : [full];
    });
  const hit = walk(ROOT).find((p) => p.toLowerCase().includes(needle.toLowerCase()));
  return hit ? { path: hit, svg: readFileSync(hit, 'utf8') } : null;
}

const SPEC: Array<[string, string, string, string, number, number]> = [
  ['fd', 'Front Door', 'networking', 'Front-Door', 60, 60],
  ['app', 'App Service', 'compute', 'App-Services', 300, 60],
  ['func', 'Functions', 'compute', 'Function-Apps', 300, 200],
  ['sql', 'Azure SQL Database', 'databases', 'SQL-Database', 560, 60],
  ['cosmos', 'Cosmos DB', 'databases', 'Azure-Cosmos-DB', 560, 200],
  ['kv', 'Key Vault', 'security', 'Key-Vaults', 820, 130],
];

const nodes: Node[] = [
  {
    id: 'zone-app',
    type: 'groupNode',
    position: { x: 20, y: 10 },
    style: { width: 740, height: 320 },
    data: { label: 'Production VNet' },
  } as Node,
  ...SPEC.map(
    ([id, label, category, , x, y]) =>
      ({
        id,
        type: 'azureNode',
        position: { x, y },
        width: 150,
        height: 90,
        data: {
          label,
          category,
          iconPath: id,
          pricing: { estimatedCost: 128, tier: 'Standard', region: 'japaneast' },
        },
      }) as Node,
  ),
];

const edges: Edge[] = [
  { id: 'e1', source: 'fd', target: 'app', label: 'HTTPS', data: { connectionType: 'network', stepNumber: 1 } },
  { id: 'e2', source: 'app', target: 'sql', label: 'TDS', data: { connectionType: 'data', stepNumber: 2 } },
  { id: 'e3', source: 'app', target: 'func', data: { connectionType: 'event', stepNumber: 3 } },
  { id: 'e4', source: 'func', target: 'cosmos', label: 'writes', data: { connectionType: 'data', stepNumber: 4 } },
  { id: 'e5', source: 'app', target: 'kv', label: 'secrets', data: { connectionType: 'identity', stepNumber: 5 } },
] as Edge[];

const presetIcons = new Map<string, string>();
for (const [id, , , needle] of SPEC) {
  const icon = findIcon(needle);
  if (icon) presetIcons.set(id, icon.svg);
  else console.warn(`no icon matched "${needle}"`);
}
console.log(`resolved ${presetIcons.size}/${SPEC.length} icons`);

const out = process.argv[2] ?? 'tmp-render/vector.svg';
const dark = process.argv[3] === 'dark';

const svg = await exportToSvg(nodes, edges, {
  isDarkMode: dark,
  title: 'Reference architecture',
  presetIcons,
});
writeFileSync(out, svg, 'utf8');
console.log(`wrote ${out} (${svg.length} bytes)`);
