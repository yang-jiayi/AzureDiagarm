import { copyFile, mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(fileURLToPath(new URL('.', import.meta.url)), '..');
const assets = [
  'iconMap.generated.json',
  'iconSvgs.generated.json',
  'pricing.generated.json',
];

await mkdir(resolve(root, 'dist'), { recursive: true });
await Promise.all(
  assets.map((asset) =>
    copyFile(resolve(root, 'src', asset), resolve(root, 'dist', asset)),
  ),
);
