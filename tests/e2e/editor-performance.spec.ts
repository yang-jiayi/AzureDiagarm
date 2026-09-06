import { expect, test, type Page, type Route } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import type { Edge, Node } from 'reactflow';
import { applyLayoutPreset } from '../../src/utils/layoutPresets';
import { buildStarterTemplate, type StarterTemplateFormat } from '../../src/services/iacRoundTrip';
import { elkLayoutFixtures } from '../fixtures/elkLayoutParity';

interface WorkerRecord {
  url: string;
  requests: number[];
  replies: number[];
  held: Array<() => void>;
  terminations: number;
}

declare global {
  interface Window {
    __elkWorkerProbe: {
      records: WorkerRecord[];
      holdNext: number;
      release(): void;
    };
  }
}

interface Diagram {
  nodes: Node[];
  edges: Edge[];
}

const layoutModule = process.env.ELK_BROWSER_HARNESS || '/src/utils/layoutPresets.ts';
const runtimeModule = process.env.ELK_BROWSER_HARNESS || '/src/utils/elkLayoutRuntime.ts';

async function prepare(page: Page) {
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
    localStorage.setItem('azure-diagram-builder.headerCollapsed.v1', '0');
    localStorage.setItem('azure-diagram-builder.focusMode.v1', '0');
    localStorage.setItem('azure-diagram-builder.canvasHintDismissed.v1', '1');
    const probe: Window['__elkWorkerProbe'] = {
      records: [], holdNext: 0,
      release() {
        for (const record of this.records) {
          for (const send of record.held.splice(0)) send();
        }
      },
    };
    window.__elkWorkerProbe = probe;
    const NativeWorker = window.Worker;
    window.Worker = class extends NativeWorker {
      private readonly record: WorkerRecord;

      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.record = { url: String(url), requests: [], replies: [], held: [], terminations: 0 };
        probe.records.push(this.record);
        this.addEventListener('message', event => {
          const id: unknown = event.data?.id;
          if (typeof id === 'number' && this.record.requests.includes(id)) {
            this.record.replies.push(id);
          }
        });
      }

      postMessage(message: unknown, transfer?: Transferable[] | StructuredSerializeOptions): void {
        const send = () => {
          if (Array.isArray(transfer)) super.postMessage(message, transfer);
          else super.postMessage(message, transfer);
        };
        if (typeof message === 'object' && message !== null
          && 'cmd' in message && message.cmd === 'layout'
          && 'id' in message && typeof message.id === 'number') {
          this.record.requests.push(message.id);
          if (probe.holdNext > 0) {
            probe.holdNext -= 1;
            this.record.held.push(send);
            return;
          }
        }
        send();
      }

      terminate(): void {
        this.record.terminations += 1;
        super.terminate();
      }
    };
  });
  await page.route('**/*', route => {
    const url = new URL(route.request().url());
    if (!['127.0.0.1', 'localhost'].includes(url.hostname)) return route.abort();
    if (url.pathname === '/api/access/me') return route.fulfill({
      json: { enabled: false, authenticated: false, allowed: true, email: null, isAdmin: false },
    });
    if (url.pathname === '/api/ai/budget') return route.fulfill({
      json: {
        available: true, limitTokens: 100000, usedTokens: 0, reservedTokens: 0, remainingTokens: 100000,
        concurrentRequests: 0, concurrentLimit: 2, resetAt: '2030-01-01T00:00:00Z', mode: 'public',
      },
    });
    if (url.pathname.startsWith('/api/')) return route.fulfill({
      status: 503, json: { error: 'Local editor performance fixture: no backend calls are permitted.' },
    });
    return route.continue();
  });
  await page.goto('/');
  await page.locator('input[aria-label="Load diagram"]').waitFor({ state: 'attached' });
}

async function frames(page: Page) {
  await page.evaluate(async () => {
    await new Promise(resolve => setTimeout(resolve, 0));
    await new Promise<void>(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
  });
}

function smallDiagram(): Diagram {
  const fixture = elkLayoutFixtures()[0];
  const nodes = fixture.nodes.slice(0, 3);
  const ids = new Set(nodes.map(node => node.id));
  return { nodes, edges: fixture.edges.filter(edge => ids.has(edge.source) && ids.has(edge.target)) };
}

async function loadDiagram(page: Page, diagram: Diagram) {
  await page.locator('input[aria-label="Load diagram"]').setInputFiles({
    name: 'editor-performance.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      ...diagram, workflow: [], viewport: { x: 30, y: 30, zoom: 1 },
      settings: {
        pricingMode: 'payg', pricingRegion: 'eastus2', stylePreset: 'detailed',
        edgeStyle: 'orthogonal', layoutEngine: 'elk', layoutPreset: 'flow-lr',
        layoutSpacing: 'comfortable', emphasizePrimaryPath: false,
        animateConnections: false, showCostBadges: false,
      },
    })),
  });
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(
    diagram.nodes.filter(node => node.type === 'azureNode').length,
  );
  await frames(page);
}

async function downloadedText(page: Page, click: () => Promise<unknown>): Promise<string> {
  const download = page.waitForEvent('download');
  await click();
  const stream = await (await download).createReadStream();
  if (!stream) throw new Error('The editor did not produce a readable download.');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}

async function saveDiagram(page: Page): Promise<Diagram> {
  await page.locator('#ribbon-tab-home').click();
  return JSON.parse(await downloadedText(page, () => page.getByRole('button', { name: 'Save', exact: true }).click()));
}

async function arrange(page: Page, preset: 'flow-lr' | 'flow-tb', emphasize = false) {
  await page.locator('#ribbon-tab-design').click();
  await page.getByRole('button', { name: 'Layout', exact: true }).click();
  await page.getByLabel('Layout preset', { exact: true }).selectOption(preset);
  await page.locator('#layoutEngine').selectOption('elk');
  await page.getByLabel('Emphasize primary path', { exact: true }).setChecked(emphasize);
  await page.getByRole('menuitem', { name: 'Apply Layout', exact: true }).click();
}

async function holdNextLayout(page: Page) {
  await page.evaluate(() => { window.__elkWorkerProbe.holdNext = 1; });
}

async function expectHeldLayout(page: Page) {
  await expect.poll(() => page.evaluate(
    () => window.__elkWorkerProbe.records.reduce((count, record) => count + record.held.length, 0),
  )).toBe(1);
}

async function expectReplies(page: Page, expected: number) {
  await expect.poll(() => page.evaluate(
    () => window.__elkWorkerProbe.records.reduce((count, record) => count + record.replies.length, 0),
  )).toBe(expected);
  await frames(page);
}

test('real ELK worker matches Node layouts and uses a same-origin JavaScript asset', async ({ page }) => {
  test.setTimeout(120_000);
  const requested: string[] = [];
  page.on('request', request => requested.push(request.url()));
  await prepare(page);
  expect(await page.evaluate(() => window.__elkWorkerProbe.records.length)).toBe(0);
  expect(requested.filter(url => url.includes('elk-api') || url.includes('elk-worker.min'))).toEqual([]);
  for (const fixture of elkLayoutFixtures()) {
    const expected = await applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
    const actual = await page.evaluate(async ({ fixture, moduleUrl }) => {
      const module = await import(moduleUrl);
      return module.applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
    }, { fixture, moduleUrl: layoutModule });
    expect(actual, fixture.name).toEqual(expected);
  }
  const records = await page.evaluate(() => window.__elkWorkerProbe.records.map(record => ({
    url: record.url, requests: record.requests.length, replies: record.replies.length,
  })));
  expect(records).toHaveLength(1);
  expect(records[0].requests).toBe(elkLayoutFixtures().length);
  expect(records[0].replies).toBe(elkLayoutFixtures().length);
  expect(new URL(records[0].url).origin).toBe(new URL(page.url()).origin);
  expect(requested.filter(url => url.includes('elk.bundled'))).toEqual([]);
  const asset = await page.request.get(records[0].url);
  expect(asset.ok()).toBe(true);
  expect(asset.headers()['content-type']).toMatch(/javascript/);
  if (process.env.ELK_BROWSER_HARNESS) {
    const distribution = await readFile(new URL('../../node_modules/elkjs/lib/elk-worker.min.js', import.meta.url));
    expect(createHash('sha256').update(await asset.body()).digest('hex'))
      .toBe(createHash('sha256').update(distribution).digest('hex'));
  }
  await page.evaluate(async moduleUrl => (await import(moduleUrl)).disposeElkLayout(), runtimeModule);
  expect(await page.evaluate(() => window.__elkWorkerProbe.records[0].terminations)).toBe(1);
});

test('worker asset failures reject, terminate, and retry only on another explicit request', async ({ page }) => {
  test.setTimeout(90_000);
  await prepare(page);
  const blockWorker = async (route: Route) => {
    const isWorkerScript = await page.evaluate(url => window.__elkWorkerProbe.records.some(
      record => new URL(record.url, window.location.href).href === url,
    ), route.request().url());
    if (isWorkerScript) {
      await route.abort();
    } else {
      await route.fallback();
    }
  };
  await page.route('**/*elk-worker.min*', blockWorker);
  const fixture = elkLayoutFixtures()[0];
  await expect(page.evaluate(async ({ fixture, moduleUrl }) => {
    const module = await import(moduleUrl);
    return module.applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  }, { fixture, moduleUrl: layoutModule })).rejects.toThrow(/ELK layout worker/);
  expect(await page.evaluate(() => window.__elkWorkerProbe.records.length)).toBe(1);
  expect(await page.evaluate(() => window.__elkWorkerProbe.records[0].terminations)).toBe(1);
  await page.unroute('**/*elk-worker.min*', blockWorker);
  const result = await page.evaluate(async ({ fixture, moduleUrl }) => {
    const module = await import(moduleUrl);
    return module.applyLayoutPreset(fixture.nodes, fixture.edges, fixture.options);
  }, { fixture, moduleUrl: layoutModule });
  expect(result.nodes).toHaveLength(fixture.nodes.length);
  expect(await page.evaluate(() => window.__elkWorkerProbe.records.length)).toBe(2);
});

test('an older worker result cannot change a newer Arrange result', async ({ page }) => {
  await prepare(page);
  await loadDiagram(page, smallDiagram());
  await holdNextLayout(page);
  await arrange(page, 'flow-lr', true);
  await expectHeldLayout(page);
  await arrange(page, 'flow-tb', false);
  await expectReplies(page, 1);
  const newest = await saveDiagram(page);
  await page.evaluate(() => window.__elkWorkerProbe.release());
  await expectReplies(page, 2);
  const afterOldReply = await saveDiagram(page);
  expect(afterOldReply.nodes).toEqual(newest.nodes);
  expect(afterOldReply.edges).toEqual(newest.edges);
});

test('dragging during real-worker layout preserves the manual move while arranging other nodes', async ({ page }) => {
  await prepare(page);
  await loadDiagram(page, smallDiagram());
  const before = await saveDiagram(page);
  await holdNextLayout(page);
  await arrange(page, 'flow-lr');
  await expectHeldLayout(page);
  const box = await page.locator('.react-flow__node[data-id="service-0"] .azure-node').boundingBox();
  if (!box) throw new Error('The draggable fixture node is not visible.');
  await page.mouse.move(box.x + 35, box.y + 40);
  await page.mouse.down();
  await page.mouse.move(box.x + 155, box.y + 80, { steps: 6 });
  await page.mouse.up();
  const edited = await saveDiagram(page);
  expect(edited.nodes[0].position).not.toEqual(before.nodes[0].position);
  await page.evaluate(() => window.__elkWorkerProbe.release());
  await expectReplies(page, 1);
  const arranged = await saveDiagram(page);
  expect(arranged.nodes[0].position).toEqual(edited.nodes[0].position);
  expect(arranged.nodes[1].position).not.toEqual(before.nodes[1].position);
  expect(arranged.nodes.map(node => node.id)).toEqual(before.nodes.map(node => node.id));
});

test('loading another document while layout is pending rejects the old lineage result', async ({ page }) => {
  await prepare(page);
  await loadDiagram(page, smallDiagram());
  await holdNextLayout(page);
  await arrange(page, 'flow-lr');
  await expectHeldLayout(page);
  const replacement = smallDiagram();
  replacement.nodes = replacement.nodes.map(node => ({
    ...node, data: { ...node.data, label: String(node.data.label).replace('Service', 'Replica') },
  }));
  await loadDiagram(page, replacement);
  await expect(page.locator('.react-flow__node[data-id="service-0"]')).toContainText('Replica 0');
  const loaded = await saveDiagram(page);
  await page.evaluate(() => window.__elkWorkerProbe.release());
  await expectReplies(page, 1);
  const afterOldReply = await saveDiagram(page);
  expect(afterOldReply.nodes).toEqual(loaded.nodes);
  expect(afterOldReply.edges).toEqual(loaded.edges);
});

test('IaC downloads stay fresh and open-modal state survives unrelated App rerenders', async ({ page, context }) => {
  test.setTimeout(90_000);
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await prepare(page);
  const initial = smallDiagram();
  await loadDiagram(page, initial);
  const initialDocument = await saveDiagram(page);
  const open = async () => {
    await page.locator('#ribbon-tab-review').click();
    await page.getByRole('button', { name: 'IaC Round-trip', exact: true }).click();
    await expect(page.getByRole('dialog', { name: 'IaC round-trip and drift', exact: true })).toBeVisible();
  };
  const download = async (format: StarterTemplateFormat) => {
    const title = format === 'bicep' ? 'Bicep' : 'Terraform';
    const card = page.locator('.iac-starter-card').filter({ has: page.getByText(title, { exact: true }) });
    return downloadedText(page, () => card.getByRole('button', { name: 'Download starter', exact: true }).click());
  };
  await open();
  for (const format of ['bicep', 'terraform'] satisfies StarterTemplateFormat[]) {
    expect(await download(format)).toBe(buildStarterTemplate(initialDocument.nodes, format, initialDocument.edges).content);
  }
  const modal = page.locator('.iac-roundtrip-modal');
  await modal.evaluate(element => element.setAttribute('data-state-probe', 'retained'));
  const copy = modal.locator('.iac-command-card .copy-button').first();
  await copy.click();
  await page.locator('#ribbon-tab-design').dispatchEvent('click');
  await expect(modal).toHaveAttribute('data-state-probe', 'retained');
  await expect(copy.locator('.lucide-check')).toBeVisible();
  await page.getByRole('button', { name: 'Close drift report', exact: true }).click();

  const changed = smallDiagram();
  changed.nodes = changed.nodes.map(node => ({
    ...node, data: { ...node.data, label: `Revised ${node.id}` },
  }));
  changed.edges = changed.edges.map(edge => ({ ...edge, data: { ...edge.data, direction: 'reverse' } }));
  await loadDiagram(page, changed);
  const changedDocument = await saveDiagram(page);
  await open();
  for (const format of ['bicep', 'terraform'] satisfies StarterTemplateFormat[]) {
    const expected = buildStarterTemplate(changedDocument.nodes, format, changedDocument.edges).content;
    expect(expected).not.toBe(buildStarterTemplate(initialDocument.nodes, format, initialDocument.edges).content);
    expect(await download(format)).toBe(expected);
  }
});
