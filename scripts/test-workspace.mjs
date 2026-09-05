import assert from 'node:assert/strict';
import { createServer } from 'node:http';
import { mkdtemp, readFile, mkdir, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { chromium, expect } from '@playwright/test';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const work = await mkdtemp(path.join(root, '.workspace-regression-'));
const browserState = path.join(work, 'node_modules', 'browser-state');
await mkdir(browserState, { recursive: true });
process.env.TEMP = browserState;
process.env.TMP = browserState;
const output = path.join(work, 'build');
const artifacts = process.env.WORKSPACE_ARTIFACT_DIR && path.resolve(process.env.WORKSPACE_ARTIFACT_DIR);
let server;
let browser;
let workspacePage;
let getBrowserDiagnostics = () => ({});
try {
  await build({
    configFile: path.join(root, 'vite.config.ts'), root, publicDir: false, logLevel: 'error',
    define: {
      'import.meta.env.VITE_AZURE_OPENAI_ENDPOINT': JSON.stringify('https://workspace-test.openai.azure.com'),
      'import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA': JSON.stringify('workspace-test-astra'),
      'import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL': JSON.stringify('workspace-test-model'),
      'import.meta.env.VITE_AZURE_OPENAI_DEPLOYMENT_GPT52': JSON.stringify('workspace-test-model'),
    },
    build: {
      outDir: output, emptyOutDir: true,
      rollupOptions: {
        preserveEntrySignatures: 'strict',
        input: { app: path.join(root, 'index.html'), 'workspace-harness': path.join(root, 'tests', 'fixtures', 'workspace-harness.ts') },
        output: { entryFileNames: chunk => chunk.name === 'workspace-harness' ? 'workspace-harness.js' : 'assets/[name]-[hash].js' },
      },
    },
  });
  const mimeTypes = { '.js': 'text/javascript', '.css': 'text/css', '.html': 'text/html', '.svg': 'image/svg+xml', '.png': 'image/png', '.json': 'application/json' };
  server = createServer(async (request, response) => {
    try {
      const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
      if (pathname === '/storage-fixtures') {
        response.writeHead(200, { 'Content-Type': 'text/html' }).end('<!doctype html><html><head><title>Workspace fixtures</title></head><body></body></html>');
        return;
      }
      const file = path.resolve(output, '.' + (pathname === '/' ? '/index.html' : pathname));
      if (!file.startsWith(output + path.sep) || !(await stat(file)).isFile()) {
        response.writeHead(404).end();
        return;
      }
      response.writeHead(200, { 'Content-Type': mimeTypes[path.extname(file)] ?? 'application/octet-stream' }).end(await readFile(file));
    } catch (error) {
      if (error.code === 'ENOENT') response.writeHead(404).end();
      else { console.error(error); response.writeHead(500).end(); }
    }
  });
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const baseUrl = `http://127.0.0.1:${address.port}`;
  browser = await chromium.launch({
    headless: true,
    ...(process.env.WORKSPACE_BROWSER_CHANNEL ? { channel: process.env.WORKSPACE_BROWSER_CHANNEL } : {}),
  });
  const context = await browser.newContext({ locale: 'en-US', viewport: { width: 1280, height: 720 }, reducedMotion: 'reduce' });
  let validationRequests = 0;
  const validationDeployments = [];
  let validationNodeId;
  let invalidValidation = false;
  let budgetUnavailable = false;
  const feedbackSubmissions = [];
  const feedbackDeletions = [];
  const unexpectedAIRequests = [];
  const startupAttempts = [];
  const mockRequest = route => {
    const url = new URL(route.request().url());
    if (url.pathname === '/api/access/me') return route.fulfill({
      json: { enabled: false, authenticated: false, allowed: true, email: null, isAdmin: false },
    });
    if (url.pathname === '/api/ai/budget') return budgetUnavailable
      ? route.fulfill({ status: 503, json: { error: 'Budget storage unavailable' } })
      : route.fulfill({ json: {
        available: true, limitTokens: 1000, usedTokens: 300, reservedTokens: 100,
        remainingTokens: 700, concurrentRequests: 1, concurrentLimit: 2,
        resetAt: '2030-01-01T00:00:00Z', mode: 'public',
      } });
    if (url.pathname === '/api/feedback/policy') return route.fulfill({
      json: { archiveEnabled: true, retentionDays: 30, emailEnabled: false },
    });
    if (url.pathname === '/api/feedback' && route.request().method() === 'POST') {
      feedbackSubmissions.push(route.request().postDataJSON());
      return route.fulfill({ json: {
        id: '00000000-0000-4000-8000-000000000001', archiveSaved: true,
        expiresAt: '2030-01-01T00:00:00Z', emailDelivered: false, canDelete: true,
      } });
    }
    if (url.pathname.startsWith('/api/feedback/') && route.request().method() === 'DELETE') {
      feedbackDeletions.push(url.pathname.split('/').pop());
      return route.fulfill({ status: 204, body: '' });
    }
    if (url.pathname === '/api/openai') {
      const body = route.request().postData() ?? '';
      if (!body.includes('Well-Architected Framework')) {
        unexpectedAIRequests.push(body.slice(0, 150));
        return route.fulfill({ status: 503, json: { error: 'Only explicit mocked validation is allowed' } });
      }
      validationRequests += 1;
      const proxyRequest = route.request().postDataJSON();
      validationDeployments.push({ deployment: proxyRequest.deployment, model: proxyRequest.body?.model });
      const finding = {
        id: 'workspace-auth-rule', severity: invalidValidation ? 'unsupported' : 'high', category: 'Identity',
        issue: 'Authentication is not configured',
        recommendation: 'Configure authentication for this service.',
        resources: ['Customer Portal'], resourceIds: [validationNodeId], source: 'ai-analysis',
      };
      return route.fulfill({ json: {
        output_text: JSON.stringify({
          overallScore: 80, summary: 'Synthetic workspace validation result.',
          pillars: ['Reliability', 'Security', 'Cost Optimization', 'Operational Excellence', 'Performance Efficiency']
            .map(pillar => ({
              // The upstream adapter normalizes severity; exceed a preserved field's bound to exercise the editor's atomic guard.
              pillar: invalidValidation && pillar === 'Reliability' ? 'Reliability'.repeat(120) : pillar,
              score: 80,
              findings: pillar === 'Security' && (validationRequests === 1 || invalidValidation) ? [finding] : [],
            })),
          quickWins: [],
        }),
        usage: { input_tokens: 10, output_tokens: 20, total_tokens: 30 },
      } });
    }
    if (url.protocol.startsWith('http') && !['localhost', '127.0.0.1'].includes(url.hostname)) return route.abort();
    return route.continue();
  };
  await context.route('**/*', mockRequest);
  const page = await context.newPage();
  workspacePage = page;
  page.setDefaultTimeout(15000);
  page.setDefaultNavigationTimeout(30000);
  const errors = [];
  const consoleErrors = [];
  let expectedReload = false;
  let protectedReloads = 0;
  getBrowserDiagnostics = () => ({ validationRequests, errors, consoleErrors: consoleErrors.slice(-12), unexpectedAIRequests, startupAttempts });
  page.on('console', message => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', error => errors.push(error.message));
  page.on('dialog', async dialog => {
    if (expectedReload && dialog.type() === 'beforeunload') {
      expectedReload = false;
      protectedReloads += 1;
      await dialog.accept();
      return;
    }
    errors.push(`Unexpected blocking dialog: ${dialog.message()}`);
    await dialog.dismiss();
  });
  const openRibbon = async task => {
    const tab = page.locator(`#ribbon-tab-${task === 'export' ? 'home' : task}`);
    if (!await tab.isVisible()) await page.locator('.mobile-command-bar [aria-controls="application-toolbar"]').click();
    await tab.click();
  };
  const undo = async () => {
    await openRibbon('design');
    await page.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).click();
    await openRibbon('home');
  };
  const redo = async () => {
    await openRibbon('design');
    await page.getByRole('button', { name: 'Redo (Ctrl+Y)', exact: true }).click();
    await openRibbon('home');
  };
  const readSavedDocument = () => page.evaluate(async () => {
    const db = await new Promise((resolve, reject) => {
      const request = indexedDB.open('AzureDiagramDrafts');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction('drafts', 'readonly');
        const request = tx.objectStore('drafts').get('local-workspace');
        tx.oncomplete = () => resolve(request.result.document);
        tx.onabort = () => reject(tx.error);
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  });
  await page.goto(`${baseUrl}/storage-fixtures`);
  const storage = await page.evaluate(async () => (await import('/workspace-harness.js')).runWorkspaceStorageFixtures());
  for (const timing of [
    { name: 'cloud-before-draft', identityDelay: 1000, cloudDelay: 0 },
    { name: 'draft-before-cloud', identityDelay: 0, cloudDelay: 1000 },
  ]) {
    const startupContext = await browser.newContext({ locale: 'en-US', viewport: { width: 1280, height: 720 }, reducedMotion: 'reduce' });
    await startupContext.addInitScript(() => {
      sessionStorage.setItem('azurediagarm.cloud-document.v1', JSON.stringify({
        documentId: 'startup-source', access: 'owner', role: 'owner',
      }));
    });
    const attempt = { name: timing.name, reads: 0, heldWrite: false };
    startupAttempts.push(attempt);
    const timestamp = new Date().toISOString();
    let startupDocument = {
      id: 'startup-source', diagramName: 'Seeded startup', createdAt: timestamp, updatedAt: timestamp,
      revision: 1, etag: '"startup-1"', access: 'owner', role: 'owner',
      owner: { id: 'fixture', email: 'startup@example.invalid' }, comments: [], shares: [],
      payload: {
        nodes: [
          { id: 'startup-a', type: 'azureNode', position: { x: 100, y: 200 }, data: { label: 'Seeded service', serviceName: 'App Service' } },
          { id: 'startup-b', type: 'azureNode', position: { x: 450, y: 200 }, data: { label: 'Seeded database', serviceName: 'SQL Database' } },
        ],
        edges: [{ id: 'startup-edge', source: 'startup-a', target: 'startup-b', type: 'editableEdge' }],
        titleBlockData: { architectureName: 'Seeded startup', author: '', date: '2026-09-05', version: '1.0' },
        viewport: { x: 35, y: 50, zoom: 0.75 },
      },
    };
    let holdEdits = false;
    let releaseEdits;
    const editGate = new Promise(resolve => { releaseEdits = resolve; });
    const writes = [];
    await startupContext.route('**/*', async route => {
      const request = route.request();
      const pathname = new URL(request.url()).pathname;
      if (pathname === '/api/access/me') {
        await new Promise(resolve => setTimeout(resolve, timing.identityDelay));
        return mockRequest(route);
      }
      if (pathname === '/api/diagrams/startup-source' && request.method() === 'GET') {
        attempt.reads += 1;
        await new Promise(resolve => setTimeout(resolve, timing.cloudDelay));
        return route.fulfill({ headers: { ETag: startupDocument.etag }, json: { document: startupDocument } });
      }
      if (pathname === '/api/diagrams/startup-source' && request.method() === 'PUT') {
        const body = request.postDataJSON();
        const operation = (async () => {
          if (holdEdits && body.payload.nodes.some(node => node.data.label === 'Unsaved startup edit')) {
            attempt.heldWrite = true;
            await editGate;
          }
          if (request.headers()['if-match'] !== startupDocument.etag) {
            return route.fulfill({ status: 412, json: { error: 'The startup fixture changed concurrently' } });
          }
          const revision = startupDocument.revision + 1;
          startupDocument = { ...startupDocument, ...body, revision, etag: `"startup-${revision}"` };
          await route.fulfill({ headers: { ETag: startupDocument.etag }, json: { document: startupDocument } });
        })();
        writes.push(operation);
        return operation;
      }
      if (pathname.startsWith('/api/diagrams')) return route.fulfill({ status: 503, json: { error: 'Only the seeded startup document is available' } });
      return mockRequest(route);
    });
    const startupPage = await startupContext.newPage();
    workspacePage = startupPage;
    startupPage.setDefaultTimeout(15000);
    startupPage.on('pageerror', error => errors.push(`Startup ${timing.name}: ${error.message}`));
    try {
      await startupPage.goto(baseUrl, { waitUntil: 'networkidle' });
      await expect(startupPage.locator('.azure-node')).toHaveCount(2);
      await expect(startupPage.locator('.react-flow__viewport')).toHaveCSS('transform', 'matrix(0.75, 0, 0, 0.75, 35, 50)');
      await expect(startupPage.getByRole('button', { name: 'Feedback', exact: true })).toBeVisible();
      await expect(startupPage.getByText('Saved on this device', { exact: true })).toBeVisible();
      const savedIds = await startupPage.evaluate(async () => {
        const database = await new Promise((resolve, reject) => {
          const request = indexedDB.open('AzureDiagramDrafts');
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        try {
          return await new Promise((resolve, reject) => {
            const transaction = database.transaction('drafts', 'readonly');
            const request = transaction.objectStore('drafts').get('local-workspace');
            transaction.oncomplete = () => resolve(request.result?.document.nodes.map(node => node.id));
            transaction.onabort = () => reject(transaction.error);
          });
        } finally { database.close(); }
      });
      assert.deepEqual(savedIds, ['startup-a', 'startup-b'], 'Both startup orderings must preserve the seeded cloud nodes in the local draft.');
      const readsBeforeLocaleChange = attempt.reads;
      holdEdits = true;
      await startupPage.locator('.node-label').first().dblclick();
      await startupPage.locator('.node-label-input').fill('Unsaved startup edit');
      await startupPage.locator('.node-label-input').press('Enter');
      await expect.poll(() => attempt.heldWrite).toBe(true);
      await startupPage.locator('.header-utility-menu > button').click();
      await startupPage.locator('.language-switch').getByRole('button', { name: '\u65e5\u672c\u8a9e', exact: true }).click();
      await startupPage.keyboard.press('Escape');
      await startupPage.waitForTimeout(500);
      await expect(startupPage.locator('.node-label').first()).toHaveText('Unsaved startup edit');
      assert.equal(attempt.reads, readsBeforeLocaleChange, 'Changing UI locale must not rehydrate a cloud document over pending edits.');
      assert.equal(readsBeforeLocaleChange, 1, 'React Flow initialization must not restart cloud startup loading.');
    } finally {
      holdEdits = false;
      releaseEdits();
      await Promise.allSettled(writes);
      await startupContext.close();
    }
  }
  workspacePage = page;
  await page.goto(baseUrl, { waitUntil: 'networkidle' });
  await expect(page.getByText('Local autosave ready', { exact: true })).toBeVisible();
  await openRibbon('home');
  await page.locator('.region-selector-button').click();
  await page.locator('.region-option').filter({ hasText: 'East US 2' }).click();
  await expect(page.locator('.region-name')).toHaveText('East US 2');
  await openRibbon('create');
  await expect(page.locator('.model-popover-label')).toHaveText('GPT-6 Astra');
  const headerHeight = await page.locator('.app-header').evaluate(element => element.getBoundingClientRect().height);
  assert.ok(headerHeight <= 220, `Create toolbar must leave room for the canvas at 1280x720, got ${headerHeight}px`);
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth > innerWidth), false);
  await page.locator('.icon-item-main').first().click();
  await expect(page.locator('.azure-node')).toHaveCount(1);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const originalLabel = await page.locator('.node-label').first().textContent();
  const originalNodeId = await page.locator('.react-flow__node').first().getAttribute('data-id');
  const originalServiceType = (await readSavedDocument()).nodes[0].data.serviceName || originalLabel;
  validationNodeId = originalNodeId;
  await page.locator('.node-label').first().dblclick();
  await page.locator('.node-label-input').fill('Customer Portal');
  await page.locator('.node-label-input').press('Enter');
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await undo();
  await expect(page.locator('.node-label').first()).toHaveText(originalLabel);
  await redo();
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await page.locator('.react-flow__node').first().click();
  await page.keyboard.press('Delete');
  await expect(page.locator('.azure-node')).toHaveCount(0);
  await page.keyboard.press('Control+z');
  await expect(page.locator('.azure-node')).toHaveCount(1);
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  expectedReload = true;
  try { await page.reload({ waitUntil: 'networkidle' }); } finally { expectedReload = false; }
  await expect(page.getByRole('dialog', { name: 'Resume your saved draft' })).toBeVisible();
  await page.getByRole('button', { name: 'Restore draft', exact: true }).click();
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await openRibbon('create');
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  const modal = page.getByRole('dialog', { name: 'AI Architecture Generator' });
  await expect(modal).toBeVisible();
  await expect.poll(() => modal.evaluate(element => element.contains(document.activeElement))).toBe(true);
  const controls = modal.locator('button:visible:not([disabled]), input:visible:not([disabled]), textarea:visible:not([disabled]), select:visible:not([disabled]), a[href]:visible, [tabindex="0"]:visible');
  await controls.last().focus();
  await page.keyboard.press('Tab');
  assert.equal(await modal.evaluate(element => element.contains(document.activeElement)), true, 'Tab remains in the modal');
  await page.keyboard.press('Escape');
  await expect(modal).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Generate Diagram', exact: true })).toBeFocused();
  await openRibbon('export');
  const picker = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Load', exact: true }).focus();
  await page.keyboard.press('Enter');
  await (await picker).setFiles([]);

  const proposal = {
    format: 'azurediagarm-ai-architecture',
    services: [
      { id: 'proposed-web', name: originalServiceType, type: originalServiceType, category: 'other' },
      { id: 'proposed-db', name: 'SQL Database', type: 'SQL Database', category: 'databases' },
    ],
    connections: [{ from: 'proposed-web', to: 'proposed-db', label: 'SQL queries', type: 'sync' }],
    groups: [],
    workflow: [{ step: 1, description: 'Query the application database', services: ['proposed-web', 'proposed-db'] }],
  };
  const loadProposal = () => page.locator('input[type="file"][aria-label="Load diagram"]').setInputFiles({
    name: 'ai-proposal.json', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(proposal)),
  });
  await loadProposal();
  const review = page.getByRole('dialog', { name: 'Review AI changes', exact: true });
  await expect(review).toBeVisible({ timeout: 60000 });
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await review.getByRole('button', { name: 'Keep current diagram', exact: true }).click();
  await expect(review).toHaveCount(0);
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');

  await loadProposal();
  await expect(review).toBeVisible({ timeout: 60000 });
  await page.evaluate(() => {
    globalThis.__workspaceOriginalPut = IDBObjectStore.prototype.put;
    IDBObjectStore.prototype.put = function (...args) {
      const request = globalThis.__workspaceOriginalPut.apply(this, args);
      if (this.transaction.db.name === 'AzureDiagramVersions') {
        request.addEventListener('success', () => this.transaction.abort());
      }
      return request;
    };
  });
  await review.getByRole('button', { name: 'Apply selected changes', exact: true }).click();
  await expect(review.getByRole('alert')).toBeVisible();
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await page.evaluate(() => {
    IDBObjectStore.prototype.put = globalThis.__workspaceOriginalPut;
    delete globalThis.__workspaceOriginalPut;
  });
  await review.getByRole('button', { name: 'Apply selected changes', exact: true }).click();
  await expect(review).toHaveCount(0);
  await expect(page.locator('.azure-node')).toHaveCount(2);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const acceptedDraft = await readSavedDocument();
  assert.equal(acceptedDraft.workflow.length, 1);
  assert.deepEqual(acceptedDraft.workflow[0].services, acceptedDraft.nodes.map(node => node.id));
  assert.ok(acceptedDraft.workflow[0].services.includes(originalNodeId), 'Workflow must reference normalized, retained node IDs.');
  const edgeLabel = page.locator('.editable-edge-label').first();
  await edgeLabel.hover();
  const edgeBounds = await edgeLabel.boundingBox();
  assert.ok(edgeBounds);
  const edgeCenter = { x: edgeBounds.x + edgeBounds.width / 2, y: edgeBounds.y + edgeBounds.height / 2 };
  await page.mouse.move(edgeCenter.x, edgeCenter.y);
  await page.mouse.down();
  await page.mouse.move(edgeCenter.x + 36, edgeCenter.y + 24, { steps: 8 });
  await page.mouse.up();
  await expect.poll(async () => (await readSavedDocument()).edges[0]?.data?.labelOffsetAuto).toBe(false);
  assert.notEqual((await readSavedDocument()).edges[0].data.labelOffsetX, acceptedDraft.edges[0].data.labelOffsetX);
  await undo();
  await expect(page.locator('.azure-node')).toHaveCount(2);
  await expect.poll(async () => (await readSavedDocument()).edges[0]?.data?.labelOffsetX).toBe(acceptedDraft.edges[0].data.labelOffsetX);
  await expect.poll(async () => (await readSavedDocument()).edges[0]?.data?.labelOffsetY).toBe(acceptedDraft.edges[0].data.labelOffsetY);
  await undo();
  await expect(page.locator('.azure-node')).toHaveCount(1);
  await expect(page.locator('.node-label').first()).toHaveText('Customer Portal');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();

  await loadProposal();
  await expect(review).toBeVisible({ timeout: 60000 });
  await review.getByRole('button', { name: 'Clear selection', exact: true }).click();
  await review.getByRole('group', { name: 'Services and groups', exact: true })
    .getByRole('checkbox', { name: /^Add SQL Database(?:\s|$)/ }).check();
  await review.getByRole('button', { name: 'Apply selected changes', exact: true }).click();
  await expect(review).toHaveCount(0);
  await expect(page.locator('.azure-node')).toHaveCount(2);
  await expect(page.locator(`.react-flow__node[data-id="${originalNodeId}"] .node-label`)).toHaveText('Customer Portal');
  await expect(page.locator('.react-flow__edge')).toHaveCount(0);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  assert.deepEqual((await readSavedDocument()).workflow, [], 'A partial proposal must not retain workflow for rejected connections.');
  await undo();
  await expect(page.locator('.azure-node')).toHaveCount(1);
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();

  await openRibbon('create');
  await page.locator(`.react-flow__node[data-id="${originalNodeId}"]`).click();
  const beforeInspector = await readSavedDocument();
  await page.getByRole('button', { name: 'Service settings', exact: true }).click();
  const inspector = page.getByRole('dialog', { name: 'Service inspector', exact: true });
  await expect(inspector).toBeVisible();
  await inspector.getByRole('spinbutton', { name: 'Quantity (instances / units)', exact: true }).fill('3');
  await inspector.getByRole('checkbox', { name: 'Use a custom monthly estimate', exact: true }).check();
  await inspector.getByRole('spinbutton', { name: 'USD per unit / month', exact: true }).fill('120');
  if (artifacts) {
    await mkdir(artifacts, { recursive: true });
    await page.screenshot({ path: path.join(artifacts, 'service-inspector.png'), fullPage: true });
  }
  await inspector.getByRole('button', { name: 'Apply', exact: true }).click();
  await expect(inspector).toHaveCount(0);
  await expect(page.locator('.azure-node .cost-badge')).toContainText('$360');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const priced = await readSavedDocument();
  assert.equal(priced.nodes[0].data.pricing.estimatedCost, 120);
  assert.equal(priced.nodes[0].data.pricing.quantity, 3);
  await undo();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  assert.deepEqual((await readSavedDocument()).nodes[0].data.pricing, beforeInspector.nodes[0].data.pricing);
  await redo();
  await expect(page.locator('.azure-node .cost-badge')).toContainText('$360');

  await page.getByRole('button', { name: 'Service settings', exact: true }).click();
  await inspector.getByRole('spinbutton', { name: 'USD per unit / month', exact: true }).fill('999');
  await page.keyboard.press('Escape');
  await expect(inspector).toHaveCount(0);
  await expect(page.locator('.azure-node .cost-badge')).toContainText('$360');

  await openRibbon('home');
  await page.locator('.region-selector-button').click();
  await page.locator('.region-option').filter({ hasText: 'Japan East' }).click();
  await expect(page.locator('.region-name')).toHaveText('Japan East');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  const regional = await readSavedDocument();
  assert.equal(regional.settings.pricingRegion, 'japaneast');
  assert.equal(regional.nodes[0].data.pricing.region, 'japaneast');
  assert.equal(regional.nodes[0].data.pricing.estimatedCost, 120);
  assert.equal(regional.nodes[0].data.pricing.quantity, 3);
  await undo();
  await expect(page.locator('.region-name')).toHaveText('East US 2');
  await expect(page.locator('.azure-node .cost-badge')).toContainText('$360');
  await redo();
  await expect(page.locator('.region-name')).toHaveText('Japan East');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();

  await expect(page.locator('.ai-budget-status')).toContainText('700 tokens left today');
  budgetUnavailable = true;
  await page.locator('.ai-budget-status').click();
  await expect(page.locator('.ai-budget-status')).toContainText('AI budget unavailable');
  budgetUnavailable = false;
  await page.locator('.ai-budget-status').click();
  await expect(page.locator('.ai-budget-status')).toContainText('700 tokens left today');

  await openRibbon('review');
  await page.getByRole('button', { name: 'Validate Architecture', exact: true }).click();
  const validation = page.getByRole('dialog', { name: /Architecture Validation/ });
  await expect(validation.locator('.finding-issue')).toContainText('Authentication is not configured', { timeout: 30000 });
  await expect(validation.locator('.source-badge')).toHaveText('AI');
  await expect.poll(async () => (await readSavedDocument()).reviewHistory?.[0]?.status).toBe('active');
  assert.equal((await readSavedDocument()).reviewHistory[0].finding.source, 'ai');
  if (artifacts) await page.screenshot({ path: path.join(artifacts, 'validation-review.png'), fullPage: true });
  await validation.getByRole('button', { name: 'Show on diagram', exact: true }).click();
  await expect(validation).toHaveCount(0);
  await expect(page.locator(`.react-flow__node[data-id="${originalNodeId}"]`)).toHaveClass(/selected/);
  await page.locator('.node-label').first().dblclick();
  await page.locator('.node-label-input').fill('Customer Portal v2');
  await page.locator('.node-label-input').press('Enter');
  await page.getByRole('button', { name: 'Revalidate Needed', exact: true }).click();
  await expect(validation.getByText('This review is out of date.', { exact: true })).toBeVisible();
  await validation.getByRole('button', { name: 'Re-evaluate now', exact: true }).click();
  const historySummary = validation.getByText('Not detected in latest review (1)', { exact: true });
  await expect(historySummary).toBeVisible({ timeout: 60000 });
  await historySummary.click();
  await expect(validation.getByText('Authentication is not configured', { exact: true })).toBeVisible();
  await expect(validation.getByText('This review is out of date.', { exact: true })).toHaveCount(0);
  await expect.poll(async () => (await readSavedDocument()).reviewHistory?.[0]?.status).toBe('not-detected');
  await validation.getByRole('button', { name: 'Close', exact: true }).first().click();
  assert.equal(validationRequests, 2);

  const historyBeforeFailure = (await readSavedDocument()).reviewHistory;
  invalidValidation = true;
  await page.getByRole('button', { name: 'Validate Architecture', exact: true }).click();
  await expect(page.locator('.workspace-notice')).toContainText('Failed to validate architecture:', { timeout: 30000 });
  await expect(validation).toHaveCount(0);
  assert.deepEqual((await readSavedDocument()).reviewHistory, historyBeforeFailure);
  await page.locator('.workspace-notice').getByRole('button', { name: 'Close', exact: true }).click();
  await page.getByRole('button', { name: 'Compare Validation', exact: true }).click();
  const comparison = page.getByRole('dialog', { name: 'Compare Validation', exact: true });
  const modelChips = comparison.locator('.compare-model-chip');
  await expect(modelChips.first()).toBeVisible();
  for (let index = 0; index < await modelChips.count(); index += 1) {
    const chip = modelChips.nth(index);
    const name = await chip.locator('.compare-model-chip-name').textContent();
    const wanted = ['GPT-6 Astra', 'GPT-5.2'].includes(name?.trim() ?? '');
    const selected = (await chip.getAttribute('class')).split(' ').includes('selected');
    if (wanted !== selected) await chip.click();
  }
  await expect(comparison.locator('.compare-model-chip.selected')).toHaveCount(2);
  await comparison.locator('.compare-run-btn').click();
  await expect(comparison.getByRole('button', { name: 'Use This Validation', exact: true })).toHaveCount(2, { timeout: 30000 });
  await comparison.getByRole('button', { name: 'Use This Validation', exact: true }).first().click();
  await expect(comparison.getByRole('alert')).toContainText('This review could not be applied.');
  await expect(validation).toHaveCount(0);
  assert.deepEqual((await readSavedDocument()).reviewHistory, historyBeforeFailure);
  await comparison.getByRole('button', { name: 'Close', exact: true }).first().click();
  await page.locator('.node-label').first().dblclick();
  await page.locator('.node-label-input').fill('Customer Portal v3');
  await page.locator('.node-label-input').press('Enter');
  await page.getByRole('button', { name: 'Compare Validation', exact: true }).click();
  await comparison.getByRole('button', { name: 'Use This Validation', exact: true }).first().click();
  await expect(comparison.getByRole('alert')).toContainText('This review could not be applied.');
  await expect(validation).toHaveCount(0);
  assert.deepEqual((await readSavedDocument()).reviewHistory, historyBeforeFailure);
  await comparison.getByRole('button', { name: 'Close', exact: true }).first().click();
  await page.locator('.workspace-notice').getByRole('button', { name: 'Close', exact: true }).click();

  invalidValidation = false;
  await page.getByRole('button', { name: 'Compare Validation', exact: true }).click();
  await comparison.getByRole('button', { name: 'New Comparison', exact: true }).click();
  await comparison.locator('.compare-run-btn').click();
  await expect(comparison.getByRole('button', { name: 'Use This Validation', exact: true })).toHaveCount(2, { timeout: 30000 });
  await comparison.getByRole('button', { name: 'Close', exact: true }).first().click();
  await page.locator('.node-label').first().dblclick();
  await page.locator('.node-label-input').fill('Customer Portal v4');
  await page.locator('.node-label-input').press('Enter');
  await page.getByRole('button', { name: 'Compare Validation', exact: true }).click();
  await comparison.getByRole('button', { name: 'Use This Validation', exact: true }).first().click();
  await expect(validation.getByText('This review is out of date.', { exact: true })).toBeVisible();
  assert.deepEqual((await readSavedDocument()).reviewHistory, historyBeforeFailure, 'Stale comparison results must not rewrite review history.');
  await validation.getByRole('button', { name: 'Close', exact: true }).first().click();
  assert.equal(validationRequests, 7);
  assert.deepEqual(validationDeployments.slice(0, 3), Array.from({ length: 3 }, () => ({
    deployment: 'workspace-test-astra', model: 'workspace-test-astra',
  })), 'Default validation must route to the configured Astra deployment, not just display its label.');
  for (const comparisonRequests of [validationDeployments.slice(3, 5), validationDeployments.slice(5, 7)]) {
    assert.deepEqual(comparisonRequests.map(request => request.deployment).sort(), ['workspace-test-astra', 'workspace-test-model']);
    assert.ok(comparisonRequests.every(request => request.model === request.deployment));
  }

  const quickFeedback = page.getByRole('dialog', { name: 'Quick feedback', exact: true });
  await quickFeedback.getByRole('radio', { name: 'Happy', exact: true }).click();
  await page.getByRole('button', { name: 'Feedback', exact: true }).click();
  const feedback = page.getByRole('dialog', { name: 'Share Feedback', exact: true });
  await expect(quickFeedback).toHaveCount(0);
  await feedback.getByRole('radio', { name: 'Happy', exact: true }).click();
  await feedback.locator('#feedback-comment').fill('Workspace interaction feedback');
  const diagnostics = feedback.getByRole('checkbox', { name: /Include optional diagnostics/ });
  await expect(diagnostics).not.toBeChecked();
  await feedback.getByText('Preview exactly what will be submitted', { exact: true }).click();
  assert.deepEqual(JSON.parse(await feedback.locator('.feedback-payload pre').textContent()).context, {});
  await page.evaluate(() => history.replaceState(null, '', '/?trace=synthetic-context#synthetic-fragment'));
  await diagnostics.check();
  const optInPreview = JSON.parse(await feedback.locator('.feedback-payload pre').textContent());
  assert.equal(optInPreview.context.url, baseUrl);
  assert.equal(optInPreview.context.serviceCount, 1);
  assert.ok(Object.keys(optInPreview.context).every(key => ['url', 'serviceCount', 'model'].includes(key)));
  await diagnostics.uncheck();
  const submittedPreview = JSON.parse(await feedback.locator('.feedback-payload pre').textContent());
  if (artifacts) await page.screenshot({ path: path.join(artifacts, 'feedback-privacy.png'), fullPage: true });
  await feedback.getByRole('button', { name: /Send feedback/i }).click();
  await expect(feedback.getByText('Thank you!', { exact: true })).toBeVisible();
  assert.deepEqual(feedbackSubmissions, [submittedPreview]);
  await feedback.getByRole('button', { name: 'Delete my archived feedback', exact: true }).click();
  await expect(feedback.getByText('Deleted from the application archive. Email copies are not deleted.', { exact: true })).toBeVisible();
  assert.deepEqual(feedbackDeletions, ['00000000-0000-4000-8000-000000000001']);
  await feedback.getByRole('button', { name: 'Done', exact: true }).click();
  await expect(quickFeedback).toHaveCount(0);
  await page.getByRole('button', { name: 'Feedback', exact: true }).click();
  await expect(diagnostics).not.toBeChecked();
  await feedback.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.evaluate(() => history.replaceState(null, '', '/'));

  await openRibbon('create');
  await page.getByRole('button', { name: 'Chat', exact: true }).click();
  await expect(page.locator('.arch-chat-panel')).toBeVisible();
  if (artifacts) {
    await mkdir(artifacts, { recursive: true });
    await page.screenshot({ path: path.join(artifacts, 'workspace-light.png'), fullPage: true });
  }
  await page.getByRole('button', { name: 'Close chat', exact: true }).click();
  const responsiveHeaderHeights = [];
  const responsiveLayoutFailures = [];
  for (const language of ['ja', 'en']) {
    await page.locator('.header-utility-menu > button').click();
    await page.locator('.language-switch').getByRole('button', { name: language === 'ja' ? '\u65e5\u672c\u8a9e' : /^EN\b/, exact: true }).click();
    await page.keyboard.press('Escape');
    for (const width of [1366, 1280, 1100, 1024, 900, 768]) {
      await page.setViewportSize({ width, height: 720 });
      for (const section of ['create', 'review', 'export']) {
        await openRibbon(section);
        const height = await page.locator('.app-header').evaluate(element => element.getBoundingClientRect().height);
        responsiveHeaderHeights.push(height);
        if (height > 220) responsiveLayoutFailures.push(`Populated ${language} ${section} toolbar at ${width}px exceeds 220px: ${height}px`);
        if (await page.evaluate(() => document.documentElement.scrollWidth > innerWidth)) {
          responsiveLayoutFailures.push(`${language} ${section} at ${width}px overflows horizontally`);
        }
      }
    }
  }
  assert.deepEqual(responsiveLayoutFailures, [], 'All 36 populated responsive headers must preserve canvas space and avoid horizontal overflow.');
  await page.setViewportSize({ width: 1280, height: 720 });
  await openRibbon('home');
  await page.getByRole('button', { name: 'Switch to Dark Mode', exact: true }).click();
  await expect(page.locator('body')).toHaveClass(/dark-mode/);
  assert.notEqual(await page.locator('.azure-node').evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(255, 255, 255)');
  if (artifacts) await page.screenshot({ path: path.join(artifacts, 'workspace-dark.png'), fullPage: true });

  await expect.poll(async () => (await readSavedDocument()).nodes[0].data.label, { timeout: 15000 }).toBe('Customer Portal v4');
  const sibling = await context.newPage();
  sibling.setDefaultTimeout(15000);
  await sibling.goto(baseUrl, { waitUntil: 'networkidle' });
  await sibling.getByRole('button', { name: 'Restore draft', exact: true }).click();
  await expect(sibling.locator('.node-label').first()).toHaveText('Customer Portal v4');
  await page.bringToFront();
  await page.locator('.node-label').first().dblclick();
  await page.locator('.node-label-input').fill('Saved by the first tab');
  await page.locator('.node-label-input').press('Enter');
  await expect.poll(async () => (await readSavedDocument()).nodes[0].data.label, { timeout: 15000 }).toBe('Saved by the first tab');
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await sibling.bringToFront();
  await sibling.locator('.node-label').first().dblclick();
  await sibling.locator('.node-label-input').fill('Conflicting second tab');
  await sibling.locator('.node-label-input').press('Enter');
  await expect(sibling.getByText('Draft not saved', { exact: true })).toBeVisible();
  const persistedLabel = (await readSavedDocument()).nodes[0].data.label;
  assert.equal(persistedLabel, 'Saved by the first tab', 'A conflicting tab must not overwrite the committed draft');
  await sibling.close({ runBeforeUnload: false });

  const cloudContext = await browser.newContext({ locale: 'en-US', viewport: { width: 1280, height: 720 }, reducedMotion: 'reduce' });
  await cloudContext.route('**/*', mockRequest);
  const cloudPage = await cloudContext.newPage();
  workspacePage = cloudPage;
  cloudPage.setDefaultTimeout(15000);
  cloudPage.on('pageerror', error => errors.push(`Cloud binding: ${error.message}`));
  cloudPage.on('dialog', async dialog => {
    errors.push(`Unexpected cloud-binding dialog: ${dialog.message()}`);
    await dialog.dismiss();
  });
  let cloudCreateSeen = false;
  let releaseCloudCreate;
  const cloudCreateGate = new Promise(resolve => { releaseCloudCreate = resolve; });
  let cloudDocument;
  await cloudPage.route('**/api/diagrams**', async route => {
    const request = route.request();
    const pathname = new URL(request.url()).pathname;
    if (request.method() === 'POST' && pathname === '/api/diagrams') {
      cloudCreateSeen = true;
      const body = request.postDataJSON();
      await cloudCreateGate;
      const now = new Date().toISOString();
      cloudDocument = {
        id: 'workspace-cloud-fixture', ...body, owner: { id: 'fixture', email: 'workspace@example.invalid' },
        createdAt: now, updatedAt: now, revision: 1, comments: [], shares: [], access: 'owner', role: 'owner', etag: '"cloud-1"',
      };
      return route.fulfill({ status: 201, headers: { ETag: cloudDocument.etag }, json: { document: cloudDocument } });
    }
    if (request.method() === 'PUT' && pathname === '/api/diagrams/workspace-cloud-fixture') {
      assert.equal(request.headers()['if-match'], cloudDocument.etag, 'Cloud updates must retain revision protection.');
      const revision = cloudDocument.revision + 1;
      cloudDocument = { ...cloudDocument, ...request.postDataJSON(), revision, etag: `"cloud-${revision}"`, updatedAt: new Date().toISOString() };
      return route.fulfill({ headers: { ETag: cloudDocument.etag }, json: { document: cloudDocument } });
    }
    return route.fulfill({ status: 503, json: { error: 'Only explicit synthetic cloud writes are available' } });
  });
  await cloudPage.goto(baseUrl, { waitUntil: 'networkidle' });
  await cloudPage.bringToFront();
  await expect(cloudPage.getByText('Local autosave ready', { exact: true })).toBeVisible();
  await cloudPage.locator('.icon-item-main').first().click();
  await expect.poll(() => cloudCreateSeen).toBe(true);
  const cloudOriginalLabel = await cloudPage.locator('.node-label').first().textContent();
  await cloudPage.locator('.node-label').first().dblclick();
  await cloudPage.locator('.node-label-input').fill('Edited while cloud creation was pending');
  await cloudPage.locator('.node-label-input').press('Enter');
  releaseCloudCreate();
  await expect.poll(() => cloudDocument?.payload?.nodes?.[0]?.data?.label).toBe('Edited while cloud creation was pending');
  assert.equal(typeof cloudDocument.payload.settings.pricingRegion, 'string');
  assert.deepEqual(cloudDocument.payload.reviewHistory, []);
  await cloudPage.locator('#ribbon-tab-design').click();
  await cloudPage.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).click();
  await expect(cloudPage.locator('.node-label').first()).toHaveText(cloudOriginalLabel);
  await cloudPage.getByRole('button', { name: 'Undo (Ctrl+Z)', exact: true }).click();
  await expect(cloudPage.locator('.azure-node')).toHaveCount(0);
  await cloudPage.getByRole('button', { name: 'Redo (Ctrl+Y)', exact: true }).click();
  await expect(cloudPage.locator('.node-label').first()).toHaveText(cloudOriginalLabel);
  await cloudContext.close();

  const costContext = await browser.newContext({ locale: 'en-US', viewport: { width: 1280, height: 720 }, reducedMotion: 'reduce' });
  await costContext.route('**/*', mockRequest);
  const costPage = await costContext.newPage();
  workspacePage = costPage;
  costPage.setDefaultTimeout(15000);
  costPage.on('pageerror', error => errors.push(`Regional export: ${error.message}`));
  costPage.on('dialog', async dialog => {
    errors.push(`Unexpected regional-export dialog: ${dialog.message()}`);
    await dialog.dismiss();
  });
  await costPage.goto(baseUrl, { waitUntil: 'networkidle' });
  await costPage.bringToFront();
  await expect(costPage.getByText('Local autosave ready', { exact: true })).toBeVisible();
  const regionalPricing = {
    quantity: 1, region: 'japaneast', unit: 'per month', lastUpdated: new Date().toISOString(),
  };
  const knownRegionalNode = {
    id: 'regional-known', type: 'azureNode', position: { x: 0, y: 0 },
    data: { label: 'Known regional service', serviceName: 'App Services', pricing: {
      ...regionalPricing, estimatedCost: 120, customPrice: 120, isCustom: true, tier: 'Custom', skuName: 'Custom',
      provenance: { kind: 'custom', source: 'user', unit: 'USD/unit/month', assumptions: [] },
    } },
  };
  const unpricedRegionalNode = {
    id: 'regional-unpriced', type: 'azureNode', position: { x: 320, y: 0 },
    data: { label: 'Unpriced regional service', serviceName: 'Regional fixture unpriced service', pricing: {
      ...regionalPricing, estimatedCost: null, isCustom: false, tier: 'Unavailable', skuName: 'Unavailable',
      provenance: { kind: 'unpriced', unit: 'USD/unit/month', assumptions: [], note: 'Synthetic service has no available rate.' },
    } },
  };
  const { default: JSZip } = await import('jszip');
  const exportRegionalFixture = async (fixtureNodes, name) => {
    await costPage.locator('input[type="file"][aria-label="Load diagram"]').setInputFiles({
      name: `regional-${name}.json`, mimeType: 'application/json',
      buffer: Buffer.from(JSON.stringify({ nodes: fixtureNodes, edges: [], workflow: [] })),
    });
    await expect(costPage.locator('.react-flow__node[data-id="regional-known"]')).toHaveCount(1);
    await expect(costPage.locator('.azure-node')).toHaveCount(fixtureNodes.length);
    await costPage.locator('#ribbon-tab-home').click();
    await costPage.getByRole('button', { name: 'Export', exact: true }).click();
    const [download] = await Promise.all([
      costPage.waitForEvent('download', { timeout: 60000 }),
      costPage.getByRole('menuitem', { name: 'Export Costs (All Formats)', exact: true }).click(),
    ]);
    const zipPath = path.join(work, `regional-${name}.zip`);
    await download.saveAs(zipPath);
    const archive = await JSZip.loadAsync(await readFile(zipPath));
    const entry = async suffix => {
      const name = Object.keys(archive.files).find(name => name.endsWith(suffix));
      assert.ok(name, `Cost archive must include ${suffix}`);
      return archive.file(name).async('string');
    };
    return {
      analysis: await entry('-analysis.md'),
      regionalCsv: await entry('-multiregion-comparison.csv'),
      breakdown: JSON.parse(await entry('.json')),
    };
  };
  const completeRegionalExport = await exportRegionalFixture([knownRegionalNode], 'complete');
  assert.equal(completeRegionalExport.breakdown.estimateCompleteness, 'complete');
  assert.match(completeRegionalExport.analysis, /\| Rank \|/);
  assert.match(completeRegionalExport.regionalCsv, /,Comparable,/);
  const partialRegionalExport = await exportRegionalFixture([knownRegionalNode, unpricedRegionalNode], 'partial');
  assert.equal(partialRegionalExport.breakdown.estimateCompleteness, 'partial');
  assert.equal(partialRegionalExport.breakdown.totalMonthlyCost, 120);
  assert.match(partialRegionalExport.analysis, /Priced subtotal/);
  assert.match(partialRegionalExport.analysis, /Partial subtotals are not ranked/);
  assert.doesNotMatch(partialRegionalExport.analysis, /\| Rank \||> Cheapest region:|\*\*Cheapest:\*\*|\*\*Potential savings:\*\*/);
  assert.match(partialRegionalExport.regionalCsv, /,Incomplete estimate,/);
  assert.doesNotMatch(partialRegionalExport.regionalCsv, /,Comparable,/);
  assert.match(partialRegionalExport.regionalCsv, /"Known regional service",regional-known,120\.00,/,
    'Individual known rates must remain available even when regional totals cannot be ranked.');
  await costContext.close();
  workspacePage = page;
  assert.deepEqual(errors, [], 'Workspace changes must not produce browser exceptions');
  assert.equal(feedbackSubmissions.length, 1, 'Opening full feedback cancels any pending quick-rating submission.');
  assert.deepEqual(unexpectedAIRequests, [], 'Browser checks must only invoke explicitly mocked AI requests');
  console.log(JSON.stringify({ storage, startupHydrations: startupAttempts, headerHeight, responsiveHeaderChecks: responsiveHeaderHeights.length, maximumResponsiveHeaderHeight: Math.max(...responsiveHeaderHeights), protectedReloads, undo: true, groupedEdgeDrag: true, draftRecovery: true, modalFocus: true, keyboardImport: true, aiReview: true, snapshotFailure: true, serviceInspector: true, regionalPricing: true, wafReview: true, staleComparison: true, invalidReviewRecovery: true, feedbackPrivacy: true, aiBudget: true, concurrentTabs: true, cloudBindingHistory: true, completeRegionalRanking: true, partialRegionalRankingSuppressed: true }, null, 2));
} catch (error) {
  console.error('Workspace browser diagnostics:', getBrowserDiagnostics());
  if (workspacePage && !workspacePage.isClosed() && artifacts) {
    console.error('Visible workspace controls:', await workspacePage.locator('button, [role="dialog"]')
      .evaluateAll(elements => elements.filter(element => element.getClientRects().length > 0)
        .map(element => element.getAttribute('aria-label') || element.textContent?.trim()).slice(0, 70))
      .catch(() => []));
    await mkdir(artifacts, { recursive: true });
    await workspacePage.screenshot({ path: path.join(artifacts, 'failure.png'), fullPage: true })
      .catch(cause => console.error('Could not capture the failed browser state:', cause));
  }
  throw error;
} finally {
  if (browser) await browser.close();
  if (server) await new Promise(resolve => server.close(resolve));
  await rm(work, { recursive: true, force: true });
}
