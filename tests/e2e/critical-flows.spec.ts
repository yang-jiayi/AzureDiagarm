import { expect, test, type Page, type Route } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

const now = '2026-08-02T00:00:00.000Z';
const wait = (milliseconds: number) => new Promise(resolve => setTimeout(resolve, milliseconds));

function summary(id: string, name: string) {
  return {
    id,
    diagramName: name,
    createdAt: now,
    updatedAt: now,
    revision: 1,
    serviceCount: 1,
    connectionCount: 0,
    commentCount: 0,
    shareCount: 0,
    access: 'owner',
    role: 'owner',
    etag: `"${id}-1"`,
  };
}

function cloudDocument(id: string, name: string) {
  return {
    ...summary(id, name),
    payload: {
      nodes: [{
        id: `${id}-node`,
        type: 'azureNode',
        position: { x: 100, y: 100 },
        data: { label: 'App Service', serviceName: 'App Service' },
      }],
      edges: [],
      titleBlockData: { architectureName: name },
    },
    owner: { id: 'owner', email: 'owner@example.com' },
    comments: [],
    shares: [],
  };
}

function cloudVersion(diagramId: string, versionId: string, notes: string) {
  return {
    versionId,
    diagramId,
    diagramName: `Diagram ${diagramId}`,
    payload: cloudDocument(diagramId, `Diagram ${diagramId}`).payload,
    notes,
    createdAt: now,
    createdByEmail: 'owner@example.com',
    sourceRevision: 1,
  };
}

function interactionCloudDocument() {
  return {
    ...cloudDocument('A', 'Interaction diagram'),
    serviceCount: 2,
    connectionCount: 1,
    payload: {
      nodes: [
        {
          id: 'node-a',
          type: 'azureNode',
          position: { x: 160, y: 220 },
          data: { label: 'App Service', serviceName: 'App Service' },
        },
        {
          id: 'node-b',
          type: 'azureNode',
          position: { x: 560, y: 220 },
          data: { label: 'Azure SQL Database', serviceName: 'Azure SQL Database' },
        },
        {
          id: 'group-a',
          type: 'groupNode',
          position: { x: 320, y: 80 },
          style: { width: 180, height: 100 },
          data: { label: 'Application layer', color: '#0078d4' },
        },
      ],
      edges: [{
        id: 'edge-ab',
        source: 'node-a',
        target: 'node-b',
        sourceHandle: 'right',
        targetHandle: 'left',
        type: 'editableEdge',
      }],
      titleBlockData: { architectureName: 'Interaction diagram' },
    },
  };
}

async function fulfillJson(
  route: Route,
  body: unknown,
  status = 200,
  headers?: Record<string, string>,
) {
  await route.fulfill({
    status,
    contentType: 'application/json',
    headers,
    body: JSON.stringify(body),
  });
}

async function initializePage(
  page: Page,
  cloudContext?: Record<string, unknown>,
) {
  await page.addInitScript((context) => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
    localStorage.setItem('azure-diagram-builder.ribbonTab.v1', 'review');
    localStorage.setItem('azure-diagram-builder.headerCollapsed.v1', '0');
    localStorage.setItem('azure-diagram-builder.canvasHintDismissed.v1', '1');
    if (context) {
      sessionStorage.setItem('azurediagarm.cloud-document.v1', JSON.stringify(context));
    } else {
      sessionStorage.removeItem('azurediagarm.cloud-document.v1');
    }
  }, cloudContext);
}

async function expectNoWcagViolations(page: Page, include?: string) {
  let builder = new AxeBuilder({ page })
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']);
  if (include) builder = builder.include(include);
  const results = await builder.analyze();
  expect(
    results.violations,
    results.violations.map(violation => (
      `${violation.id}: ${violation.nodes.map(node => node.target.join(' ')).join(', ')}`
    )).join('\n'),
  ).toEqual([]);
}

async function openInteractionDiagram(
  page: Page,
  onSave?: (payload: Record<string, unknown>) => void,
) {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  const document = interactionCloudDocument();

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: true,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: true,
        allowed: true,
      });
      return;
    }
    if (path === '/api/access/users' && method === 'GET') {
      await fulfillJson(route, { users: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, document, 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      onSave?.(body.payload || {});
      await fulfillJson(route, {
        ...document,
        diagramName: body.diagramName || document.diagramName,
        payload: body.payload || document.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toBeVisible();
}

async function openAiGenerator(page: Page) {
  await page.getByRole('tab', { name: 'Create' }).click();
  await page.getByRole('button', { name: 'Generate with AI' }).click();
  const modal = page.locator('.ai-architecture-modal');
  await expect(modal).toBeFocused();
  return modal;
}

test('primary application shell meets WCAG A and AA checks', async ({ page }) => {
  await initializePage(page);
  await page.route('**/api/**', async (route) => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  await expect(page.getByRole('region', { name: 'Architecture canvas' })).toBeVisible();
  const paletteTabs = page.getByRole('tablist', { name: 'Icon library views' });
  const allTab = paletteTabs.getByRole('tab', { name: /^All/ });
  const favoritesTab = paletteTabs.getByRole('tab', { name: /^Favorites/ });
  const collectionsTab = paletteTabs.getByRole('tab', { name: /^Collections/ });
  await allTab.focus();
  await page.keyboard.press('ArrowRight');
  await expect(favoritesTab).toBeFocused();
  await expect(favoritesTab).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('End');
  await expect(collectionsTab).toBeFocused();
  await expect(page.locator('#palette-view-panel')).toHaveAttribute(
    'aria-labelledby',
    'palette-view-tab-collections',
  );
  await page.keyboard.press('Home');
  await expect(allTab).toBeFocused();
  await expectNoWcagViolations(page);
});

test('canvas context menus and modal focus are keyboard safe', async ({ page }) => {
  let lastSavedLabel = '';
  let lastSavedGroupLabel = '';
  let lastSavedNodeCount = -1;
  await openInteractionDiagram(page, (payload) => {
    const savedNodes = Array.isArray(payload.nodes)
      ? payload.nodes as Array<{ id?: string; data?: { label?: string } }>
      : [];
    const savedNode = savedNodes.find(node => node.id === 'node-a');
    const savedGroup = savedNodes.find(node => node.id === 'group-a');
    lastSavedLabel = String(savedNode?.data?.label || '');
    lastSavedGroupLabel = String(savedGroup?.data?.label || '');
    lastSavedNodeCount = savedNodes.length;
  });
  await expectNoWcagViolations(page);
  const nodeA = page.locator('[data-testid="rf__node-node-a"]');
  const nodeB = page.locator('[data-testid="rf__node-node-b"]');
  const nodeALabel = nodeA.locator('[data-node-keyboard-target]');
  const edge = page.locator('[data-testid="rf__edge-edge-ab"]');

  await nodeALabel.focus();
  await expect(nodeA).toHaveClass(/selected/);
  const beforeMove = await nodeA.boundingBox();
  await page.keyboard.press('ArrowRight');
  await expect.poll(async () => (await nodeA.boundingBox())?.x || 0)
    .toBeGreaterThan(beforeMove?.x || 0);
  await page.keyboard.press('F2');
  const nodeLabelInput = nodeA.locator('.node-label-input');
  await nodeLabelInput.fill('Renamed App Service');
  await nodeLabelInput.press('Enter');
  await expect(nodeALabel).toHaveText('Renamed App Service');
  await expect(nodeALabel).toBeFocused();
  await expect.poll(() => lastSavedLabel, { timeout: 5_000 }).toBe('Renamed App Service');
  await page.keyboard.press('F2');
  await nodeA.locator('.node-label-input').fill('Discarded label');
  await page.keyboard.press('Escape');
  await expect(nodeALabel).toHaveText('Renamed App Service');
  await expect(nodeALabel).toBeFocused();

  const groupA = page.locator('[data-testid="rf__node-group-a"]');
  const groupLabel = groupA.locator('[data-node-keyboard-target]');
  await groupLabel.focus();
  await page.keyboard.press('F2');
  const groupLabelInput = groupA.locator('.group-label-input');
  await groupLabelInput.fill('Renamed application layer');
  await groupLabelInput.press('Enter');
  await expect(groupLabel).toHaveText('Renamed application layer');
  await expect(groupLabel).toBeFocused();
  await expect.poll(() => lastSavedGroupLabel, { timeout: 5_000 })
    .toBe('Renamed application layer');
  await page.keyboard.press('F2');
  await groupA.locator('.group-label-input').fill('Discarded group label');
  await page.keyboard.press('Escape');
  await expect(groupLabel).toHaveText('Renamed application layer');
  await expect(groupLabel).toBeFocused();
  expect(lastSavedGroupLabel).toBe('Renamed application layer');

  await nodeA.click();
  await nodeB.click({ modifiers: ['Control'] });
  await expect(nodeA).toHaveClass(/selected/);
  await expect(nodeB).toHaveClass(/selected/);

  await nodeA.click({ button: 'right' });
  const nodeMenu = page.getByRole('menu', { name: 'Service actions' });
  await expect(nodeMenu.getByRole('menuitem', { name: 'Duplicate service' })).toBeFocused();
  await expectNoWcagViolations(page, '.node-context-menu');
  await expect(page.locator('.react-flow__node.selected')).toHaveCount(1);
  await expect(nodeA).toHaveClass(/selected/);
  await page.keyboard.press('End');
  await expect(nodeMenu.getByRole('menuitem', { name: 'Delete service' })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(nodeMenu.getByRole('menuitem', { name: 'Duplicate service' })).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await nodeMenu.getByRole('menuitem', { name: 'Set cost estimate' }).press('Enter');
  const pricingEditor = page.locator('.npe-modal');
  await expect(pricingEditor).toBeFocused();
  await expectNoWcagViolations(page, '.npe-modal');
  await page.keyboard.press('Escape');
  await expect(pricingEditor).toBeHidden();
  await expect(nodeALabel).toBeFocused();

  await page.keyboard.press('Shift+F10');
  await expect(nodeMenu).toBeVisible();
  await page.setViewportSize({ width: 1500, height: 900 });
  await expect(nodeMenu).toBeHidden();
  await expect(nodeALabel).toBeFocused();

  const edgeLabel = page.locator('[data-edge-label-id="edge-ab"]');
  await edgeLabel.focus();
  await page.keyboard.press('Shift+F10');
  const edgeMenu = page.getByRole('menu', { name: 'Connection actions' });
  await expect(edgeMenu.getByRole('menuitem', { name: 'One-way (Forward)' })).toBeFocused();
  await expectNoWcagViolations(page, '.edge-context-menu');
  await page.keyboard.press('Escape');
  await expect(edgeMenu).toBeHidden();
  await expect(edgeLabel).toBeFocused();

  await edge.focus();
  await page.keyboard.press('Shift+F10');
  await expect(edgeMenu.getByRole('menuitem', { name: 'One-way (Forward)' })).toBeFocused();
  await page.keyboard.press('End');
  await expect(edgeMenu.getByRole('menuitem', { name: 'Delete connection' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(edgeMenu).toBeHidden();
  await expect(edge).toBeFocused();
  await page.keyboard.press('Delete');
  await expect(edge).toBeHidden();

  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await expect(canvas).toBeFocused();
  await canvas.focus();
  await page.keyboard.press('Shift+F10');
  const canvasMenu = page.getByRole('menu', { name: 'Canvas actions' });
  await expect(canvasMenu.getByRole('menuitem', { name: 'Add layer here' })).toBeFocused();
  await expectNoWcagViolations(page, '.pane-context-menu');
  await page.keyboard.press('ArrowDown');
  await expect(canvasMenu.getByRole('menuitem', { name: 'Fit diagram to view' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(canvasMenu).toBeHidden();
  await expect(canvas).toBeFocused();

  await nodeA.click();
  await page.evaluate(() => {
    const dialog = document.createElement('div');
    dialog.dataset.testNonModalDialog = 'true';
    dialog.setAttribute('role', 'dialog');
    dialog.textContent = 'Non-modal notice';
    document.body.appendChild(dialog);
  });
  await page.keyboard.press('Delete');
  await expect(nodeA).toBeVisible();
  await page.evaluate(() => {
    document.querySelector('[data-test-non-modal-dialog]')?.remove();
  });

  const accessButton = page.getByRole('button', { name: 'Access', exact: true });
  await accessButton.click();
  const accessModal = page.locator('.access-modal');
  await expect(accessModal).toBeFocused();
  await expectNoWcagViolations(page, '.access-modal');
  await page.keyboard.press('Shift+Tab');
  await expect(accessModal.locator(':focus')).toHaveCount(1);
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  await page.keyboard.press('Delete');
  await expect(nodeA).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(accessModal).toBeHidden();
  await expect(accessButton).toBeFocused();

  await nodeB.locator('[data-node-keyboard-target]').focus();
  await page.keyboard.press('Shift+F10');
  await nodeMenu.getByRole('menuitem', { name: 'Delete service' }).click();
  await expect(nodeB).toBeHidden();
  await expect(canvas).toBeFocused();

  await nodeALabel.focus();
  await page.keyboard.press('Delete');
  await expect(nodeA).toBeHidden();
  await groupLabel.focus();
  await page.keyboard.press('Delete');
  await expect(page.locator('.react-flow__node')).toHaveCount(0);
  await expect.poll(() => lastSavedNodeCount, { timeout: 5_000 }).toBe(0);
});

test('AI generation cannot be dismissed while work is active', async ({ page }) => {
  await initializePage(page);
  await page.route('**/api/**', async (route) => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/openai') {
      await wait(600);
      await fulfillJson(route, {
        error: {
          source: 'azure_openai',
          code: 'upstream_unavailable',
          message: 'Test generation failure',
        },
      }, 502);
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  const modal = await openAiGenerator(page);
  await expectNoWcagViolations(page, '.ai-architecture-modal');
  await modal.getByLabel('Architecture Description or Modification')
    .fill('Create a small web application');
  await modal.getByRole('button', { name: 'Generate Architecture' }).click();
  await expect(modal).toHaveAttribute('aria-busy', 'true');
  await expect(modal.locator('.modal-close')).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(modal).toBeVisible();
  await page.locator('.modal-overlay').dispatchEvent('click');
  await expect(modal).toBeVisible();

  await expect(modal).toHaveAttribute('aria-busy', 'false', { timeout: 5_000 });
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
});

test('image analysis is single-flight and the reference viewer is keyboard safe', async ({ page }) => {
  await initializePage(page);
  await page.clock.install({ time: new Date(now) });
  let imageAnalysisCalls = 0;
  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/openai') {
      const isImageAnalysis = (request.postData() || '').includes('input_image');
      if (isImageAnalysis) {
        imageAnalysisCalls += 1;
        await wait(500);
        await fulfillJson(route, {
          model: 'playwright-gpt-5-6-sol',
          output_text: 'An App Service hosts the web application.',
          usage: { input_tokens: 20, output_tokens: 10, total_tokens: 30 },
        });
        return;
      }
      await fulfillJson(route, {
        model: 'playwright-gpt-5-6-sol',
        output_text: JSON.stringify({
          architectureName: 'Uploaded diagram',
          groups: [],
          services: [{
            id: 'web',
            name: 'App Service',
            type: 'App Service',
            category: 'compute',
            description: 'Hosts the web application',
          }],
          connections: [],
          workflow: [],
        }),
        usage: { input_tokens: 30, output_tokens: 20, total_tokens: 50 },
      });
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  const modal = await openAiGenerator(page);
  const fileInput = modal.locator('input[type="file"]');
  const pixelPng = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=',
    'base64',
  );
  await fileInput.setInputFiles({
    name: 'first.png',
    mimeType: 'image/png',
    buffer: pixelPng,
  });
  await expect(fileInput).toBeDisabled();
  await expect(modal).toHaveAttribute('aria-busy', 'true');
  await expect(modal.locator('.modal-close')).toBeDisabled();
  await expect(modal.locator('.modal-footer-actions').getByRole('button', { name: 'Cancel' }))
    .toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(modal).toBeVisible();
  await page.locator('.modal-overlay').dispatchEvent('click');
  await expect(modal).toBeVisible();
  await page.evaluate((base64) => {
    const bytes = Uint8Array.from(atob(base64), character => character.charCodeAt(0));
    const transfer = new DataTransfer();
    transfer.items.add(new File([bytes], 'second.png', { type: 'image/png' }));
    document.querySelector('.image-drop-zone')?.dispatchEvent(new DragEvent('drop', {
      bubbles: true,
      cancelable: true,
      dataTransfer: transfer,
    }));
  }, pixelPng.toString('base64'));
  await expect(modal.getByLabel('Architecture Description or Modification'))
    .toHaveValue(/Analyzed from uploaded diagram[\s\S]*App Service/, { timeout: 5_000 });
  await expect(modal).toHaveAttribute('aria-busy', 'false');
  expect(imageAnalysisCalls).toBe(1);

  await modal.getByRole('button', { name: 'Generate Architecture' }).click();
  await expect(page.locator('[data-testid="rf__node-web"]')).toBeVisible({ timeout: 10_000 });
  await fileInput.setInputFiles({
    name: 'after-generation.png',
    mimeType: 'image/png',
    buffer: pixelPng,
  });
  await expect(modal).toHaveAttribute('aria-busy', 'true');
  await page.clock.fastForward(46_000);
  await expect(modal).toBeVisible();
  await expect(modal).toHaveAttribute('aria-busy', 'false', { timeout: 5_000 });
  expect(imageAnalysisCalls).toBe(2);
  await modal.locator('.modal-footer-actions').getByRole('button', { name: 'Close' }).click();
  await expect(modal).toBeHidden();

  const expandReference = page.getByRole('button', { name: 'Expand reference image' });
  await expandReference.click();
  const referenceDialog = page.getByRole('dialog', { name: 'Reference Diagram' });
  await expect(referenceDialog).toBeFocused();
  await expectNoWcagViolations(page);
  await page.keyboard.press('Shift+Tab');
  await expect(referenceDialog.locator(':focus')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(referenceDialog).toBeHidden();
  await expect(expandReference).toBeFocused();
});

test('cloud workspace ignores stale details and destructive completions', async ({ page }) => {
  await initializePage(page);
  let aVersionListCalls = 0;
  let bVersionListCalls = 0;
  let bRevision = 1;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();

    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await wait(20);
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'DELETE') {
      await wait(500);
      await route.fulfill({ status: 204, body: '' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await wait(20);
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      bRevision += 1;
      await fulfillJson(route, {
        ...cloudDocument('B', body.diagramName || 'Diagram B'),
        diagramName: body.diagramName || 'Diagram B',
        payload: body.payload,
        revision: bRevision,
        etag: `"B-${bRevision}"`,
      }, 200, { etag: `"B-${bRevision}"` });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      aVersionListCalls += 1;
      await wait(20);
      await fulfillJson(route, {
        versions: [cloudVersion('A', `va-${aVersionListCalls}`, 'A version')],
      });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'GET') {
      bVersionListCalls += 1;
      await wait(bVersionListCalls === 1 ? 500 : 20);
      await fulfillJson(route, {
        versions: [cloudVersion('B', 'vb', 'B version')],
      });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      if (path.includes('/B/') && bVersionListCalls === 1) await wait(500);
      else await wait(20);
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions/vb' && method === 'GET') {
      await wait(500);
      await fulfillJson(route, { version: cloudVersion('B', 'vb', 'B version') });
      return;
    }

    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  const diagramA = modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram A' });
  const diagramB = modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' });

  await expect(modal.getByRole('heading', { name: 'Diagram A' })).toBeVisible();
  await diagramB.click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  await diagramA.click();
  await expect(modal.getByText('A version')).toBeVisible();
  await page.waitForTimeout(650);
  await expect(modal.getByRole('heading', { name: 'Diagram A' })).toBeVisible();
  await expect(modal.getByText('B version')).toHaveCount(0);

  await diagramB.click();
  await expect(modal.getByText('B version')).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await diagramA.click();
  await expect(modal.getByRole('heading', { name: 'Diagram A' })).toBeVisible();
  await page.waitForTimeout(650);
  await expect(modal).toBeVisible();

  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Delete' }).click();
  await diagramB.click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  await modal.getByRole('button', { name: 'Open' }).click();
  await page.waitForTimeout(650);
  await expect(cloudButton).toHaveClass(/btn-active/);

  await cloudButton.click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Open' })).toBeEnabled();
});

test('failed shared startup load blocks autosave of persisted local settings', async ({ page }) => {
  await initializePage(page);
  const shareToken = 's'.repeat(43);
  let sharedLoads = 0;
  let createAttempts = 0;
  await page.addInitScript(() => {
    localStorage.setItem('azurediagarm.pricing-scenarios.v1', JSON.stringify([{
      id: 'persisted-custom',
      name: 'Persisted custom pricing',
      kind: 'custom',
      pricingMode: 'payg',
      capacityMultiplier: 1.2,
      usageMultiplier: 0.8,
      discountPercent: 5,
      supportPercent: 0,
      currency: 'JPY',
      exchangeRate: 150,
    }]));
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'viewer@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}` && method === 'GET') {
      sharedLoads += 1;
      await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, cloudDocument('unexpected', 'Unexpected'), 201, {
        etag: '"unexpected-1"',
      });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto(`/#share-${shareToken}`);
  await expect.poll(() => sharedLoads).toBeGreaterThan(0);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Local only');
  await expect(cloudButton).toHaveAttribute('title', 'Temporary storage outage');
  await page.clock.fastForward(5_000);
  await page.waitForTimeout(100);
  expect(createAttempts).toBe(0);
  await expect(cloudButton).toHaveAttribute('title', 'Temporary storage outage');
});

test('invalid share link remains an error across StrictMode initialization', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let staleContextLoads = 0;
  let createAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      staleContextLoads += 1;
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, cloudDocument('unexpected', 'Unexpected'), 201, {
        etag: '"unexpected-1"',
      });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/#share-bad');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Local only');
  await expect(cloudButton).toHaveAttribute('title', 'The shared diagram link is invalid.');
  await page.clock.fastForward(5_000);
  await page.waitForTimeout(100);
  expect(staleContextLoads).toBe(0);
  expect(createAttempts).toBe(0);
  await expect(cloudButton).toHaveAttribute('title', 'The shared diagram link is invalid.');
});

test('a stale collaboration session cannot resurrect remotely deleted content', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let etagSequence = 1;
  let currentEtag = '"A-1"';
  let currentPayload = interactionCloudDocument().payload;
  let currentComments: Array<Record<string, unknown>> = [];
  let commentCalls = 0;
  let preCommentEtag = '';
  let staleSaveAttempts = 0;
  let resurrectionAccepted = false;
  let failReload = false;

  const currentDocument = () => ({
    ...interactionCloudDocument(),
    payload: currentPayload,
    comments: currentComments,
    revision: etagSequence,
    updatedAt: now,
    etag: currentEtag,
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      if (failReload) {
        await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
        return;
      }
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [{
          ...summary('A', 'Interaction diagram'),
          revision: etagSequence,
          etag: currentEtag,
        }],
      });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, {
        versions: [cloudVersion('A', 'collaboration-snapshot', 'Before collaboration')],
      });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentCalls += 1;
      if (commentCalls === 1) {
        currentComments = [{
          commentId: 'comment-old',
          message: 'Delayed review',
          authorEmail: 'owner@example.com',
          createdAt: now,
        }];
        etagSequence += 1;
        currentEtag = `"A-${etagSequence}"`;
        const staleMetadataResponse = currentDocument();
        const staleMetadataEtag = currentEtag;
        await wait(1_200);
        await fulfillJson(route, staleMetadataResponse, 201, { etag: staleMetadataEtag });
        return;
      }
      preCommentEtag = currentEtag;
      currentPayload = { ...currentPayload, nodes: [], edges: [] };
      currentComments = [{
        commentId: 'comment-new',
        message: 'Remote review',
        authorEmail: 'owner@example.com',
        createdAt: now,
      }];
      etagSequence += 1;
      currentEtag = `"A-${etagSequence}"`;
      await fulfillJson(route, currentDocument(), 201, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const requestEtag = request.headers()['if-match'] || '';
      const body = JSON.parse(request.postData() || '{}');
      if (preCommentEtag) staleSaveAttempts += 1;
      if (requestEtag !== currentEtag) {
        await fulfillJson(route, {
          error: 'The document was modified by another request. Reload and retry.',
        }, 412);
        return;
      }
      currentPayload = body.payload;
      etagSequence += 1;
      currentEtag = `"A-${etagSequence}"`;
      resurrectionAccepted = currentPayload.nodes.length > 0 && preCommentEtag !== '';
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const staleNode = page.locator('[data-testid="rf__node-node-a"]');
  await expect(staleNode).toBeVisible();
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/, { timeout: 5_000 });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Delayed review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentCalls).toBe(1);
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
  await expect(cloudButton).toBeFocused();

  await cloudButton.click();
  await modal.getByPlaceholder('Add a review comment...').fill('Remote review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
  await expect(modal.getByRole('button', { name: 'Restore' })).toBeDisabled();
  await page.waitForTimeout(1_300);
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  failReload = true;
  await modal.getByRole('button', { name: 'Load cloud copy' }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Load cloud copy' })).toBeEnabled();
  await modal.getByRole('button', { name: 'Save as copy' }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Save as copy' })).toBeEnabled();
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
  await expect(cloudButton).toBeFocused();

  const staleLabel = staleNode.locator('[data-node-keyboard-target]');
  await staleLabel.focus();
  await expect(staleLabel).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await page.waitForTimeout(2_300);
  expect(preCommentEtag).not.toBe('');
  expect(staleSaveAttempts).toBe(0);
  expect(resurrectionAccepted).toBe(false);
  expect(currentPayload.nodes).toHaveLength(0);
});

test('metadata success reconciles the current ETag after the modal closes', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentEtag = '"A-1"';
  let currentPayload = cloudDocument('A', 'Diagram A').payload;
  let updateAttempts = 0;
  let commentStarted = 0;
  let commentCompleted = 0;
  let postCommentSaveEtag = '';

  const currentDocument = () => ({
    ...cloudDocument('A', 'Diagram A'),
    payload: currentPayload,
    comments: commentCompleted > 0 ? [{
      commentId: 'comment-after-close',
      message: 'Closed modal review',
      authorEmail: 'owner@example.com',
      createdAt: now,
    }] : [],
    revision: Number(currentEtag.match(/\d+/)?.[0] || 1),
    etag: currentEtag,
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const requestEtag = request.headers()['if-match'] || '';
      if (commentCompleted > 0) postCommentSaveEtag = requestEtag;
      if (requestEtag !== currentEtag) {
        await fulfillJson(route, { error: 'The document changed.' }, 412);
        return;
      }
      const body = JSON.parse(request.postData() || '{}');
      currentPayload = body.payload;
      currentEtag = `"A-${updateAttempts + 1}"`;
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentStarted += 1;
      await wait(800);
      commentCompleted += 1;
      currentEtag = '"A-3"';
      await fulfillJson(route, currentDocument(), 201, { etag: currentEtag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Closed modal review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentStarted).toBe(1);
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
  await expect.poll(() => commentCompleted).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  expect(postCommentSaveEtag).toBe('"A-3"');
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/);
});

test('out-of-order metadata responses cannot roll back the current ETag', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentPayload = cloudDocument('A', 'Diagram A').payload;
  let currentEtag = '"A-1"';
  let currentRevision = 1;
  let commentAttempts = 0;
  let commentCompletions = 0;
  let updateAttempts = 0;
  let postCommentSaveEtag = '';

  const documentAt = (
    revision: number,
    etag: string,
    comments: Array<Record<string, unknown>>,
  ) => ({
    ...cloudDocument('A', 'Diagram A'),
    payload: currentPayload,
    comments,
    revision,
    etag,
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, documentAt(currentRevision, currentEtag, []), 200, {
        etag: currentEtag,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const requestEtag = request.headers()['if-match'] || '';
      if (commentCompletions >= 2) postCommentSaveEtag = requestEtag;
      if (requestEtag !== currentEtag) {
        await fulfillJson(route, { error: 'The document changed.' }, 412);
        return;
      }
      const body = JSON.parse(request.postData() || '{}');
      currentPayload = body.payload;
      currentRevision += 1;
      currentEtag = `"A-${currentRevision}"`;
      await fulfillJson(route, documentAt(currentRevision, currentEtag, []), 200, {
        etag: currentEtag,
      });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      currentRevision += 1;
      currentEtag = `"A-${currentRevision}"`;
      const responseRevision = currentRevision;
      const responseEtag = currentEtag;
      const responseComments = Array.from({ length: commentAttempts }, (_, index) => ({
        commentId: `comment-${index + 1}`,
        message: `Review ${index + 1}`,
        authorEmail: 'owner@example.com',
        createdAt: now,
      }));
      if (commentAttempts === 1) await wait(1_200);
      commentCompletions += 1;
      await fulfillJson(
        route,
        documentAt(responseRevision, responseEtag, responseComments),
        201,
        { etag: responseEtag },
      );
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Review 1');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await page.keyboard.press('Escape');
  await cloudButton.click();
  await modal.getByPlaceholder('Add a review comment...').fill('Review 2');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentCompletions).toBe(2);
  await page.keyboard.press('Escape');
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  expect(postCommentSaveEtag).toBe('"A-4"');
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/);
});

test('a stale metadata failure cannot conflict a newer success on the same document', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentPayload = cloudDocument('A', 'Diagram A').payload;
  let currentRevision = 1;
  let currentEtag = '"A-1"';
  let updateAttempts = 0;
  let commentAttempts = 0;
  let staleFailureCompleted = 0;
  let postCommentSaveEtag = '';

  const currentDocument = () => ({
    ...cloudDocument('A', 'Diagram A'),
    payload: currentPayload,
    revision: currentRevision,
    etag: currentEtag,
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const requestEtag = request.headers()['if-match'] || '';
      if (commentAttempts >= 2) postCommentSaveEtag = requestEtag;
      if (requestEtag !== currentEtag) {
        await fulfillJson(route, { error: 'The document changed.' }, 412);
        return;
      }
      const body = JSON.parse(request.postData() || '{}');
      currentPayload = body.payload;
      currentRevision += 1;
      currentEtag = `"A-${currentRevision}"`;
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      if (commentAttempts === 1) {
        await wait(1_200);
        staleFailureCompleted += 1;
        await fulfillJson(route, {
          error: 'The document was modified by another request. Reload and retry.',
        }, 412);
        return;
      }
      currentRevision += 1;
      currentEtag = `"A-${currentRevision}"`;
      await fulfillJson(route, {
        ...currentDocument(),
        comments: [{
          commentId: 'newer-comment',
          message: 'Newer success',
          authorEmail: 'owner@example.com',
          createdAt: now,
        }],
      }, 201, { etag: currentEtag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Older failure');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await page.keyboard.press('Escape');
  await cloudButton.click();
  await modal.getByPlaceholder('Add a review comment...').fill('Newer success');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect(modal.getByText('Newer success')).toBeVisible();
  await expect.poll(() => staleFailureCompleted).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toHaveCount(0);
  await page.keyboard.press('Escape');
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  expect(postCommentSaveEtag).toBe('"A-3"');
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/);
});

test('reload cannot clear a newer conflict raised while it is in flight', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;
  let commentAttempts = 0;
  let delayedFailureCompleted = 0;
  let delayReload = false;
  let reloadLoads = 0;

  const revisionTwoDocument = {
    ...cloudDocument('A', 'Diagram A'),
    revision: 2,
    etag: '"A-2"',
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      if (delayReload) {
        reloadLoads += 1;
        await wait(1_500);
      }
      await fulfillJson(route, revisionTwoDocument, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...revisionTwoDocument,
        payload: body.payload,
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      if (commentAttempts === 1) {
        await wait(800);
        delayedFailureCompleted += 1;
      }
      await fulfillJson(route, {
        error: 'The document was modified by another request. Reload and retry.',
      }, 412);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Delayed conflict');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await page.keyboard.press('Escape');
  await cloudButton.click();
  await modal.getByPlaceholder('Add a review comment...').fill('Immediate conflict');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  delayReload = true;
  await modal.getByRole('button', { name: 'Load cloud copy' }).click();
  await expect.poll(() => delayedFailureCompleted).toBe(1);
  await expect.poll(() => reloadLoads).toBe(1);
  await page.waitForTimeout(1_600);
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
});

test('normalized shared viewer metadata does not create a false conflict', async ({ page }) => {
  const shareToken = 'v'.repeat(43);
  await initializePage(page, {
    documentId: 'A',
    access: 'shared',
    role: 'viewer',
    shareToken,
  });
  let commentAttempts = 0;
  const sharedDocument = {
    ...cloudDocument('A', 'Shared diagram'),
    access: 'shared',
    role: 'viewer',
    etag: '"A-1"',
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'viewer@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}` && method === 'GET') {
      await fulfillJson(route, sharedDocument, 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [] });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}/versions` && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}/comments` && method === 'POST') {
      commentAttempts += 1;
      await fulfillJson(route, {
        ...sharedDocument,
        comments: [{
          commentId: 'viewer-comment',
          message: 'Viewer review',
          authorEmail: 'viewer@example.com',
          createdAt: now,
        }],
        revision: 2,
        etag: '"A-2"',
      }, 201, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Cloud read-only');
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Viewer review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toHaveCount(0);
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Cloud read-only');
  await expect(modal.getByText('Viewer review')).toBeVisible();
});

test('metadata write conflicts enter the sticky cloud conflict state', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let commentAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, {
        versions: [cloudVersion('A', 'metadata-conflict', 'Before concurrent update')],
      });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      if (commentAttempts === 1) {
        await fulfillJson(route, { error: 'The diagram has reached the comment limit.' }, 409);
        return;
      }
      await fulfillJson(route, {
        error: 'The document was modified by another request. Reload and retry.',
      }, 412);
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Concurrent review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toHaveCount(0);
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeEnabled();
  await modal.getByPlaceholder('Add a review comment...').fill('Retry concurrent review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(2);
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
  await expect(modal.getByRole('button', { name: 'Restore' })).toBeDisabled();
  await page.keyboard.press('Escape');
  await cloudButton.click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
});

test('metadata conflict after a prerequisite save uses the saved revision', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentRevision = 1;
  let currentEtag = '"A-1"';
  let currentPayload = cloudDocument('A', 'Diagram A').payload;
  let updateAttempts = 0;
  let commentAttempts = 0;

  const currentDocument = () => ({
    ...cloudDocument('A', 'Diagram A'),
    payload: currentPayload,
    revision: currentRevision,
    etag: currentEtag,
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      currentPayload = body.payload;
      currentRevision += 1;
      currentEtag = `"A-${currentRevision}"`;
      await fulfillJson(route, currentDocument(), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      await fulfillJson(route, {
        error: 'The document was modified by another request. Reload and retry.',
      }, 412);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Conflict after save');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
});

test('metadata action stops when save replaces a remotely deleted document', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;
  let replacementCreated = 0;
  let commentAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: replacementCreated > 0 ? [summary('B', 'Diagram A')] : [summary('A', 'Diagram A')],
      });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/versions$/.test(path) && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts === 1) {
        await fulfillJson(route, {
          ...cloudDocument('A', body.diagramName || 'Diagram A'),
          payload: body.payload,
          revision: 2,
          etag: '"A-2"',
        }, 200, { etag: '"A-2"' });
      } else {
        await fulfillJson(route, { error: 'Not found' }, 404);
      }
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      replacementCreated += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('B', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 1,
        etag: '"B-1"',
      }, 201, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Do not send to deleted A');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => replacementCreated, { timeout: 5_000 }).toBe(1);
  expect(commentAttempts).toBe(0);
  await expect.poll(async () => (
    JSON.parse(await page.evaluate(() => (
      sessionStorage.getItem('azurediagarm.cloud-document.v1') || '{}'
    ))).documentId
  )).toBe('B');
  await expect(modal).toBeVisible();
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/);
});

test('share refresh failure blocks saves until the ETag is reconciled', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let shareCreated = false;
  let initialUpdates = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      if (shareCreated) {
        await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
      } else {
        await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      }
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      initialUpdates += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'POST') {
      shareCreated = true;
      await fulfillJson(route, {
        share: {
          shareId: 'share-1',
          role: 'viewer',
          createdAt: now,
          createdByEmail: 'owner@example.com',
        },
        token: 't'.repeat(43),
        url: `https://example.test/#share-${'t'.repeat(43)}`,
      }, 201);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => initialUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByRole('button', { name: 'Create link' }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
});

test('a stale metadata failure cannot conflict a newly opened diagram', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let commentStarted = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (/^\/api\/diagrams\/[AB]$/.test(path) && method === 'GET') {
      const id = path.endsWith('/B') ? 'B' : 'A';
      await fulfillJson(route, cloudDocument(id, `Diagram ${id}`), 200, {
        etag: `"${id}-1"`,
      });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/versions$/.test(path) && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentStarted += 1;
      await wait(1_200);
      await fulfillJson(route, {
        error: 'The document was modified by another request. Reload and retry.',
      }, 412);
      return;
    }
    if (/^\/api\/diagrams\/[AB]$/.test(path) && method === 'PUT') {
      const id = path.endsWith('/B') ? 'B' : 'A';
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument(id, body.diagramName || `Diagram ${id}`),
        payload: body.payload,
        revision: 2,
        etag: `"${id}-2"`,
      }, 200, { etag: `"${id}-2"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Delayed concurrent review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentStarted).toBe(1);
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  await modal.getByRole('button', { name: 'Open', exact: true }).click();
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-B-node"]')).toBeVisible();
  await page.waitForTimeout(1_400);
  await expect(cloudButton).toHaveAccessibleName(/Cloud saved/, { timeout: 5_000 });
});

test('a metadata 404 preserves the current local draft as a conflict', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let commentAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      commentAttempts += 1;
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Review after remote deletion');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Save as copy' })).toBeEnabled();
});

test('current diagram detail 404 enters conflict before navigation', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts > 1) await wait(1_200);
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
  await page.waitForTimeout(1_300);
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
});

test('an in-flight save failure cannot hide a newer detail conflict', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts > 1) {
        await wait(1_200);
        await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
        return;
      }
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await page.waitForTimeout(1_300);
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByText('Temporary storage outage')).toHaveCount(0);
});

test('an empty save cannot recreate a document deleted remotely', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentEtag = '"A-1"';
  let emptySaveAttempts = 0;
  let createAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      if (body.payload?.nodes?.length === 0) {
        emptySaveAttempts += 1;
        await fulfillJson(route, { error: 'Not found' }, 404);
        return;
      }
      currentEtag = '"A-2"';
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: currentEtag,
      }, 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, {
        ...cloudDocument('replacement', 'Replacement'),
        payload: JSON.parse(request.postData() || '{}').payload,
      }, 201, { etag: '"replacement-1"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const node = page.locator('[data-testid="rf__node-A-node"]');
  await expect(node).toBeVisible();
  await node.locator('[data-node-keyboard-target]').focus();
  await page.keyboard.press('Delete');
  await expect(node).toBeHidden();
  await expect.poll(() => emptySaveAttempts, { timeout: 5_000 }).toBe(1);
  await page.waitForTimeout(500);
  expect(createAttempts).toBe(0);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Sync conflict');
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Save as copy' })).toBeEnabled();
  let confirmationCount = 0;
  page.on('dialog', async (dialog) => {
    confirmationCount += 1;
    if (confirmationCount === 1) await dialog.accept();
    else await dialog.dismiss();
  });
  await modal.getByRole('button', { name: 'New diagram' }).click();
  await expect.poll(() => confirmationCount).toBe(2);
  await expect(modal).toBeVisible();
  expect(createAttempts).toBe(0);
});

test('a stale 404 cannot replace a newer queued zero-node save', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateStarted = 0;
  let createAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateStarted += 1;
      await wait(1_200);
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, cloudDocument('replacement', 'Replacement'), 201, {
        etag: '"replacement-1"',
      });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const node = page.locator('[data-testid="rf__node-A-node"]');
  await expect(node).toBeVisible();
  const nodeTarget = node.locator('[data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateStarted, { timeout: 5_000 }).toBe(1);
  await nodeTarget.focus();
  await page.keyboard.press('Delete');
  await expect(node).toBeHidden();
  await page.waitForTimeout(2_500);
  expect(createAttempts).toBe(0);
});

test('remote deletion replacement is not deduplicated after reverting an edit', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;
  let replacementPayload: { nodes?: Array<{ position?: { x?: number } }> } | null = null;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts === 1) {
        await fulfillJson(route, {
          ...cloudDocument('A', body.diagramName || 'Diagram A'),
          payload: body.payload,
          revision: 2,
          etag: '"A-2"',
        }, 200, { etag: '"A-2"' });
        return;
      }
      await wait(1_200);
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      const body = JSON.parse(request.postData() || '{}');
      replacementPayload = body.payload;
      await fulfillJson(route, {
        ...cloudDocument('replacement', body.diagramName || 'Diagram A'),
        payload: body.payload,
        etag: '"replacement-1"',
      }, 201, { etag: '"replacement-1"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  await nodeTarget.focus();
  await page.keyboard.press('ArrowLeft');
  await expect.poll(() => replacementPayload, { timeout: 5_000 }).not.toBeNull();
  expect(replacementPayload?.nodes?.[0]?.position?.x).toBe(100);
});

test('reverting while a save is in flight persists the reverted payload', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;
  let remoteNodeX = 100;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts === 2) await wait(1_200);
      remoteNodeX = Number(body.payload?.nodes?.[0]?.position?.x);
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: updateAttempts + 1,
        etag: `"A-${updateAttempts + 1}"`,
      }, 200, { etag: `"A-${updateAttempts + 1}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  await nodeTarget.focus();
  await page.keyboard.press('ArrowLeft');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(3);
  expect(remoteNodeX).toBe(100);
});

test('deleting during initial cloud creation persists the final empty canvas', async ({ page }) => {
  await initializePage(page);
  let createStarted = 0;
  let emptyUpdateStarted = 0;
  let remotePayload: Record<string, unknown> | null = null;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createStarted += 1;
      const body = JSON.parse(request.postData() || '{}');
      remotePayload = body.payload;
      await wait(1_500);
      await fulfillJson(route, {
        ...cloudDocument('created', body.diagramName || 'Created diagram'),
        diagramName: body.diagramName || 'Created diagram',
        payload: body.payload,
        revision: 1,
        etag: '"created-1"',
      }, 201, { etag: '"created-1"' });
      return;
    }
    if (path === '/api/diagrams/created' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      remotePayload = body.payload;
      if (body.payload?.nodes?.length === 0) emptyUpdateStarted += 1;
      await fulfillJson(route, {
        ...cloudDocument('created', body.diagramName || 'Created diagram'),
        diagramName: body.diagramName || 'Created diagram',
        payload: body.payload,
        revision: 2,
        etag: '"created-2"',
      }, 200, { etag: '"created-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menu', { name: 'Canvas actions' })
    .getByRole('menuitem', { name: 'Add layer here' })
    .click();
  const group = page.locator('.react-flow__node').first();
  await expect(group).toBeVisible();
  await expect.poll(() => createStarted, { timeout: 5_000 }).toBe(1);

  await group.locator('[data-node-keyboard-target]').focus();
  await page.keyboard.press('Delete');
  await expect(page.locator('.react-flow__node')).toHaveCount(0);
  await expect.poll(() => emptyUpdateStarted, { timeout: 5_000 }).toBe(1);
  expect((remotePayload as { nodes?: unknown[] } | null)?.nodes).toHaveLength(0);
});

test('clearing a failed new draft cancels its cloud retry', async ({ page }) => {
  await initializePage(page);
  let createAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menu', { name: 'Canvas actions' })
    .getByRole('menuitem', { name: 'Add layer here' })
    .click();
  const group = page.locator('.react-flow__node').first();
  await expect(group).toBeVisible();
  await expect.poll(() => createAttempts, { timeout: 5_000 }).toBe(1);
  await group.locator('[data-node-keyboard-target]').focus();
  await page.keyboard.press('Delete');
  await expect(page.locator('.react-flow__node')).toHaveCount(0);
  await expect(page.getByRole('button', { name: /^Cloud workspace:/ }))
    .toHaveAccessibleName('Cloud workspace: Cloud');
  await page.clock.fastForward(16_000);
  await page.waitForTimeout(100);
  expect(createAttempts).toBe(1);
});

test('non-retryable client save errors are not retried', async ({ page }) => {
  await initializePage(page);
  let createAttempts = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      createAttempts += 1;
      await fulfillJson(route, { error: 'The diagram payload is invalid.' }, 400);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Shift+F10');
  await page.getByRole('menu', { name: 'Canvas actions' })
    .getByRole('menuitem', { name: 'Add layer here' })
    .click();
  await expect.poll(() => createAttempts, { timeout: 5_000 }).toBe(1);
  await expect(page.getByRole('button', { name: /^Cloud workspace:/ }))
    .toHaveAccessibleName('Cloud workspace: Local only');
  await page.clock.fastForward(16_000);
  await page.waitForTimeout(100);
  expect(createAttempts).toBe(1);
});

test('metadata, validation, and pricing-only drafts persist with zero nodes', async ({ page }) => {
  await initializePage(page);
  const createdPayloads: Record<string, unknown>[] = [];

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      const body = JSON.parse(request.postData() || '{}');
      createdPayloads.push(body.payload);
      await fulfillJson(route, {
        ...cloudDocument(`draft-${createdPayloads.length}`, body.diagramName || 'Draft'),
        diagramName: body.diagramName || 'Metadata only',
        payload: body.payload,
        etag: `"draft-${createdPayloads.length}"`,
      }, 201, { etag: `"draft-${createdPayloads.length}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await page.locator('input[accept=".json"]').setInputFiles({
    name: 'blank-title-draft.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [],
      edges: [],
      titleBlockData: { architectureName: '   ' },
    })),
  });
  await page.waitForTimeout(2_300);
  expect(createdPayloads).toHaveLength(0);

  await page.locator('input[accept=".json"]').setInputFiles({
    name: 'metadata-only-draft.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [],
      edges: [],
      architecturePrompt: '',
      originalPrompt: '',
      workflow: [],
      titleBlockData: {
        architectureName: 'Metadata-only draft',
        author: 'Architecture team',
        version: '2.0',
        date: '2026-08-03',
      },
    })),
  });
  await expect.poll(() => createdPayloads.length, { timeout: 5_000 }).toBe(1);
  expect((createdPayloads[0] as { nodes?: unknown[] }).nodes).toHaveLength(0);
  expect(
    (createdPayloads[0] as { titleBlockData?: { architectureName?: string } })
      ?.titleBlockData?.architectureName,
  ).toBe('Metadata-only draft');

  await page.locator('input[accept=".json"]').setInputFiles({
    name: 'validation-only-draft.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [],
      edges: [],
      validationScore: 0,
    })),
  });
  await expect.poll(() => createdPayloads.length, { timeout: 5_000 }).toBe(2);
  expect((createdPayloads[1] as { nodes?: unknown[] }).nodes).toHaveLength(0);
  expect((createdPayloads[1] as { validationScore?: number }).validationScore).toBe(0);

  await page.locator('input[accept=".json"]').setInputFiles({
    name: 'pricing-only-draft.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [],
      edges: [],
      pricingScenarios: [{
        id: 'custom-pricing',
        name: 'Custom pricing scenario',
        kind: 'custom',
        pricingMode: 'payg',
        capacityMultiplier: 1.5,
        usageMultiplier: 0.75,
        discountPercent: 10,
        supportPercent: 5,
        currency: 'JPY',
        exchangeRate: 150,
      }],
    })),
  });
  await expect.poll(() => createdPayloads.length, { timeout: 5_000 }).toBe(3);
  expect((createdPayloads[2] as { nodes?: unknown[] }).nodes).toHaveLength(0);
  expect(
    (createdPayloads[2] as { pricingScenarios?: Array<{ name?: string }> })
      .pricingScenarios?.[0]?.name,
  ).toBe('Custom pricing scenario');
});

test('snapshot restore cannot overtake a newly detected cloud conflict', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let currentEtag = '"A-1"';
  let conflictingSaveStarted = 0;
  let snapshotFetches = 0;
  const snapshotVersion = {
    ...cloudVersion('A', 'conflict-snapshot', 'Remote snapshot'),
    payload: {
      ...cloudDocument('A', 'Diagram A').payload,
      nodes: [{
        id: 'snapshot-node',
        type: 'azureNode',
        position: { x: 300, y: 200 },
        data: { label: 'Snapshot Node', serviceName: 'App Service' },
      }],
      edges: [],
    },
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: currentEtag });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/versions/conflict-snapshot' && method === 'GET') {
      snapshotFetches += 1;
      await wait(1_800);
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      const nodeX = Number(body.payload?.nodes?.[0]?.position?.x || 0);
      if (nodeX > 100) {
        conflictingSaveStarted += 1;
        await wait(4_000);
        await fulfillJson(route, {
          error: 'The document was modified by another request. Reload and retry.',
        }, 412);
        return;
      }
      currentEtag = '"A-2"';
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: currentEtag,
      }, 200, { etag: currentEtag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const originalNode = page.locator('[data-testid="rf__node-A-node"]');
  await expect(originalNode).toBeVisible();
  const originalLabel = originalNode.locator('[data-node-keyboard-target]');
  await originalLabel.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => conflictingSaveStarted, { timeout: 5_000 }).toBe(1);

  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  let confirmationCount = 0;
  page.on('dialog', async (dialog) => {
    confirmationCount += 1;
    if (confirmationCount === 1) await dialog.accept();
    else await dialog.dismiss();
  });
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible({ timeout: 5_000 });
  await expect(modal.getByRole('alert')).toContainText(
    'The document was modified by another request. Reload and retry.',
    { timeout: 5_000 },
  );
  await expect(modal).toBeVisible();
  expect(snapshotFetches).toBe(0);
  await expect(originalNode).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-snapshot-node"]')).toHaveCount(0);
});

test('snapshot restore backs up the current cloud revision before replacement', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  const operationOrder: string[] = [];
  let backupNotes = '';
  let restoredDiagramName = '';
  const snapshotVersion = {
    ...cloudVersion('A', 'restore-snapshot', 'Restore target'),
    diagramName: 'Historical Diagram A',
    payload: {
      ...cloudDocument('A', 'Diagram A').payload,
      titleBlockData: { architectureName: 'Historical Diagram A' },
      nodes: [{
        id: 'snapshot-node',
        type: 'azureNode',
        position: { x: 300, y: 200 },
        data: { label: 'Snapshot Node', serviceName: 'App Service' },
      }],
      edges: [],
    },
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/versions/restore-snapshot' && method === 'GET') {
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'POST') {
      operationOrder.push('backup');
      backupNotes = JSON.parse(request.postData() || '{}').notes || '';
      await fulfillJson(route, {
        version: cloudVersion('A', 'automatic-backup', backupNotes),
      }, 201);
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      if (body.payload?.nodes?.some((node: { id?: string }) => node.id === 'snapshot-node')) {
        operationOrder.push('restore-save');
        restoredDiagramName = body.diagramName || '';
      }
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-snapshot-node"]')).toBeVisible();
  await expect.poll(() => operationOrder.includes('restore-save'), { timeout: 5_000 }).toBe(true);
  expect(operationOrder).toEqual(['backup', 'restore-save']);
  expect(backupNotes).toContain('Automatic backup');
  expect(restoredDiagramName).toBe('Historical Diagram A');
});

test('edits made while snapshot restore verifies are queued', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let revision = 1;
  let restoreVerificationStarted = 0;
  let manualSnapshots = 0;
  let releaseRestoreVerification = () => {};
  const restoreVerificationGate = new Promise<void>((resolve) => {
    releaseRestoreVerification = resolve;
  });
  const restoredPositions: number[] = [];
  const restoredEtags: string[] = [];
  const snapshotVersion = {
    ...cloudVersion('A', 'edit-during-restore', 'Edit during restore'),
    payload: {
      ...cloudDocument('A', 'Diagram A').payload,
      nodes: [{
        id: 'restore-edit-node',
        type: 'azureNode',
        position: { x: 300, y: 200 },
        data: { label: 'Restore Edit Node', serviceName: 'App Service' },
      }],
      edges: [],
    },
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/versions/edit-during-restore' && method === 'GET') {
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'POST') {
      const notes = JSON.parse(request.postData() || '{}').notes || '';
      if (notes === 'Saved during restore') manualSnapshots += 1;
      await fulfillJson(route, {
        version: cloudVersion('A', 'edit-during-restore-backup', notes),
      }, 201);
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      const restoredNode = body.payload?.nodes?.find(
        (node: { id?: string }) => node.id === 'restore-edit-node',
      );
      if (restoredNode) {
        restoredEtags.push(request.headers()['if-match'] || '');
        restoredPositions.push(Number(restoredNode.position?.x || 0));
        if (restoreVerificationStarted === 0) {
          restoreVerificationStarted += 1;
          await restoreVerificationGate;
        }
      }
      revision += 1;
      const etag = `"A-${revision}"`;
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision,
        etag,
      }, 200, { etag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => revision, { timeout: 5_000 }).toBeGreaterThan(1);
  await page.getByRole('button', { name: /^Cloud workspace:/ }).click();
  const modal = page.locator('.cloud-workspace-modal');
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect.poll(() => restoreVerificationStarted, { timeout: 5_000 }).toBe(1);

  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
  await page.waitForTimeout(100);
  const restoredNode = page.locator('[data-testid="rf__node-restore-edit-node"]');
  await expect(restoredNode).toBeVisible();
  const beforeMove = await restoredNode.boundingBox();
  const restoredNodeTarget = restoredNode.locator('[data-node-keyboard-target]');
  await restoredNodeTarget.focus();
  await expect(restoredNodeTarget).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect.poll(async () => (await restoredNode.boundingBox())?.x || 0)
    .toBeGreaterThan(beforeMove?.x || 0);

  await page.getByRole('button', { name: 'Snapshot', exact: true }).click();
  const snapshotModal = page.getByRole('dialog', { name: 'Save Snapshot' });
  await snapshotModal.getByPlaceholder(/Before adding authentication/).fill('Saved during restore');
  await snapshotModal.getByRole('button', { name: 'Save Snapshot' }).click();
  releaseRestoreVerification();

  await expect.poll(
    () => restoredPositions.some(position => position > 300),
    { timeout: 7_000 },
  ).toBe(true);
  await expect.poll(() => manualSnapshots, { timeout: 7_000 }).toBe(1);
  expect(restoredEtags.length).toBeGreaterThan(1);
  expect(restoredEtags[1]).not.toBe(restoredEtags[0]);
});

test('snapshot restore can switch from the current diagram to another diagram', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let diagramAUpdates = 0;
  let diagramBUpdates = 0;
  const snapshotVersion = {
    ...cloudVersion('B', 'cross-document-snapshot', 'Cross-document restore'),
    payload: {
      ...cloudDocument('B', 'Diagram B').payload,
      nodes: [{
        id: 'cross-document-node',
        type: 'azureNode',
        position: { x: 380, y: 220 },
        data: { label: 'Cross-document Snapshot', serviceName: 'App Service' },
      }],
      edges: [],
    },
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions/cross-document-snapshot' && method === 'GET') {
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'POST') {
      await fulfillJson(route, {
        version: cloudVersion('B', 'cross-document-backup', 'Automatic backup'),
      }, 201);
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      diagramAUpdates += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: diagramAUpdates + 1,
        etag: `"A-${diagramAUpdates + 1}"`,
      }, 200, { etag: `"A-${diagramAUpdates + 1}"` });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'PUT') {
      diagramBUpdates += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('B', body.diagramName || 'Diagram B'),
        payload: body.payload,
        revision: 2,
        etag: '"B-2"',
      }, 200, { etag: '"B-2"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-cross-document-node"]')).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-A-node"]')).toHaveCount(0);
  await expect.poll(() => diagramBUpdates, { timeout: 5_000 }).toBe(1);
});

test('identical cross-document snapshot restore still verifies the target ETag', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let diagramAUpdates = 0;
  let diagramBUpdates = 0;
  const snapshotVersion = {
    ...cloudVersion('B', 'identical-snapshot', 'Identical cached snapshot'),
    payload: cloudDocument('B', 'Diagram B').payload,
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions/identical-snapshot' && method === 'GET') {
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'POST') {
      await fulfillJson(route, {
        version: cloudVersion('B', 'identical-backup', 'Automatic backup'),
      }, 201);
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      diagramAUpdates += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: diagramAUpdates + 1,
        etag: `"A-${diagramAUpdates + 1}"`,
      }, 200, { etag: `"A-${diagramAUpdates + 1}"` });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'PUT') {
      diagramBUpdates += 1;
      await fulfillJson(route, {
        error: 'The document was modified by another request. Reload and retry.',
      }, 412);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  await page.getByRole('button', { name: /^Cloud workspace:/ }).click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();

  await expect.poll(() => diagramBUpdates, { timeout: 5_000 }).toBe(1);
  await expect(modal).toBeVisible();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal).toContainText(
    'The document was modified by another request. Reload and retry.',
  );
});

test('snapshot restore keeps a replacement current diagram selected', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let diagramAUpdates = 0;
  let replacementCreated = 0;
  let snapshotFetches = 0;
  const snapshotVersion = {
    ...cloudVersion('B', 'replacement-race-snapshot', 'Replacement race restore'),
    payload: cloudDocument('B', 'Diagram B').payload,
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (path === '/api/diagrams/C/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (/^\/api\/diagrams\/[ABC]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/B/versions/replacement-race-snapshot' && method === 'GET') {
      snapshotFetches += 1;
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      diagramAUpdates += 1;
      if (diagramAUpdates > 1) {
        await fulfillJson(route, { error: 'Document not found.' }, 404);
        return;
      }
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      replacementCreated += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('C', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 1,
        etag: '"C-1"',
      }, 201, { etag: '"C-1"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  await page.getByRole('button', { name: /^Cloud workspace:/ }).click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();

  await expect.poll(() => replacementCreated, { timeout: 5_000 }).toBe(1);
  await expect(modal).toBeVisible();
  await expect(modal.locator('.cloud-document-list button')).toHaveCount(3);
  await expect(modal.locator('.cloud-document-list button').first()).toHaveClass(/selected/);
  await expect(modal.getByRole('heading', { name: 'Diagram A' })).toBeVisible();
  await expect(modal.getByRole('alert')).toContainText(
    'The original cloud diagram was deleted. Your work was saved as a replacement',
  );
  expect(snapshotFetches).toBe(0);
});

test('viewer snapshot restore uses a detached safety copy', async ({ page }) => {
  const shareToken = 'w'.repeat(43);
  await initializePage(page, {
    documentId: 'A',
    access: 'shared',
    role: 'viewer',
    shareToken,
  });
  let safetyCopies = 0;
  let snapshotFetches = 0;
  const snapshotVersion = {
    ...cloudVersion('A', 'viewer-snapshot', 'Viewer restore target'),
    payload: {
      ...cloudDocument('A', 'Shared diagram').payload,
      nodes: [{
        id: 'viewer-snapshot-node',
        type: 'azureNode',
        position: { x: 360, y: 220 },
        data: { label: 'Viewer Snapshot', serviceName: 'App Service' },
      }],
      edges: [],
    },
  };
  const sharedDocument = {
    ...cloudDocument('A', 'Shared diagram'),
    access: 'shared',
    role: 'viewer',
    etag: '"A-1"',
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'viewer@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}` && method === 'GET') {
      await fulfillJson(route, sharedDocument, 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [] });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}/versions` && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (
      path === `/api/diagrams/shared/${shareToken}/versions/viewer-snapshot`
      && method === 'GET'
    ) {
      snapshotFetches += 1;
      await wait(500);
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      safetyCopies += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('viewer-copy', body.diagramName || 'Shared diagram'),
        payload: body.payload,
        etag: '"viewer-copy-1"',
      }, 201, { etag: '"viewer-copy-1"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect.poll(() => safetyCopies).toBe(1);
  await expect.poll(() => snapshotFetches).toBe(1);
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-viewer-snapshot-node"]')).toBeVisible();
  await expect(cloudButton).toHaveAccessibleName('Cloud workspace: Cloud read-only');
});

test('discarding before snapshot restore cancels the failed save retry', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let updateAttempts = 0;
  let versionFetchStarted = 0;
  let releaseVersion = () => {};
  const versionGate = new Promise<void>((resolve) => {
    releaseVersion = resolve;
  });
  const snapshotVersion = cloudVersion('A', 'discard-retry-snapshot', 'Restore target');

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [snapshotVersion] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts === 2) {
        await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
        return;
      }
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/versions/discard-retry-snapshot' && method === 'GET') {
      versionFetchStarted += 1;
      await versionGate;
      await fulfillJson(route, { version: snapshotVersion });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'POST') {
      await fulfillJson(route, {
        version: cloudVersion('A', 'discard-retry-backup', 'Automatic backup'),
      }, 201);
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  page.on('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect.poll(() => updateAttempts).toBe(2);
  await expect.poll(() => versionFetchStarted).toBe(1);
  await page.clock.fastForward(16_000);
  await page.waitForTimeout(100);
  expect(updateAttempts).toBe(2);
  releaseVersion();
  await expect(modal).toBeHidden();
  expect(updateAttempts).toBe(3);
});

test('opening another cloud diagram requires saving or explicit discard', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let failedSaves = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      failedSaves += 1;
      await fulfillJson(route, { error: 'Temporary storage outage' }, 503);
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/versions$/.test(path) && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const nodeA = page.locator('[data-testid="rf__node-A-node"]');
  await expect(nodeA).toBeVisible();
  await nodeA.locator('[data-node-keyboard-target]').focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => failedSaves, { timeout: 5_000 }).toBeGreaterThan(0);

  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();

  let discardConfirmations = 0;
  page.on('dialog', async (dialog) => {
    discardConfirmations += 1;
    if (discardConfirmations === 1) await dialog.dismiss();
    else await dialog.accept();
  });
  await modal.getByRole('button', { name: 'Open', exact: true }).click();
  await expect.poll(() => discardConfirmations).toBe(1);
  await expect(modal).toBeVisible();
  await expect(nodeA).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-B-node"]')).toHaveCount(0);

  await modal.getByRole('button', { name: 'Open', exact: true }).click();
  await expect.poll(() => discardConfirmations).toBe(2);
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-B-node"]')).toBeVisible();
  await expect(nodeA).toHaveCount(0);
});

test('opening another diagram verifies unchanged cloud state before discard', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let diagramAUpdates = 0;
  let replacementCopies = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [summary('A', 'Diagram A'), summary('B', 'Diagram B')],
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/B' && method === 'GET') {
      await fulfillJson(route, cloudDocument('B', 'Diagram B'), 200, { etag: '"B-1"' });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/versions$/.test(path) && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (/^\/api\/diagrams\/[AB]\/shares$/.test(path) && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      diagramAUpdates += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (diagramAUpdates === 1) {
        await fulfillJson(route, {
          ...cloudDocument('A', body.diagramName || 'Diagram A'),
          payload: body.payload,
          revision: 2,
          etag: '"A-2"',
        }, 200, { etag: '"A-2"' });
        return;
      }
      await fulfillJson(route, { error: 'Not found' }, 404);
      return;
    }
    if (path === '/api/diagrams/B' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('B', body.diagramName || 'Diagram B'),
        payload: body.payload,
        revision: 2,
        etag: '"B-2"',
      }, 200, { etag: '"B-2"' });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      replacementCopies += 1;
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('replacement', body.diagramName || 'Diagram A'),
        payload: body.payload,
        etag: '"replacement-1"',
      }, 201, { etag: '"replacement-1"' });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.locator('.cloud-document-list button').filter({ hasText: 'Diagram B' }).click();
  await expect(modal.getByRole('heading', { name: 'Diagram B' })).toBeVisible();
  await modal.getByRole('button', { name: 'Open', exact: true }).click();
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-B-node"]')).toBeVisible();
  expect(diagramAUpdates).toBe(2);
  expect(replacementCopies).toBe(1);
});

test('discarding a current conflict reloads remote instead of cached content', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let serveRemoteDocument = false;
  let remoteLoads = 0;
  let updateAttempts = 0;
  const remoteDocument = {
    ...cloudDocument('A', 'Diagram A'),
    payload: {
      ...cloudDocument('A', 'Diagram A').payload,
      nodes: [{
        id: 'remote-node',
        type: 'azureNode',
        position: { x: 420, y: 220 },
        data: { label: 'Remote Node', serviceName: 'App Service' },
      }],
      edges: [],
    },
    revision: 3,
    etag: '"A-3"',
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      const result = serveRemoteDocument ? remoteDocument : cloudDocument('A', 'Diagram A');
      if (serveRemoteDocument) remoteLoads += 1;
      await fulfillJson(route, result, 200, {
        etag: serveRemoteDocument ? '"A-3"' : '"A-1"',
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      updateAttempts += 1;
      const body = JSON.parse(request.postData() || '{}');
      if (updateAttempts === 2) {
        serveRemoteDocument = true;
        await fulfillJson(route, {
          error: 'The document was modified by another request. Reload and retry.',
        }, 412);
        return;
      }
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: updateAttempts + 1,
        etag: `"A-${updateAttempts + 1}"`,
      }, 200, { etag: `"A-${updateAttempts + 1}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Open', exact: true }).click();
  await expect(modal).toBeHidden();
  await expect.poll(() => remoteLoads).toBeGreaterThan(0);
  await expect(page.locator('[data-testid="rf__node-remote-node"]')).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-A-node"]')).toHaveCount(0);
});

test('access management ignores a stale load from a previous open cycle', async ({ page }) => {
  await initializePage(page);
  let accessListCalls = 0;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: true,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: true,
        allowed: true,
      });
      return;
    }
    if (path === '/api/access/users' && method === 'GET') {
      accessListCalls += 1;
      if (accessListCalls === 1) await wait(500);
      else await wait(20);
      await fulfillJson(route, {
        users: [{
          email: accessListCalls === 1 ? 'old@example.com' : 'new@example.com',
          addedAt: now,
          addedBy: 'owner',
          isAdmin: false,
          immutable: false,
        }],
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [] });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await page.getByRole('button', { name: 'Access', exact: true }).click();
  const modal = page.locator('.access-modal');
  await expect(modal).toBeVisible();
  await modal.getByRole('button', { name: 'Close' }).click();
  await page.getByRole('button', { name: 'Access', exact: true }).click();
  await expect(modal.getByText('new@example.com')).toBeVisible();
  await page.waitForTimeout(650);
  await expect(modal.getByText('new@example.com')).toBeVisible();
  await expect(modal.getByText('old@example.com')).toHaveCount(0);
});

test('snapshot modal cannot close while its cloud target is being saved', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let snapshotNotes = '';

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Diagram A'),
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      }, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'POST') {
      snapshotNotes = JSON.parse(request.postData() || '{}').notes || '';
      await wait(500);
      await fulfillJson(route, {
        version: cloudVersion('A', 'snapshot-a', snapshotNotes),
      });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect(page.getByRole('button', { name: /^Cloud workspace:/ })).toHaveClass(/btn-active/);
  await page.getByRole('button', { name: 'Snapshot' }).click();
  const modal = page.locator('.save-snapshot-modal');
  await modal.getByLabel('Notes (optional) Describe what makes this version special').fill('Before change');
  await modal.getByRole('button', { name: 'Save Snapshot' }).click();
  await page.keyboard.press('Escape');
  await expect(modal).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Close' })).toBeDisabled();
  await expect(modal).toBeHidden({ timeout: 5_000 });
  expect(snapshotNotes).toBe('Before change');
});

test('diagram imports are atomic and AI imports save pricing to a new cloud document', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  const dialogMessages: string[] = [];
  const importedPayloads: Record<string, any>[] = [];
  let sourceUpdateAttempts = 0;
  let importedRevision = 1;

  page.on('dialog', async (dialog) => {
    dialogMessages.push(dialog.message());
    await dialog.accept();
  });

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Source diagram'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      sourceUpdateAttempts += 1;
      await fulfillJson(route, { error: 'The source diagram must not be updated.' }, 500);
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      const body = JSON.parse(request.postData() || '{}');
      importedPayloads.push(body.payload || {});
      await fulfillJson(route, {
        ...cloudDocument('IMPORTED', body.diagramName || 'Imported AI diagram'),
        diagramName: body.diagramName || 'Imported AI diagram',
        payload: body.payload,
        revision: importedRevision,
        etag: `"IMPORTED-${importedRevision}"`,
      }, 201, { etag: `"IMPORTED-${importedRevision}"` });
      return;
    }
    if (path === '/api/diagrams/IMPORTED' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      importedRevision += 1;
      importedPayloads.push(body.payload || {});
      await fulfillJson(route, {
        ...cloudDocument('IMPORTED', body.diagramName || 'Imported AI diagram'),
        diagramName: body.diagramName || 'Imported AI diagram',
        payload: body.payload,
        revision: importedRevision,
        etag: `"IMPORTED-${importedRevision}"`,
      }, 200, { etag: `"IMPORTED-${importedRevision}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const sourceNode = page.locator('[data-testid="rf__node-A-node"]');
  const cloudButton = page.getByRole('button', { name: /^Cloud workspace:/ });
  const fileInput = page.locator('input[accept=".json"]');
  await expect(sourceNode).toBeVisible();
  await expect(cloudButton).toHaveClass(/btn-active/);

  await fileInput.setInputFiles({
    name: 'invalid-diagram.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [],
      edges: 'invalid',
    })),
  });
  await expect.poll(() => dialogMessages.length).toBe(1);
  await expect(sourceNode).toBeVisible();
  await expect(cloudButton).toHaveClass(/btn-active/);
  expect(sourceUpdateAttempts).toBe(0);

  await fileInput.setInputFiles({
    name: 'invalid-ai-diagram.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      format: 'azurediagarm-ai-architecture',
      services: [],
      connections: [],
      groups: [],
    })),
  });
  await expect.poll(() => dialogMessages.length).toBe(2);
  await expect(sourceNode).toBeVisible();
  await expect(cloudButton).toHaveClass(/btn-active/);
  expect(sourceUpdateAttempts).toBe(0);

  await fileInput.setInputFiles({
    name: 'valid-ai-diagram.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      format: 'azurediagarm-ai-architecture',
      architectureName: 'Imported AI diagram',
      metadata: { prompt: 'Import an App Service architecture' },
      services: [{
        id: 'imported-app',
        name: 'App Service',
        type: 'App Service',
        category: 'app services',
      }],
      connections: [],
      groups: [],
      workflow: [],
    })),
  });

  await expect(page.locator('[data-testid="rf__node-imported-app"]')).toBeVisible({
    timeout: 10_000,
  });
  await expect(sourceNode).toHaveCount(0);
  await expect.poll(() => importedPayloads.length, { timeout: 5_000 }).toBeGreaterThan(0);
  await expect.poll(
    () => importedPayloads.some(payload => payload.nodes?.some(
      (node: any) => node.id === 'imported-app'
        && typeof node.data?.pricing?.estimatedCost === 'number'
        && node.data?.pricing?.region === 'japaneast',
    )),
    { timeout: 10_000 },
  ).toBe(true);
  expect(sourceUpdateAttempts).toBe(0);
});

test('service and layer ungroup actions preserve absolute positions', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  const document = interactionCloudDocument();
  document.payload.nodes = [
    {
      id: 'group-a',
      type: 'groupNode',
      position: { x: 320, y: 80 },
      style: { width: 420, height: 260 },
      data: { label: 'Application layer', color: '#0078d4' },
    },
    {
      id: 'node-a',
      type: 'azureNode',
      position: { x: 20, y: 60 },
      parentNode: 'group-a',
      extent: 'parent',
      data: { label: 'App Service', serviceName: 'App Service' },
    },
    {
      id: 'node-b',
      type: 'azureNode',
      position: { x: 220, y: 60 },
      parentNode: 'group-a',
      extent: 'parent',
      data: { label: 'Azure SQL Database', serviceName: 'Azure SQL Database' },
    },
  ];
  document.payload.edges = [];
  const savedPayloads: Record<string, any>[] = [];
  let revision = 1;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, document, 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      revision += 1;
      savedPayloads.push(body.payload || {});
      await fulfillJson(route, {
        ...document,
        diagramName: body.diagramName || document.diagramName,
        payload: body.payload,
        revision,
        etag: `"A-${revision}"`,
      }, 200, { etag: `"A-${revision}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const nodeA = page.locator('[data-testid="rf__node-node-a"]');
  const nodeB = page.locator('[data-testid="rf__node-node-b"]');
  const group = page.locator('[data-testid="rf__node-group-a"]');
  await expect(nodeA).toBeVisible();
  await expect(nodeB).toBeVisible();
  await expect(group).toBeVisible();

  await nodeA.click({ button: 'right' });
  await page.getByRole('menu', { name: 'Service actions' })
    .getByRole('menuitem', { name: 'Remove from layer' })
    .click();
  await expect.poll(() => savedPayloads.length, { timeout: 5_000 }).toBe(1);
  const detachedNode = savedPayloads[0].nodes.find((node: any) => node.id === 'node-a');
  expect(detachedNode.parentNode).toBeUndefined();
  expect(detachedNode.extent).toBeUndefined();
  expect(detachedNode.position).toEqual({ x: 340, y: 140 });

  await group.locator('.group-node-header').click({ button: 'right' });
  await page.getByRole('menu', { name: 'Layer actions' })
    .getByRole('menuitem', { name: /Ungroup layer/ })
    .click();
  await expect(group).toHaveCount(0);
  await expect(nodeA).toBeVisible();
  await expect(nodeB).toBeVisible();
  await expect.poll(() => savedPayloads.length, { timeout: 5_000 }).toBe(2);
  const finalPayload = savedPayloads[1];
  expect(finalPayload.nodes.some((node: any) => node.id === 'group-a')).toBe(false);
  const ungroupedNode = finalPayload.nodes.find((node: any) => node.id === 'node-b');
  expect(ungroupedNode.parentNode).toBeUndefined();
  expect(ungroupedNode.extent).toBeUndefined();
  expect(ungroupedNode.position).toEqual({ x: 540, y: 140 });
});

test('deployment guide accordions are keyboard accessible and stale results are discarded', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let guideRequests = 0;
  let importedRevision = 1;
  const guide = {
    title: 'Deploy the test architecture',
    overview: 'Deploy one App Service.',
    prerequisites: ['Azure CLI'],
    estimatedTime: '10 minutes',
    deploymentSteps: [{
      step: 1,
      title: 'Prepare resources',
      description: 'Select the subscription and resource group.',
      commands: ['az account show'],
      notes: ['Use the intended subscription.'],
    }],
    configuration: [],
    postDeployment: ['Verify the application endpoint'],
    troubleshooting: [],
    estimatedCost: '$10/month',
    bicepTemplates: [{
      name: 'Main',
      description: 'Deploys the application.',
      filename: 'main.bicep',
      content: 'param location string',
    }],
  };

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Guide source'), 200, { etag: '"A-1"' });
      return;
    }
    if (path === '/api/docs-search' && method === 'POST') {
      await fulfillJson(route, { results: [] });
      return;
    }
    if (path === '/api/openai' && method === 'POST') {
      guideRequests += 1;
      if (guideRequests === 2) await wait(700);
      await fulfillJson(route, {
        model: 'playwright-gpt-5-6-sol',
        output_text: JSON.stringify(guide),
        usage: { input_tokens: 30, output_tokens: 20, total_tokens: 50 },
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'POST') {
      const body = JSON.parse(request.postData() || '{}');
      await fulfillJson(route, {
        ...cloudDocument('IMPORTED', body.diagramName || 'Replacement diagram'),
        diagramName: body.diagramName || 'Replacement diagram',
        payload: body.payload,
        revision: importedRevision,
        etag: `"IMPORTED-${importedRevision}"`,
      }, 201, { etag: `"IMPORTED-${importedRevision}"` });
      return;
    }
    if (path === '/api/diagrams/IMPORTED' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      importedRevision += 1;
      await fulfillJson(route, {
        ...cloudDocument('IMPORTED', body.diagramName || 'Replacement diagram'),
        diagramName: body.diagramName || 'Replacement diagram',
        payload: body.payload,
        revision: importedRevision,
        etag: `"IMPORTED-${importedRevision}"`,
      }, 200, { etag: `"IMPORTED-${importedRevision}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const generateButton = page.getByTitle('Generate comprehensive deployment guide');
  await generateButton.click();
  const modal = page.locator('.deployment-modal');
  await expect(modal.getByText('Deploy the test architecture')).toBeVisible({
    timeout: 5_000,
  });

  const stepToggle = modal.locator('.step-header');
  await expect(stepToggle).toHaveAttribute('aria-expanded', 'true');
  await stepToggle.focus();
  await page.keyboard.press('Enter');
  await expect(stepToggle).toHaveAttribute('aria-expanded', 'false');
  await expect(modal.locator('#deployment-step-0')).toHaveCount(0);
  await stepToggle.press('Enter');
  await expect(stepToggle).toHaveAttribute('aria-expanded', 'true');
  await expect(modal.locator('#deployment-step-0')).toBeVisible();

  const bicepToggle = modal.locator('.bicep-template-toggle');
  await expect(bicepToggle).toHaveAttribute('aria-expanded', 'true');
  await bicepToggle.focus();
  await page.keyboard.press('Enter');
  await expect(bicepToggle).toHaveAttribute('aria-expanded', 'false');
  await bicepToggle.press('Enter');
  await expect(bicepToggle).toHaveAttribute('aria-expanded', 'true');
  await expect(modal.getByRole('button', { name: 'Download main.bicep' })).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Copy to clipboard' })).toHaveCount(2);
  await expectNoWcagViolations(page, '.deployment-modal');

  await modal.getByRole('button', { name: 'Close' }).last().click();
  await generateButton.click();
  await expect(modal.getByText('Generating comprehensive deployment guide...')).toBeVisible();
  await page.locator('input[accept=".json"]').setInputFiles({
    name: 'replacement-diagram.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify({
      nodes: [{
        id: 'replacement-node',
        type: 'azureNode',
        position: { x: 200, y: 180 },
        data: {
          label: 'Azure SQL Database',
          serviceName: 'Azure SQL Database',
        },
      }],
      edges: [],
      titleBlockData: { architectureName: 'Replacement diagram' },
    })),
  });
  await expect(page.locator('[data-testid="rf__node-replacement-node"]')).toBeVisible();
  await expect(modal).toBeHidden();
  await page.waitForTimeout(900);
  await expect(page.getByTitle('Open last deployment guide')).toHaveCount(0);
});
