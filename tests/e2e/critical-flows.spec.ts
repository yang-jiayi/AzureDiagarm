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
  await openInteractionDiagram(page, (payload) => {
    const savedNodes = Array.isArray(payload.nodes)
      ? payload.nodes as Array<{ id?: string; data?: { label?: string } }>
      : [];
    const savedNode = savedNodes.find(node => node.id === 'node-a');
    const savedGroup = savedNodes.find(node => node.id === 'group-a');
    lastSavedLabel = String(savedNode?.data?.label || '');
    lastSavedGroupLabel = String(savedGroup?.data?.label || '');
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
