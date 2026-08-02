import { expect, test, type Page, type Route } from '@playwright/test';

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
  await page.getByRole('button', { name: /Access$/ }).click();
  const modal = page.locator('.access-modal');
  await expect(modal).toBeVisible();
  await modal.getByRole('button', { name: 'Close' }).click();
  await page.getByRole('button', { name: /Access$/ }).click();
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
