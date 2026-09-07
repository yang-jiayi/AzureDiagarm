import { expect, test, type Locator, type Page, type Route } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { PRICING_DATA_AS_OF } from '../../src/data/azurePricing';
import type { CloudDiagramDocument } from '../../src/services/cloudDiagramService';

const now = '2026-08-02T00:00:00.000Z';
const wait = (milliseconds: number) => new Promise(resolve => setTimeout(resolve, milliseconds));
const visualSnapshotStylePath = fileURLToPath(new URL('./visual-snapshot.css', import.meta.url));

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
  options?: { showCanvasHint?: boolean; focusMode?: boolean },
) {
  await page.addInitScript(({ context, showCanvasHint, focusMode }) => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
    localStorage.setItem('azure-diagram-builder.ribbonTab.v1', 'review');
    localStorage.setItem('azure-diagram-builder.headerCollapsed.v1', '0');
    if (!sessionStorage.getItem('playwright.focus-mode-seeded')) {
      localStorage.setItem('azure-diagram-builder.focusMode.v1', focusMode ? '1' : '0');
      sessionStorage.setItem('playwright.focus-mode-seeded', '1');
    }
    if (showCanvasHint) {
      localStorage.removeItem('azure-diagram-builder.canvasHintDismissed.v1');
    } else {
      localStorage.setItem('azure-diagram-builder.canvasHintDismissed.v1', '1');
    }
    if (context) {
      sessionStorage.setItem('azurediagarm.cloud-document.v1', JSON.stringify(context));
    } else {
      sessionStorage.removeItem('azurediagarm.cloud-document.v1');
    }
  }, {
    context: cloudContext,
    showCanvasHint: options?.showCanvasHint === true,
    focusMode: options?.focusMode === true,
  });
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

async function expectReadableContrast(locator: Locator, minimumRatio = 4.5) {
  await expect(locator).toBeVisible();
  const readContrast = () => locator.evaluate((element) => {
    type Rgba = [number, number, number, number];

    const parseColor = (value: string): Rgba | null => {
      const match = value.match(
        /rgba?\((\d+(?:\.\d+)?)[, ]+(\d+(?:\.\d+)?)[, ]+(\d+(?:\.\d+)?)(?:[, /]+(\d+(?:\.\d+)?))?\)/,
      );
      if (!match) return null;
      return [
        Number(match[1]),
        Number(match[2]),
        Number(match[3]),
        match[4] === undefined ? 1 : Number(match[4]),
      ];
    };

    const blend = (foreground: Rgba, background: Rgba): Rgba => {
      const alpha = foreground[3] + background[3] * (1 - foreground[3]);
      if (alpha === 0) return [0, 0, 0, 0];
      return [
        (foreground[0] * foreground[3]
          + background[0] * background[3] * (1 - foreground[3])) / alpha,
        (foreground[1] * foreground[3]
          + background[1] * background[3] * (1 - foreground[3])) / alpha,
        (foreground[2] * foreground[3]
          + background[2] * background[3] * (1 - foreground[3])) / alpha,
        alpha,
      ];
    };

    const luminance = ([red, green, blue]: Rgba) => {
      const channels = [red, green, blue].map((value) => {
        const channel = value / 255;
        return channel <= 0.03928
          ? channel / 12.92
          : ((channel + 0.055) / 1.055) ** 2.4;
      });
      return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
    };

    const surfaces: Rgba[] = [];
    for (let current: Element | null = element; current; current = current.parentElement) {
      const background = parseColor(getComputedStyle(current).backgroundColor);
      if (background && background[3] > 0) surfaces.push(background);
    }

    let background: Rgba = [255, 255, 255, 1];
    surfaces.reverse().forEach((surface) => {
      background = blend(surface, background);
    });

    const style = getComputedStyle(element);
    const parsedForeground = parseColor(style.color);
    if (!parsedForeground) {
      return {
        ratio: 0,
        color: style.color,
        background: style.backgroundColor,
        opacity: style.opacity,
        label: element.textContent?.trim() || element.getAttribute('aria-label') || element.tagName,
      };
    }

    const foreground = blend(parsedForeground, background);
    const foregroundLuminance = luminance(foreground);
    const backgroundLuminance = luminance(background);
    const ratio = (Math.max(foregroundLuminance, backgroundLuminance) + 0.05)
      / (Math.min(foregroundLuminance, backgroundLuminance) + 0.05);

    return {
      ratio,
      color: style.color,
      background: style.backgroundColor,
      opacity: style.opacity,
      label: element.textContent?.trim() || element.getAttribute('aria-label') || element.tagName,
    };
  });

  let result = await readContrast();
  await expect.poll(async () => {
    result = await readContrast();
    return result.ratio;
  }, {
    message: `${result.label}: ${result.color} on ${result.background} (opacity ${result.opacity})`,
    timeout: 1_000,
  }).toBeGreaterThanOrEqual(minimumRatio);
}

async function expectSelectedMenuState(locator: Locator) {
  await expect(locator).toHaveCSS('background-color', 'rgb(15, 108, 189)');
  await expect(locator).toHaveCSS('color', 'rgb(255, 255, 255)');
  await expectReadableContrast(locator);
}

async function expectNotToOverlap(first: Locator, second: Locator) {
  await expect(first).toBeVisible();
  await expect(second).toBeVisible();
  await expect.poll(async () => {
    const [firstBox, secondBox] = await Promise.all([first.boundingBox(), second.boundingBox()]);
    if (!firstBox || !secondBox) return true;

    return firstBox.x < secondBox.x + secondBox.width
      && firstBox.x + firstBox.width > secondBox.x
      && firstBox.y < secondBox.y + secondBox.height
      && firstBox.y + firstBox.height > secondBox.y;
  }).toBe(false);
}

async function openInteractionDiagram(
  page: Page,
  onSave?: (payload: Record<string, unknown>) => void,
  showCanvasHint = false,
  documentOverride?: ReturnType<typeof interactionCloudDocument>,
) {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  }, { showCanvasHint });
  const document = documentOverride ?? interactionCloudDocument();

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
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  const modal = page.locator('.ai-architecture-modal');
  await expect(modal).toBeFocused();
  return modal;
}

function getCloudWorkspaceButton(page: Page) {
  return page.getByRole('button', {
    name: /^Cloud workspace:/,
    includeHidden: true,
  });
}

async function editDiagramAuthor(page: Page, author = 'Cloud regression setup') {
  await expect(page.getByTestId('rf__node-A-node')).toBeVisible();
  const titleBlock = page.locator('.title-block');
  await expect(titleBlock).toBeVisible();
  const toggle = titleBlock.locator('.title-block-toggle');
  if (await toggle.getAttribute('aria-expanded') === 'false') await toggle.click();
  await titleBlock.locator('.title-block-display').press('Enter');
  const authorField = titleBlock.getByLabel('Author:', { exact: true });
  await expect(authorField).not.toHaveValue(author);
  await authorField.fill(author);
  await titleBlock.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(titleBlock.locator('.title-block-display')).toContainText(author);
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
  await expect(page.locator('.react-flow__controls')).toHaveCount(0);
  await expect(page.locator('.legend')).toHaveCount(0);
  await expect(page.getByRole('button', { name: 'Feedback' })).toHaveCount(0);
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

test('More options restores Escape focus without stealing outside focus', async ({ page }) => {
  await initializePage(page);
  await page.route('**/api/**', async route => {
    const access = new URL(route.request().url()).pathname === '/api/access/me';
    await fulfillJson(route, access
      ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
      : { error: 'Not found' }, access ? 200 : 404);
  });
  await page.goto('/');

  const trigger = page.getByRole('button', { name: 'More', exact: true });
  const popover = page.getByRole('dialog', { name: 'More application options' });
  await trigger.click();
  await popover.getByRole('button', { name: 'Resume recent work' }).focus();
  await page.keyboard.press('Escape');
  await expect(popover).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(trigger).toHaveAttribute('aria-expanded', 'false');

  await trigger.click();
  const outside = page.locator('.search-box input');
  await outside.focus();
  await page.keyboard.press('Escape');
  await expect(popover).toBeHidden();
  await expect(outside).toBeFocused();
});

for (const touch of [false, true]) {
  test.describe(`palette targets ${touch ? 'touch' : 'desktop'}`, () => {
    test.use({
      hasTouch: touch,
      viewport: touch ? { width: 390, height: 844 } : { width: 1600, height: 900 },
    });

    test('folds controls upward while retaining usable icons and view preferences', async ({ page }, testInfo) => {
      await initializePage(page);
      await page.route('**/api/**', async route => {
        const access = new URL(route.request().url()).pathname === '/api/access/me';
        await fulfillJson(route, access
          ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
          : { error: 'Not found' }, access ? 200 : 404);
      });
      await page.goto('/');
      const openPalette = async () => {
        if (touch) {
          await page.getByRole('navigation', { name: 'Mobile command bar' })
            .getByRole('button', { name: 'Services' }).click();
        } else {
          await page.getByRole('button', { name: 'Open services panel' }).click();
        }
      };
      if (touch) await openPalette();
      const palette = page.locator('.icon-palette');
      const toggle = palette.locator('.palette-controls-toggle');
      const search = palette.locator('.search-box input');
      await expect(palette.locator('.icon-image').first()).toBeVisible();
      await expect(toggle).toHaveAccessibleName('Microsoft Services: Collapse icon controls');
      await expect(toggle).toHaveAttribute('aria-expanded', 'true');
      const toggleBounds = await toggle.boundingBox();
      expect(toggleBounds?.height).toBeGreaterThanOrEqual(44);

      const measureIconSpace = () => palette.evaluate(element => {
        const content = element.querySelector('.palette-content');
        const viewport = element.querySelector('.virtualized-icons-viewport');
        if (!content || !viewport) throw new Error('Expected the populated icon list');
        const contentBounds = content.getBoundingClientRect();
        const viewportBounds = viewport.getBoundingClientRect();
        const top = Math.max(contentBounds.top, viewportBounds.top, 0);
        const bottom = Math.min(contentBounds.bottom, viewportBounds.bottom, window.innerHeight);
        const fullIcons = [...viewport.querySelectorAll('.icon-item')].filter(icon => {
          const bounds = icon.getBoundingClientRect();
          return bounds.top >= top && bounds.bottom <= bottom;
        }).length;
        return { height: contentBounds.height, visibleHeight: bottom - top, fullIcons };
      });
      const expanded = await measureIconSpace();
      await testInfo.attach('expanded-palette', {
        body: await palette.screenshot(), contentType: 'image/png',
      });
      await toggle.focus();
      await page.keyboard.press('Enter');
      await expect(toggle).toBeFocused();
      await expect(toggle).toHaveAttribute('aria-expanded', 'false');
      await expect(search).toBeHidden();
      await expect(palette.getByRole('tablist')).toHaveCount(0);
      await expect(palette.locator('.palette-recommended-intro')).toHaveCount(0);
      await expect(palette.locator('#palette-view-panel')).toHaveAttribute(
        'aria-labelledby', 'palette-collapsed-summary',
      );
      const focusClearance = await toggle.evaluate(element => {
        const summary = document.getElementById('palette-collapsed-summary');
        if (!summary) throw new Error('Expected the visible palette summary');
        const style = getComputedStyle(element);
        return summary.getBoundingClientRect().top - element.getBoundingClientRect().bottom
          - Number.parseFloat(style.outlineWidth) - Number.parseFloat(style.outlineOffset);
      });
      expect(focusClearance, 'the keyboard focus ring must not cover the view summary').toBeGreaterThanOrEqual(0);
      await expect.poll(async () => (await measureIconSpace()).fullIcons)
        .toBeGreaterThan(expanded.fullIcons);
      const folded = await measureIconSpace();
      expect(folded.height - expanded.height).toBeGreaterThan(160);
      expect(folded.visibleHeight - expanded.visibleHeight).toBeGreaterThan(100);
      expect(folded.fullIcons).toBeGreaterThan(expanded.fullIcons);
      await testInfo.attach('palette-space', {
        body: JSON.stringify({ expanded, folded }), contentType: 'application/json',
      });
      await testInfo.attach('folded-palette', {
        body: await palette.screenshot(), contentType: 'image/png',
      });
      await page.keyboard.press('Tab');
      await expect(palette.getByRole('button', { name: 'Close services panel' })).toBeFocused();
      await page.keyboard.press('Tab');
      await expect(palette.locator('.virtualized-icons-viewport')).toBeFocused();
      for (const dark of [false, true]) {
        await page.evaluate(enabled => document.body.classList.toggle('dark-mode', enabled), dark);
        await expectNoWcagViolations(page, '.icon-palette');
      }

      if (!touch) {
        await page.setViewportSize({ width: 1600, height: 1400 });
        await expect.poll(async () => (await measureIconSpace()).visibleHeight).toBeGreaterThan(750);
        const tallFolded = await measureIconSpace();
        await toggle.press('Space');
        await expect(toggle).toHaveAttribute('aria-expanded', 'true');
        const tallExpanded = await measureIconSpace();
        expect(tallFolded.visibleHeight - tallExpanded.visibleHeight).toBeGreaterThan(100);
        expect(tallFolded.fullIcons).toBeGreaterThan(tallExpanded.fullIcons);
        await testInfo.attach('tall-palette-space', {
          body: JSON.stringify({ expanded: tallExpanded, folded: tallFolded }),
          contentType: 'application/json',
        });
        await page.setViewportSize({ width: 1600, height: 900 });
      } else {
        await toggle.press('Space');
      }
      await expect(toggle).toHaveAttribute('aria-expanded', 'true');
      await palette.getByRole('tab', { name: /^All/ }).click();
      await search.fill('App Services');
      await palette.getByRole('button', { name: 'List view', exact: true }).click();
      await expect(palette.getByRole('button', { name: 'Add App Services to the canvas' })).toHaveCount(1);
      await toggle.click();
      await expect(palette.locator('.palette-collapsed-summary')).toContainText('All');
      await expect(palette.locator('.palette-collapsed-summary')).toContainText('App Services');
      await palette.getByRole('button', { name: 'Close services panel' }).click();
      await openPalette();
      await expect(toggle).toHaveAttribute('aria-expanded', 'false');
      await expect(search).toHaveValue('App Services');
      await toggle.click();
      await expect(palette.getByRole('tab', { name: /^All/ })).toHaveAttribute('aria-selected', 'true');
      await expect(palette.getByRole('button', { name: 'List view', exact: true }))
        .toHaveAttribute('aria-pressed', 'true');
      await expect(search).toHaveValue('App Services');
      await toggle.click();
      await expect.poll(() => page.evaluate(() => (
        localStorage.getItem('azure-diagram-builder.paletteControlsCollapsed.v1')
      ))).toBe('true');

      await page.reload();
      if (touch) await openPalette();
      await expect(toggle).toHaveAttribute('aria-expanded', 'false');
      await expect(palette).toHaveClass(/palette-layout-list/);
      await expect(palette.locator('.icon-image').first()).toBeVisible();
      await toggle.click();
      await expect(search).toBeVisible();
      await expect.poll(() => page.evaluate(() => (
        localStorage.getItem('azure-diagram-builder.paletteControlsCollapsed.v1')
      ))).toBe('false');
      await palette.locator('.icon-item-main').first().click();
      await expect(page.locator('.react-flow__node')).toHaveCount(1);
    });

    test('keeps the selected collection and Japanese disclosure usable at narrow widths', async ({ page }, testInfo) => {
      await initializePage(page);
      await page.setViewportSize({ width: 1280, height: 800 });
      await page.route('**/api/**', async route => {
        const access = new URL(route.request().url()).pathname === '/api/access/me';
        await fulfillJson(route, access
          ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
          : { error: 'Not found' }, access ? 200 : 404);
      });
      await page.goto('/');
      await page.getByRole('button', { name: 'More', exact: true }).click();
      await page.getByRole('button', { name: '日本語', exact: true }).click();
      await page.keyboard.press('Escape');
      if (touch) {
        await page.setViewportSize({ width: 320, height: 640 });
        await page.locator('.mobile-command-bar').getByRole('button', { name: 'サービス' }).click();
      }
      const palette = page.locator('.icon-palette');
      const toggle = palette.locator('.palette-controls-toggle');
      await expect(toggle).toHaveAccessibleName(/アイコンの操作エリアを折りたたむ$/);
      await palette.getByRole('tab', { name: /^コレクション/ }).click();
      for (const name of ['Backend', 'Data']) {
        page.once('dialog', dialog => dialog.accept(name));
        await palette.getByRole('button', { name: '新規', exact: true }).click();
      }
      const collection = palette.getByRole('combobox', { name: '選択中のコレクション' });
      await collection.selectOption({ label: 'Backend (0)' });
      const selected = await collection.inputValue();
      await toggle.click();
      await expect(toggle).toHaveAccessibleName(/アイコンの操作エリアを展開$/);
      await expect(collection).toHaveValue(selected);
      await expect(palette.locator('.palette-collapsed-summary')).toContainText('コレクション');
      expect(await palette.evaluate(element => element.scrollWidth <= element.clientWidth)).toBe(true);
      expect(await toggle.locator('span').evaluate(element => element.scrollWidth <= element.clientWidth))
        .toBe(true);
      await expectNoWcagViolations(page, '.icon-palette');
      await testInfo.attach('folded-palette-japanese', {
        body: await palette.screenshot(), contentType: 'image/png',
      });
      await toggle.click();
      await expect(palette.getByRole('tab', { name: /^コレクション/ }))
        .toHaveAttribute('aria-selected', 'true');
      await expect(collection).toHaveValue(selected);
    });

    test('separate insertion from organization with synchronized virtual rows', async ({ page }) => {
      await initializePage(page);
      await page.route('**/api/**', async route => {
        const access = new URL(route.request().url()).pathname === '/api/access/me';
        await fulfillJson(route, access
          ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
          : { error: 'Not found' }, access ? 200 : 404);
      });
      await page.goto('/');
      if (touch) {
        await page.getByRole('navigation', { name: 'Mobile command bar' })
          .getByRole('button', { name: 'Services' }).click();
      }
      await expect(page.locator('.icon-image').first()).toBeVisible();

      for (const layout of ['Grid view', 'List view']) {
        await page.getByRole('button', { name: layout, exact: true }).click();
        const viewport = page.locator('.virtualized-icons-viewport').first();
        const card = viewport.locator('.icon-item').first();
        const main = card.locator('.icon-item-main');
        const actions = card.locator('.icon-workspace-action');
        const mainBounds = await main.boundingBox();
        expect(mainBounds).not.toBeNull();

        for (const action of await actions.all()) {
          const bounds = await action.boundingBox();
          expect(bounds).not.toBeNull();
          expect(bounds!.width).toBeGreaterThanOrEqual(touch ? 44 : 24);
          expect(bounds!.height).toBeGreaterThanOrEqual(touch ? 44 : 24);
          const overlapWidth = Math.min(bounds!.x + bounds!.width, mainBounds!.x + mainBounds!.width)
            - Math.max(bounds!.x, mainBounds!.x);
          const overlapHeight = Math.min(bounds!.y + bounds!.height, mainBounds!.y + mainBounds!.height)
            - Math.max(bounds!.y, mainBounds!.y);
          expect(overlapWidth <= 0 || overlapHeight <= 0).toBe(true);
          expect(await action.evaluate(element => {
            const rect = element.getBoundingClientRect();
            return element.contains(document.elementFromPoint(
              rect.x + rect.width / 2, rect.y + rect.height / 2,
            ));
          })).toBe(true);
        }

        const metrics = await viewport.evaluate(element => {
          const window = element.querySelector<HTMLElement>('.virtualized-icons-window')!;
          const cards = [...window.querySelectorAll<HTMLElement>('.icon-item')];
          const columns = getComputedStyle(window).gridTemplateColumns.split(' ').length;
          return {
            rowHeight: Number.parseFloat(getComputedStyle(element).getPropertyValue('--palette-row-height')),
            gap: Number.parseFloat(getComputedStyle(window).rowGap),
            cardHeight: cards[0].getBoundingClientRect().height,
            rowSpacing: cards[columns].offsetTop - cards[0].offsetTop,
          };
        });
        expect(metrics.cardHeight + metrics.gap).toBe(metrics.rowHeight);
        expect(metrics.rowSpacing).toBe(metrics.rowHeight);

        const favorite = actions.first();
        const wasFavorite = await favorite.getAttribute('aria-pressed') === 'true';
        const favoriteBounds = (await favorite.boundingBox())!;
        if (touch) {
          await page.touchscreen.tap(
            favoriteBounds.x + favoriteBounds.width / 2, favoriteBounds.y + favoriteBounds.height / 2,
          );
        } else {
          await page.mouse.click(
            favoriteBounds.x + favoriteBounds.width / 2, favoriteBounds.y + favoriteBounds.height / 2,
          );
        }
        await expect(favorite).toHaveAttribute('aria-pressed', String(!wasFavorite));
        await expect(page.locator('.react-flow__node')).toHaveCount(0);

        const collectionBounds = (await actions.last().boundingBox())!;
        if (touch) {
          await page.touchscreen.tap(
            collectionBounds.x + collectionBounds.width / 2, collectionBounds.y + collectionBounds.height / 2,
          );
        } else {
          await page.mouse.click(
            collectionBounds.x + collectionBounds.width / 2, collectionBounds.y + collectionBounds.height / 2,
          );
        }
        const collection = page.locator('.collection-assignment-panel');
        await expect(collection).toBeVisible();
        await expect(page.locator('.react-flow__node')).toHaveCount(0);
        await collection.getByRole('button', { name: 'Close', exact: true }).click();

        const results = await new AxeBuilder({ page })
          .include('.virtualized-icons-viewport')
          .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa'])
          .analyze();
        expect(results.violations).toEqual([]);
      }

      const addService = page.locator('.icon-item-main').first();
      const addBounds = (await addService.boundingBox())!;
      if (touch) {
        await page.touchscreen.tap(addBounds.x + addBounds.width / 2, addBounds.y + addBounds.height / 2);
      } else {
        await page.mouse.click(addBounds.x + addBounds.width / 2, addBounds.y + addBounds.height / 2);
      }
      await expect(page.locator('.react-flow__node')).toHaveCount(1);
    });
  });
}

test('cloud empty and recovery guidance keeps readable theme contrast', async ({ page }) => {
  await initializePage(page);
  let failList = false;
  await page.route('**/api/**', async route => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false, authenticated: true, allowed: true, isAdmin: false,
        email: 'owner@example.com',
      });
      return;
    }
    if (path === '/api/diagrams' && route.request().method() === 'GET') {
      await fulfillJson(route, failList
        ? { error: 'Cloud diagrams are temporarily unavailable.' }
        : { documents: [] }, failList ? 503 : 200);
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });
  await page.goto('/');
  await page.locator('.document-status').click();
  const modal = page.getByRole('dialog', { name: 'Cloud workspace' });
  const refresh = modal.getByRole('button', { name: 'Refresh cloud diagrams' });

  for (const dark of [false, true]) {
    await page.evaluate(enabled => document.body.classList.toggle('dark-mode', enabled), dark);
    failList = false;
    await refresh.click();
    const guidance = modal.locator('.cloud-empty-state p');
    await expect(guidance).toHaveCount(2);
    for (const paragraph of await guidance.all()) await expectReadableContrast(paragraph);
    await expectNoWcagViolations(page, '.cloud-workspace-modal');

    failList = true;
    await refresh.click();
    const error = modal.getByRole('alert');
    await expect(error).toContainText('Cloud diagrams are temporarily unavailable.');
    await expectReadableContrast(error);
    await expectNoWcagViolations(page, '.cloud-workspace-modal');
  }
});

test('compact workspace chrome preserves canvas space and visible mobile commands', async ({ page }) => {
  await initializePage(page);
  await page.route('**/api/**', async route => {
    const access = new URL(route.request().url()).pathname === '/api/access/me';
    await fulfillJson(route, access
      ? { enabled: false, authenticated: false, isAdmin: false, allowed: true }
      : { error: 'Not found' }, access ? 200 : 404);
  });
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.goto('/');
  await page.getByRole('tab', { name: 'Home', exact: true }).click();

  for (const viewport of [
    { width: 1600, height: 900, maximumChrome: 190 },
    { width: 1366, height: 768, maximumChrome: 190 },
    { width: 768, height: 1024, maximumChrome: 235 },
    { width: 390, height: 844, maximumChrome: 210 },
  ]) {
    await page.setViewportSize({ width: viewport.width, height: viewport.height });
    await expect.poll(async () => (
      (await page.locator('.canvas-container').boundingBox())?.y ?? Infinity
    )).toBeLessThan(viewport.maximumChrome);
    expect(await page.evaluate(() => document.documentElement.scrollWidth))
      .toBeLessThanOrEqual(viewport.width);
    const targets = await page.locator('.ribbon-tab, .ribbon-command-strip .btn, .mobile-command-bar button, .workflow-stepper button')
      .evaluateAll(elements => elements
        .filter(element => element.getBoundingClientRect().width > 0)
        .map(element => element.getBoundingClientRect().height));
    expect(targets.length).toBeGreaterThan(0);
    expect(Math.min(...targets)).toBeGreaterThanOrEqual(44);
  }

  const mobileCommands = page.locator('.mobile-command-bar button').first();
  await mobileCommands.click();
  const sheet = page.getByRole('dialog', { name: 'Ribbon commands' });
  await expect(sheet).toBeVisible();
  // Inspect before locator.click can scroll clipped commands into view.
  const initialTabs = await sheet.getByRole('tab').evaluateAll(elements => elements.map(element => {
    const bounds = element.getBoundingClientRect();
    const hit = document.elementFromPoint(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
    return bounds.top >= 0 && bounds.bottom <= innerHeight
      && (hit === element || element.contains(hit));
  }));
  expect(initialTabs).toEqual([true, true, true, true]);
  await sheet.getByRole('tab', { name: 'Create', exact: true }).click();
  await expect(sheet.getByRole('button', { name: 'Add Group', exact: true })).toBeVisible();
  await expectNoWcagViolations(page, '.mobile-ribbon-drawer');
  await page.keyboard.press('Escape');
  await expect(mobileCommands).toBeFocused();

  await page.locator('.header-utility-button').click();
  await page.getByRole('button', { name: '日本語', exact: true }).click();
  await page.keyboard.press('Escape');
  await page.setViewportSize({ width: 768, height: 1024 });
  await page.getByRole('tab', { name: 'ホーム', exact: true }).click();
  await expect.poll(async () => (
    (await page.locator('.canvas-container').boundingBox())?.y ?? Infinity
  )).toBeLessThan(235);
});

for (const touch of [false, true]) {
  test.describe(`populated pricing toolbar ${touch ? 'touch' : 'desktop'}`, () => {
    test.use({ hasTouch: touch, viewport: { width: 1280, height: 720 } });

    test('retains compact dated pricing and primary commands in Japanese and English', async ({ page }) => {
      await initializePage(page);
      await page.clock.setFixedTime(new Date(
        Date.parse(`${PRICING_DATA_AS_OF}T00:00:00Z`) + 53 * 86_400_000,
      ));
      await page.route('**/api/**', async route => {
        const access = new URL(route.request().url()).pathname === '/api/access/me';
        await fulfillJson(route, access
          ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
          : { error: 'Not found' }, access ? 200 : 404);
      });
      await page.goto('/');
      await page.locator('.search-box input').fill('App Services');
      await page.getByRole('button', { name: /Add App Services to the canvas/i }).click();
      await page.locator('#ribbon-tab-create').click();
      await page.locator('.azure-node').click();
      await page.getByRole('button', { name: 'Service settings', exact: true }).click();
      const inspector = page.getByRole('dialog', { name: 'Service inspector', exact: true });
      await inspector.getByRole('spinbutton', { name: 'Quantity (instances / units)', exact: true }).fill('3');
      await inspector.getByRole('checkbox', { name: 'Use a custom monthly estimate', exact: true }).check();
      await inspector.getByRole('spinbutton', { name: 'USD per unit / month', exact: true }).fill('120');
      await inspector.getByRole('button', { name: 'Apply', exact: true }).click();
      await expect(inspector).toBeHidden();
      await expect(page.locator('.cost-badge')).toContainText('$360');
      await page.locator('#ribbon-tab-home').click();
      await page.locator('.region-selector-button').click();
      await page.locator('.region-option').filter({ hasText: 'Japan East' }).click();

      for (const language of ['ja', 'en']) {
        await page.locator('.header-utility-menu > button').click();
        await page.locator('.language-switch').getByRole('button', {
          name: language === 'ja' ? '日本語' : /^EN\b/, exact: true,
        }).click();
        await page.keyboard.press('Escape');
        for (const width of [900, 768]) {
          await page.setViewportSize({ width, height: 720 });
          await page.locator('#ribbon-tab-home').click();
          await expect(page.locator('.cost-indicator')).toContainText('$360.00/mo');
          await expect(page.locator('.pricing-scenario-launch')).toBeVisible();
          await expect(page.locator('.pricing-mode-btn')).toHaveCount(2);
          const stamp = page.locator('.pricing-freshness');
          const date = stamp.locator('.pricing-freshness-label');
          const age = stamp.locator('.pricing-freshness-age');
          await expect(date).toBeVisible();
          await expect(age).toContainText(language === 'ja' ? '53日前' : '53 days old');
          const [dateBounds, ageBounds, stampBounds, headerBounds] = await Promise.all(
            [date, age, stamp, page.locator('.app-header')].map(locator => locator.boundingBox()),
          );
          expect(ageBounds!.y).toBeGreaterThanOrEqual(dateBounds!.y + dateBounds!.height);
          expect(stampBounds!.height).toBeLessThanOrEqual(44);
          expect(headerBounds!.height).toBeLessThanOrEqual(220);
          expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(width);
          const controls = await page.locator(
            '.ribbon-command-strip .btn:visible, .ribbon-command-strip .region-selector-button:visible, .ribbon-command-strip .cost-visibility-toggle:visible',
          ).evaluateAll(elements => elements.map(element => {
            const bounds = element.getBoundingClientRect();
            return { width: bounds.width, height: bounds.height };
          }));
          expect(controls.length).toBeGreaterThanOrEqual(8);
          expect(controls.every(control => control.width >= 44 && control.height >= 44)).toBe(true);
          if (width === 768) await expectNoWcagViolations(page, '.pricing-freshness');
        }
      }
      await page.getByRole('button', { name: 'Switch to Dark Mode', exact: true }).click();
      await expectNoWcagViolations(page, '.pricing-freshness');
    });
  });
}

test('workflow sidebar follows workspace bounds and defaults closed on narrow screens', async ({ page }) => {
  const base = interactionCloudDocument();
  const cloudWithWorkflow = {
    ...base,
    payload: {
      ...base.payload,
      workflow: [{ step: 1, description: 'Application handles the request.', services: ['node-a'] }],
    },
  };
  await openInteractionDiagram(page, undefined, false, cloudWithWorkflow);
  const panel = page.locator('.workflow-panel');
  await expect(panel).toHaveClass(/expanded/);
  const toggle = panel.getByRole('button', { name: 'Collapse workflow', exact: true });
  await expect.poll(() => toggle.evaluate(element => {
    const bounds = element.getBoundingClientRect();
    const workspace = document.querySelector('.workspace')!.getBoundingClientRect();
    const hit = document.elementFromPoint(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2);
    return bounds.top >= workspace.top && (hit === element || element.contains(hit));
  })).toBe(true);
  const canvas = await page.locator('.canvas-container').boundingBox();
  const sidebar = await panel.boundingBox();
  expect(canvas!.x + canvas!.width).toBeLessThanOrEqual(sidebar!.x + 1);

  const step = panel.locator('.workflow-step').first();
  await step.press('Enter');
  await expect(step).toHaveAttribute('aria-pressed', 'true');
  await step.hover();
  await page.locator('.canvas-container').hover();
  await expect(step).toHaveAttribute('aria-pressed', 'true');
  await toggle.click();
  await expect(panel).toHaveClass(/collapsed/);
  await panel.getByRole('button', { name: 'Expand workflow', exact: true }).click();
  await expect(step).toHaveAttribute('aria-pressed', 'true');

  await page.setViewportSize({ width: 768, height: 1024 });
  await expect(panel).toHaveClass(/collapsed/);
  await expect.poll(async () => (
    (await page.locator('.canvas-container').boundingBox())?.width ?? 0
  )).toBeGreaterThanOrEqual(400);
  await expect(panel.getByRole('button', { name: 'Expand workflow', exact: true })).toBeVisible();
});

test('workflow stepper has stable light, dark, mobile, and forced-colors visuals', async ({ page }) => {
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

  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.goto('/');
  await page.evaluate(() => document.fonts.ready);

  const stepper = page.getByRole('navigation', { name: 'Architecture delivery workflow' });
  const steps = stepper.getByRole('button');
  await expect(stepper).toBeVisible();
  await expect(steps).toHaveCount(4);
  await expect(steps.nth(0)).toHaveAttribute('aria-current', 'step');
  await expect(steps.nth(1)).toBeDisabled();
  await expect.soft(stepper).toHaveScreenshot('workflow-stepper-light.png', {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
    scale: 'css',
    stylePath: visualSnapshotStylePath,
  });

  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByRole('button', { name: 'Switch to Dark Mode' }).click();
  await expect(page.locator('body')).toHaveClass(/dark-mode/);
  await expect.soft(stepper).toHaveScreenshot('workflow-stepper-dark.png', {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
    scale: 'css',
    stylePath: visualSnapshotStylePath,
  });

  await page.getByRole('button', { name: 'Switch to Light Mode' }).click();
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(stepper).toHaveCSS('overflow-x', 'hidden');
  await expect.poll(() => stepper.evaluate((element) => (
    element.scrollWidth <= element.clientWidth
  ))).toBe(true);
  await expect.soft(stepper).toHaveScreenshot('workflow-stepper-mobile.png', {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
    scale: 'css',
    stylePath: visualSnapshotStylePath,
  });

  await page.setViewportSize({ width: 1600, height: 1000 });
  await page.emulateMedia({ forcedColors: 'active', reducedMotion: 'reduce' });
  expect(await page.evaluate(() => matchMedia('(forced-colors: active)').matches)).toBe(true);
  await expect(steps.nth(0)).toHaveCSS('box-shadow', 'none');
  await expect.soft(stepper).toHaveScreenshot('workflow-stepper-forced-colors.png', {
    animations: 'disabled',
    caret: 'hide',
    maxDiffPixelRatio: 0.01,
    scale: 'css',
    stylePath: visualSnapshotStylePath,
  });
});

test('feedback launcher is styled before the lazy modal chunk is opened', async ({ page }) => {
  await openInteractionDiagram(page);
  const launcher = page.getByRole('button', { name: 'Feedback' });
  await expect(launcher).toBeVisible();
  await expect(launcher).toHaveCSS('position', 'absolute');
  await expect(launcher).toHaveCSS('border-radius', '999px');
  expect((await launcher.boundingBox())?.width).toBeLessThan(180);

  await page.setViewportSize({ width: 390, height: 844 });
  await expect(launcher).toHaveCSS('width', '44px');
  await expect(launcher).toHaveCSS('height', '44px');
});

test('shared feedback dialog follows the selected app theme and restores focus', async ({ page }) => {
  await openInteractionDiagram(page);
  const launcher = page.getByRole('button', { name: 'Feedback' });

  await launcher.click();
  let dialog = page.getByRole('dialog', { name: 'Share Feedback' });
  await expect(dialog).toBeFocused();
  await expect(dialog).toHaveCSS('background-color', 'rgb(255, 255, 255)');
  await expectReadableContrast(dialog.getByRole('heading', { name: 'Share Feedback' }));
  await expectNoWcagViolations(page, '.feedback-modal');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(launcher).toBeFocused();

  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByRole('button', { name: 'Switch to Dark Mode' }).click();
  await launcher.click();
  dialog = page.getByRole('dialog', { name: 'Share Feedback' });
  await expect(dialog).toHaveCSS('background-color', 'rgb(39, 51, 61)');
  await expectReadableContrast(dialog.getByRole('heading', { name: 'Share Feedback' }));
  await expectNoWcagViolations(page, '.feedback-modal');
});

test('About dialog exposes attribution and repository details accessibly', async ({ page }) => {
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
  await expect(page.getByText('Community project', { exact: true })).toBeVisible();
  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' }).fill('about');
  await palette.getByRole('option', { name: /^About this application/ }).click();

  const dialog = page.getByRole('dialog', { name: 'Microsoft Product Architecture Diagram Builder' });
  await expect(dialog).toBeFocused();
  await expect(dialog.getByText('Arturo Quiroga')).toBeVisible();
  await expect(dialog.getByText('Swarm Data SE, Jiayi Yang')).toBeVisible();
  await expect(dialog.getByText(/independent, community-maintained fork/)).toBeVisible();
  await expect(dialog.getByText(/not an official Microsoft product/)).toBeVisible();
  await expect(dialog.getByRole('link', { name: /yang-jiayi\/AzureDiagarm/ }))
    .toHaveAttribute('href', 'https://github.com/yang-jiayi/AzureDiagarm');
  await expectNoWcagViolations(page, '.about-dialog');

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await page.getByRole('button', { name: 'More', exact: true }).click();
  await page.getByRole('button', { name: '日本語' }).click();
  await expect(page.getByText('コミュニティ版', { exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: /テンプレートを見る/ })).toBeVisible();
});

test('template gallery previews, filters, and applies a starter architecture', async ({ page }) => {
  await initializePage(page);
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
    if (path === '/api/diagrams' && request.method() === 'POST') {
      const body = request.postDataJSON() as {
        diagramName: string;
        payload: ReturnType<typeof interactionCloudDocument>['payload'];
      };
      await fulfillJson(route, {
        ...cloudDocument('template-cloud', body.diagramName),
        diagramName: body.diagramName,
        payload: body.payload,
      }, 200, { etag: '"template-cloud-1"' });
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  await page.getByRole('tab', { name: 'Create' }).click();
  await page.getByRole('button', { name: 'Templates', exact: true }).click();

  const gallery = page.getByRole('dialog', { name: 'Choose a template' });
  await expect(gallery).toBeFocused();
  await expect(gallery.getByRole('option')).toHaveCount(4);
  await expect(gallery.getByRole('option', { name: /Secure web application/ }))
    .toHaveAttribute('aria-selected', 'true');

  await gallery.getByPlaceholder('Search by workload or pattern').fill('serverless');
  const eventTemplate = gallery.getByRole('option', { name: /Event-driven integration/ });
  await expect(eventTemplate).toBeVisible();
  await expect(gallery.getByRole('option')).toHaveCount(1);
  await eventTemplate.click();
  await expect(gallery.getByRole('heading', { name: 'Event-driven integration' })).toBeVisible();
  await gallery.getByRole('button', { name: 'Use this template' }).click();

  await expect(gallery).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(5);
  await expect(page.locator('.react-flow__node-azureNode img.node-icon')).toHaveCount(5);
  await expect(page.locator('.react-flow__node-azureNode', { hasText: 'API Management' }))
    .toHaveCount(1);
  await expect(page.locator('.react-flow__node-azureNode', { hasText: 'Azure Cosmos DB' }))
    .toHaveCount(1);
});

test('diagram history, document status, privacy review, and threat overlay stay available', async ({ page }) => {
  await openInteractionDiagram(page);

  await expect(page.getByRole('button', { name: /^Document status: Saved/ })).toBeVisible();

  await page.getByRole('tab', { name: 'Design' }).click();
  const undo = page.getByRole('button', { name: 'Undo' });
  const redo = page.getByRole('button', { name: 'Redo' });
  await expect(undo).toBeDisabled();
  const node = page.locator('[data-testid="rf__node-node-a"]');
  const initial = await node.boundingBox();
  expect(initial).not.toBeNull();
  if (!initial) return;

  await page.mouse.move(initial.x + initial.width / 2, initial.y + initial.height / 2);
  await page.mouse.down();
  await page.mouse.move(initial.x + initial.width / 2 + 160, initial.y + initial.height / 2 + 80, {
    steps: 8,
  });
  await page.mouse.up();
  await expect(
    page.getByRole('button', { name: /^Document status: Saving/ }),
  ).toBeVisible({ timeout: 1_000 });
  await expect(undo).toBeEnabled();
  const moved = await node.boundingBox();
  expect(moved).not.toBeNull();
  expect(Math.abs((moved?.x || 0) - initial.x)).toBeGreaterThan(80);

  await undo.click();
  await expect.poll(async () => (await node.boundingBox())?.x ?? -1).toBeCloseTo(initial.x, 0);
  await expect(redo).toBeEnabled();
  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+Y');
  await expect.poll(async () => (await node.boundingBox())?.x ?? -1)
    .toBeCloseTo(moved?.x || 0, 0);

  await page.getByRole('tab', { name: 'Review' }).click();
  await page.getByRole('button', { name: 'Privacy' }).click();
  const privacyDialog = page.getByRole('dialog', { name: 'Ready to share' });
  await expect(privacyDialog).toBeFocused();
  await expect(privacyDialog.getByText('No known sensitive patterns found')).toBeVisible();
  await privacyDialog.getByRole('button', { name: 'Done' }).click();

  const threats = page.getByRole('button', { name: 'Threats' });
  await threats.click();
  await expect(threats).toHaveAttribute('aria-pressed', 'true');
  const threatOverlay = page.getByRole('complementary', { name: 'Threat model overlay' });
  await expect(threatOverlay).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-node-b"]'))
    .toHaveCSS('outline-color', 'rgb(217, 119, 6)');

  await page.evaluate(() => {
    const state = window as Window & {
      __sawThreatOverlayExport?: boolean;
      __threatOverlayExportClipped?: boolean;
      __threatOverlayExportObserver?: MutationObserver;
    };
    state.__sawThreatOverlayExport = false;
    state.__threatOverlayExportClipped = false;
    state.__threatOverlayExportObserver?.disconnect();
    const observer = new MutationObserver(() => {
      const overlay = document.querySelector<HTMLElement>('[data-export-threat-overlay="true"]');
      if (overlay) {
        const frame = overlay.parentElement;
        const overlayBounds = overlay.getBoundingClientRect();
        const frameBounds = frame?.getBoundingClientRect();
        state.__sawThreatOverlayExport = true;
        state.__threatOverlayExportClipped = (
          overlay.scrollHeight > overlay.clientHeight + 1
          || Boolean(frameBounds && overlayBounds.bottom > frameBounds.bottom + 1)
        );
        observer.disconnect();
      }
    });
    state.__threatOverlayExportObserver = observer;
    observer.observe(document.body, { childList: true, subtree: true });
  });
  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');
  const downloadPromise = page.waitForEvent('download');
  await palette.getByRole('option', { name: /^Export PNG/ }).click();
  await downloadPromise;
  expect(await page.evaluate(() => (
    (window as Window & { __sawThreatOverlayExport?: boolean }).__sawThreatOverlayExport
  ))).toBe(true);
  expect(await page.evaluate(() => (
    (window as Window & { __threatOverlayExportClipped?: boolean })
      .__threatOverlayExportClipped
  ))).toBe(false);

  await threatOverlay.getByRole('button', { name: 'Hide threat model overlay' }).click();
  await expect(threatOverlay).toBeHidden();
});

test('mobile users can edit a selected service without opening a desktop inspector', async ({ page }) => {
  await openInteractionDiagram(page);
  await page.setViewportSize({ width: 390, height: 844 });

  await page.locator('[data-testid="rf__node-node-a"]').click();
  const launcher = page.getByRole('button', { name: 'Edit selected node: App Service' });
  await expect(launcher).toBeVisible();
  await launcher.click();

  const editor = page.getByRole('dialog', { name: 'Edit selected node' });
  await expect(editor).toBeVisible();
  await editor.getByLabel('Display name').fill('Customer API');
  await editor.getByLabel('Description').fill('Handles customer requests');
  await editor.getByRole('button', { name: 'Save changes' }).click();

  await expect(editor).toBeHidden();
  await expect(page.getByRole('button', { name: 'Edit selected node: Customer API' })).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toContainText('Customer API');

  await page.getByRole('button', { name: 'Edit selected node: Customer API' }).click();
  await editor.getByRole('button', { name: 'Cost' }).click();
  const pricingEditor = page.getByRole('dialog', { name: /Cost settings — App Service/ });
  await expect(pricingEditor).toBeVisible();
  await pricingEditor.getByRole('button', { name: 'Close' }).click();

  const group = page.locator('[data-testid="rf__node-group-a"]');
  await group.locator('.group-node-header').click();
  await page.getByRole('button', { name: 'Edit selected node: Application layer' }).click();
  const groupEditor = page.getByRole('dialog', { name: 'Edit selected node' });
  await groupEditor.locator('input[type="color"]').fill('#ef4444');
  await groupEditor.getByRole('button', { name: 'Save changes' }).click();
  await page.locator('.react-flow__pane').click({ position: { x: 20, y: 20 } });
  await expect(group.locator('.group-node')).toHaveCSS('border-color', 'rgb(239, 68, 68)');
});

test('cloud share links wait for privacy preflight confirmation', async ({ page }) => {
  const document = interactionCloudDocument();
  const remoteComment = {
    commentId: 'comment-privacy',
    message: 'client_secret=comment~Secret1234',
    authorEmail: 'reviewer@example.com',
    createdAt: now,
  };
  const historicalVersion = cloudVersion('A', 'version-privacy', 'Security review');
  (historicalVersion.payload.nodes[0].data as Record<string, unknown>).description =
    'password=history!Secret5678';
  await openInteractionDiagram(page, undefined, false, document);
  let shareRequests = 0;
  let versionDetailRequests = 0;
  let latestDocumentRequests = 0;
  let persistedDocument = document;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [summary('A', document.diagramName)] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      latestDocumentRequests += 1;
      const share = {
        shareId: 'share-1',
        role: 'viewer',
        createdAt: now,
      };
      await fulfillJson(route, {
        ...persistedDocument,
        revision: shareRequests > 0 ? 4 : 3,
        comments: [remoteComment],
        shares: shareRequests > 0 ? [share] : [],
      }, 200, { etag: shareRequests > 0 ? '"A-4"' : '"A-3"' });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      persistedDocument = {
        ...persistedDocument,
        diagramName: body.diagramName || persistedDocument.diagramName,
        payload: body.payload || persistedDocument.payload,
        revision: 2,
        etag: '"A-2"',
      };
      await fulfillJson(route, persistedDocument, 200, { etag: '"A-2"' });
      return;
    }
    if (path === '/api/diagrams/A/versions' && method === 'GET') {
      await fulfillJson(route, {
        versions: [{
          versionId: historicalVersion.versionId,
          diagramId: historicalVersion.diagramId,
          diagramName: historicalVersion.diagramName,
          notes: historicalVersion.notes,
          createdAt: historicalVersion.createdAt,
          createdByEmail: historicalVersion.createdByEmail,
          sourceRevision: historicalVersion.sourceRevision,
        }],
      });
      return;
    }
    if (path === '/api/diagrams/A/versions/version-privacy' && method === 'GET') {
      versionDetailRequests += 1;
      await fulfillJson(route, { version: historicalVersion });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'GET') {
      await fulfillJson(route, { shares: [] });
      return;
    }
    if (path === '/api/diagrams/A/shares' && method === 'POST') {
      shareRequests += 1;
      await fulfillJson(route, {
        result: {
          token: 'privacy-approved-share',
          url: 'https://example.test/#share-privacy-approved-share',
          share: {
            shareId: 'share-1',
            role: 'viewer',
            createdAt: now,
          },
        },
      });
      return;
    }
    await route.fallback();
  });

  await page.getByRole('button', { name: /^Document status:/ }).click();
  const cloudWorkspace = page.getByRole('dialog', { name: 'Cloud workspace' });
  await expect(cloudWorkspace.getByRole('button', { name: 'Create link' })).toBeVisible();
  await cloudWorkspace.getByRole('button', { name: 'Create link' }).click();

  const privacyDialog = page.getByRole('dialog', { name: 'Review sensitive information' });
  await expect(privacyDialog).toBeVisible();
  await expect(privacyDialog.getByText('Credential', { exact: true })).toHaveCount(2);
  await expect(
    privacyDialog.getByText('share.current.comments[0].message', { exact: true }),
  ).toBeVisible();
  await expect(
    privacyDialog.getByText('share.versions[0].payload.nodes[0].data.description', {
      exact: true,
    }),
  ).toBeVisible();
  await expect(privacyDialog).not.toContainText('comment~Secret1234');
  await expect(privacyDialog).not.toContainText('history!Secret5678');
  expect(versionDetailRequests).toBe(1);
  expect(latestDocumentRequests).toBe(1);
  expect(shareRequests).toBe(0);
  await page.keyboard.press('Escape');
  await expect(privacyDialog).toBeHidden();
  await expect(cloudWorkspace).toBeVisible();
  expect(shareRequests).toBe(0);

  await expect(cloudWorkspace.getByRole('button', { name: 'Create link' })).toBeEnabled();
  await cloudWorkspace.getByRole('button', { name: 'Create link' }).click();
  await privacyDialog.getByRole('button', { name: 'Proceed as shown' }).click();
  await expect.poll(() => shareRequests).toBe(1);
  await expect(cloudWorkspace.getByLabel('New share URL'))
    .toHaveValue('https://example.test/#share-privacy-approved-share');
});

for (const [width, allowBYO] of [[1600, true], [1600, false], [390, true], [390, false]] as const) {
  test(`managed Astra workspace preserves BYO access and retires comparisons at ${width}px with BYO policy ${allowBYO}`, async ({ page }, testInfo) => {
    await initializePage(page);
    await page.addInitScript(() => {
      localStorage.setItem('azure-diagrams-model-settings', JSON.stringify({
        version: 3, model: 'gpt-5.6-sol', reasoningEffort: 'high',
        featureOverrides: { blueprint: { model: 'gpt-5.6-luna', reasoningEffort: 'medium' } },
      }));
      localStorage.setItem('azure-diagrams-byo-ai-settings', JSON.stringify({
        version: 1, enabled: true, provider: 'openai', model: 'customer-owned-model',
        baseUrl: 'https://api.openai.com/v1', apiFormat: 'responses',
      }));
    });
    const requests: Record<string, unknown>[] = [];
    await page.route('**/api/**', async route => {
      const path = new URL(route.request().url()).pathname;
      if (path === '/api/access/me') {
        await fulfillJson(route, { enabled: false, authenticated: false, allowed: true, isAdmin: false });
      } else if (path === '/api/runtime-config') {
        await fulfillJson(route, { features: { bringYourOwnAI: allowBYO } });
      } else if (path === '/api/openai') {
        requests.push(route.request().postDataJSON() as Record<string, unknown>);
        await fulfillJson(route, { error: { message: 'Mocked request completed without a generated graph.' } }, 400);
      } else {
        await fulfillJson(route, { error: 'Not found' }, 404);
      }
    });
    await page.setViewportSize({ width, height: 900 });
    await page.goto('/');
    const retired = /Compare Models|Compare Validation/;
    const mobileCommands = page.getByRole('navigation', { name: 'Mobile command bar' });
    const ribbonSheet = page.getByRole('dialog', { name: 'Ribbon commands' });
    const moreTrigger = page.getByRole('button', {
      name: width < 640 ? 'More application options' : 'More',
      exact: true,
    });
    const moreOptions = page.getByRole('dialog', { name: 'More application options' });
    await moreTrigger.click();
    await expect(moreOptions).toBeVisible();
    await expect(moreOptions.getByRole('button', { name: retired })).toHaveCount(0);
    await expect(moreOptions.getByRole('button', { name: 'AI connections', exact: true })).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(moreOptions).toBeHidden();
    await expect(moreTrigger).toBeFocused();
    await page.getByRole('region', { name: 'Architecture canvas' }).focus();
    await page.keyboard.press('Control+K');
    const commands = page.getByTestId('command-palette');
    for (const query of ['compare models', 'compare validation']) {
      await commands.getByRole('combobox', { name: 'Search commands and services' }).fill(query);
      await expect(commands.getByRole('option', { name: retired })).toHaveCount(0);
    }
    await commands.getByRole('combobox', { name: 'Search commands and services' }).fill('custom AI');
    const connectionCommand = commands.getByRole('option', { name: /AI connections/ });
    await expect(connectionCommand).toBeVisible();
    await connectionCommand.click();
    const connectionDialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
    await expect(connectionDialog).toBeVisible();
    await expect(connectionDialog).toContainText('customer-owned-model');
    await expect(connectionDialog.getByLabel('API key (tab memory only)', { exact: true })).toHaveValue('');
    await expect(connectionDialog.getByRole('button', { name: 'Test connection', exact: true })).toBeDisabled();
    await expect(connectionDialog.getByRole('button', { name: 'Save profile', exact: true })).toBeEnabled();
    if (!allowBYO) await expect(connectionDialog).toContainText('disabled by the application administrator');
    await expectNoWcagViolations(page, '.byo-ai-dialog');
    await connectionDialog.getByRole('button', { name: 'Use managed Astra', exact: true }).click();
    await connectionDialog.getByRole('button', { name: 'Done', exact: true }).click();
    await expect(connectionDialog).toBeHidden();
    if (width < 640) {
      await mobileCommands.getByRole('button', { name: 'Review', exact: true }).click();
      await expect(ribbonSheet).toBeVisible();
    }
    await page.getByRole('tab', { name: 'Create', exact: true }).click();
    await expect(page.getByRole('button', { name: retired })).toHaveCount(0);
    const trigger = page.locator('.model-popover-trigger');
    await expect(trigger).toContainText('GPT-6 Astra');
    await expect(trigger).not.toContainText('Custom');
    await trigger.click();
    const settings = page.getByRole('dialog', { name: 'AI model settings', exact: true });
    await expect(settings).not.toContainText(/GPT-5|Alternative model/i);
    await expect(settings.getByRole('combobox', { name: 'AI connection', exact: true })).toHaveValue('');
    await expect(settings.locator('.msp-model-btn, .msp-byo-card')).toHaveCount(0);
    await settings.locator('.msp-advanced-settings summary').click();
    await expect(settings.locator('.astra-reasoning-settings').getByRole('combobox')).toHaveCount(4);
    await expect(settings.getByRole('combobox', { name: 'Blueprint Diagrams - Reasoning effort' })).toHaveValue('medium');
    await expect(settings.getByRole('button', { name: 'High', exact: true })).toHaveAttribute('aria-pressed', 'true');
    await expectNoWcagViolations(page, '.toolbar-dropdown-menu--model-settings');
    await settings.getByRole('button', { name: 'Close', exact: true }).click();
    await expect(settings).toBeHidden();
    if (width < 640) {
      await expect(ribbonSheet).toBeHidden();
      const reopenCommands = mobileCommands.getByRole('button', { name: 'Create', exact: true });
      await expect(reopenCommands).toBeFocused();
      await reopenCommands.click();
      await expect(ribbonSheet).toBeVisible();
    }
    await page.getByRole('tab', { name: 'Review', exact: true }).click();
    await expect(page.getByRole('button', { name: retired })).toHaveCount(0);
    if (width < 640) {
      await page.keyboard.press('Escape');
      await expect(ribbonSheet).toBeHidden();
      await expect(mobileCommands.getByRole('button', { name: 'Review', exact: true })).toBeFocused();
    }
    if (width === 1600) {
      const generator = await openAiGenerator(page);
      await generator.getByLabel('Architecture Description or Modification').fill('Create a small web application');
      await generator.getByRole('button', { name: 'Continue to output' }).click();
      await expect(generator.getByRole('combobox', { name: /Select AI model/ })).toHaveCount(0);
      await generator.getByRole('button', { name: 'Generate Architecture' }).click();
      if (testInfo.project.metadata.managedAstraConfigured === false) {
        await expect(generator.getByRole('alert')).toHaveText(
          'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment.',
        );
        await expect(generator.getByRole('button', { name: 'Retry generation' })).toBeEnabled();
        expect(requests).toHaveLength(0);
      } else {
        await expect.poll(() => requests.length).toBe(1);
        expect(requests[0]).toMatchObject({
          apiFormat: 'responses',
          deployment: 'playwright-gpt-6-astra',
          body: { model: 'playwright-gpt-6-astra' },
        });
        expect(requests[0]).not.toHaveProperty('byo');
      }
    }
  });
}

test('BYO connection manager follows the canonical compact breakpoint without overflow', async ({ page }) => {
  await initializePage(page);
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagrams-byo-ai-settings', JSON.stringify({
      version: 1, enabled: true, provider: 'openai', model: 'customer-owned-model',
      baseUrl: 'https://api.openai.com/v1', apiFormat: 'responses',
    }));
  });
  await page.route('**/api/**', async route => {
    const path = new URL(route.request().url()).pathname;
    if (path === '/api/access/me') {
      await fulfillJson(route, { enabled: false, authenticated: false, allowed: true, isAdmin: false });
    } else if (path === '/api/runtime-config') {
      await fulfillJson(route, { features: { bringYourOwnAI: true } });
    } else {
      await fulfillJson(route, { error: 'Not found' }, 404);
    }
  });
  await page.setViewportSize({ width: 640, height: 900 });
  await page.goto('/');
  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const commands = page.getByTestId('command-palette');
  await commands.getByRole('combobox', { name: 'Search commands and services' }).fill('custom AI');
  await commands.getByRole('option', { name: /AI connections/ }).click();
  const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
  const form = dialog.locator('.byo-ai-form-grid');
  await expect(dialog).toBeVisible();
  await expect(form).toHaveCount(2);
  for (const [width, columns] of [[640, 1], [641, 2]] as const) {
    await page.setViewportSize({ width, height: 900 });
    await expect.poll(() => form.evaluateAll(elements => elements.map(
      element => getComputedStyle(element).gridTemplateColumns.split(/\s+/).length,
    ))).toEqual([columns, columns]);
    expect(await dialog.evaluate(element => element.scrollWidth <= element.clientWidth + 1)).toBe(true);
    await expectNoWcagViolations(page, '.byo-ai-dialog');
  }
});

for (const [provider, apiFormat] of [
  ['openai', 'responses'],
  ['azure-openai', 'chat-completions'],
] as const) {
  test(`BYO full application tests, selects, and generates with ${provider} ${apiFormat}`, async ({ page }) => {
    await initializePage(page);
    const endpoint = provider === 'openai' ? 'https://api.openai.com' : 'https://byo-e2e.openai.azure.com';
    const model = `customer-${apiFormat}-deployment`;
    const key = 'sk-byo-e2e-local-fixture-key';
    const requests: Record<string, unknown>[] = [];
    await page.route('**/*', async route => {
      const hostname = new URL(route.request().url()).hostname;
      if (hostname === '127.0.0.1' || hostname === 'localhost') await route.continue();
      else await route.abort('blockedbyclient');
    });
    await page.route('**/api/**', async route => {
      const path = new URL(route.request().url()).pathname;
      if (path === '/api/access/me') {
        await fulfillJson(route, { enabled: false, authenticated: false, allowed: true, isAdmin: false });
      } else if (path === '/api/runtime-config') {
        await fulfillJson(route, { features: { bringYourOwnAI: true } });
      } else if (path === '/api/openai') {
        const payload = route.request().postDataJSON() as Record<string, unknown>;
        requests.push(payload);
        const text = JSON.stringify(requests.length === 1 ? { status: 'ok' } : {
          architectureName: 'BYO captured connection',
          groups: [],
          services: [{
            id: 'byo-web',
            name: 'App Service',
            type: 'App Service',
            category: 'compute',
            description: 'Generated using the explicitly selected user connection',
          }],
          connections: [],
          workflow: [],
        });
        await fulfillJson(route, payload.apiFormat === 'responses' ? {
          status: 'completed',
          model,
          output_text: text,
          usage: { input_tokens: 30, output_tokens: 20, total_tokens: 50 },
        } : {
          model,
          choices: [{ finish_reason: 'stop', message: { role: 'assistant', content: text } }],
          usage: { prompt_tokens: 30, completion_tokens: 20, total_tokens: 50 },
        });
      } else {
        await fulfillJson(route, { error: 'Not found' }, 404);
      }
    });
    await page.goto('/');
    await page.getByRole('button', { name: 'More', exact: true }).click();
    await page.getByRole('dialog', { name: 'More application options' })
      .getByRole('button', { name: 'AI connections', exact: true }).click();
    const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
    await dialog.getByLabel('Profile name', { exact: true }).fill('Full application BYO');
    await dialog.getByLabel('Provider', { exact: true }).selectOption(provider);
    if (provider === 'azure-openai') {
      await dialog.getByLabel('Endpoint origin', { exact: true }).fill(endpoint);
    }
    await dialog.getByLabel('Model / deployment', { exact: true }).fill(model);
    await dialog.getByLabel('API format', { exact: true }).selectOption(apiFormat);
    await dialog.getByLabel('Supports reasoning', { exact: true }).check();
    await dialog.getByLabel('Profile reasoning effort', { exact: true }).selectOption('max');
    await dialog.getByLabel('Maximum output tokens', { exact: true }).fill('32000');
    await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
    await dialog.getByLabel('API key (tab memory only)', { exact: true }).fill(key);
    await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
    await expect(dialog.locator('.byo-ai-status--verified')).toBeVisible();
    await expect.poll(() => requests.length).toBe(1);
    expect(requests[0]).toMatchObject({ apiFormat, deployment: model, byo: { provider, endpoint, apiKey: key } });
    expect(await page.evaluate(() => JSON.parse(localStorage.getItem('azure-diagrams-byo-ai-settings') || '{}').activeProfileId))
      .toBeNull();
    await dialog.getByRole('button', { name: 'Use this profile', exact: true }).click();
    await dialog.getByRole('button', { name: 'Done', exact: true }).click();
    const generator = await openAiGenerator(page);
    await generator.getByLabel('Architecture Description or Modification').fill('Create one Azure App Service using this connection.');
    await generator.getByRole('button', { name: 'Continue to output' }).click();
    await generator.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
    const review = page.getByRole('dialog', { name: 'Review AI changes', exact: true });
    await expect(review).toBeVisible();
    await expect(page.getByTestId('rf__node-byo-web')).toHaveCount(0);
    await review.getByRole('button', { name: 'Apply selected changes', exact: true }).click();
    await expect(review).toBeHidden();
    await expect(page.getByTestId('rf__node-byo-web')).toBeVisible();
    await expect(page.locator('.model-generation-badge-text strong')).toContainText(model);
    expect(requests.length).toBeGreaterThanOrEqual(2);
    expect(requests[1]).toMatchObject({
      apiFormat, deployment: model, byo: { provider, endpoint, apiKey: key },
      body: apiFormat === 'responses'
        ? { model, max_output_tokens: 32000, reasoning: { effort: 'max' } }
        : { model, max_completion_tokens: 32000, reasoning_effort: 'max' },
    });
    expect(await page.evaluate(() => JSON.stringify({
      local: { ...localStorage }, session: { ...sessionStorage },
    }))).not.toContain(key);
    const requestsBeforeReload = requests.length;
    await page.reload();
    await page.getByRole('button', { name: 'More', exact: true }).click();
    await page.getByRole('dialog', { name: 'More application options' })
      .getByRole('button', { name: 'AI connections', exact: true }).click();
    await expect(dialog.locator('.byo-ai-status--key-required')).toBeVisible();
    await expect(dialog.getByLabel('API key (tab memory only)', { exact: true })).toHaveValue('');
    await expect(dialog.getByRole('button', { name: 'Test connection', exact: true })).toBeDisabled();
    await expect(dialog.getByRole('button', { name: 'Profile selected', exact: true })).toBeDisabled();
    expect(await page.evaluate(() => JSON.parse(localStorage.getItem('azure-diagrams-byo-ai-settings') || '{}').activeProfileId))
      .not.toBeNull();
    expect(requests).toHaveLength(requestsBeforeReload);
    for (const request of requests) {
      expect(request).toMatchObject({ apiFormat, deployment: model, byo: { provider, endpoint, apiKey: key } });
    }
  });
}

test('PNG export contains rendered diagram content instead of a blank canvas', async ({ page }, testInfo) => {
  await openInteractionDiagram(page);
  await page.waitForTimeout(500);
  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');

  const downloadPromise = page.waitForEvent('download');
  await palette.getByRole('option', { name: /^Export PNG/ }).click();
  const download = await downloadPromise;
  const outputPath = testInfo.outputPath('diagram-export.png');
  await download.saveAs(outputPath);
  const png = await readFile(outputPath);
  expect(png.byteLength).toBeGreaterThan(20_000);

  const metrics = await page.evaluate(async (dataUrl) => {
    const image = new Image();
    image.src = dataUrl;
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) throw new Error('Canvas context unavailable');
    context.drawImage(image, 0, 0);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    const colors = new Set<string>();
    const stride = Math.max(4, Math.floor(pixels.length / 20_000 / 4) * 4);
    for (let index = 0; index < pixels.length; index += stride) {
      colors.add(`${pixels[index]},${pixels[index + 1]},${pixels[index + 2]},${pixels[index + 3]}`);
    }
    return {
      width: canvas.width,
      height: canvas.height,
      sampledColors: colors.size,
    };
  }, `data:image/png;base64,${png.toString('base64')}`);

  expect(metrics.width).toBeGreaterThan(600);
  expect(metrics.height).toBeGreaterThan(400);
  expect(metrics.sampledColors).toBeGreaterThan(40);
});

for (const varied of [false, true]) {
  test(`PNG export legend preserves ${varied ? 'mixed-style disclosure' : 'authored paint'} and collapsed metadata`, async ({ page }, testInfo) => {
    const base = interactionCloudDocument();
    const document = {
      ...base,
      payload: {
        ...base.payload,
        titleBlockData: {
          architectureName: 'Authored connection styles',
          author: 'Export reviewer',
          date: '2026-09-07',
          version: '1.2',
        },
        edges: [
          {
            ...base.payload.edges[0],
            data: { connectionType: 'security' },
            style: { stroke: '#006d77', strokeDasharray: '10 2 3 2', opacity: 0.45 },
          },
          ...(varied ? [{
            ...base.payload.edges[0],
            id: 'default-security',
            data: { connectionType: 'security' },
          }] : []),
        ],
      },
    };
    await openInteractionDiagram(page, undefined, false, document);
    await expect(page.locator('.title-block')).toHaveClass(/collapsed/);
    await page.evaluate(() => {
      const observer = new MutationObserver(() => {
        const legend = window.document.querySelector<HTMLElement>('[data-export-legend]');
        const host = legend?.parentElement;
        const swatch = legend?.querySelector('svg');
        if (!legend || !host || !swatch) return;
        const line = swatch.querySelector('path');
        const box = host.getBoundingClientRect();
        const sample = swatch.getBoundingClientRect();
        window.document.body.dataset.capturedPngLegend = JSON.stringify({
          label: legend.querySelector('strong')?.textContent,
          color: line?.getAttribute('stroke') ?? null,
          dash: line?.getAttribute('stroke-dasharray') ?? null,
          opacity: line?.getAttribute('opacity') ?? null,
          text: host.textContent,
          width: box.width,
          height: box.height,
          sampleX: sample.left - box.left + 4,
          gapX: sample.left - box.left + 11,
          sampleY: sample.top - box.top + 6,
        });
        observer.disconnect();
      });
      observer.observe(window.document.body, { childList: true, subtree: true });
    });

    await page.getByRole('region', { name: 'Architecture canvas' }).focus();
    await page.keyboard.press('Control+K');
    const palette = page.getByTestId('command-palette');
    await palette.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');
    const downloadPromise = page.waitForEvent('download');
    await palette.getByRole('option', { name: /^Export PNG/ }).click();
    const download = await downloadPromise;
    const outputPath = testInfo.outputPath('authored-legend.png');
    await download.saveAs(outputPath);
    const png = await readFile(outputPath);
    expect(png.subarray(0, 8).toString('hex')).toBe('89504e470d0a1a0a');

    const captured = await page.evaluate(() => JSON.parse(window.document.body.dataset.capturedPngLegend || 'null'));
    expect(captured).not.toBeNull();
    expect(captured.label).toBe(varied ? 'Security (varied)' : 'Security');
    expect(captured.color).toBe(varied ? null : 'rgb(0, 109, 119)');
    expect(captured.dash?.replaceAll('px', '').split(/[,\s]+/).map(Number) ?? null)
      .toEqual(varied ? null : [10, 2, 3, 2]);
    expect(captured.opacity).toBe(varied ? null : '0.45');
    expect(captured.text).toContain('Export reviewer');
    expect(captured.text).toContain('2026-09-07');
    expect(captured.text).toContain('v1.2');

    const pixels = await page.evaluate(async ({ dataUrl, capture }) => {
      const image = new Image();
      image.src = dataUrl;
      await image.decode();
      const canvas = window.document.createElement('canvas');
      canvas.width = image.naturalWidth;
      canvas.height = image.naturalHeight;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      if (!context) throw new Error('Canvas context unavailable');
      context.drawImage(image, 0, 0);
      const sample = (x: number) => Array.from(context.getImageData(
        Math.floor(x / capture.width * canvas.width),
        Math.floor(capture.sampleY / capture.height * canvas.height),
        1, 1,
      ).data);
      return { line: sample(capture.sampleX), gap: sample(capture.gapX) };
    }, { dataUrl: `data:image/png;base64,${png.toString('base64')}`, capture: captured });
    if (varied) {
      expect(pixels.line.slice(0, 3).every(channel => channel >= 235)).toBe(true);
    } else {
      expect(pixels.line[0]).toBeGreaterThan(100);
      expect(pixels.line[0]).toBeLessThan(170);
      expect(pixels.line[1] - pixels.line[0]).toBeGreaterThan(35);
      expect(pixels.line[2] - pixels.line[0]).toBeGreaterThan(35);
      expect(pixels.gap.slice(0, 3).every(channel => channel >= 235)).toBe(true);
    }
  });
}

test('PNG export expands to include a manually offset edge label', async ({ page }, testInfo) => {
  const document = interactionCloudDocument();
  document.payload.edges = [{
    ...document.payload.edges[0],
    label: 'Reads application data',
    data: {
      labelOffsetX: 1100,
      labelOffsetY: 0,
      labelOffsetAuto: false,
    },
  }];
  await openInteractionDiagram(page, undefined, false, document);
  await expect(page.locator('.editable-edge-label-shell'))
    .toHaveAttribute('data-label-offset-x', '1100');

  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');

  const downloadPromise = page.waitForEvent('download');
  await palette.getByRole('option', { name: /^Export PNG/ }).click();
  const download = await downloadPromise;
  const outputPath = testInfo.outputPath('diagram-export-with-offset-label.png');
  await download.saveAs(outputPath);
  const png = await readFile(outputPath);
  const width = await page.evaluate(async (dataUrl) => {
    const image = new Image();
    image.src = dataUrl;
    await image.decode();
    return image.naturalWidth;
  }, `data:image/png;base64,${png.toString('base64')}`);

  expect(width).toBeGreaterThan(3_000);
});

test('automatic edge label offsets refresh after a connected node moves', async ({ page }) => {
  const document = interactionCloudDocument();
  document.payload.nodes = [
    ...document.payload.nodes,
    {
      id: 'node-blocker',
      type: 'azureNode',
      position: { x: 360, y: 220 },
      data: { label: 'API Management', serviceName: 'API Management' },
    },
  ];
  document.payload.edges = [{
    ...document.payload.edges[0],
    label: 'Reads application data',
    data: {
      labelOffsetX: 0,
      labelOffsetY: 0,
      labelOffsetAuto: true,
    },
  }];
  await openInteractionDiagram(page, undefined, false, document);

  const edgeLabelShell = page.locator('.editable-edge-label-shell');
  await expect(edgeLabelShell).toHaveAttribute('data-label-offset-auto', 'true');
  await expect.poll(async () => {
    const offsetX = Number(await edgeLabelShell.getAttribute('data-label-offset-x'));
    const offsetY = Number(await edgeLabelShell.getAttribute('data-label-offset-y'));
    return Math.hypot(offsetX, offsetY);
  }).toBeGreaterThan(0);

  const [sourceBox, targetBox] = await Promise.all([
    page.locator('[data-testid="rf__node-node-a"]').boundingBox(),
    page.locator('[data-testid="rf__node-node-b"]').boundingBox(),
  ]);
  expect(sourceBox).not.toBeNull();
  expect(targetBox).not.toBeNull();
  if (!sourceBox || !targetBox) return;

  await page.mouse.move(
    targetBox.x + targetBox.width / 2,
    targetBox.y + targetBox.height / 2,
  );
  await page.mouse.down();
  await page.mouse.move(
    sourceBox.x + sourceBox.width / 2,
    targetBox.y + targetBox.height / 2 + 320,
    { steps: 8 },
  );
  await page.mouse.up();

  await expect.poll(async () => {
    const offsetX = Number(await edgeLabelShell.getAttribute('data-label-offset-x'));
    const offsetY = Number(await edgeLabelShell.getAttribute('data-label-offset-y'));
    return Math.hypot(offsetX, offsetY);
  }).toBe(0);
});

test('manual service insertion stays readable and avoids overlapping nodes', async ({ page }) => {
  await initializePage(page, undefined, { showCanvasHint: true });
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
  const search = page.locator('.search-box input');
  await search.fill('Azure OpenAI');
  await page.getByRole('button', { name: /Add Azure OpenAI to the canvas/i }).click();
  await expect(page.locator('.react-flow__node')).toHaveCount(1);
  await page.getByRole('note', { name: 'Canvas navigation tips' })
    .getByRole('button', { name: 'Fit to view' })
    .click();
  await page.waitForTimeout(450);
  const singleNodeScale = await page.locator('.react-flow__viewport').evaluate((element) => {
    const transform = new DOMMatrixReadOnly(getComputedStyle(element).transform);
    return transform.a;
  });
  expect(singleNodeScale).toBeLessThanOrEqual(1.21);
  await page.getByRole('button', { name: 'Fit diagram to view' }).click();
  await page.waitForTimeout(300);
  const controlFitScale = await page.locator('.react-flow__viewport').evaluate((element) => {
    const transform = new DOMMatrixReadOnly(getComputedStyle(element).transform);
    return transform.a;
  });
  expect(controlFitScale).toBeLessThanOrEqual(1.21);

  await search.fill('App Services');
  const appServices = page.getByRole('button', { name: /Add App Services to the canvas/i });
  await expect(appServices).toHaveCount(1);
  await appServices.click();
  await expect(page.locator('.react-flow__node')).toHaveCount(2);

  const scale = await page.locator('.react-flow__viewport').evaluate((element) => {
    const transform = new DOMMatrixReadOnly(getComputedStyle(element).transform);
    return transform.a;
  });
  expect(scale).toBeLessThanOrEqual(1.21);

  const nodeBoxes = await page.locator('.react-flow__node').evaluateAll((elements) => (
    elements.map((element) => {
      const rect = element.getBoundingClientRect();
      return { left: rect.left, right: rect.right, top: rect.top, bottom: rect.bottom };
    })
  ));
  const [first, second] = nodeBoxes;
  const overlaps = first.left < second.right
    && first.right > second.left
    && first.top < second.bottom
    && first.bottom > second.top;
  expect(overlaps).toBe(false);

  const costBadge = page.locator('.cost-badge').first();
  await expect(costBadge).toBeVisible();
  const badgeFontSize = await costBadge.evaluate((element) => (
    Number.parseFloat(getComputedStyle(element).fontSize)
  ));
  expect(badgeFontSize).toBeLessThan(12);
});

test('service discovery recommends common services and persists a deduplicated layout', async ({ page }) => {
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
  const recommendedTab = page.getByRole('tab', { name: /^Recommended/ });
  await expect(recommendedTab).toHaveAttribute('aria-selected', 'true');
  await expect(page.getByText('Common building blocks for starting an Azure architecture.'))
    .toBeVisible();

  await page.getByRole('tab', { name: /^All/ }).click();
  const search = page.locator('.search-box input');
  await search.fill('App Services');
  const appServices = page.getByRole('button', { name: /Add App Services to the canvas/i });
  await expect(appServices).toHaveCount(1);
  await expect(appServices.locator('.icon-label mark')).toHaveText(['App', 'Services']);

  await page.getByRole('button', { name: 'List view' }).click();
  await expect(page.locator('.icon-palette')).toHaveClass(/palette-layout-list/);
  await expect.poll(() => page.evaluate(() => (
    localStorage.getItem('azure-diagram-builder.paletteLayout.v1')
  ))).toBe('list');

  await page.reload();
  await expect(page.getByRole('region', { name: 'Architecture canvas' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'List view' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.getByRole('button', { name: 'Grid view' }).click();
  await page.getByRole('tab', { name: /^All/ }).click();
  await page.locator('.search-box input').fill('Front Door CDN');
  const longLabel = page
    .getByRole('button', { name: /Add Front Door And CDN Profiles to the canvas/i })
    .locator('.icon-label');
  await expect(longLabel).toBeVisible();
  expect(await longLabel.evaluate(element => (
    getComputedStyle(element).getPropertyValue('-webkit-line-clamp')
  ))).toBe('2');
});

test('architecture chat docks without covering the canvas and persists keyboard resizing', async ({ page }) => {
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
  await page.getByRole('button', { name: 'Guided Chat' }).click();

  const app = page.locator('.app');
  const workspace = page.locator('.workspace');
  const chat = page.getByRole('complementary', { name: 'Architecture Chat' });
  const resizer = page.getByRole('separator', { name: 'Resize Architecture Chat' });
  await expect(app).toHaveClass(/chat-open/);
  await expect(chat).toBeVisible();
  await expect(app).toHaveCSS('padding-right', '460px');
  await expect(chat).toHaveCSS('transform', 'none');
  await expect(page.getByRole('button', { name: 'Open services panel' })).toBeVisible();

  const dockedBounds = await Promise.all([workspace, chat].map(locator => locator.boundingBox()));
  expect(dockedBounds[0]).not.toBeNull();
  expect(dockedBounds[1]).not.toBeNull();
  expect((dockedBounds[0]?.x || 0) + (dockedBounds[0]?.width || 0))
    .toBeLessThanOrEqual((dockedBounds[1]?.x || 0) + 1);

  await resizer.focus();
  await page.keyboard.press('ArrowLeft');
  await expect(resizer).toHaveAttribute('aria-valuenow', '484');
  await expect.poll(() => page.evaluate(() => (
    localStorage.getItem('azure-diagram-builder.chatPanelWidth.v1')
  ))).toBe('484');

  await page.reload();
  await expect(page.getByRole('region', { name: 'Architecture canvas' })).toBeVisible();
  await page.getByRole('button', { name: 'Guided Chat' }).click();
  await expect(page.getByRole('separator', { name: 'Resize Architecture Chat' }))
    .toHaveAttribute('aria-valuenow', '484');

  await page.evaluate(() => {
    localStorage.setItem('azure-diagram-builder.chatPanelWidth.v1', '720');
  });
  await page.setViewportSize({ width: 1200, height: 800 });
  await page.reload();
  await page.getByRole('button', { name: 'Guided Chat' }).click();
  const clampedChat = page.getByRole('complementary', { name: 'Architecture Chat' });
  await expect(page.locator('.app')).toHaveCSS('padding-right', '680px');
  await expect(clampedChat).toHaveCSS('width', '680px');
  await expect(clampedChat).toHaveCSS('transform', 'none');
  const clampedBounds = await Promise.all([workspace, clampedChat].map(
    locator => locator.boundingBox(),
  ));
  expect((clampedBounds[0]?.x || 0) + (clampedBounds[0]?.width || 0))
    .toBeLessThanOrEqual((clampedBounds[1]?.x || 0) + 1);
});

test('compact chat and services panels provide dismissible backdrops', async ({ page }) => {
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

  await page.setViewportSize({ width: 900, height: 800 });
  await page.goto('/');
  await page.getByRole('button', { name: 'Guided Chat' }).click();

  const app = page.locator('.app');
  const chatBackdrop = page.locator('.arch-chat-backdrop');
  const compactChat = page.getByRole('dialog', { name: 'Architecture Chat' });
  await expect(chatBackdrop).toBeVisible();
  await expect(compactChat).toHaveAttribute('aria-modal', 'true');
  await expect.poll(
    () => compactChat.evaluate(dialog => dialog.contains(document.activeElement)),
  ).toBe(true);
  await expect(page.locator('.app-header')).toHaveAttribute('inert', '');
  await expect(page.locator('.workspace')).toHaveAttribute('inert', '');
  await expect(app).toHaveCSS('padding-right', '0px');
  await expect(page.getByRole('separator', { name: 'Resize Architecture Chat' })).toBeHidden();
  await page.keyboard.press('Tab');
  expect(await compactChat.evaluate((dialog) => dialog.contains(document.activeElement))).toBe(true);
  await chatBackdrop.click({ position: { x: 40, y: 400 } });
  await expect(compactChat).toBeHidden();
  await expect(page.locator('.workspace')).not.toHaveAttribute('inert', '');
  await expect(page.getByRole('region', { name: 'Architecture canvas' })).toBeFocused();

  await page.setViewportSize({ width: 390, height: 844 });
  const mobileBar = page.getByRole('navigation', { name: 'Mobile command bar' });
  const servicesButton = mobileBar.getByRole('button', { name: 'Services' });
  await expect(servicesButton).toBeVisible();
  await servicesButton.click();

  const paletteBackdrop = page.locator('.palette-backdrop');
  await expect(paletteBackdrop).toBeVisible();
  await expect(page.getByRole('dialog', { name: 'Microsoft Services' })).toBeVisible();
  await expect(page.locator('.app-header')).toHaveAttribute('inert', '');
  await expect(page.locator('.canvas-container')).toHaveAttribute('inert', '');
  const backdropBounds = await paletteBackdrop.boundingBox();
  expect(backdropBounds).not.toBeNull();
  await paletteBackdrop.click({
    position: {
      x: Math.max(1, (backdropBounds?.width || 1) - 5),
      y: Math.max(1, (backdropBounds?.height || 1) / 2),
    },
  });
  await expect(page.getByRole('dialog', { name: 'Microsoft Services' })).toBeHidden();
  await expect(servicesButton).toBeFocused();

  await mobileBar.getByRole('button', { name: 'Search' }).click();
  const commandPalette = page.getByTestId('command-palette');
  await commandPalette.getByRole('option', { name: /Open cloud workspace/ }).click();
  const cloudDrawer = page.getByRole('dialog', { name: 'Cloud workspace' });
  await expect(cloudDrawer).toBeVisible();
  await expect(cloudDrawer).toHaveAttribute('data-placement', 'bottom');
  await expect(page.locator('.cloud-workspace-overlay')).toBeVisible();
  await expect.poll(async () => {
    const bounds = await cloudDrawer.boundingBox();
    return bounds ? Math.round(844 - (bounds.y + bounds.height)) : -1;
  }).toBe(0);
  await page.keyboard.press('Escape');
  await expect(cloudDrawer).toBeHidden();
  await expect(mobileBar.getByRole('button', { name: 'Search' })).toBeFocused();
});

test('mobile command bar opens one keyboard-safe ribbon bottom sheet', async ({ page }) => {
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

  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');

  const mobileBar = page.getByRole('navigation', { name: 'Mobile command bar' });
  const commandButton = mobileBar.getByRole('button').first();
  await expect(mobileBar).toBeVisible();
  await commandButton.click();

  const ribbonSheet = page.getByRole('dialog', { name: 'Ribbon commands' });
  await expect(ribbonSheet).toBeVisible();
  await expect(page.locator('#application-toolbar')).toHaveCount(1);
  await expect(page.locator('.mobile-command-bar')).toHaveAttribute('inert', '');
  await expect(page.locator('.workspace')).toHaveAttribute('inert', '');
  await expect.poll(async () => {
    const bounds = await ribbonSheet.boundingBox();
    return bounds ? Math.round(844 - (bounds.y + bounds.height)) : -1;
  }).toBe(0);

  await ribbonSheet.getByRole('tab', { name: 'Create' }).click();
  await expect(ribbonSheet).toBeVisible();
  await ribbonSheet.getByRole('button', { name: 'Add Group' }).click();
  await expect(ribbonSheet).toBeHidden();
  await expect(page.locator('.react-flow__node-groupNode')).toHaveCount(1);
  await expect(commandButton).toBeFocused();

  await commandButton.click();
  await expect(ribbonSheet).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(ribbonSheet).toBeHidden();
  await expect(commandButton).toBeFocused();
  await expect(page.locator('.workspace')).not.toHaveAttribute('inert', '');

  const mobileFocus = mobileBar.getByRole('button', { name: 'Focus' });
  await mobileFocus.click();
  const exitFocus = page.getByRole('button', { name: 'Exit Focus' });
  await expect(exitFocus).toBeFocused();
  await exitFocus.click();
  await expect(mobileFocus).toBeFocused();
});

test('mobile canvas-first mode collapses chrome after creation begins', async ({ page }) => {
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

  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await expect(page.locator('.workflow-stepper')).not.toHaveClass(/workflow-stepper--collapsed/);

  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' })
    .fill('App Services');
  await palette.getByRole('option', { name: /App Services/ }).click();

  const app = page.locator('.app');
  const stepper = page.locator('.workflow-stepper');
  await expect(app).toHaveClass(/mobile-canvas-first/);
  await expect(stepper).toHaveClass(/workflow-stepper--collapsed/);
  await expect(page.locator('.icon-palette')).toHaveClass(/collapsed/);
  const collapsedHeight = await stepper.evaluate(element => element.getBoundingClientRect().height);
  expect(collapsedHeight).toBeLessThan(48);

  const summary = stepper.locator('.workflow-stepper-summary');
  await expect(summary).toHaveAttribute('aria-expanded', 'false');
  await summary.click();
  await expect(stepper).not.toHaveClass(/workflow-stepper--collapsed/);
  await expect(stepper.getByRole('button', { name: 'Collapse workflow' })).toBeVisible();
  await expectNoWcagViolations(page, '.workflow-stepper');

  await stepper.getByRole('button', { name: 'Collapse workflow' }).click();
  await expect(stepper).toHaveClass(/workflow-stepper--collapsed/);
});

test('diagram quality doctor finds and safely repairs visual issues', async ({ page }) => {
  const document = interactionCloudDocument();
  document.payload.nodes = [
    {
      id: 'node-a',
      type: 'azureNode',
      position: { x: 180, y: 220 },
      style: {
        width: 160,
        height: 128,
        color: '#777777',
        backgroundColor: '#888888',
      },
      data: {
        label: 'An exceptionally long Azure service label that needs more room',
        serviceName: 'App Service',
      },
    },
    {
      id: 'node-b',
      type: 'azureNode',
      position: { x: 220, y: 250 },
      style: { width: 160, height: 128 },
      data: { label: 'Azure SQL Database', serviceName: 'Azure SQL Database' },
    },
  ];
  document.payload.edges = [];
  await openInteractionDiagram(page, undefined, false, document);

  await page.getByRole('button', { name: 'More', exact: true }).click();
  await page.getByRole('button', { name: 'Diagram Quality Doctor' }).click();

  const doctor = page.getByRole('dialog', { name: 'Diagram Quality Doctor' });
  await expect(doctor).toBeVisible();
  await expect(doctor).toBeFocused();
  await expect(page.locator('#root')).toHaveAttribute('inert', '');
  await expect(page.locator('#root')).toHaveAttribute('aria-hidden', 'true');
  await page.keyboard.press('Shift+Tab');
  await expect(doctor.locator(':focus')).toHaveCount(1);
  await expect(doctor.getByText('Overlapping diagram elements')).toBeVisible();
  await expect(doctor.getByText('Crowded service label')).toBeVisible();
  await expect(doctor.getByText('Low node text contrast')).toBeVisible();
  await expect(doctor.getByText('Unconnected service').first()).toBeVisible();
  await expect(doctor.getByRole('button', { name: /Apply selected fixes/ })).toBeEnabled();
  await expectNoWcagViolations(page, '.quality-doctor-dialog');

  await doctor.getByRole('button', { name: /Apply selected fixes/ }).click();
  await expect(doctor).toBeHidden();
  await expect(page.locator('#root')).not.toHaveAttribute('inert', '');
  await expect(page.locator('#root')).not.toHaveAttribute('aria-hidden', 'true');

  const firstNode = page.locator('[data-testid="rf__node-node-a"]');
  const secondNode = page.locator('[data-testid="rf__node-node-b"]');
  await expectNotToOverlap(firstNode, secondNode);
  await expect(firstNode.locator('.node-label')).toHaveCSS('max-width', '260px');
  await expect(firstNode).not.toHaveCSS('color', 'rgb(119, 119, 119)');
  await expect(firstNode).not.toHaveCSS('background-color', 'rgb(136, 136, 136)');
});

test('diagram quality doctor stays usable on a mobile canvas', async ({ page }) => {
  const document = interactionCloudDocument();
  document.payload.nodes[1].position = { x: 190, y: 235 };
  await openInteractionDiagram(page, undefined, false, document);
  await page.setViewportSize({ width: 390, height: 844 });

  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' })
    .fill('quality doctor');
  await palette.getByRole('option', { name: /^Run Diagram Quality Doctor/ }).click();

  const doctor = page.getByRole('dialog', { name: 'Diagram Quality Doctor' });
  await expect(doctor).toBeVisible();
  await expect(doctor.getByRole('button', { name: /Apply selected fixes/ })).toBeVisible();
  const bounds = await doctor.boundingBox();
  expect(bounds?.width).toBe(390);
  expect(bounds?.height).toBe(844);
  await expectNoWcagViolations(page, '.quality-doctor-dialog');
  await page.keyboard.press('Escape');
  await expect(doctor).toBeHidden();
  await expect(canvas).toBeFocused();
});

test('cloud review supports anchored comments resolution requests and reports', async ({ page }, testInfo) => {
  let current: CloudDiagramDocument = {
    ...interactionCloudDocument(),
    access: 'owner',
    role: 'owner',
    etag: '"A-1"',
    review: { status: 'draft', updatedAt: now },
  };
  const updateCurrent = (changes: Partial<CloudDiagramDocument>) => {
    const revision = current.revision + 1;
    current = {
      ...current,
      ...changes,
      revision,
      updatedAt: new Date(Date.parse(now) + revision * 1_000).toISOString(),
      etag: `"A-${revision}"`,
    };
  };

  await initializePage(page, { documentId: 'A', access: 'owner', role: 'owner' });
  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: true,
        authenticated: true,
        email: 'owner@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, {
        documents: [{
          ...summary('A', current.diagramName),
          revision: current.revision,
          commentCount: current.comments.length,
          openCommentCount: current.comments.filter(comment => !comment.resolved).length,
          reviewStatus: current.review?.status,
          etag: current.etag,
        }],
      });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = request.postDataJSON() as {
        diagramName: string;
        payload: CloudDiagramDocument['payload'];
      };
      updateCurrent({ diagramName: body.diagramName, payload: body.payload });
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
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
    if (path === '/api/diagrams/A/comments' && method === 'POST') {
      const body = request.postDataJSON() as {
        message: string;
        anchor?: { type: 'canvas' | 'node' | 'edge'; targetId?: string; label?: string };
      };
      updateCurrent({
        comments: [
          ...current.comments,
          {
            commentId: 'comment-1',
            message: body.message,
            authorEmail: 'owner@example.com',
            authorId: 'owner',
            createdAt: now,
            anchor: body.anchor,
            resolved: false,
          },
        ],
      });
      await fulfillJson(route, { document: current }, 201, { etag: current.etag });
      return;
    }
    if (path === '/api/diagrams/A/comments/comment-1' && method === 'PATCH') {
      const body = request.postDataJSON() as { resolved: boolean };
      updateCurrent({
        comments: current.comments.map(comment => (
          comment.commentId === 'comment-1'
            ? {
                ...comment,
                resolved: body.resolved,
                resolvedAt: body.resolved ? now : undefined,
                resolvedByEmail: body.resolved ? 'owner@example.com' : undefined,
              }
            : comment
        )),
      });
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
      return;
    }
    if (path === '/api/diagrams/A/review' && method === 'PATCH') {
      updateCurrent({
        review: {
          status: 'in_review',
          requestedAt: now,
          requestedByEmail: 'owner@example.com',
          updatedAt: now,
        },
      });
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toBeVisible();
  await getCloudWorkspaceButton(page).click();

  let workspace = page.getByRole('dialog', { name: 'Cloud workspace' });
  const commentsCard = workspace.locator('.cloud-comments-card');
  await commentsCard.getByRole('combobox', { name: 'Attach to' }).selectOption({ label: 'App Service' });
  await commentsCard.getByPlaceholder('Add a review comment...').fill('Check the public endpoint.');
  await commentsCard.getByRole('button', { name: 'Comment' }).click();
  await expect(commentsCard.getByText('Check the public endpoint.')).toBeVisible();
  await expectNoWcagViolations(page, '.cloud-workspace-modal');

  await commentsCard.getByRole('button', { name: 'App Service' }).click();
  await expect(workspace).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toHaveClass(/selected/);

  await getCloudWorkspaceButton(page).click();
  workspace = page.getByRole('dialog', { name: 'Cloud workspace' });
  const reopenedComments = workspace.locator('.cloud-comments-card');
  await reopenedComments.getByRole('button', { name: 'Resolve' }).click();
  await expect(reopenedComments.getByText('Check the public endpoint.')).toBeHidden();
  await reopenedComments.getByRole('button', { name: 'Show resolved comments' }).click();
  await expect(reopenedComments.getByText('Resolved', { exact: true })).toBeVisible();

  const reviewCard = workspace.locator('.cloud-review-card');
  await reviewCard.getByRole('button', { name: 'Request review' }).click();
  await expect(reviewCard.getByText('In review', { exact: true })).toBeVisible();

  const downloadPromise = page.waitForEvent('download');
  await reviewCard.getByRole('button', { name: 'Download report' }).click();
  const download = await downloadPromise;
  const reportPath = testInfo.outputPath('cloud-review.md');
  await download.saveAs(reportPath);
  const report = await readFile(reportPath, 'utf8');
  expect(report).toContain('Review status:** In review');
  expect(report).toContain('Resolved comments:** 1');
  expect(report).toContain('Check the public endpoint.');
});

test('shared reviewers can approve the requested cloud revision', async ({ page }) => {
  const shareToken = 'a'.repeat(43);
  let current: CloudDiagramDocument = {
    ...interactionCloudDocument(),
    access: 'shared',
    role: 'viewer',
    etag: '"A-3"',
    review: {
      status: 'in_review',
      requestedAt: now,
      requestedByEmail: 'owner@example.com',
      updatedAt: now,
    },
  };
  await initializePage(page, {
    documentId: 'A',
    access: 'shared',
    role: 'viewer',
    shareToken,
  });
  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: true,
        authenticated: true,
        email: 'reviewer@example.com',
        isAdmin: false,
        allowed: true,
      });
      return;
    }
    if (path === '/api/diagrams' && method === 'GET') {
      await fulfillJson(route, { documents: [] });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}` && method === 'GET') {
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}/versions` && method === 'GET') {
      await fulfillJson(route, { versions: [] });
      return;
    }
    if (path === `/api/diagrams/shared/${shareToken}/review` && method === 'PATCH') {
      const body = request.postDataJSON() as { note?: string };
      current = {
        ...current,
        revision: current.revision + 1,
        etag: '"A-4"',
        review: {
          ...current.review!,
          status: 'approved',
          decidedAt: now,
          decidedByEmail: 'reviewer@example.com',
          decisionNote: body.note,
          updatedAt: now,
        },
      };
      await fulfillJson(route, { document: current }, 200, { etag: current.etag });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toBeVisible();
  await getCloudWorkspaceButton(page).click();
  const workspace = page.getByRole('dialog', { name: 'Cloud workspace' });
  const reviewCard = workspace.locator('.cloud-review-card');
  await reviewCard.getByLabel('Decision note (optional)').fill('Security review completed.');
  await reviewCard.getByRole('button', { name: 'Approve revision' }).click();
  await expect(reviewCard.getByText('Approved', { exact: true })).toBeVisible();
  await expect(reviewCard.getByText('Security review completed.')).toBeVisible();
  await expectNoWcagViolations(page, '.cloud-review-card');
});

test('multi-selection bulk edit applies grouping styling tags region and pricing', async ({ page }) => {
  let latestSavedPayload: Record<string, unknown> | null = null;
  const document = interactionCloudDocument();
  document.payload.nodes.push({
    id: 'group-b',
    type: 'groupNode',
    position: { x: 820, y: 560 },
    style: { width: 180, height: 100 },
    data: { label: 'Data layer', color: '#10b981' },
  });
  await openInteractionDiagram(page, payload => {
    latestSavedPayload = payload;
  }, false, document);

  const firstNode = page.locator('[data-testid="rf__node-node-a"]');
  const secondNode = page.locator('[data-testid="rf__node-node-b"]');
  const groupNode = page.locator('[data-testid="rf__node-group-a"]');
  const secondGroupNode = page.locator('[data-testid="rf__node-group-b"]');
  await firstNode.click();
  await secondNode.click({ modifiers: ['Control'] });

  const toolbar = page.locator('.alignment-toolbar');
  await expect(toolbar).toBeVisible();
  await expect(toolbar).toContainText('2 selected');
  await toolbar.getByRole('button', { name: 'Bulk edit' }).click();

  const editor = page.getByRole('dialog', { name: 'Bulk edit selected items' });
  await expect(editor).toBeVisible();
  await editor.getByLabel('Group', { exact: true }).selectOption({ label: 'Application layer' });
  await editor.getByLabel('Azure region').selectOption('eastus2');
  await editor.getByLabel('Service style').selectOption('presentation');
  await editor.getByLabel('Tags (comma-separated)').fill('production, critical');
  await editor.getByLabel('Quantity').fill('2');
  await editor.getByLabel('Custom monthly unit price (USD)').fill('125');
  await expectNoWcagViolations(page, '.bulk-edit-popover');
  await editor.getByRole('button', { name: 'Apply to selection' }).click();
  await expect(editor).toBeHidden();

  await expect(firstNode.locator('.azure-node')).toHaveClass(/style-presentation/);
  await expect(secondNode.locator('.azure-node')).toHaveClass(/style-presentation/);
  await expect(firstNode.locator('.node-tags')).toContainText('production');
  await expect(firstNode.locator('.node-tags')).toContainText('critical');

  await expect.poll(() => {
    const payload = latestSavedPayload as {
      nodes?: Array<{
        id?: string;
        parentNode?: string;
        data?: {
          groupId?: string;
          stylePreset?: string;
          tags?: string[];
          customColor?: { name?: string };
          pricing?: {
            estimatedCost?: number;
            quantity?: number;
            region?: string;
            isCustom?: boolean;
          };
        };
      }>;
    } | null;
    const savedFirst = payload?.nodes?.find(node => node.id === 'node-a');
    const savedSecond = payload?.nodes?.find(node => node.id === 'node-b');
    return {
      firstParent: savedFirst?.parentNode,
      secondParent: savedSecond?.parentNode,
      firstGroupId: savedFirst?.data?.groupId,
      style: savedFirst?.data?.stylePreset,
      tags: savedFirst?.data?.tags,
      price: savedFirst?.data?.pricing?.estimatedCost,
      quantity: savedFirst?.data?.pricing?.quantity,
      region: savedFirst?.data?.pricing?.region,
      custom: savedFirst?.data?.pricing?.isCustom,
    };
  }).toEqual({
    firstParent: 'group-a',
    secondParent: 'group-a',
    firstGroupId: 'group-a',
    style: 'presentation',
    tags: ['production', 'critical'],
    price: 125,
    quantity: 2,
    region: 'eastus2',
    custom: true,
  });

  await toolbar.getByRole('button', { name: 'Align left' }).click();
  await expect.poll(async () => {
    const [firstBox, secondBox] = await Promise.all([
      firstNode.boundingBox(),
      secondNode.boundingBox(),
    ]);
    return firstBox && secondBox ? Math.abs(firstBox.x - secondBox.x) : 999;
  }).toBeLessThan(1);

  await groupNode.click();
  await secondGroupNode.click({ modifiers: ['Control'] });
  await expect(toolbar).toContainText('2 selected');
  await toolbar.getByRole('button', { name: 'Bulk edit' }).click();
  const groupEditor = page.getByRole('dialog', { name: 'Bulk edit selected items' });
  await groupEditor.getByLabel('Group color', { exact: true }).selectOption('Blue');
  await groupEditor.getByRole('button', { name: 'Apply to selection' }).click();
  await expect(groupEditor).toBeHidden();
  await expect.poll(() => {
    const payload = latestSavedPayload as {
      nodes?: Array<{ id?: string; data?: { customColor?: { name?: string } } }>;
    } | null;
    return payload?.nodes
      ?.filter(node => node.id === 'group-a' || node.id === 'group-b')
      .map(node => node.data?.customColor?.name);
  }).toEqual(['Blue', 'Blue']);
});

test('bulk editor remains usable on a mobile canvas', async ({ page }) => {
  await openInteractionDiagram(page);
  await page.setViewportSize({ width: 390, height: 844 });
  const firstNode = page.locator('[data-testid="rf__node-node-a"]');
  const secondNode = page.locator('[data-testid="rf__node-node-b"]');
  await firstNode.click();
  await secondNode.click({ modifiers: ['Control'] });

  const toolbar = page.locator('.alignment-toolbar');
  await expect(toolbar).toBeVisible();
  await toolbar.getByRole('button', { name: 'Bulk edit' }).click();
  const editor = page.getByRole('dialog', { name: 'Bulk edit selected items' });
  await expect(editor).toBeVisible();
  await expect(editor.getByLabel('Tags (comma-separated)')).toBeVisible();
  await expect(editor.getByRole('button', { name: 'Apply to selection' })).toBeVisible();
  const bounds = await editor.boundingBox();
  expect(bounds?.width).toBeLessThanOrEqual(370);
  await expectNoWcagViolations(page, '.bulk-edit-popover');
});

test('cloud workspace owns opaque readable modal surfaces', async ({ page }) => {
  await initializePage(page);
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
      await fulfillJson(route, { documents: [summary('A', 'Diagram A')] });
      return;
    }
    if (path === '/api/diagrams/A' && method === 'GET') {
      await fulfillJson(route, cloudDocument('A', 'Diagram A'), 200, { etag: '"A-1"' });
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
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  await getCloudWorkspaceButton(page).click();

  const modal = page.getByRole('dialog', { name: 'Cloud workspace' });
  const surfaces = modal.locator(
    ':scope, :scope > .modal-header, .cloud-document-details, :scope > .modal-actions',
  );
  await expect(modal.getByRole('heading', { name: 'Diagram A' })).toBeVisible();

  for (let index = 0; index < await surfaces.count(); index += 1) {
    const alpha = await surfaces.nth(index).evaluate((element) => {
      const color = getComputedStyle(element).backgroundColor;
      const channels = color
        .replace(/^rgba?\(/, '')
        .replace(/\)$/, '')
        .split(/[\s,/]+/)
        .filter(Boolean);
      return channels.length > 3 ? Number(channels[3]) : 1;
    });
    expect(alpha).toBe(1);
  }

  await expectReadableContrast(modal.getByRole('heading', { name: 'Cloud workspace' }));
  await expectReadableContrast(modal.getByRole('button', { name: 'Open', exact: true }));
  await expectReadableContrast(modal.getByRole('button', { name: 'Close', exact: true }).last());
});

test('command palette adds services and focus mode persists until Escape', async ({ page }) => {
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
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Control+K');

  const palette = page.getByTestId('command-palette');
  const search = palette.getByRole('combobox', { name: 'Search commands and services' });
  await expect(palette).toBeVisible();
  await expect(search).toBeFocused();
  await expect(page.locator('.app-header')).toHaveAttribute('inert', '');
  await expect(page.locator('.workspace')).toHaveAttribute('inert', '');
  await expectNoWcagViolations(page, '[data-testid="command-palette"]');

  await search.fill('App Services');
  await palette.getByRole('option', { name: /App Services/ }).click();
  await expect(palette).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(1);
  // Recovery requires a committed draft, not just a rendered node.
  await expect(page.getByRole('status').filter({ hasText: 'Saved on this device' })).toBeVisible();

  await canvas.focus();
  await page.keyboard.press('Control+K');
  await palette.getByRole('combobox', { name: 'Search commands and services' })
    .fill('focus mode');
  await palette.getByRole('option', { name: /^Enter focus mode/ }).click();

  await expect(page.locator('.app')).toHaveClass(/focus-mode/);
  await expect(page.locator('.app-header')).toBeHidden();
  await expect(page.getByRole('button', { name: 'Exit Focus' })).toBeFocused();
  await expect(page.locator('.nav-minimap')).toBeHidden();
  await expect(page.locator('.title-block')).toBeHidden();
  await expect(page.locator('.legend')).toBeHidden();
  await expect.poll(() => page.evaluate(() => (
    localStorage.getItem('azure-diagram-builder.focusMode.v1')
  ))).toBe('1');

  await page.reload();
  await expect(page.locator('.app')).toHaveClass(/focus-mode/);
  const recovery = page.getByRole('dialog', { name: 'Resume your saved draft' });
  await expect(recovery).toBeVisible();
  await expect(recovery).toHaveAttribute('aria-modal', 'true');
  await page.keyboard.press('Escape');
  await expect(recovery).toBeVisible();
  await expect(page.locator('.app')).toHaveClass(/focus-mode/);
  await recovery.getByRole('button', { name: 'Restore draft', exact: true }).click();
  await expect(recovery).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(1);
  const exitFocus = page.getByRole('button', { name: 'Exit Focus' });
  await expect(exitFocus).toBeVisible();
  await exitFocus.focus();
  await page.keyboard.press('Escape');
  await expect(page.locator('.app')).not.toHaveClass(/focus-mode/);
  await expect(page.locator('.app-header')).toBeVisible();
  await expect(canvas).toBeFocused();
  await expect.poll(() => page.evaluate(() => (
    localStorage.getItem('azure-diagram-builder.focusMode.v1')
  ))).toBe('0');
});

test('recent work restores an interrupted local diagram after reload', async ({ page }) => {
  await initializePage(page);
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
    if (path === '/api/diagrams' && request.method() === 'GET') {
      await fulfillJson(route, { error: 'Temporary cloud outage' }, 503);
      return;
    }
    if (path === '/api/diagrams' && request.method() === 'POST') {
      await fulfillJson(route, { error: 'Temporary cloud outage' }, 503);
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  const canvas = page.getByRole('region', { name: 'Architecture canvas' });
  await canvas.focus();
  await page.keyboard.press('Control+K');
  const palette = page.getByTestId('command-palette');
  await palette.getByRole('combobox', { name: 'Search commands and services' })
    .fill('App Services');
  await palette.getByRole('option', { name: /App Services/ }).click();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(1);
  await page.waitForTimeout(1_200);

  await page.getByRole('button', { name: 'More' }).click();
  await page.getByRole('button', { name: 'Resume recent work' }).click();
  const currentRecentWork = page.getByRole('dialog', { name: 'Recent work' });
  await expect(currentRecentWork.getByRole('button', { name: 'Return' })).toBeVisible();
  await currentRecentWork.getByRole('button', { name: 'Close recent work' }).click();

  await page.evaluate(() => {
    sessionStorage.removeItem('azurediagarm.recent-work-session.v1');
  });
  await page.reload();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(0);

  // Resolve the new autosave prompt explicitly before exercising the separate
  // recent-work catalog. Discarding the active draft must not erase its archive.
  const recovery = page.getByRole('dialog', { name: 'Resume your saved draft' });
  await expect(recovery).toBeVisible();
  await recovery.getByRole('button', { name: 'Start without this draft' }).click();
  await expect(recovery.getByRole('button', { name: 'Delete draft and start new' })).toBeVisible();
  await recovery.getByRole('button', { name: 'Keep draft' }).click();
  await expect(recovery.getByRole('button', { name: 'Restore draft', exact: true })).toBeEnabled();
  await recovery.getByRole('button', { name: 'Start without this draft' }).click();
  await recovery.getByRole('button', { name: 'Delete draft and start new' }).click();
  await expect(recovery).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(0);

  await page.getByRole('button', { name: 'More' }).click();
  await page.getByRole('button', { name: 'Resume recent work' }).click();
  const recentWork = page.getByRole('dialog', { name: 'Recent work' });
  await expect(recentWork).toBeVisible();
  await expect(recentWork.getByText('Recovered sessions')).toBeVisible();
  await expectNoWcagViolations(page, '.recent-work-modal');
  await recentWork.getByRole('button', { name: 'Resume' }).click();

  await expect(recentWork).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(1);
});

test('version history compares and selectively restores diagram elements', async ({ page }) => {
  // Adds two services through the command palette, snapshots, diffs, and then
  // selectively restores — enough steps to exceed the default budget on a cold
  // CI runner.
  test.slow();
  await initializePage(page);
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
    if (path === '/api/diagrams' && request.method() === 'POST') {
      await fulfillJson(route, { error: 'Temporary cloud outage' }, 503);
      return;
    }
    await fulfillJson(route, { error: 'Not found' }, 404);
  });

  await page.goto('/');
  const addService = async (name: string) => {
    const canvas = page.getByRole('region', { name: 'Architecture canvas' });
    await canvas.focus();
    await page.keyboard.press('Control+K');
    const palette = page.getByTestId('command-palette');
    await palette.getByRole('combobox', { name: 'Search commands and services' }).fill(name);
    await palette.getByRole('option', { name: new RegExp(name) }).first().click();
  };

  await addService('App Services');
  await page.getByRole('button', { name: 'Snapshot', exact: true }).click();
  const snapshotModal = page.getByRole('dialog', { name: 'Save Snapshot' });
  await snapshotModal.locator('#snapshot-notes').fill('Baseline snapshot');
  await snapshotModal.getByRole('button', { name: 'Save Snapshot' }).click();
  await expect(snapshotModal).toBeHidden();

  await addService('Storage Accounts');
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(2);
  await page.getByTitle('View version history').click();

  const history = page.getByRole('dialog', { name: 'Version History' });
  await history.locator('.version-item').filter({ hasText: 'Baseline snapshot' }).click();
  const comparison = history.getByRole('region', { name: 'Visual version comparison' });
  await expect(comparison).toBeVisible();
  await expect(comparison.getByText('−1')).toBeVisible();
  await expectNoWcagViolations(page, '.version-history-modal');

  await comparison.getByRole('button', { name: 'Select all' }).click();
  page.on('dialog', dialog => dialog.accept());
  await comparison.getByRole('button', { name: /Apply selected/ }).click();

  await expect(history).toBeHidden();
  await expect(page.locator('.react-flow__node-azureNode')).toHaveCount(1);
});

test('canvas uses neutral defaults and brand emphasis only for selection and flow', async ({ page }) => {
  await openInteractionDiagram(page);

  const node = page.locator('[data-testid="rf__node-node-a"]');
  const nodeCard = node.locator('.azure-node');
  const group = page.locator('[data-testid="rf__node-group-a"] .group-node');
  const edge = page.locator('[data-testid="rf__edge-edge-ab"]');
  const edgePath = edge.locator('.react-flow__edge-path');
  const edgeLabel = page.locator('[data-edge-label-id="edge-ab"]');

  await expect(edgePath).toHaveCSS('stroke', 'rgb(100, 116, 139)');
  await expect(edgeLabel).toHaveCSS('background-color', 'rgba(255, 255, 255, 0.92)');
  await expect(group).toHaveCSS('background-color', 'rgba(107, 114, 128, 0.08)');

  await node.click();
  // Solid brand blue, not the 30 %-alpha soft token: composited over the white
  // card that resolved to ~1.3 : 1, far below the 3 : 1 WCAG needs for a
  // non-text indicator. #0f6cbd on white is 5.4 : 1.
  await expect(nodeCard).toHaveCSS('outline-color', 'rgb(15, 108, 189)');
  await expect(node).toHaveCSS('box-shadow', 'none');

  await edge.locator('.react-flow__edge-interaction').click({ force: true });
  await expect(edge).toHaveClass(/selected/);
  await expect(edgePath).toHaveCSS('stroke', 'rgb(15, 108, 189)');
});

test('flat diagram styling preserves authored colors, paths and stationary connection anchors', async ({ page }) => {
  const base = interactionCloudDocument();
  const customColor = { bg: 'rgba(139, 0, 0, 0.1)', border: '#8b0000', header: '#8b0000' };
  await openInteractionDiagram(page, undefined, false, {
    ...base,
    payload: {
      ...base.payload,
      nodes: base.payload.nodes.map(node => node.type === 'groupNode'
        ? { ...node, data: { ...node.data, customColor } }
        : node),
      edges: base.payload.edges.map(edge => ({
        ...edge,
        label: 'Authored connection',
        style: { stroke: '#b45309', strokeWidth: 3, strokeDasharray: '9, 4', opacity: 0.8 },
        data: { connectionType: 'security', pathStyle: 'straight', direction: 'reverse' },
      })),
    },
  });

  const node = page.getByTestId('rf__node-node-a').locator('.azure-node');
  const path = page.getByTestId('rf__edge-edge-ab').locator('.react-flow__edge-path');
  const group = page.getByTestId('rf__node-group-a').locator('.group-node');
  const route = await path.getAttribute('d');
  const bounds = await node.boundingBox();

  await expect(node).toHaveCSS('border-radius', '6px');
  await expect(node).toHaveCSS('box-shadow', 'none');
  await node.hover();
  await expect(node).toHaveCSS('transform', 'none');
  expect(await node.boundingBox()).toEqual(bounds);
  await expect(path).toHaveAttribute('d', route!);
  await expect(path).toHaveCSS('stroke', 'rgb(180, 83, 9)');
  await expect(path).toHaveCSS('stroke-width', '3px');
  await expect(path).toHaveCSS('stroke-dasharray', '9px, 4px');
  await expect(path).toHaveCSS('opacity', '0.8');
  await expect(group).toHaveCSS('background-color', customColor.bg);
  await expect(group).toHaveCSS('border-color', 'rgb(139, 0, 0)');
  await expect(group.locator('.group-label')).toHaveCSS('color', 'rgb(139, 0, 0)');

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.getByRole('button', { name: 'Switch to Dark Mode' }).click();
  await expect(path).toHaveCSS('stroke', 'rgb(180, 83, 9)');
  await expect(path).toHaveCSS('stroke-dasharray', '9px, 4px');
  await expect(path).toHaveCSS('opacity', '0.8');
  await expect(path).toHaveAttribute('d', route!);
  await expect(group).toHaveCSS('background-color', customColor.bg);
  await expect(group).toHaveCSS('border-color', 'rgb(139, 0, 0)');
  await expect(group.locator('.group-label')).toHaveCSS('color', 'rgb(139, 0, 0)');
});

test('restored connection styles use per-field defaults without losing solid overrides or authored markers', async ({ page }) => {
  const base = interactionCloudDocument();
  const edges = [
    {
      ...base.payload.edges[0],
      id: 'edge-solid-zero',
      label: 'Authored solid connection',
      style: { stroke: '#b45309', strokeDasharray: 'none', opacity: 0 },
      markerStart: { type: 'arrowclosed', color: '#2563eb' },
      markerEnd: { type: 'arrowclosed', color: '#be185d' },
      data: {
        connectionType: 'optional', pathStyle: 'orthogonal', direction: 'bidirectional',
        labelOffsetX: 38, labelOffsetY: -18, labelOffsetAuto: false,
      },
    },
    {
      ...base.payload.edges[0],
      id: 'edge-partial',
      style: { stroke: '#15803d', strokeWidth: 3 },
      markerEnd: { type: 'arrowclosed' },
      data: { connectionType: 'optional', pathStyle: 'smooth', direction: 'forward' },
    },
    {
      ...base.payload.edges[0],
      id: 'edge-semantic',
      markerStart: { type: 'arrowclosed' },
      data: { connectionType: 'security', pathStyle: 'straight', direction: 'reverse' },
    },
  ];
  let savedEdges: Record<string, unknown>[] = [];
  await openInteractionDiagram(page, payload => {
    savedEdges = Array.isArray(payload.edges) ? payload.edges as Record<string, unknown>[] : [];
  }, false, { ...base, connectionCount: edges.length, payload: { ...base.payload, edges } });

  const paths = edges.map(edge => page.getByTestId(`rf__edge-${edge.id}`).locator('.react-flow__edge-path'));
  for (const path of paths) await expect(path).toHaveAttribute('d', /^M/);
  const routes = await Promise.all(paths.map(path => path.getAttribute('d')));
  const markerFill = (path: Locator, attribute: 'marker-start' | 'marker-end') =>
    path.evaluate((element, name) => {
      const id = element.getAttribute(name)?.match(/^url\(["']?#(.*?)["']?\)$/)?.[1];
      const shape = id ? document.getElementById(id)?.querySelector('polyline, path') : null;
      return shape ? getComputedStyle(shape).fill : null;
    }, attribute);

  for (const dark of [false, true]) {
    await page.evaluate(enabled => document.body.classList.toggle('dark-mode', enabled), dark);
    await expect(paths[0]).toHaveCSS('stroke', 'rgb(180, 83, 9)');
    await expect(paths[0]).toHaveCSS('stroke-dasharray', 'none');
    await expect(paths[0]).toHaveCSS('opacity', '0');
    expect(await markerFill(paths[0], 'marker-start')).toBe('rgb(37, 99, 235)');
    expect(await markerFill(paths[0], 'marker-end')).toBe('rgb(190, 24, 93)');
    await expect(paths[1]).toHaveCSS('stroke', 'rgb(21, 128, 61)');
    await expect(paths[1]).toHaveCSS('stroke-width', '3px');
    await expect(paths[1]).toHaveCSS('stroke-dasharray', '2px, 4px');
    await expect(paths[1]).toHaveCSS('opacity', '0.68');
    expect(await markerFill(paths[1], 'marker-end')).toBe('rgb(21, 128, 61)');
    await expect(paths[2]).toHaveCSS('stroke', 'rgb(220, 38, 38)');
    await expect(paths[2]).toHaveCSS('stroke-dasharray', '2px, 3px');
    await expect(paths[2]).toHaveCSS('opacity', '1');
    expect(await markerFill(paths[2], 'marker-start')).toBe('rgb(220, 38, 38)');
    expect(await Promise.all(paths.map(path => path.getAttribute('d')))).toEqual(routes);
  }

  await page.locator('[data-edge-label-id="edge-solid-zero"]').press('Enter');
  const input = page.locator('.editable-edge-label-input');
  await input.fill('Restored explicit style');
  await input.press('Enter');
  await expect.poll(
    () => savedEdges.find(edge => edge.id === 'edge-solid-zero')?.label,
    { timeout: 5_000 },
  ).toBe('Restored explicit style');
  expect(savedEdges.find(edge => edge.id === 'edge-solid-zero')).toMatchObject({
    ...edges[0],
    label: 'Restored explicit style',
  });
  expect(savedEdges.find(edge => edge.id === 'edge-partial')).toMatchObject({
    ...edges[1],
    style: { ...edges[1].style, strokeDasharray: '2, 4', opacity: 0.68 },
    markerEnd: { type: 'arrowclosed', color: '#15803d' },
  });
  expect(savedEdges.find(edge => edge.id === 'edge-semantic')).toMatchObject({
    ...edges[2],
    style: { stroke: '#dc2626', strokeDasharray: '2, 3' },
    markerStart: { type: 'arrowclosed', color: '#dc2626' },
  });
  expect(await Promise.all(paths.map(path => path.getAttribute('d')))).toEqual(routes);
});

test('automatic group headings stay neutral and readable without changing the zone tint', async ({ page }) => {
  const base = interactionCloudDocument();
  await openInteractionDiagram(page, undefined, false, {
    ...base,
    payload: {
      ...base.payload,
      nodes: base.payload.nodes.map(node => node.type === 'groupNode'
        ? { ...node, data: { ...node.data, label: 'Data layer' } }
        : node),
    },
  });
  const group = page.getByTestId('rf__node-group-a').locator('.group-node');
  for (const dark of [false, true]) {
    await page.evaluate(enabled => document.body.classList.toggle('dark-mode', enabled), dark);
    await expect(group).toHaveCSS('background-color', 'rgba(16, 185, 129, 0.1)');
    await expectReadableContrast(group.locator('.group-label'));
    await expectReadableContrast(group.locator('.fit-to-content-button'), 3);
    await expectReadableContrast(group.locator('.color-picker-button'), 3);
  }
});

test('Japanese service labels avoid orphan characters without changing diagram geometry', async ({ page }) => {
  const base = interactionCloudDocument();
  await openInteractionDiagram(page, undefined, false, {
    ...base,
    payload: {
      ...base.payload,
      nodes: base.payload.nodes.map(node => node.id === 'node-a'
        ? { ...node, data: { ...node.data, label: '顧客向け Web アプリ' } }
        : node),
    },
  });
  const node = page.getByTestId('rf__node-node-a');
  const label = node.locator('.node-label');
  const path = page.getByTestId('rf__edge-edge-ab').locator('.react-flow__edge-path');
  const lines = () => label.evaluate(element => {
    const text = element.firstChild;
    if (!(text instanceof Text)) throw new Error('Expected the authored service label');
    const rows = new Map<string, string>();
    let offset = 0;
    for (const character of text.data) {
      const range = document.createRange();
      range.setStart(text, offset);
      range.setEnd(text, offset + character.length);
      offset += character.length;
      const top = range.getBoundingClientRect().top.toFixed(2);
      rows.set(top, (rows.get(top) ?? '') + character);
    }
    return [...rows.values()].map(row => row.trim()).filter(Boolean);
  });

  const greedyStyle = await page.addStyleTag({ content: '.node-label { text-wrap: wrap; }' });
  await expect(label).toHaveCSS('text-wrap', 'wrap');
  // System fonts split the greedy word at different characters.
  const greedyLines = await lines();
  expect(greedyLines).toHaveLength(2);
  expect(greedyLines.join('')).toBe('顧客向け Web アプリ');
  expect(greedyLines.every(line => !line.includes('アプリ'))).toBe(true);
  const beforeNode = await node.boundingBox();
  const beforeLabel = await label.boundingBox();
  const beforeRoute = await path.getAttribute('d');
  const beforeTypography = await label.evaluate(element => {
    const style = getComputedStyle(element);
    return [style.fontFamily, style.fontSize, style.lineHeight, style.padding, style.maxWidth];
  });

  await greedyStyle.evaluate(element => element.remove());
  await expect(label).toHaveCSS('text-wrap', 'balance');
  expect(await lines()).toEqual(['顧客向け', 'Web アプリ']);
  expect(await node.boundingBox()).toEqual(beforeNode);
  expect(await label.boundingBox()).toEqual(beforeLabel);
  await expect(path).toHaveAttribute('d', beforeRoute!);
  expect(await label.evaluate(element => {
    const style = getComputedStyle(element);
    return [style.fontFamily, style.fontSize, style.lineHeight, style.padding, style.maxWidth];
  })).toEqual(beforeTypography);
});

test('mobile minimap drawing and hit area stay inside its frame with accessible canvas controls', async ({ page }) => {
  await openInteractionDiagram(page);
  await page.setViewportSize({ width: 390, height: 844 });
  const miniMap = page.locator('.nav-minimap');
  const svg = miniMap.locator('svg').first();
  await expect(svg).toHaveAttribute('width', '140');
  await expect(svg).toHaveAttribute('height', '100');
  const [frame, drawing] = await Promise.all([miniMap.boundingBox(), svg.boundingBox()]);
  expect(frame).not.toBeNull();
  expect(drawing).not.toBeNull();
  expect(drawing!.x).toBeGreaterThanOrEqual(frame!.x);
  expect(drawing!.y).toBeGreaterThanOrEqual(frame!.y);
  expect(drawing!.x + drawing!.width).toBeLessThanOrEqual(frame!.x + frame!.width);
  expect(drawing!.y + drawing!.height).toBeLessThanOrEqual(frame!.y + frame!.height);
  for (const button of await page.locator('.react-flow__controls-button, .title-block-toggle').all()) {
    const box = await button.boundingBox();
    expect(box!.width).toBeGreaterThanOrEqual(44);
    expect(box!.height).toBeGreaterThanOrEqual(44);
  }
  await page.getByRole('navigation', { name: 'Mobile command bar' })
    .getByRole('button', { name: 'Services', exact: true }).click();
  for (const button of await page.locator('.palette-close-toggle, .palette-layout-switch button').all()) {
    const box = await button.boundingBox();
    expect(box!.width).toBeGreaterThanOrEqual(44);
    expect(box!.height).toBeGreaterThanOrEqual(44);
  }
});

test('mobile fit shows a wide authored diagram without changing its route', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  const base = interactionCloudDocument();
  await openInteractionDiagram(page, undefined, false, {
    ...base,
    payload: {
      ...base.payload,
      nodes: base.payload.nodes.map(node => node.id === 'node-b'
        ? { ...node, position: { x: 1500, y: 220 } }
        : node),
    },
  });
  const path = page.getByTestId('rf__edge-edge-ab').locator('.react-flow__edge-path');
  const route = await path.getAttribute('d');
  await page.locator('.react-flow__controls-fitview').click();
  await expect.poll(() => page.locator('.canvas-container').evaluate(canvas => {
    const bounds = canvas.getBoundingClientRect();
    return [...canvas.querySelectorAll('.react-flow__node')].every(node => {
      const box = node.getBoundingClientRect();
      return box.left >= bounds.left && box.right <= bounds.right
        && box.top >= bounds.top && box.bottom <= bounds.bottom;
    });
  })).toBe(true);
  await expect(path).toHaveAttribute('d', route!);
});

test('diagram metadata starts compact and keeps keyboard editing available', async ({ page }) => {
  await openInteractionDiagram(page);
  const metadata = page.locator('.title-block');
  const toggle = metadata.locator('.title-block-toggle');
  await expect(metadata).toHaveClass(/collapsed/);
  await expect(metadata.locator('.title-block-label')).toHaveText('Interaction diagram');
  await expect(toggle).toHaveAttribute('aria-expanded', 'false');
  await toggle.press('Enter');
  await expect(toggle).toHaveAttribute('aria-expanded', 'true');
  await metadata.locator('.title-block-display').press('Enter');
  await metadata.getByLabel('Author:', { exact: true }).fill('Diagram author');
  await metadata.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(metadata.locator('.title-block-display')).toContainText('Diagram author');
  await toggle.press('Enter');
  await expect(metadata).toHaveClass(/collapsed/);
  await expect(metadata.locator('.title-block-label')).toHaveText('Interaction diagram');
});

test('canvas chrome keeps navigation, metadata, controls, and feedback in separate zones', async ({ page }) => {
  await openInteractionDiagram(page, undefined, true);

  const controls = page.locator('.react-flow__controls');
  const legend = page.locator('.legend.collapsed');
  const miniMap = page.locator('.nav-minimap');
  const feedback = page.getByRole('button', { name: 'Feedback' });
  const navigationHint = page.getByRole('note', { name: 'Canvas navigation tips' });
  const titleBlock = page.locator('.title-block');

  await expect(navigationHint).toHaveCSS('transform', 'none');
  await expectNotToOverlap(controls, legend);
  await expectNotToOverlap(miniMap, feedback);
  await expectNotToOverlap(navigationHint, titleBlock);

  await page.setViewportSize({ width: 390, height: 844 });
  await expect(page.getByRole('navigation', { name: 'Mobile command bar' })).toBeVisible();
  await expectNotToOverlap(controls, legend);
  await expectNotToOverlap(miniMap, feedback);
  await expectNotToOverlap(navigationHint, titleBlock);
});

test('menu states keep readable contrast in light and dark themes', async ({ page }) => {
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

  await page.setViewportSize({ width: 1502, height: 768 });
  await page.goto('/');

  const checkMenuStates = async (expectedColorScheme: 'light' | 'dark') => {
    await page.getByRole('tab', { name: 'Home' }).click();
    const regionButton = page.locator('.region-selector-button');
    await regionButton.click();
    await expect(regionButton).toHaveAttribute('aria-expanded', 'true');
    await expectSelectedMenuState(regionButton);
    const selectedRegion = page.locator('.region-option.selected');
    await expectReadableContrast(selectedRegion.locator('.region-display-name'));
    await expectReadableContrast(selectedRegion.locator('.region-location'));
    await regionButton.click();

    await page.getByRole('tab', { name: 'Create' }).click();
    const modelButton = page.locator('.model-popover-trigger');
    await modelButton.click();
    await expect(modelButton).toHaveAttribute('aria-expanded', 'true');
    await expectSelectedMenuState(modelButton);
    const modelMenu = page.locator('.toolbar-dropdown-menu--model-settings');
    await expectReadableContrast(modelMenu.locator('.toolbar-dropdown-heading').first());
    await expectReadableContrast(modelMenu.locator('.msp-managed-note strong'));
    const selectedReasoning = modelMenu.locator('.msp-reasoning-btn.active');
    await expect(selectedReasoning).toHaveCSS('color', 'rgb(255, 255, 255)');
    await expectReadableContrast(selectedReasoning);
    await modelMenu.locator('.msp-close-btn').click();

    await page.getByRole('tab', { name: 'Design' }).click();
    await expect(page.locator('body')).toHaveCSS('color-scheme', expectedColorScheme);

    const designTab = page.getByRole('tab', { name: 'Design' });
    const flowMotion = page.getByRole('button', { name: 'Flow motion' });
    const selectButton = page.getByRole('button', { name: 'Select', exact: true });
    const collapseGroups = page.getByRole('button', { name: 'Collapse Groups' });

    await expectReadableContrast(designTab);
    await expectSelectedMenuState(flowMotion);
    await flowMotion.hover();
    await expectSelectedMenuState(flowMotion);

    for (const disabledControl of [selectButton, collapseGroups]) {
      await expect(disabledControl).toBeDisabled();
      await expect(disabledControl).toHaveCSS('opacity', '1');
      await expectReadableContrast(disabledControl);
    }

    const layoutButton = page.getByRole('button', { name: 'Layout', exact: true });
    await layoutButton.click();
    await expect(layoutButton).toHaveAttribute('aria-expanded', 'true');
    await expectSelectedMenuState(layoutButton);

    const layoutMenu = page.getByRole('menu', { name: 'Layout options' });
    for (const select of await layoutMenu.locator('select').all()) {
      await expectReadableContrast(select);
    }
    await expectReadableContrast(layoutMenu.locator('.toolbar-dropdown-heading'));
    await expectReadableContrast(layoutMenu.locator('.toolbar-dropdown-hint'));
    const applyLayout = layoutMenu.getByRole('menuitem', { name: 'Apply Layout' });
    await expect(applyLayout).toBeDisabled();
    await expect(applyLayout).toHaveCSS('opacity', '1');
    await expectReadableContrast(applyLayout);
    await page.keyboard.press('Escape');

    const styleButton = page.getByRole('button', { name: 'Style', exact: true });
    await styleButton.click();
    await expectSelectedMenuState(styleButton);
    const styleMenu = page.getByRole('menu', { name: 'Style preset options' });
    const selectedStyle = styleMenu.getByRole('menuitem', { name: 'Detailed (Default)' });
    await selectedStyle.hover();
    await expectSelectedMenuState(selectedStyle);
    await expectReadableContrast(
      styleMenu.getByRole('menuitem', { name: 'Presentation (Professional)' }),
    );
    await page.keyboard.press('Escape');
  };

  await checkMenuStates('light');
  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByRole('button', { name: 'Switch to Dark Mode' }).click();
  await checkMenuStates('dark');
  await expectNoWcagViolations(page, '.app-header');
});

test('dark context and alignment menus remain readable', async ({ page }) => {
  await openInteractionDiagram(page);
  await page.getByRole('tab', { name: 'Home' }).click();
  await page.getByRole('button', { name: 'Switch to Dark Mode' }).click();

  const nodeA = page.locator('[data-testid="rf__node-node-a"]');
  const nodeB = page.locator('[data-testid="rf__node-node-b"]');
  await nodeA.click();
  await nodeB.click({ modifiers: ['Control'] });

  const alignmentToolbar = page.locator('.alignment-toolbar');
  await expect(alignmentToolbar).toBeVisible();
  await expectReadableContrast(alignmentToolbar.locator('.toolbar-label').first());
  await expectReadableContrast(alignmentToolbar.locator('.toolbar-info'));
  const alignLeft = alignmentToolbar.getByRole('button', { name: 'Align Left' });
  await expectReadableContrast(alignLeft);
  await alignLeft.hover();
  await expectReadableContrast(alignLeft);

  await nodeA.click({ button: 'right' });
  const nodeMenu = page.getByRole('menu', { name: 'Service actions' });
  await expectReadableContrast(nodeMenu.locator('.context-menu-header'));
  const duplicate = nodeMenu.getByRole('menuitem', { name: 'Duplicate service' });
  await expect(duplicate).toBeFocused();
  await expectReadableContrast(duplicate);
  const deleteService = nodeMenu.getByRole('menuitem', { name: 'Delete service' });
  await deleteService.hover();
  await expectReadableContrast(deleteService);
  await expectNoWcagViolations(page, '.node-context-menu');
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
  const pricingEditor = page.getByRole('dialog', { name: 'Service inspector' });
  await expect(pricingEditor.getByLabel('Label', { exact: true })).toBeFocused();
  await expect(pricingEditor.getByLabel('Quantity')).toBeEnabled();
  await expectNoWcagViolations(page, '.service-inspector');
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
  await expect(accessModal.getByRole('button', { name: 'Close', exact: true })).toBeFocused();
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
  await modal.getByRole('button', { name: 'Continue to output' }).click();
  await modal.getByRole('button', { name: 'Generate Architecture' }).click();
  await expect(modal).toHaveAttribute('aria-busy', 'true');
  await expect(modal.locator('.modal-close')).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(modal).toBeVisible();
  await page.locator('.ai-generator-overlay').dispatchEvent('click');
  await expect(modal).toBeVisible();

  await expect(modal).toHaveAttribute('aria-busy', 'false', { timeout: 5_000 });
  await page.keyboard.press('Escape');
  await expect(modal).toBeHidden();
});

test('AI generator opens from the mobile start card above application chrome', async ({ page }) => {
  await initializePage(page);
  await page.setViewportSize({ width: 390, height: 844 });
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
  const startChooser = page.locator('.start-chooser');
  await startChooser.getByRole('button', { name: 'Generate Diagram' }).click();

  const modal = page.locator('.ai-architecture-modal');
  await expect(modal).toBeFocused();
  await expect(modal.getByRole('button', { name: '1. Brief' }))
    .toHaveAttribute('aria-current', 'step');
  await expectNoWcagViolations(page, '.ai-architecture-modal');

  const closeIsTopmost = await modal.locator('.modal-close').evaluate((button) => {
    const box = button.getBoundingClientRect();
    const topmost = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2);
    return topmost === button || button.contains(topmost);
  });
  expect(closeIsTopmost).toBe(true);

  await modal.locator('.modal-close').click();
  await expect(modal).toBeHidden();

  await startChooser.getByRole('button', { name: 'Generate Diagram' }).click();
  await page.locator('.ai-generator-overlay').click({ position: { x: 4, y: 4 } });
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
          model: 'playwright-gpt-6-astra',
          output_text: 'An App Service hosts the web application.',
          usage: { input_tokens: 20, output_tokens: 10, total_tokens: 30 },
        });
        return;
      }
      await fulfillJson(route, {
        model: 'playwright-gpt-6-astra',
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
  await expect(modal.locator('.modal-footer-actions').getByRole('button', { name: 'Cancel request', exact: true }))
    .toBeEnabled();
  await page.keyboard.press('Escape');
  await expect(modal).toBeVisible();
  await page.locator('.ai-generator-overlay').dispatchEvent('click');
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

  await modal.getByRole('button', { name: 'Continue to output' }).click();
  await modal.getByRole('button', { name: 'Generate Architecture' }).click();
  const changeReview = page.getByRole('dialog', { name: 'Review AI changes' });
  await expect(changeReview).toBeVisible();
  await expect(page.locator('[data-testid="rf__node-web"]')).toHaveCount(0);
  await changeReview.getByRole('button', { name: 'Apply selected changes' }).click();
  await expect(page.locator('[data-testid="rf__node-web"]')).toBeVisible({ timeout: 10_000 });
  await modal.getByRole('button', { name: '1. Brief' }).click();
  await expect(fileInput).toBeEnabled();
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
  await modal.locator('.modal-footer-actions').getByRole('button', { name: 'Cancel' }).click();
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
  const cloudButton = getCloudWorkspaceButton(page);
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
  const cloudButton = getCloudWorkspaceButton(page);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Local only');
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
  const cloudButton = getCloudWorkspaceButton(page);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Local only');
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
  const cloudButton = getCloudWorkspaceButton(page);
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/, { timeout: 5_000 });
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/);
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/);
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/);
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
  let reloadCompleted = 0;
  let releaseDelayedFailure = () => {};
  const delayedFailureGate = new Promise<void>((resolve) => {
    releaseDelayedFailure = resolve;
  });

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
        reloadCompleted += 1;
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
        await delayedFailureGate;
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect.poll(() => reloadLoads).toBe(1);
  releaseDelayedFailure();
  await expect.poll(() => delayedFailureCompleted).toBe(1);
  await expect.poll(() => reloadCompleted, { timeout: 5_000 }).toBe(1);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  const cloudButton = getCloudWorkspaceButton(page);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Cloud read-only');
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Viewer review');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toHaveCount(0);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Cloud read-only');
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
  const cloudButton = getCloudWorkspaceButton(page);
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Conflict after save');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  // Keep A current until the metadata action starts, even if autosave runs first.
  let releaseReplacementSave: () => void = () => {};
  const replacementSaveGate = new Promise<void>((resolve) => { releaseReplacementSave = resolve; });

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
        await replacementSaveGate;
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Do not send to deleted A');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  releaseReplacementSave();
  await expect.poll(() => replacementCreated, { timeout: 5_000 }).toBe(1);
  expect(commentAttempts).toBe(0);
  await expect.poll(async () => (
    JSON.parse(await page.evaluate(() => (
      sessionStorage.getItem('azurediagarm.cloud-document.v1') || '{}'
    ))).documentId
  )).toBe('B');
  await expect(modal).toBeVisible();
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/);
});

test('share refresh failure blocks saves until the ETag is reconciled', async ({ page }) => {
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let shareCreated = false;
  let initialUpdates = 0;
  let storedDocument = cloudDocument('A', 'Diagram A');

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
        await fulfillJson(route, storedDocument, 200, { etag: storedDocument.etag });
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
      storedDocument = {
        ...storedDocument,
        diagramName: body.diagramName || storedDocument.diagramName,
        payload: body.payload,
        revision: 2,
        etag: '"A-2"',
      };
      await fulfillJson(route, storedDocument, 200, { etag: '"A-2"' });
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
  await editDiagramAuthor(page);
  await expect.poll(() => initialUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByRole('button', { name: 'Create link' }).click();
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect(cloudButton).toHaveAttribute('aria-label', /Cloud saved/, { timeout: 5_000 });
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
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await modal.getByPlaceholder('Add a review comment...').fill('Review after remote deletion');
  await modal.getByRole('button', { name: 'Comment', exact: true }).click();
  await expect.poll(() => commentAttempts).toBe(1);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await expect(modal.getByRole('button', { name: 'Open', exact: true })).toBeDisabled();
  await page.waitForTimeout(1_300);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const nodeTarget = page.locator('[data-testid="rf__node-A-node"] [data-node-keyboard-target]');
  await nodeTarget.focus();
  await page.keyboard.press('ArrowRight');
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(2);
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  await expect(modal.getByText('A newer cloud revision exists')).toBeVisible();
  await page.waitForTimeout(1_300);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  const cloudButton = getCloudWorkspaceButton(page);
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Sync conflict');
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
  await editDiagramAuthor(page);
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
  await editDiagramAuthor(page);
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
  await expect(getCloudWorkspaceButton(page))
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
  await expect(getCloudWorkspaceButton(page))
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
  expect(createdPayloads[1].validationSourceFingerprint).toBeNull();

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

  const cloudButton = getCloudWorkspaceButton(page);
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
  const cloudButton = getCloudWorkspaceButton(page);
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
  await editDiagramAuthor(page);
  await expect.poll(() => revision, { timeout: 5_000 }).toBeGreaterThan(1);
  await getCloudWorkspaceButton(page).click();
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
  await editDiagramAuthor(page);
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await editDiagramAuthor(page);
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  await getCloudWorkspaceButton(page).click();
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
  await editDiagramAuthor(page);
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  await getCloudWorkspaceButton(page).click();
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
  const cloudButton = getCloudWorkspaceButton(page);
  await cloudButton.click();
  const modal = page.locator('.cloud-workspace-modal');
  page.once('dialog', dialog => dialog.accept());
  await modal.getByRole('button', { name: 'Restore' }).click();
  await expect.poll(() => safetyCopies).toBe(1);
  await expect.poll(() => snapshotFetches).toBe(1);
  await expect(modal).toBeHidden();
  await expect(page.locator('[data-testid="rf__node-viewer-snapshot-node"]')).toBeVisible();
  await expect(cloudButton).toHaveAttribute('aria-label', 'Cloud workspace: Cloud read-only');
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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

  const cloudButton = getCloudWorkspaceButton(page);
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
  await editDiagramAuthor(page);
  await expect.poll(() => diagramAUpdates, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await editDiagramAuthor(page);
  await expect.poll(() => updateAttempts, { timeout: 5_000 }).toBe(1);
  const cloudButton = getCloudWorkspaceButton(page);
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
  await expect(getCloudWorkspaceButton(page)).toHaveClass(/btn-active/);
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
  const importedPayloads: Record<string, any>[] = [];
  let sourceUpdateAttempts = 0;
  let importedRevision = 1;

  page.on('dialog', async (dialog) => {
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
  const cloudButton = getCloudWorkspaceButton(page);
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
  const importError = page.getByRole('alert').filter({ hasText: 'Error loading diagram file' });
  await expect(importError).toBeVisible();
  await expect(fileInput).toHaveValue('');
  await expect(sourceNode).toBeVisible();
  await expect(cloudButton).toHaveClass(/btn-active/);
  expect(sourceUpdateAttempts).toBe(0);
  await importError.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(importError).toBeHidden();

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
  await expect(importError).toBeVisible();
  await expect(fileInput).toHaveValue('');
  await expect(sourceNode).toBeVisible();
  await expect(cloudButton).toHaveClass(/btn-active/);
  expect(sourceUpdateAttempts).toBe(0);
  await importError.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(importError).toBeHidden();

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

  const changeReview = page.getByRole('dialog', { name: 'Review AI changes' });
  await expect(changeReview).toBeVisible();
  await expect(sourceNode).toBeVisible();
  expect(sourceUpdateAttempts).toBe(0);
  expect(importedPayloads).toHaveLength(0);
  await changeReview.getByRole('button', { name: 'Apply selected changes' }).click();
  await expect(page.locator('[data-testid="rf__node-imported-app"]')).toBeVisible({
    timeout: 10_000,
  });
  const promptBanner = page.locator('.canvas-prompt-banner');
  await expect(promptBanner).toContainText('Import an App Service architecture');
  await expect(page.locator('.icon-palette')).toHaveCSS('width', '0px');
  const [initialPromptBox, canvasBox] = await Promise.all([
    promptBanner.boundingBox(),
    page.getByRole('region', { name: 'Architecture canvas' }).boundingBox(),
  ]);
  expect(initialPromptBox).not.toBeNull();
  expect(canvasBox).not.toBeNull();
  if (initialPromptBox && canvasBox) {
    await page.mouse.move(initialPromptBox.x + 24, initialPromptBox.y + 16);
    await page.mouse.down();
    await page.mouse.move(initialPromptBox.x + 84, initialPromptBox.y + 56, { steps: 4 });
    await page.mouse.up();
    const movedPromptBox = await promptBanner.boundingBox();
    expect(movedPromptBox).not.toBeNull();
    if (movedPromptBox) {
      expect(Math.abs(movedPromptBox.x - initialPromptBox.x - 60)).toBeLessThan(8);
      expect(Math.abs(movedPromptBox.y - initialPromptBox.y - 40)).toBeLessThan(8);
      expect(movedPromptBox.x).toBeGreaterThanOrEqual(canvasBox.x + 7);
      expect(movedPromptBox.y).toBeGreaterThanOrEqual(canvasBox.y + 7);
      expect(movedPromptBox.x + movedPromptBox.width).toBeLessThanOrEqual(
        canvasBox.x + canvasBox.width - 7,
      );
      expect(movedPromptBox.y + movedPromptBox.height).toBeLessThanOrEqual(
        canvasBox.y + canvasBox.height - 7,
      );
    }
  }
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
  test.slow();
  await initializePage(page, {
    documentId: 'A',
    access: 'owner',
    role: 'owner',
  });
  let guideRequests = 0;
  let sourceRevision = 1;
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
    if (path === '/api/diagrams/A' && method === 'PUT') {
      const body = JSON.parse(request.postData() || '{}');
      sourceRevision += 1;
      await fulfillJson(route, {
        ...cloudDocument('A', body.diagramName || 'Guide source'),
        diagramName: body.diagramName || 'Guide source',
        payload: body.payload,
        revision: sourceRevision,
        etag: `"A-${sourceRevision}"`,
      }, 200, { etag: `"A-${sourceRevision}"` });
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
        model: 'playwright-gpt-6-astra',
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
  await expect(page.locator('[data-testid="rf__node-A-node"]')).toBeVisible();
  await expect(page.locator('.title-block')).toContainText('Guide source');
  const generateButton = page.getByTitle('Generate comprehensive deployment guide');
  await generateButton.click();
  const modal = page.locator('.deployment-modal');
  await expect(modal.getByText('Deploy the test architecture')).toBeVisible({
    timeout: 15_000,
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

test('nodes can be connected using only the keyboard', async ({ page }) => {
  await openInteractionDiagram(page);

  const edges = page.locator('.react-flow__edge');
  await expect(edges).toHaveCount(1);

  const sourceLabel = page.locator('[data-testid="rf__node-node-b"] [data-node-keyboard-target]');
  const targetLabel = page.locator('[data-testid="rf__node-node-a"] [data-node-keyboard-target]');
  const liveRegion = page.getByTestId('live-region-polite');

  // Step 1: arm the connection on the source node.
  await sourceLabel.focus();
  await page.keyboard.press('c');
  await expect(page.locator('[data-testid="rf__node-node-b"] .azure-node'))
    .toHaveAttribute('data-connect-source', 'true');
  await expect(liveRegion).toContainText('Connection started from Azure SQL Database');

  // Escape must abandon the attempt without creating an edge.
  await page.keyboard.press('Escape');
  await expect(page.locator('[data-connect-source="true"]')).toHaveCount(0);
  await expect(liveRegion).toContainText('Connection cancelled.');
  await expect(edges).toHaveCount(1);

  // Step 2: arm again and complete on the target node.
  await sourceLabel.focus();
  await page.keyboard.press('c');
  await targetLabel.focus();
  await page.keyboard.press('c');

  await expect(edges).toHaveCount(2);
  await expect(liveRegion).toContainText('Connected Azure SQL Database to App Service');
  await expect(page.locator('[data-connect-source="true"]')).toHaveCount(0);
});

test('keyboard connections reject groups, duplicates and stale sources', async ({ page }) => {
  let latestSavedPayload: Record<string, unknown> | null = null;
  await openInteractionDiagram(page, payload => {
    latestSavedPayload = payload;
  });

  const edges = page.locator('.react-flow__edge');
  await expect(edges).toHaveCount(1);

  const nodeA = page.locator('[data-testid="rf__node-node-a"] [data-node-keyboard-target]');
  const nodeB = page.locator('[data-testid="rf__node-node-b"] [data-node-keyboard-target]');
  const groupA = page.locator('[data-testid="rf__node-group-a"] [data-node-keyboard-target]');
  const liveRegion = page.getByTestId('live-region-polite');

  // Group boxes render no connection handles, so React Flow would keep the edge
  // in state and never draw it. Arming on one must be refused outright.
  await groupA.focus();
  await page.keyboard.press('c');
  await expect(liveRegion).toContainText('Group boxes cannot be connected');
  await expect(page.locator('[data-connect-source="true"]')).toHaveCount(0);
  await expect(edges).toHaveCount(1);

  // node-a is already connected to node-b on the same handles; addEdge would
  // drop the duplicate, which must not be announced as a new connection.
  await nodeA.focus();
  await page.keyboard.press('c');
  await nodeB.focus();
  await page.keyboard.press('c');
  await expect(liveRegion).toContainText('is already connected to');
  await expect(edges).toHaveCount(1);

  // A source deleted mid-flow must not produce an edge with a dangling endpoint.
  await nodeA.focus();
  await page.keyboard.press('c');
  await expect(page.locator('[data-testid="rf__node-node-a"] .azure-node'))
    .toHaveAttribute('data-connect-source', 'true');
  await page.locator('[data-testid="rf__node-node-a"]').click();
  await page.keyboard.press('Delete');
  await expect(page.locator('[data-testid="rf__node-node-a"]')).toHaveCount(0);

  await nodeB.focus();
  await page.keyboard.press('c');
  await expect(page.getByTestId('live-region-assertive')).toContainText('no longer on the canvas');

  // A dangling edge never reaches the DOM — React Flow's renderer returns null
  // when an endpoint is missing, which is exactly the ghost-edge mechanism this
  // test exists to catch. It has to be asserted against persisted state.
  await expect.poll(() => {
    const payload = latestSavedPayload as {
      nodes?: Array<{ id?: string }>;
      edges?: Array<{ source?: string; target?: string }>;
    } | null;
    if (!payload?.edges || !payload?.nodes) return null;
    const ids = new Set(payload.nodes.map(node => node.id));
    return payload.edges.filter(edge => !ids.has(edge.source) || !ids.has(edge.target)).length;
  }, { timeout: 15000 }).toBe(0);
});

/**
 * The ungroup button is driven by the node's parentNode, which lives on the
 * React Flow node rather than in `data`. Reading it with a non-reactive
 * getNode() snapshot meant the memo comparator saw no change and the button
 * stayed on screen for a node that had already left its group.
 */
test('the ungroup button disappears as soon as the node leaves its group', async ({ page }) => {
  await initializePage(page, { documentId: 'A', access: 'owner', role: 'owner' });
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
  ];
  document.payload.edges = [];
  let revision = 1;

  await page.route('**/api/**', async (route) => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    const method = request.method();
    if (path === '/api/access/me') {
      await fulfillJson(route, {
        enabled: false, authenticated: true, email: 'owner@example.com', isAdmin: false, allowed: true,
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
      await fulfillJson(route, {
        ...document, payload: body.payload, revision, etag: `"A-${revision}"`,
      }, 200, { etag: `"A-${revision}"` });
      return;
    }
    await fulfillJson(route, { error: `Unhandled test endpoint: ${method} ${path}` }, 404);
  });

  await page.goto('/');
  const node = page.locator('[data-testid="rf__node-node-a"]');
  await expect(node).toBeVisible();

  const ungroupButton = node.locator('button.ungroup-button');
  await expect(ungroupButton).toHaveCount(0);

  await node.locator('[data-node-keyboard-target]').click();
  await expect(ungroupButton).toBeVisible();

  await ungroupButton.click();

  // The node stays selected, its data object is untouched and its id is
  // unchanged, so nothing the memo comparator looks at changes here. Only a
  // reactive read of the store removes the button.
  await expect(ungroupButton).toHaveCount(0);
  await expect(node).toBeVisible();
});
