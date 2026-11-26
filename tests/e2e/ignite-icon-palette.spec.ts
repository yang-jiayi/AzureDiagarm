import { expect, test } from '@playwright/test';

type Box = { x: number; y: number; width: number; height: number };

function overlaps(left: Box, right: Box): boolean {
  return left.x < right.x + right.width
    && left.x + left.width > right.x
    && left.y < right.y + right.height
    && left.y + left.height > right.y;
}

test('unofficial IQ badges stay accessible and within palette cards in both layouts and languages', async ({ page }, testInfo) => {
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
  });
  await page.goto('/');
  await page.locator('#palette-view-tab-catalog').click();
  const search = page.locator('input[type="search"]');
  await search.fill('IQ');
  await expect(page.locator('.icon-item-main')).toHaveCount(5);

  for (const language of ['en', 'ja'] as const) {
    if (language === 'ja') {
      await page.getByRole('button', { name: 'More', exact: true }).click();
      await page.getByRole('dialog', { name: 'More application options' })
        .getByRole('button', { name: '日本語' }).click();
    }

    for (const layout of ['grid', 'list'] as const) {
      await page.locator('.palette-layout-switch button').nth(layout === 'grid' ? 0 : 1).click();
      const cards = page.locator('.icon-item').filter({
        has: page.locator('.icon-drawing-kind--unofficial'),
      });
      await expect(cards).toHaveCount(4);

      for (const card of await cards.all()) {
        const main = card.locator('.icon-item-main');
        const badge = card.locator('.icon-drawing-kind--unofficial');
        const icon = card.locator('.icon-image');
        const label = card.locator('.icon-label');
        const actions = card.locator('.icon-item-actions');
        const fullDescription = language === 'en'
          ? 'Unofficial concept symbol, not a product logo.'
          : '非公式の概念記号です。製品ロゴではありません。';
        const descriptionId = await main.getAttribute('aria-describedby');
        expect(descriptionId).toBeTruthy();
        await expect(page.locator(`[id="${descriptionId}"]`)).toHaveText(fullDescription);
        await expect(badge).toHaveText(language === 'en' ? 'Unofficial' : '非公式');
        await expect(badge).toHaveAttribute('title', fullDescription);
        await expect(icon).toBeVisible();
        await expect(actions.getByRole('button')).toHaveCount(2);

        const cardBox = await card.boundingBox() as Box;
        const mainBox = await main.boundingBox() as Box;
        const badgeBox = await badge.boundingBox() as Box;
        const iconBox = await icon.boundingBox() as Box;
        const labelBox = await label.boundingBox() as Box;
        const actionsBox = await actions.boundingBox() as Box;
        for (const box of [mainBox, actionsBox]) {
          expect(box.x).toBeGreaterThanOrEqual(cardBox.x);
          expect(box.y).toBeGreaterThanOrEqual(cardBox.y);
          expect(box.x + box.width).toBeLessThanOrEqual(cardBox.x + cardBox.width + 1);
          expect(box.y + box.height).toBeLessThanOrEqual(cardBox.y + cardBox.height + 1);
        }
        for (const [name, box] of [
          ['badge', badgeBox],
          ['icon', iconBox],
          ['label', labelBox],
        ] as const) {
          expect(box.x, `${language}/${layout} ${name} x`).toBeGreaterThanOrEqual(mainBox.x);
          expect(box.y, `${language}/${layout} ${name} y`).toBeGreaterThanOrEqual(mainBox.y);
          expect(box.x + box.width, `${language}/${layout} ${name} right`)
            .toBeLessThanOrEqual(mainBox.x + mainBox.width + 1);
          expect(box.y + box.height, `${language}/${layout} ${name} bottom`)
            .toBeLessThanOrEqual(mainBox.y + mainBox.height + 1);
        }
        expect(overlaps(mainBox, actionsBox)).toBe(false);
        expect(overlaps(iconBox, labelBox)).toBe(false);
        expect(overlaps(labelBox, badgeBox)).toBe(false);
        expect(overlaps(badgeBox, actionsBox)).toBe(false);

        const textMetrics = await label.evaluate(element => ({
          clientHeight: element.clientHeight,
          clientWidth: element.clientWidth,
          scrollWidth: element.scrollWidth,
        }));
        expect(textMetrics.clientHeight).toBeGreaterThan(0);
        expect(textMetrics.scrollWidth).toBeLessThanOrEqual(textMetrics.clientWidth + 1);
      }

      if (language === 'en' && layout === 'grid') {
        await page.screenshot({
          path: testInfo.outputPath('ignite-iq-palette.png'),
          fullPage: true,
        });
      }
    }
  }
});

test('Ignite product icons remain searchable, addable, and exportable', async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
  });
  await page.goto('/');
  await page.locator('#palette-view-tab-catalog').click();
  const search = page.locator('input[type="search"]');
  const items = page.locator('.icon-item-main');

  const targets = [
    { query: 'Microsoft IQ', name: 'Add Microsoft IQ to the canvas' },
    { query: 'Fabric Apps', name: 'Add Fabric Apps to the canvas' },
    { query: 'HorizonDB', name: 'Add Azure Horizondb to the canvas' },
    { query: 'HorizonDB', name: 'Add PostgreSQL Database (generic symbol) to the canvas' },
  ];
  for (const [index, target] of targets.entries()) {
    await search.fill(target.query);
    expect(await items.count()).toBeGreaterThan(0);
    await page.getByRole('button', { name: target.name, exact: true }).click();
    await expect(page.locator('.react-flow__node')).toHaveCount(index + 1);
  }

  await page.getByRole('region', { name: 'Architecture canvas' }).focus();
  await page.keyboard.press('Control+K');
  const commandPalette = page.getByTestId('command-palette');
  await commandPalette.getByRole('combobox', { name: 'Search commands and services' }).fill('export png');
  const downloadPromise = page.waitForEvent('download');
  await commandPalette.getByRole('option', { name: /^Export PNG/ }).click();
  const preflight = page.getByRole('dialog', { name: 'Review data handling' });
  await expect(preflight).toBeVisible();
  await preflight.getByRole('button', { name: 'Proceed as shown' }).click();
  const download = await downloadPromise;
  expect(await download.failure()).toBeNull();
  expect(download.suggestedFilename()).toMatch(/\.png$/i);
});
