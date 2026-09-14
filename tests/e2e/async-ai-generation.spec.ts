import { expect, test, type Download } from '@playwright/test';
import AxeBuilder from '@axe-core/playwright';

async function expectPng(download: Download) {
  expect(download.suggestedFilename()).toMatch(/\.png$/);
  const stream = await download.createReadStream();
  if (!stream) throw new Error('PNG download has no readable body.');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const bytes = Buffer.concat(chunks);
  expect([...bytes.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
  expect(bytes.readUInt32BE(16)).toBeGreaterThan(600);
  expect(bytes.readUInt32BE(20)).toBeGreaterThan(300);
  expect(bytes.length).toBeGreaterThan(5000);
}

test('MAX blueprint generation displays async progress and delivers a real PNG from the separately retrieved result', async ({ page }) => {
  await page.addInitScript(() => {
    localStorage.setItem('azure-diagram-builder.language.v1', 'en');
    localStorage.setItem('azure-diagram-builder.headerCollapsed.v1', '0');
    localStorage.setItem('azure-diagram-builder.focusMode.v1', '0');
    localStorage.setItem('azure-diagram-builder.canvasHintDismissed.v1', '1');
    localStorage.setItem('aiGenerator.mode', 'blueprint');
    localStorage.setItem('azure-diagrams-model-settings', JSON.stringify({
      version: 3, model: 'gpt-6-astra', reasoningEffort: 'max',
    }));
  });
  let id = '';
  let ready = false;
  const posts: Array<{ body: { reasoning: unknown; max_output_tokens: unknown } }> = [];
  const paths: string[] = [];
  const pageErrors: string[] = [];
  page.on('pageerror', error => pageErrors.push(error.message));
  const job = (status: string) => ({
    job: { id, status, elapsedMs: status === 'queued' ? 0 : 250000, deadlineAt: Date.now() + 900000, pollAfterMs: 1000 },
  });
  await page.route('**/api/**', async route => {
    const request = route.request();
    const path = new URL(request.url()).pathname;
    paths.push(path);
    let body: unknown = {};
    let status = 200;
    if (path === '/api/access/me') body = { enabled: false, authenticated: false, allowed: true, isAdmin: false };
    else if (path === '/api/runtime-config') body = { features: { bringYourOwnAI: true } };
    else if (path === '/api/ai/budget') body = {
      available: true, concurrentRequests: 0, concurrentLimit: 2, remainingTokens: 250000, limitTokens: 250000, usedTokens: 0, reservedTokens: 0,
    };
    else if (path === '/api/openai') {
      expect(request.headers().prefer).toBe('respond-async');
      id = request.headers()['idempotency-key'];
      posts.push(request.postDataJSON());
      body = job('queued');
      status = 202;
    } else if (path === `/api/openai/jobs/${id}/result`) {
      expect(request.postData()).toBeNull();
      body = {
        status: 'completed',
        output_text: JSON.stringify({
          title: '非同期ジョブの構成図',
          canvas: { width: 900, height: 550 },
          zones: [{ id: 'azure', label: 'Azure', kind: 'azure', x: 40, y: 70, width: 780, height: 380 }],
          nodes: [
            { id: 'functions', name: 'Azure Functions', category: 'compute', x: 160, y: 200, zone: 'azure' },
            { id: 'database', name: 'Azure Cosmos DB', category: 'databases', x: 540, y: 200, zone: 'azure' },
          ],
          edges: [{ id: 'data', from: 'functions', to: 'database', label: 'データ保存', step: 1 }],
          workflow: [{ step: 1, description: 'データを安全に保存する' }],
        }),
        usage: { input_tokens: 100, output_tokens: 50, total_tokens: 150 },
      };
    } else if (path === `/api/openai/jobs/${id}`) {
      expect(request.postData()).toBeNull();
      body = job(ready ? 'succeeded' : 'running');
    } else {
      body = { error: 'Not found' };
      status = 404;
    }
    await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
  });
  await page.goto('/');
  await page.getByRole('tab', { name: 'Create', exact: true }).click();
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  const generator = page.locator('.ai-architecture-modal');
  await generator.getByLabel('Architecture Description or Modification').fill('Create a Functions and Cosmos DB architecture.');
  await generator.getByRole('button', { name: 'Continue to output', exact: true }).click();
  await generator.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
  const progress = generator.locator('[data-ai-job-progress="blueprint"]');
  await expect(progress).toContainText('Generating in background');
  await expect(progress).toContainText('250s elapsed');
  await expect(generator.getByRole('button', { name: 'Cancel request', exact: true })).toBeEnabled();
  const accessibility = await new AxeBuilder({ page }).include('.ai-architecture-modal')
    .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa']).analyze();
  expect(accessibility.violations).toEqual([]);
  expect(posts).toHaveLength(1);
  expect(posts[0].body.reasoning).toEqual({ effort: 'max' });
  expect(posts[0].body.max_output_tokens).toBe(32000);
  const downloading = page.waitForEvent('download', { predicate: value => value.suggestedFilename().endsWith('.png') });
  ready = true;
  await expectPng(await downloading);
  await expect(page.locator('[data-bp-arch-export-host]')).toHaveCount(0);
  expect(paths).toContain(`/api/openai/jobs/${id}/result`);
  expect(posts).toHaveLength(1);
  expect(pageErrors).toEqual([]);
});

test('detached editorial PNG rendering has its own language context and cleans up after download', async ({ page }) => {
  await page.route('**/api/**', route => {
    const path = new URL(route.request().url()).pathname;
    const body = path === '/api/access/me'
      ? { enabled: false, authenticated: false, allowed: true, isAdmin: false }
      : path === '/api/runtime-config' ? { features: { bringYourOwnAI: true } }
        : { available: true, concurrentRequests: 0, concurrentLimit: 2, remainingTokens: 250000, limitTokens: 250000, usedTokens: 0, reservedTokens: 0 };
    return route.fulfill({ contentType: 'application/json', body: JSON.stringify(body) });
  });
  await page.goto('/');
  const downloading = page.waitForEvent('download', { predicate: value => value.suggestedFilename() === 'async-editorial-test.png' });
  await page.evaluate(async () => {
    const exporter = await import('/src/utils/exportReferencePng.ts');
    await exporter.exportReferenceArchitectureAsPng({
      title: '日本語のリファレンス構成',
      stages: [{ id: 'compute', label: '処理', services: [{ id: 'functions', name: 'Azure Functions', category: 'compute' }] }],
      connections: [],
    }, { fileName: 'async-editorial-test' });
  });
  await expectPng(await downloading);
  await expect(page.locator('[data-ref-arch-export-host]')).toHaveCount(0);
});
