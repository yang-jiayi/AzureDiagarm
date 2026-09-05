import test, { before, after } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { build } from 'esbuild';
import { chromium, type Browser, type Page } from 'playwright';

let browser: Browser;
let script: string;
const inspectorCss = readFileSync(new URL('../src/components/ServiceInspector.css', import.meta.url), 'utf8');
before(async () => {
  const mocks: Record<string, string> = {
    LanguageContext: `export const useLanguage=()=>({language:window.h.language,t:x=>x,translate:x=>x});`,
    regionalPricingService: `
      export const getActiveRegion=()=> 'eastus2';
      export const preloadCommonServices=async()=>{};
      export const getRegionalServicePricing=async()=>{
        if(window.h.deferPrices) await new Promise(resolve=>window.h.resolvePrices=resolve);
        return {tiers:window.h.tiers,defaultTier:'B1',calculationType:'hourly'};
      };
    `,
    architectureValidator: `export const formatValidationReport=()=> '# Review report';`,
    validationDisplayStore: `
      import {useState} from 'react';
      export const useValidationDisplayPrefs=()=>useState({showNumericScore:false});
    `,
  };
  const bundle = await build({
    stdin: {
      contents: `
        import React from 'react';
        import {createRoot} from 'react-dom/client';
        import Inspector from './src/components/ServiceInspector';
        import ValidationModal from './src/components/ValidationModal';
        import NodePricingEditor from './src/components/NodePricingEditor';
        import {refreshAllNodePricing} from './src/services/costEstimationService';
        const h=window.h;
        const root=createRoot(document.getElementById('root'));
        h.refresh=async region=>{h.refreshed=await refreshAllNodePricing([h.node],region);};
        h.render=()=>{
          const close=()=>{h.closed++;h.open=false;h.render();};
          if(!h.open) {root.render(null);return;}
          if(h.surface==='validation') root.render(<ValidationModal
            isOpen validation={h.validation} isStale={h.stale}
            onClose={close} onRevalidate={()=>h.revalidations++}
            reviewHistory={h.history} onFocusResources={resources=>h.focused.push(resources)}
            onApplyRecommendations={findings=>h.applies.push(findings)}/>);
          else if(h.surface==='node-pricing') root.render(<NodePricingEditor
            serviceType={h.node.data.serviceName||h.node.data.label} pricing={h.node.data.pricing}
            onClose={close} onApply={pricing=>h.applies.push(pricing)}/>);
          else root.render(<Inspector node={h.node} onClose={close}
            onUpdateNode={(id,data)=>h.applies.push({id,data})}/>);
        };
        h.render();
      `,
      resolveDir: process.cwd(), loader: 'tsx',
    },
    bundle: true, write: false, platform: 'browser', format: 'iife', logLevel: 'silent',
    define: { 'process.env.NODE_ENV': '"test"', 'import.meta.env': '{}' },
    plugins: [{
      name: 'inspector-test-boundaries',
      setup(builder) {
        builder.onResolve({ filter: /\.css$/ }, args => ({ path: args.path, namespace: 'no-css' }));
        builder.onLoad({ filter: /.*/, namespace: 'no-css' }, () => ({ contents: '' }));
        builder.onResolve({ filter: /\/(LanguageContext|regionalPricingService|architectureValidator|validationDisplayStore)$/ },
          args => ({ path: args.path.split('/').at(-1)!, namespace: 'boundary' }));
        builder.onLoad({ filter: /.*/, namespace: 'boundary' }, args => ({
          contents: mocks[args.path], loader: 'js', resolveDir: process.cwd(),
        }));
      },
    }],
  });
  script = bundle.outputFiles[0].text;
  browser = await chromium.launch({ headless: true });
});
after(async () => { await browser?.close(); });

async function setup(t: { after: (callback: () => Promise<void>) => void }, overrides: Record<string, unknown> = {}): Promise<Page> {
  const page = await browser.newPage();
  t.after(() => page.close());
  const errors: string[] = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.route('http://127.0.0.1/inspector-fixture', route => route.fulfill({
    contentType: 'text/html',
    body: '<html><body><button id="opener">Open</button><div id="root"></div></body></html>',
  }));
  await page.goto('http://127.0.0.1/inspector-fixture');
  await page.addStyleTag({ content: inspectorCss });
  await page.locator('#opener').focus();
  await page.evaluate(overrides => {
    const provenance = { kind: 'official-meter', source: 'azure-retail-prices',
      asOf: '2026-07-01', unit: '1 Hour', meterId: 'meter-1', assumptions: [{ label: 'Monthly operation', value: 730, unit: 'hours/month' }] };
    const finding = { severity: 'high', category: 'Availability', issue: 'Single instance',
      recommendation: 'Add redundancy', resources: ['Web'], resourceIds: ['node-1'], source: 'rule-based' };
    (window as any).h = {
      surface: 'inspector', language: 'en', open: true, closed: 0, applies: [], focused: [], revalidations: 0,
      tiers: [
        { name: 'B1', skuName: 'B1', monthlyPrice: 73, unit: '1 Hour', provenance },
        { name: 'P1', skuName: 'P1', monthlyPrice: 146, unit: '1 Hour', provenance },
      ],
      node: {
        id: 'node-1', type: 'azureNode', position: { x: 0, y: 0 },
        data: { label: 'App Service', description: 'Original', icon: 'app.svg', groupId: 'group-1',
          pricing: { tier: 'B1', skuName: 'B1', quantity: 1, estimatedCost: 73, region: 'eastus2',
            isCustom: false, unit: '1 Hour', lastUpdated: '2026-07-01', provenance } },
      },
      stale: true,
      validation: { overallScore: 70, summary: 'Design review', timestamp: '2026-09-01T00:00:00Z',
        pillars: [{ pillar: 'Reliability', score: 70, findings: [finding, { ...finding, issue: 'Architecture-wide finding', resources: [], resourceIds: [], source: 'ai' }] }],
        quickWins: [{ ...finding, issue: 'Quick win', recommendation: 'Add monitoring', source: 'ai' }] },
      history: [{ key: 'old', pillar: 'Security', status: 'not-detected', firstSeenAt: 100, lastSeenAt: 100,
        lastReviewedAt: 200, notDetectedAt: 200, finding: { ...finding, issue: 'Old finding' } }],
      ...overrides,
    };
  }, overrides);
  await page.addScriptTag({ content: script });
  assert.deepEqual(errors, [], 'browser fixture must bootstrap without runtime errors');
  return page;
}

test('the existing cost editor preserves unpriced values while editing quantities', async t => {
  const page = await setup(t, {
    surface: 'node-pricing',
    tiers: [{ name: 'B1', skuName: 'B1', monthlyPrice: null, unit: 'instance/month' }],
    node: {
      id: 'unpriced', type: 'azureNode', position: { x: 0, y: 0 },
      data: { label: 'App Service', pricing: {
        tier: 'B1', skuName: 'B1', quantity: 2, estimatedCost: null, region: 'eastus2',
        isCustom: false, unit: 'instance/month', lastUpdated: '',
      } },
    },
  });
  await page.getByLabel('Tier / SKU').waitFor();
  assert.match(await page.locator('.npe-preview').innerText(), /Unpriced/);
  assert.doesNotMatch(await page.locator('.npe-preview').innerText(), /Free|\$0/);
  const quantity = page.getByLabel('Instances / units');
  const apply = page.getByRole('button', { name: 'Apply', exact: true });
  for (const invalid of ['', '0', '1.5', '100001']) {
    await quantity.fill(invalid);
    assert.equal(await apply.isDisabled(), true);
    assert.equal(await quantity.inputValue(), invalid, 'invalid input must not silently become another quantity');
  }
  await quantity.fill('100000');
  await apply.click();
  const applied = await page.evaluate(() => (window as any).h.applies);
  assert.equal(applied.length, 1);
  assert.equal(applied[0].estimatedCost, null);
  assert.equal(applied[0].quantity, 100000);
});

test('selecting an unpriced SKU cannot reuse the previous priced estimate', async t => {
  const page = await setup(t, {
    surface: 'node-pricing',
    tiers: [
      { name: 'B1', skuName: 'B1', monthlyPrice: 73, unit: 'instance/month' },
      { name: 'P1', skuName: 'P1', monthlyPrice: null, unit: 'instance/month' },
    ],
  });
  await page.getByLabel('Tier / SKU').selectOption('P1');
  assert.match(await page.locator('.npe-preview').innerText(), /Unpriced/);
  assert.doesNotMatch(await page.locator('.npe-preview').innerText(), /\$73|Free/);
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  const applied = await page.evaluate(() => (window as any).h.applies[0]);
  assert.equal(applied.estimatedCost, null);
  assert.equal(applied.skuName, 'P1');
});

test('an unknown price needs an explicit custom value before it can become zero', async t => {
  const page = await setup(t, {
    surface: 'node-pricing', tiers: [],
    node: {
      id: 'unpriced', type: 'azureNode', position: { x: 0, y: 0 },
      data: { label: 'App Service', pricing: {
        tier: 'Unknown', skuName: 'Unknown', quantity: 2, estimatedCost: null, region: 'eastus2',
        isCustom: false, unit: 'instance/month', lastUpdated: '',
      } },
    },
  });
  await page.getByLabel('Override with a custom monthly price (per unit)').check();
  const customPrice = page.getByLabel('Custom price (USD / month / unit)');
  assert.equal(await customPrice.inputValue(), '');
  assert.equal(await page.getByRole('button', { name: 'Apply', exact: true }).isDisabled(), true);
  await customPrice.fill('0');
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  const applied = await page.evaluate(() => (window as any).h.applies[0]);
  assert.equal(applied.estimatedCost, 0);
  assert.equal(applied.quantity, 2);
  assert.equal(applied.provenance.kind, 'custom');
});

test('inspector previews SKU and quantity once and applies exactly one merged payload', async t => {
  const page = await setup(t);
  await page.getByRole('button', { name: 'Apply', exact: true }).waitFor();
  await page.waitForFunction(() => !(document.querySelector('select') as HTMLSelectElement).disabled);
  await page.getByLabel('Label', { exact: true }).fill('Customer API');
  await page.getByLabel('Description', { exact: true }).fill('New description');
  await page.getByLabel('Tier / SKU').selectOption('P1');
  await page.getByLabel('Quantity (instances / units)').fill('3');
  assert.match(await page.locator('.service-inspector-preview').innerText(), /\$438\.00/);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  const applied = await page.evaluate(() => (window as any).h.applies);
  assert.equal(applied.length, 1);
  assert.equal(applied[0].id, 'node-1');
  assert.equal(applied[0].data.label, 'Customer API');
  assert.equal(applied[0].data.pricing.estimatedCost, 146);
  assert.equal(applied[0].data.pricing.quantity, 3);
  assert.equal(applied[0].data.icon, 'app.svg');
  assert.equal(applied[0].data.groupId, 'group-1');
  assert.equal(applied[0].data.serviceName, 'App Service');
  assert.equal(await page.getByRole('dialog').count(), 0);
});

test('cancel and Escape discard edits, restore focus, and ignore late price responses', async t => {
  const page = await setup(t, { deferPrices: true });
  await page.getByLabel('Label', { exact: true }).fill('Unsaved');
  await page.keyboard.press('Escape');
  await page.evaluate(() => (window as any).h.resolvePrices());
  assert.equal(await page.getByRole('dialog').count(), 0);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.locator('#opener').evaluate(element => element === document.activeElement), true);
  await page.evaluate(() => { const h = (window as any).h; h.open = true; h.deferPrices = false; h.render(); });
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
});

test('invalid quantities and custom estimates cannot apply', async t => {
  const page = await setup(t);
  const apply = page.getByRole('button', { name: 'Apply', exact: true });
  for (const quantity of ['0', '1.5', '100001', '']) {
    await page.getByLabel('Quantity (instances / units)').fill(quantity);
    assert.equal(await apply.isDisabled(), true);
  }
  await page.getByLabel('Quantity (instances / units)').fill('2');
  await page.getByLabel('Use a custom monthly estimate').check();
  await page.getByLabel('USD per unit / month').fill('-1');
  assert.equal(await apply.isDisabled(), true);
  await page.getByLabel('USD per unit / month').fill('12.50');
  assert.match(await page.locator('.service-inspector-preview').innerText(), /\$25\.00/);
  await apply.click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[0].data.pricing.provenance.kind), 'custom');
});

test('legacy provenance stays unknown and edits do not relabel it as an official price', async t => {
  const page = await setup(t);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.open = false; h.render();
  });
  await page.getByRole('dialog').waitFor({ state: 'detached' });
  await page.evaluate(() => {
    const h = (window as any).h;
    delete h.node.data.pricing.provenance;
    h.open = true; h.render();
  });
  await page.getByText('Unknown source (imported)').waitFor();
  assert.match(await page.locator('.service-inspector-preview').innerText(), /Price data date: Unknown/);
  await page.getByLabel('Quantity (instances / units)').fill('2');
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[0].data.pricing.provenance.kind), 'unknown');
});

test('only substantiated meter usage can be edited and unknown consumption is not shown as free', async t => {
  const page = await setup(t, {
    node: { id: 'usage', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'OpenAI' } },
    tiers: [{ name: 'Token meter', skuName: 'Token meter', monthlyPrice: null, unit: '1K',
      usage: { amount: null, unit: '1K', unitPrice: 2 },
      provenance: { kind: 'unpriced', source: 'azure-retail-prices', unit: '1K', assumptions: [] } }],
  });
  await page.getByLabel('Monthly usage per instance').waitFor();
  assert.match(await page.locator('.service-inspector-preview').innerText(), /not treated as zero/);
  await page.getByLabel('Monthly usage per instance').fill('5');
  await page.getByLabel('Quantity (instances / units)').fill('2');
  assert.match(await page.locator('.service-inspector-preview').innerText(), /\$20\.00/);
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  const applied = await page.evaluate(() => (window as any).h.applies[0].data.pricing);
  assert.equal(applied.estimatedCost, 10);
  assert.equal(applied.usage.amount, 5);
  assert.equal(applied.provenance.kind, 'usage-estimate');
});

test('inspector rejects outside node-data changes rather than overwriting another edit', async t => {
  const page = await setup(t);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.node = { ...h.node, data: { ...h.node.data, description: 'External update' } }; h.render();
  });
  await page.getByText('This service changed while the inspector was open.', { exact: false }).waitFor();
  assert.equal(await page.getByRole('button', { name: 'Apply', exact: true }).isDisabled(), true);
});

test('region refresh preserves a chosen usage amount and quantity while replacing its source rate', async t => {
  const page = await setup(t, {
    node: {
      id: 'usage', type: 'azureNode', position: { x: 0, y: 0 },
      data: { label: 'Token API', serviceName: 'OpenAI',
        pricing: { tier: 'Friendly token plan', skuName: 'Token meter', quantity: 3, estimatedCost: 10,
          isCustom: false, isUsageBased: true, unit: '1K', region: 'eastus2', lastUpdated: '',
          usage: { amount: 5, unit: '1K', unitPrice: 2 },
          provenance: { kind: 'usage-estimate', source: 'azure-retail-prices', unit: '1K', assumptions: [] } } },
    },
    tiers: [{ name: 'Token meter', skuName: 'Token meter', monthlyPrice: null, unit: '1K',
      usage: { amount: null, unit: '1K', unitPrice: 4 },
      provenance: { kind: 'unpriced', source: 'azure-retail-prices', unit: '1K', assumptions: [] } }],
  });
  await page.evaluate(() => (window as any).h.refresh('japaneast'));
  const pricing = await page.evaluate(() => (window as any).h.refreshed[0].data.pricing);
  assert.equal(pricing.quantity, 3);
  assert.equal(pricing.usage.amount, 5);
  assert.equal(pricing.usage.unitPrice, 4);
  assert.equal(pricing.estimatedCost, 20);
  assert.equal(pricing.region, 'japaneast');
  await page.evaluate(() => {
    const h = (window as any).h;
    h.node = h.refreshed[0];
    h.tiers[0].provenance.meterName = 'Different usage meter';
    return h.refresh('canadacentral');
  });
  assert.equal(await page.evaluate(() => (window as any).h.refreshed[0].data.pricing.usage.amount), null);
  assert.equal(await page.evaluate(() => (window as any).h.refreshed[0].data.pricing.estimatedCost), null);
});

test('WAF shows stale state, reevaluation, collapsed non-remediation history, and rule/AI badges', async t => {
  const page = await setup(t, { surface: 'validation' });
  await page.getByText('This review is out of date.').waitFor();
  assert.equal(await page.locator('details').evaluate(element => element.hasAttribute('open')), false);
  assert.equal(await page.locator('.source-badge.rule-based').count(), 2);
  assert.equal(await page.locator('.source-badge.ai').count(), 2);
  await page.getByRole('button', { name: 'Re-evaluate now' }).click();
  assert.equal(await page.evaluate(() => (window as any).h.revalidations), 1);
  await page.getByText('Not detected in latest review (1)').click();
  assert.match(await page.locator('details').innerText(), /Absence is not proof of remediation/);
  assert.match(await page.locator('details').innerText(), /Old finding/);
});

test('WAF focus uses stable IDs, disables unscoped finding, and preserves checkbox/apply/download flow', async t => {
  const page = await setup(t, { surface: 'validation' });
  const focuses = page.getByRole('button', { name: 'Show on diagram', exact: true });
  assert.equal(await focuses.nth(1).isDisabled(), true);
  await page.getByRole('checkbox', { name: 'Select recommendation: Single instance', exact: true }).check();
  await page.getByRole('checkbox', { name: 'Select recommendation: Quick win', exact: true }).check();
  await page.getByRole('button', { name: 'Apply 2 recommendations', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[0].length), 2);
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download Report', exact: true }).click();
  assert.match((await download).suggestedFilename(), /\.md$/);
  await focuses.first().click();
  assert.deepEqual(await page.evaluate(() => (window as any).h.focused), [['node-1']]);
  assert.equal(await page.getByRole('dialog').count(), 0);
});

test('new Japanese UI copy is localized and modal focus remains contained', async t => {
  const page = await setup(t, { language: 'ja' });
  await page.getByRole('heading', { name: 'サービスの詳細設定' }).waitFor();
  await page.waitForFunction(() => !(document.querySelector('select') as HTMLSelectElement).disabled);
  await page.getByLabel('ラベル', { exact: true }).focus();
  await page.keyboard.press('Shift+Tab');
  assert.equal(await page.getByRole('button', { name: '詳細設定を閉じる' }).evaluate(element => element === document.activeElement), true);
  await page.keyboard.press('Shift+Tab');
  assert.equal(await page.getByRole('button', { name: '適用', exact: true }).evaluate(element => element === document.activeElement), true);
  await page.getByRole('button', { name: 'キャンセル', exact: true }).click();
  await page.evaluate(() => { const h = (window as any).h; h.open = true; h.surface = 'validation'; h.render(); });
  await page.getByText('このレビューは最新ではありません。').waitFor();
  assert.equal(await page.getByRole('button', { name: '図で表示', exact: true }).count(), 3);
});

test('inspector remains a blocking slide-over with contained Tab and backdrop cancellation', async t => {
  const page = await setup(t);
  const dialog = page.getByRole('dialog');
  await dialog.waitFor();
  assert.equal(await dialog.getAttribute('aria-modal'), 'true');
  assert.equal(await page.locator('.service-inspector-overlay').evaluate(element => getComputedStyle(element).position), 'fixed');
  const opener = await page.locator('#opener').boundingBox();
  assert.ok(opener);
  assert.equal(await page.evaluate(({ x, y }) =>
    document.elementFromPoint(x, y)?.id === 'opener',
  { x: opener.x + 2, y: opener.y + 2 }), false, 'backdrop must block background pointer interaction');
  await page.getByLabel('Label', { exact: true }).fill('Unsaved background test');
  for (let i = 0; i < 20; i++) {
    await page.keyboard.press('Tab');
    assert.equal(await page.evaluate(() => !!document.activeElement?.closest('.service-inspector')), true);
  }
  await page.mouse.click(2, 2);
  await dialog.waitFor({ state: 'detached' });
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.locator('#opener').evaluate(element => element === document.activeElement), true);
});

test('inspector preserves platform-specific IDs for otherwise identical App Service SKUs', async t => {
  const page = await setup(t, {
    tiers: [
      { id: 'Linux::S1', name: 'S1 (Linux)', skuName: 'S1', monthlyPrice: 73, unit: '1 Hour' },
      { id: 'Windows::S1', name: 'S1 (Windows)', skuName: 'S1', monthlyPrice: 146, unit: '1 Hour' },
    ],
  });
  await page.waitForFunction(() => !(document.querySelector('select') as HTMLSelectElement).disabled);
  await page.getByLabel('Tier / SKU').selectOption('Windows::S1');
  await page.getByLabel('Quantity (instances / units)').fill('3');
  assert.match(await page.locator('.service-inspector-preview').innerText(), /\$438\.00/);
  await page.getByRole('button', { name: 'Apply', exact: true }).click();
  const pricing = await page.evaluate(() => (window as any).h.applies[0].data.pricing);
  assert.equal(pricing.tierId, 'Windows::S1');
  assert.equal(pricing.skuName, 'S1');
  assert.equal(pricing.estimatedCost, 146);
});

test('WAF retains remote evidence, remediation, reference links and individual apply actions', async t => {
  const page = await setup(t, { surface: 'validation' });
  await page.evaluate(() => {
    const h = (window as any).h;
    const finding = {
      ...h.validation.pillars[0].findings[0],
      evidence: ['The diagram contains one instance.'],
      remediation: ['Add a second instance.', 'Configure health checks.'],
      referenceUrl: 'https://learn.microsoft.com/azure/well-architected/reliability/',
      applyAction: { type: 'add-service', label: 'Add another instance', serviceType: 'App Service' },
    };
    h.validation = { ...h.validation, pillars: [{ pillar: 'Reliability', score: 70, findings: [finding] }] };
    h.render();
  });
  await page.getByText('The diagram contains one instance.').waitFor();
  assert.match(await page.locator('.finding-remediation').innerText(), /Configure health checks/);
  assert.equal(await page.getByRole('link', { name: 'Microsoft Learn', exact: true }).getAttribute('href'),
    'https://learn.microsoft.com/azure/well-architected/reliability/');
  await page.getByRole('button', { name: 'Add another instance', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][0].applyAction.serviceType), 'App Service');
  await page.getByRole('checkbox', { name: 'Select recommendation: Single instance', exact: true }).check();
  await page.evaluate(() => {
    const h = (window as any).h;
    h.validation = { ...h.validation, summary: 'Fresh report' }; h.render();
  });
  await page.getByText('Fresh report', { exact: true }).waitFor();
  assert.equal(await page.getByRole('checkbox', { name: 'Select recommendation: Single instance', exact: true }).isChecked(), false);
});
