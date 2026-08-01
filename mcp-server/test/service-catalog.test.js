import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

import {
  SERVICE_CATALOG,
  resolvePricingServiceName,
  resolveServiceName,
} from '../dist/serviceCatalog.js';
import { buildIconFileToTypeMap, importArchitecture } from '../dist/importer.js';
import { estimateServiceCost } from '../dist/pricing.js';

const generatedIconMap = JSON.parse(
  readFileSync(new URL('../dist/iconMap.generated.json', import.meta.url), 'utf8'),
);

test('generated services and aliases are available through the MCP catalog', () => {
  assert.ok(SERVICE_CATALOG['Microsoft Foundry']);
  assert.equal(resolveServiceName('Microsoft Foundry'), 'Microsoft Foundry');
  assert.equal(resolveServiceName('Azure AI Foundry'), 'Microsoft Foundry');
  assert.equal(resolveServiceName('ロードバランサー'), 'Load Balancer');
  assert.equal(resolveServiceName('負荷分散'), 'Load Balancer');
});

test('generated services retain pricing metadata', () => {
  for (const key of ['AML Managed Compute', 'Virtual Machine Scale Sets']) {
    const service = SERVICE_CATALOG[key];
    assert.equal(service.hasPricingData, true);
    assert.equal(service.pricingServiceName, 'Virtual Machines');
    assert.equal(
      estimateServiceCost({ pricingServiceName: service.pricingServiceName }).hasPricingData,
      true,
    );
  }
});

test('Fabric capacity owns the numeric Fabric pricing data', () => {
  assert.equal(SERVICE_CATALOG['Microsoft Fabric'].hasPricingData, false);
  assert.equal(resolvePricingServiceName('Microsoft Fabric'), null);
  assert.equal(
    SERVICE_CATALOG['Microsoft Fabric Capacity'].pricingServiceName,
    'Microsoft Fabric',
  );
  const pricingName = resolvePricingServiceName('Microsoft Fabric Capacity');
  assert.equal(pricingName, 'Microsoft Fabric');
  assert.equal(
    estimateServiceCost({ pricingServiceName: pricingName }).hasPricingData,
    true,
  );
  assert.equal(
    estimateServiceCost({
      pricingServiceName: pricingName,
      region: 'unsupported-region',
    }).hasPricingData,
    false,
  );
});

test('React Flow imports prefer the persisted canonical service name', () => {
  const result = importArchitecture({
    nodes: [{
      id: 'vm-prod',
      type: 'azureNode',
      data: {
        label: 'vm-prod',
        serviceName: 'Virtual Machines',
        iconPath: '/Azure_Public_Service_Icons/Icons/compute/virtual-machines.svg',
      },
    }],
    edges: [],
  }, {
    iconFileToType: { 'virtual-machines': 'AML Managed Compute' },
  });
  assert.equal(result.services[0].name, 'vm-prod');
  assert.equal(result.services[0].type, 'Virtual Machines');
});

test('React Flow imports do not guess a service from an ambiguous icon', () => {
  const iconFileToType = buildIconFileToTypeMap(generatedIconMap);
  assert.equal(iconFileToType['virtual-machines'], undefined);
  assert.equal(iconFileToType['app-service'], 'App Service');

  const result = importArchitecture({
    nodes: [{
      id: 'legacy-vm',
      type: 'azureNode',
      data: {
        label: 'Virtual Machines',
        iconPath: '/Azure_Public_Service_Icons/Icons/compute/virtual-machines.svg',
      },
    }],
    edges: [],
  }, { iconFileToType });
  assert.equal(result.services[0].type, 'Virtual Machines');
});

test('generated canonical names and aliases resolve without legacy conflicts', () => {
  assert.equal(resolveServiceName('Azure API for FHIR'), 'Azure API for FHIR');
  assert.equal(resolveServiceName('Legacy FHIR API'), 'Azure API for FHIR');
  assert.equal(resolveServiceName('Azure Health Data Services'), 'Azure Health Data Services');
  assert.equal(resolveServiceName('Health Data Services'), 'Azure Health Data Services');
  assert.equal(resolveServiceName('FHIR'), 'Azure Health Data Services');
  assert.equal(resolveServiceName('FHIR Service'), 'Azure Health Data Services');
});

test('generated alias collisions follow the web catalog source order', () => {
  assert.equal(resolveServiceName('Power BI'), 'Power BI Embedded');
  assert.equal(resolveServiceName('Power BI (Fabric)'), 'Fabric Power BI Workload');
  assert.equal(resolveServiceName('Purview'), 'Microsoft Purview');
  assert.equal(resolveServiceName('Purview (Fabric)'), 'Fabric Purview Workload');
  assert.equal(resolveServiceName('Dashboard'), 'Azure Dashboard');
  assert.equal(resolveServiceName('Fabric Dashboard'), 'Power BI Dashboard');
});

test('non-conflicting legacy aliases remain available', () => {
  assert.equal(resolveServiceName('Azure Fabric'), 'Microsoft Fabric');
});

test('the MCP catalog mirrors every generated service and alias deterministically', () => {
  const canonicalNames = new Set(
    Object.keys(SERVICE_CATALOG).map(key => key.toLowerCase()),
  );
  const expectedAliases = new Map();

  for (const [key, generated] of Object.entries(generatedIconMap)) {
    const service = SERVICE_CATALOG[key];
    assert.ok(service, `Missing generated service ${key}`);
    assert.equal(resolveServiceName(key.toUpperCase()), key);
    assert.equal(service.displayName, generated.displayName);
    assert.equal(service.category, generated.category);
    assert.equal(service.hasPricingData, generated.hasPricingData);
    assert.equal(service.pricingServiceName, generated.pricingServiceName);
    assert.equal(service.isUsageBased, generated.isUsageBased);
    assert.equal(service.costRange, generated.costRange);

    for (const alias of [generated.displayName, ...generated.aliases]) {
      const normalized = alias.trim().toLowerCase();
      if (
        normalized
        && !canonicalNames.has(normalized)
        && !expectedAliases.has(normalized)
      ) {
        expectedAliases.set(normalized, key);
      }
    }
  }

  for (const [alias, key] of expectedAliases) {
    assert.equal(resolveServiceName(alias), key, `Unexpected owner for alias ${alias}`);
  }
});
