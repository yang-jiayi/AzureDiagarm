import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import {
  policySetNameForAssignments,
  verifyAccessStore,
  verifyAfdCustomDomain,
  verifyNoReplicaReduction,
  verifyPreservedRuntimeConfiguration,
  verifyRuntimeConfiguration,
  verifyStoragePolicyExemption,
} from '../scripts/verify-deployment-preflight.mjs';

const accessEndpoint = 'https://access.table.example.net/';
const budgetEndpoint = 'https://budget.table.example.net/';
const keyVaultId = '/subscriptions/sub/resourceGroups/rg/providers/Microsoft.KeyVault/vaults/access';
const policySetId = '/providers/Microsoft.Authorization/policySetDefinitions/1f3afdf9-d0c9-4c3d-847f-89da613e70a8';
const assignmentScope = '/providers/Microsoft.Management/managementGroups/02310481-6752-4d53-8fc4-f597a8e047f4';
const assignmentId = `${assignmentScope}/providers/Microsoft.Authorization/policyAssignments/Azure_Security_Baseline`;
const storageScope = '/subscriptions/sub/resourceGroups/rg/providers/Microsoft.Storage/storageAccounts/diagrams';
const exemptionId = `${storageScope}/providers/Microsoft.Authorization/policyExemptions/AzureDiagarm-Storage-NSP`;
const tableRuntime = {
  AZURE_TABLES_ACCESS_ENDPOINT: accessEndpoint,
  MCP_ENABLED: 'true',
  ALLOW_BYO_AI_ENDPOINTS: 'false',
  AI_BUDGET_STORE: 'table',
  AI_DAILY_TOKEN_BUDGET: '250000',
  AI_MAX_CONCURRENT_REQUESTS: '2',
  AZURE_TABLES_BUDGET_ENDPOINT: budgetEndpoint,
};
const assignments = [{
  name: 'Azure_Security_Baseline',
  id: assignmentId,
  properties: {
    scope: assignmentScope,
    policyDefinitionId: policySetId,
  },
}];
const policySet = {
  id: policySetId,
  properties: {
    policyDefinitions: [
      { policyDefinitionReferenceId: 'storageAccountsShouldRestrictNetworkAccessUsingVirtualNetworkRulesMonitoringEffect' },
      { policyDefinitionReferenceId: 'storageAccountShouldUseAPrivateLinkConnectionMonitoringEffect' },
      { policyDefinitionReferenceId: 'unrelatedPolicy' },
    ],
  },
};
const policyReferences = policySet.properties.policyDefinitions.slice(0, 2)
  .map(definition => definition.policyDefinitionReferenceId);
const validExemption = {
  id: exemptionId,
  properties: {
    policyAssignmentId: assignmentId,
    exemptionCategory: 'Mitigated',
    expiresOn: '2027-08-04T00:00:00Z',
    policyDefinitionReferenceIds: policyReferences.map(reference => reference.toLowerCase()),
  },
};

test('validates exactly one access store and retains the selected Tables or Key Vault path', () => {
  assert.deepEqual(verifyAccessStore({ AZURE_TABLES_ACCESS_ENDPOINT: accessEndpoint }), {
    type: 'tables',
    value: accessEndpoint,
  });
  assert.deepEqual(verifyAccessStore({ AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: keyVaultId }), {
    type: 'key-vault',
    value: keyVaultId,
  });
  for (const env of [
    {},
    { AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: keyVaultId, AZURE_TABLES_ACCESS_ENDPOINT: accessEndpoint },
    { AZURE_TABLES_ACCESS_ENDPOINT: 'http://access.table.example.net' },
    { AZURE_TABLES_ACCESS_ENDPOINT: 'https://user:pass@access.table.example.net' },
    { AZURE_TABLES_ACCESS_ENDPOINT: 'https://access.table.example.net/path' },
    { AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: 'not-a-resource-id' },
  ]) {
    assert.throws(() => verifyAccessStore(env));
  }
});

test('requires explicit production runtime settings and bounded numeric budgets', () => {
  assert.deepEqual(verifyRuntimeConfiguration(tableRuntime), {
    accessStore: { type: 'tables', value: accessEndpoint },
    budgetEndpoint,
  });
  for (const [name, value, message] of [
    ['AZURE_ACCESS_KEY_VAULT_RESOURCE_ID', keyVaultId, /Exactly one access-list store/],
    ['MCP_ENABLED', undefined, /MCP_ENABLED/],
    ['MCP_ENABLED', 'false', /MCP_ENABLED/],
    ['ALLOW_BYO_AI_ENDPOINTS', undefined, /ALLOW_BYO_AI_ENDPOINTS/],
    ['ALLOW_BYO_AI_ENDPOINTS', 'true', /ALLOW_BYO_AI_ENDPOINTS/],
    ['AI_BUDGET_STORE', undefined, /AI_BUDGET_STORE/],
    ['AI_BUDGET_STORE', 'cosmos', /AI_BUDGET_STORE/],
    ['AI_DAILY_TOKEN_BUDGET', undefined, /AI_DAILY_TOKEN_BUDGET/],
    ['AI_DAILY_TOKEN_BUDGET', '0', /AI_DAILY_TOKEN_BUDGET/],
    ['AI_DAILY_TOKEN_BUDGET', '1.2', /AI_DAILY_TOKEN_BUDGET/],
    ['AI_MAX_CONCURRENT_REQUESTS', undefined, /AI_MAX_CONCURRENT_REQUESTS/],
    ['AI_MAX_CONCURRENT_REQUESTS', '9007199254740992', /AI_MAX_CONCURRENT_REQUESTS/],
    ['AZURE_TABLES_BUDGET_ENDPOINT', undefined, /AZURE_TABLES_BUDGET_ENDPOINT/],
  ] as const) {
    const invalid = { ...tableRuntime, [name]: value };
    if (value === undefined) delete invalid[name as keyof typeof invalid];
    assert.throws(() => verifyRuntimeConfiguration(invalid), message);
  }
});

test('the Windows-safe runtime-config CLI validates explicit settings', () => {
  const result = spawnSync(process.execPath, [
    fileURLToPath(new URL('../scripts/verify-deployment-preflight.mjs', import.meta.url)),
    'runtime-config',
  ], {
    encoding: 'utf8',
    env: { ...process.env, ...tableRuntime },
    timeout: 30_000,
  });
  assert.equal(result.error, undefined);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /Production runtime configuration validated/);
});

test('requires deployment access and budget stores to preserve current runtime selection', () => {
  const currentEnv = Object.entries(tableRuntime).map(([name, value]) => ({ name, value }));
  assert.doesNotThrow(() => verifyPreservedRuntimeConfiguration(tableRuntime, currentEnv));
  assert.throws(
    () => verifyPreservedRuntimeConfiguration(
      { ...tableRuntime, AZURE_ACCESS_KEY_VAULT_RESOURCE_ID: keyVaultId },
      currentEnv,
    ),
    /Exactly one access-list store/,
  );
  assert.throws(
    () => verifyPreservedRuntimeConfiguration(
      { ...tableRuntime, AI_DAILY_TOKEN_BUDGET: '300000' },
      currentEnv,
    ),
    /does not preserve the current production value/,
  );
  assert.throws(
    () => verifyPreservedRuntimeConfiguration(tableRuntime, currentEnv.slice(1)),
    /Exactly one access-list store/,
  );
});

test('accepts only the approved and provisioned matching AFD domain with TLS 1.2', () => {
  const domain = {
    id: '/subscriptions/sub/resourceGroups/rg/providers/Microsoft.Cdn/profiles/afd/customDomains/app',
    hostName: 'app.example.net',
    domainValidationState: 'Approved',
    provisioningState: 'Succeeded',
    tlsSettings: { minimumTlsVersion: 'TLS12', certificateType: 'ManagedCertificate' },
  };
  assert.equal(verifyAfdCustomDomain([domain], 'APP.EXAMPLE.NET'), domain.id);
  assert.equal(verifyAfdCustomDomain([
    { ...domain, tlsSettings: { ...domain.tlsSettings, certificateType: 'CustomerCertificate' } },
  ], 'app.example.net'), domain.id);
  for (const [changedDomain, host] of [
    [{ ...domain, domainValidationState: 'Pending' }, 'app.example.net'],
    [{ ...domain, provisioningState: 'Updating' }, 'app.example.net'],
    [{ ...domain, tlsSettings: { ...domain.tlsSettings, minimumTlsVersion: 'TLS10' } }, 'app.example.net'],
    [{ ...domain, tlsSettings: { ...domain.tlsSettings, certificateType: 'Unknown' } }, 'app.example.net'],
    [domain, 'other.example.net'],
  ] as const) {
    assert.throws(() => verifyAfdCustomDomain([changedDomain], host));
  }
  assert.throws(() => verifyAfdCustomDomain([domain, domain], 'app.example.net'), /approved TLS requirements/);
});

test('resolves and verifies the actual assigned policy initiative and preserves reference spelling', () => {
  assert.equal('scope' in validExemption, false);
  assert.equal(policySetNameForAssignments(assignments), policySetId.split('/').at(-1));
  const childAssignment = {
    ...assignments[0],
    id: `${storageScope}/providers/Microsoft.Authorization/policyAssignments/Azure_Security_Baseline`,
    properties: {
      ...assignments[0].properties,
      scope: storageScope,
    },
  };
  assert.equal(policySetNameForAssignments([...assignments, childAssignment]), policySetId.split('/').at(-1));
  assert.deepEqual(
    verifyStoragePolicyExemption({
      assignments,
      policySet,
      exemption: validExemption,
      storageScope,
      now: Date.parse('2026-06-01T00:00:00Z'),
    }),
    policyReferences,
  );
});

test('fails closed for missing or mismatched assignment, initiative, exemption, or references', () => {
  const mismatches = [
    { assignments: [], policySet, exemption: validExemption, storageScope },
    {
      assignments: [{
        ...assignments[0],
        properties: { ...assignments[0].properties, policyDefinitionId: `${policySetId}-other` },
      }],
      policySet,
      exemption: validExemption,
      storageScope,
    },
    {
      assignments: [{
        ...assignments[0],
        id: `${assignmentScope}/providers/Microsoft.Authorization/policyAssignments/child`,
        properties: { ...assignments[0].properties, scope: `${storageScope}/child` },
      }],
      policySet,
      exemption: validExemption,
      storageScope,
    },
    { assignments, policySet: { ...policySet, id: `${policySetId}-other` }, exemption: validExemption, storageScope },
    {
      assignments,
      policySet: { ...policySet, properties: { policyDefinitions: [{ policyDefinitionReferenceId: 'unrelatedPolicy' }] } },
      exemption: validExemption,
      storageScope,
    },
    {
      assignments,
      policySet: {
        ...policySet,
        properties: {
          policyDefinitions: [
            ...policySet.properties.policyDefinitions,
            { policyDefinitionReferenceId: 'storageAccountsShouldUseAPrivateLinkConnectionAuditEffect' },
          ],
        },
      },
      exemption: validExemption,
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: { ...validExemption, properties: { ...validExemption.properties, policyAssignmentId: `${assignmentId}-other` } },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: { ...validExemption, id: `${storageScope}/providers/Microsoft.Authorization/policyExemptions/other` },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        id: '/subscriptions/sub/resourceGroups/rg/providers/Microsoft.Storage/storageAccounts/other/providers/Microsoft.Authorization/policyExemptions/AzureDiagarm-Storage-NSP',
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        id: '/subscriptions/other/resourceGroups/rg/providers/Microsoft.Storage/storageAccounts/diagrams/providers/Microsoft.Authorization/policyExemptions/AzureDiagarm-Storage-NSP',
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        id: '/subscriptions/sub/resourceGroups/rg/providers/Microsoft.Authorization/policyExemptions/AzureDiagarm-Storage-NSP',
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        id: `${storageScope}/child/providers/Microsoft.Authorization/policyExemptions/AzureDiagarm-Storage-NSP`,
        scope: storageScope,
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        properties: { ...validExemption.properties, exemptionCategory: 'Waiver' },
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        properties: { ...validExemption.properties, expiresOn: '2026-01-01T00:00:00Z' },
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        properties: { ...validExemption.properties, policyDefinitionReferenceIds: [policyReferences[0]] },
      },
      storageScope,
    },
    {
      assignments,
      policySet,
      exemption: {
        ...validExemption,
        properties: { ...validExemption.properties, policyDefinitionReferenceIds: undefined },
      },
      storageScope,
    },
  ];
  for (const input of mismatches) {
    assert.throws(() => verifyStoragePolicyExemption({
      ...input,
      now: Date.parse('2026-06-01T00:00:00Z'),
    }));
  }
});

test('never silently reduces the current maximum replica count', () => {
  assert.doesNotThrow(() => verifyNoReplicaReduction('2', '2'));
  assert.doesNotThrow(() => verifyNoReplicaReduction('2', '3'));
  for (const current of ['3', 'unknown', '0', '2.0', '2e0', '']) {
    assert.throws(() => verifyNoReplicaReduction(current, '2'));
  }
  assert.throws(() => verifyNoReplicaReduction('2', '2.0'));
});
