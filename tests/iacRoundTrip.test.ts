import test from 'node:test';
import assert from 'node:assert/strict';
import type { Node } from 'reactflow';
import {
  buildIaCBaseline,
  buildStarterTemplate,
  compareDiagramToBaseline,
  parseDeploymentPlan,
  restoreIaCBaseline,
} from '../src/services/iacRoundTrip.ts';

function azureNode(id: string, label: string, serviceName = label): Node {
  return {
    id,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: {
      label,
      serviceName,
    },
  } as unknown as Node;
}

test('buildIaCBaseline parses Bicep resources without executing expressions', () => {
  const baseline = buildIaCBaseline({
    format: 'bicep',
    importedAt: '2026-08-02T00:00:00.000Z',
    files: [{
      name: 'main.bicep',
      text: `
resource stg 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: 'stgacct01'
  location: resourceGroup().location
}

resource kv 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: '\${namePrefix}-kv'
  location: resourceGroup().location
}
      `,
    }],
  });

  assert.equal(baseline.resourceCount, 2);
  assert.equal(baseline.formatLabel, 'Bicep');
  assert.equal(baseline.resources[0].mappedService, 'Key Vault');
  assert.equal(baseline.resources[0].approximation, 'type-only');
  assert.equal(baseline.resources[1].mappedService, 'Storage Account');
  assert.equal(baseline.resources[1].resourceName, 'stgacct01');
  assert.deepEqual(restoreIaCBaseline(baseline), baseline);
});

test('compareDiagramToBaseline normalizes Terraform aliases and identifies new nodes', () => {
  const baseline = buildIaCBaseline({
    format: 'terraform-hcl',
    files: [{
      name: 'main.tf',
      text: `
resource "azurerm_linux_web_app" "api" {
  name = "orders-api"
}
resource "azurerm_storage_account" "logs" {
  name = "orderslogs"
}
      `,
    }],
  });

  const report = compareDiagramToBaseline([
    azureNode('app', 'App Service'),
    azureNode('db', 'Azure Cache for Redis'),
  ], baseline);

  assert.ok(report);
  assert.equal(report.matched.length, 1);
  assert.equal(report.matched[0].baseline.mappedService, 'App Service');
  assert.equal(report.sourceOnly.length, 1);
  assert.equal(report.diagramOnly.length, 1);
});

test('parseDeploymentPlan supports Azure what-if and Terraform plan JSON', () => {
  const terraform = parseDeploymentPlan('tfplan.json', JSON.stringify({
    resource_changes: [
      {
        address: 'azurerm_storage_account.logs',
        type: 'azurerm_storage_account',
        name: 'logs',
        change: { actions: ['create'] },
      },
      {
        address: 'azurerm_linux_web_app.api',
        type: 'azurerm_linux_web_app',
        name: 'api',
        change: { actions: ['delete', 'create'] },
      },
    ],
  }));
  assert.equal(terraform.kind, 'terraform-plan');
  assert.equal(terraform.changeCounts.create, 1);
  assert.equal(terraform.changeCounts.replace, 1);

  const whatIf = parseDeploymentPlan('whatif.json', JSON.stringify({
    properties: {
      changes: [
        {
          changeType: 'Modify',
          resourceId: '/subscriptions/123/resourceGroups/demo/providers/Microsoft.Storage/storageAccounts/demo',
          after: {
            resourceType: 'Microsoft.Storage/storageAccounts',
            name: 'demo',
          },
        },
      ],
    },
  }));
  assert.equal(whatIf.kind, 'azure-what-if');
  assert.equal(whatIf.changeCounts.update, 1);
  assert.equal(whatIf.changes[0].resourceName, 'demo');
});

test('buildStarterTemplate keeps unsupported services as TODO comments', () => {
  const bicep = buildStarterTemplate([
    azureNode('storage', 'Storage Account'),
    azureNode('lb', 'Load Balancer'),
  ], 'bicep');

  assert.match(bicep.content, /Microsoft\.Storage\/storageAccounts/);
  assert.match(bicep.content, /TODO: Model unsupported service "Load Balancer"/);
  assert.equal(bicep.supportedResourceCount, 1);
  assert.equal(bicep.todoCount, 1);
});

test('buildStarterTemplate requires secure database password inputs', () => {
  const nodes = [
    azureNode('sql', 'SQL Server'),
    azureNode('postgresql', 'PostgreSQL'),
    azureNode('mysql', 'MySQL'),
  ];
  const bicep = buildStarterTemplate(nodes, 'bicep');
  const terraform = buildStarterTemplate(nodes, 'terraform');

  assert.doesNotMatch(bicep.content, /ChangeMe-UseKeyVault/);
  assert.match(bicep.content, /@secure\(\)\nparam sqlAdministratorPassword string/);
  assert.match(bicep.content, /@secure\(\)\nparam postgresqlAdministratorPassword string/);
  assert.match(bicep.content, /@secure\(\)\nparam mysqlAdministratorPassword string/);
  assert.match(bicep.content, /administratorLoginPassword: sqlAdministratorPassword/);

  assert.doesNotMatch(terraform.content, /ChangeMe-UseKeyVault/);
  assert.match(terraform.content, /variable "sql_administrator_password" \{[\s\S]*?sensitive\s+= true/);
  assert.match(terraform.content, /variable "postgresql_administrator_password" \{[\s\S]*?sensitive\s+= true/);
  assert.match(terraform.content, /variable "mysql_administrator_password" \{[\s\S]*?sensitive\s+= true/);
  assert.match(terraform.content, /administrator_login_password = var\.sql_administrator_password/);
});
