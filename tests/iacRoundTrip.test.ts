import test from 'node:test';
import assert from 'node:assert/strict';
import type { Edge, Node } from 'reactflow';
import { lookupServiceMeta } from '../src/services/armExtractor.ts';
import {
  buildIaCBaseline,
  buildStarterTemplate,
  compareDiagramToBaseline,
  listBicepSupportedTypes,
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

function edge(id: string, source: string, target: string): Edge {
  return { id, source, target } as Edge;
}

/**
 * The diagram is the only place the user states how the services relate, so
 * dropping that on export left them with a template whose resources deploy in
 * an arbitrary order.
 */
test('buildStarterTemplate turns diagram connections into deployment ordering', () => {
  const nodes = [
    azureNode('web', 'App Service'),
    azureNode('store', 'Storage Account'),
    azureNode('vault', 'Key Vault'),
  ];
  // The app talks to storage and Key Vault, so both must exist first.
  const edges = [edge('e1', 'web', 'store'), edge('e2', 'web', 'vault')];

  const bicep = buildStarterTemplate(nodes, 'bicep', edges);
  const storeSymbol = /resource (\w+) 'Microsoft\.Storage\/storageAccounts/.exec(bicep.content)?.[1];
  const vaultSymbol = /resource (\w+) 'Microsoft\.KeyVault\/vaults/.exec(bicep.content)?.[1];
  assert.ok(storeSymbol && vaultSymbol);
  // Ordering follows emission order, which is sorted, so the output is stable.
  assert.match(
    bicep.content,
    new RegExp(`resource \\w+ 'Microsoft\\.Web/sites[\\s\\S]*?\\n  dependsOn: \\[\\n    ${vaultSymbol}\\n    ${storeSymbol}\\n  \\]\\n\\}`),
  );
  assert.equal(bicep.dependencyCount, 2);

  const terraform = buildStarterTemplate(nodes, 'terraform', edges);
  assert.match(
    terraform.content,
    /resource "azurerm_linux_web_app" "[^"]+" \{[\s\S]*?\n  depends_on = \[\n    azurerm_key_vault\.\w+,\n    azurerm_storage_account\.\w+,\n  \]\n\}/,
  );
  assert.equal(terraform.dependencyCount, 2);
});

test('buildStarterTemplate emits no ordering when the diagram has no connections', () => {
  const template = buildStarterTemplate([azureNode('store', 'Storage Account')], 'bicep');
  assert.equal(template.dependencyCount, 0);
  assert.doesNotMatch(template.content, /dependsOn/);
});

/**
 * Both languages reject a dependency cycle outright, so emitting one would
 * hand the user a template that cannot deploy at all — strictly worse than
 * emitting no ordering.
 */
test('buildStarterTemplate breaks dependency cycles instead of emitting them', () => {
  const nodes = [
    azureNode('a', 'Storage Account'),
    azureNode('b', 'Key Vault'),
    azureNode('c', 'Container Registry'),
  ];
  const cyclic = [edge('e1', 'a', 'b'), edge('e2', 'b', 'c'), edge('e3', 'c', 'a')];

  const bicep = buildStarterTemplate(nodes, 'bicep', cyclic);
  assert.equal(bicep.dependencyCount, 2, 'exactly one back edge must be dropped');

  // Verify the emitted ordering really is acyclic rather than merely smaller.
  const blocks = [...bicep.content.matchAll(/resource (\w+) '[^']+' = \{\n([\s\S]*?)\n\}/g)];
  const graph = new Map<string, string[]>();
  for (const block of blocks) {
    const clause = /\n {2}dependsOn: \[\n([\s\S]*?)\n {2}\]/.exec(block[2]);
    graph.set(block[1], clause ? [...clause[1].matchAll(/^ {4}(\w+)$/gm)].map((m) => m[1]) : []);
  }
  const state = new Map<string, number>();
  const visit = (symbol: string): void => {
    assert.notEqual(state.get(symbol), 1, `cycle reached ${symbol}`);
    if (state.get(symbol) === 2) return;
    state.set(symbol, 1);
    for (const next of graph.get(symbol) ?? []) visit(next);
    state.set(symbol, 2);
  };
  for (const symbol of graph.keys()) visit(symbol);
  assert.equal(graph.size, 3);
});

test('buildStarterTemplate ignores connections to services it could not model', () => {
  const nodes = [azureNode('web', 'App Service'), azureNode('lb', 'Load Balancer')];
  const template = buildStarterTemplate(nodes, 'bicep', [
    edge('e1', 'web', 'lb'),
    edge('e2', 'web', 'missing-node'),
    edge('e3', 'web', 'web'),
  ]);
  assert.equal(template.dependencyCount, 0);
  assert.doesNotMatch(template.content, /dependsOn/);
});

test('every Bicep type listed as supported actually emits a resource block', () => {
  for (const armType of listBicepSupportedTypes()) {
    const meta = lookupServiceMeta(armType);
    assert.ok(meta?.name, `${armType} has no service metadata, so the palette cannot produce it`);
    const template = buildStarterTemplate([azureNode('n1', meta.name)], 'bicep');
    assert.equal(
      template.todoCount,
      0,
      `${armType} is listed as supported but "${meta.name}" degraded to a TODO comment`,
    );
  }
});

/**
 * Bicep has no decimal literal, so `cpu: 0.5` in the Container Apps stub never
 * compiled — `az bicep build` rejected it with BCP020. Numeric fractions have
 * to go through json(). Nothing in the type system catches this because the
 * template is assembled as strings.
 */
test('generated Bicep contains no bare decimal literal', () => {
  const nodes = listBicepSupportedTypes().map((armType, index) => {
    const name = lookupServiceMeta(armType)?.name;
    assert.ok(name, `${armType} has no service metadata`);
    return azureNode(`n${index}`, name);
  });
  const content = buildStarterTemplate(nodes, 'bicep').content;
  const offenders = [...content.matchAll(/^\s*[\w']+:\s*-?\d+\.\d+\s*$/gm)].map((m) => m[0].trim());
  assert.deepEqual(offenders, [], 'Wrap fractional values in json(\'...\') so Bicep can parse them.');
  assert.match(content, /cpu: json\('0\.5'\)/);
});

/**
 * Bicep cannot infer a lower bound through take(), so without the constraint
 * and the suppression every export opens with BCP334 warnings.
 */
test('generated Bicep constrains namePrefix and suppresses the unprovable length warning', () => {
  const content = buildStarterTemplate([
    azureNode('store', 'Storage Account'),
    azureNode('acr', 'Container Registry'),
  ], 'bicep').content;

  assert.match(content, /@minLength\(3\)\n@maxLength\(12\)\nparam namePrefix string/);

  const lines = content.split('\n');
  const truncated = lines
    .map((line, index) => ({ line, index }))
    .filter((entry) => /^ {2}name: take\(/.test(entry.line));
  assert.ok(truncated.length >= 2, 'no truncated name expressions found, so the guard is stale');
  for (const entry of truncated) {
    assert.match(
      lines[entry.index - 1],
      /^ {2}#disable-next-line BCP334/,
      `"${entry.line.trim()}" would emit an unsuppressed BCP334 warning`,
    );
  }
});

test('every generated resource block closes every brace it opens', () => {
  const nodes = listBicepSupportedTypes().map((armType, index) => (
    azureNode(`n${index}`, lookupServiceMeta(armType)!.name)
  ));
  const edges = nodes.slice(1).map((node, index) => edge(`e${index}`, node.id, nodes[index].id));

  for (const format of ['bicep', 'terraform'] as const) {
    const content = buildStarterTemplate(nodes, format, edges).content;
    let depth = 0;
    let bracket = 0;
    for (const line of content.split('\n')) {
      const code = line.replace(/'[^']*'/g, "''").replace(/"[^"]*"/g, '""');
      for (const character of code) {
        if (character === '{') depth += 1;
        else if (character === '}') depth -= 1;
        else if (character === '[') bracket += 1;
        else if (character === ']') bracket -= 1;
        assert.ok(depth >= 0 && bracket >= 0, `${format} closes a delimiter that was never opened: ${line}`);
      }
    }
    assert.equal(depth, 0, `${format} leaves ${depth} brace(s) open`);
    assert.equal(bracket, 0, `${format} leaves ${bracket} bracket(s) open`);
  }
});
