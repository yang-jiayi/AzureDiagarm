/**
 * Generates a starter template covering every supported service, with a
 * dependency chain through all of them, and puts it through the real Bicep
 * and Terraform compilers.
 *
 * Deliberately NOT part of `npm test`: it needs the Azure CLI and Terraform
 * on PATH and downloads the azurerm provider. The offline structural guards
 * live in tests/iacRoundTrip.test.ts; this is the belt-and-braces check to run
 * after touching the stub generators.
 *
 *   npm run verify:iac
 */
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { Node, Edge } from 'reactflow';
import { buildStarterTemplate, listBicepSupportedTypes } from '../src/services/iacRoundTrip';
import { lookupServiceMeta } from '../src/services/armExtractor';

const armTypes = listBicepSupportedTypes();
const nodes: Node[] = armTypes.map((armType, index) => {
  const name = lookupServiceMeta(armType)?.name;
  if (!name) throw new Error(`${armType} is listed as supported but has no service metadata.`);
  return {
    id: `n${index}`,
    type: 'azureNode',
    position: { x: 0, y: 0 },
    data: { label: name, serviceName: name },
  } as unknown as Node;
});

// Chain every resource so each block carries a dependency clause.
const edges: Edge[] = nodes.slice(1).map((node, index) => ({
  id: `e${index}`,
  source: node.id,
  target: nodes[index].id,
}));

const workspace = mkdtempSync(join(tmpdir(), 'azd-iac-'));
let failures = 0;

function run(command: string, args: string[], cwd: string): { ok: boolean; output: string } {
  try {
    const output = execFileSync(command, args, { cwd, encoding: 'utf8', stdio: 'pipe', shell: true });
    return { ok: true, output };
  } catch (error) {
    const shell = error as { stdout?: string; stderr?: string; message?: string };
    return { ok: false, output: `${shell.stdout ?? ''}${shell.stderr ?? ''}` || shell.message || 'failed' };
  }
}

try {
  const bicep = buildStarterTemplate(nodes, 'bicep', edges);
  const terraform = buildStarterTemplate(nodes, 'terraform', edges);
  writeFileSync(join(workspace, 'main.bicep'), bicep.content);
  writeFileSync(join(workspace, 'main.tf'), terraform.content);

  console.log(
    `${armTypes.length} supported type(s); bicep: ${bicep.supportedResourceCount} resources, `
    + `${bicep.dependencyCount} dependencies, ${bicep.todoCount} TODO; `
    + `terraform: ${terraform.supportedResourceCount} resources, `
    + `${terraform.dependencyCount} dependencies, ${terraform.todoCount} TODO`,
  );

  if (bicep.todoCount > 0 || terraform.todoCount > 0) {
    console.error('  FAIL a supported type degraded to a TODO comment');
    failures += 1;
  }

  const bicepBuild = run('az', ['bicep', 'build', '--file', 'main.bicep', '--stdout'], workspace);
  if (!bicepBuild.ok) {
    console.error(`  FAIL az bicep build\n${bicepBuild.output}`);
    failures += 1;
  } else if (/\bWarning BCP\d+/.test(bicepBuild.output)) {
    console.error(`  FAIL az bicep build emitted warnings\n${bicepBuild.output}`);
    failures += 1;
  } else {
    console.log('  ok   az bicep build (no errors, no warnings)');
  }

  const init = run('terraform', ['init', '-backend=false', '-input=false', '-no-color'], workspace);
  if (!init.ok) {
    console.error(`  FAIL terraform init\n${init.output}`);
    failures += 1;
  } else {
    const validate = run('terraform', ['validate', '-no-color'], workspace);
    if (!validate.ok) {
      console.error(`  FAIL terraform validate\n${validate.output}`);
      failures += 1;
    } else {
      console.log('  ok   terraform validate');
    }
  }
} finally {
  rmSync(workspace, { recursive: true, force: true });
}

console.log(failures === 0 ? '\nAll IaC template checks passed.' : `\n${failures} check(s) failed.`);
if (failures > 0) process.exitCode = 1;
