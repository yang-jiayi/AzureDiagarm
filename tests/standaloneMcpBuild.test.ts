import assert from 'node:assert/strict';
import { readFileSync, statSync } from 'node:fs';
import { posix, resolve } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const repoRoot = fileURLToPath(new URL('../', import.meta.url));
const diskPath = (relativePath: string) => resolve(repoRoot, ...relativePath.split('/'));
const dockerfile = readFileSync(diskPath('mcp-server/Dockerfile'), 'utf8');
const manifest = JSON.parse(readFileSync(diskPath('mcp-server/package.json'), 'utf8')) as {
  scripts: Record<string, string>;
};

function requiredBuildInputs(): Set<string> {
  const inputs = new Set([
    'mcp-server/package.json',
    'mcp-server/package-lock.json',
    'mcp-server/tsconfig.json',
    'mcp-server/src',
  ]);
  const visitedScripts = new Set<string>();

  function traceModule(modulePath: string): void {
    if (inputs.has(modulePath)) return;
    inputs.add(modulePath);
    const source = ts.createSourceFile(
      modulePath,
      readFileSync(diskPath(modulePath), 'utf8'),
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.JS,
    );
    function visit(node: ts.Node): void {
      if (ts.isImportDeclaration(node) && ts.isStringLiteral(node.moduleSpecifier)) {
        const specifier = node.moduleSpecifier.text;
        if (specifier.startsWith('.')) {
          traceModule(posix.normalize(posix.join(posix.dirname(modulePath), specifier)));
        }
      }
      if (
        ts.isCallExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === 'resolve' &&
        node.arguments[0] &&
        ts.isIdentifier(node.arguments[0]) &&
        node.arguments[0].text === 'repoRoot'
      ) {
        const parts: string[] = [];
        for (const argument of node.arguments.slice(1)) {
          if (!ts.isStringLiteral(argument)) break;
          parts.push(argument.text);
        }
        assert.ok(parts.length > 0, `${modulePath}: review a new computed repository input`);
        inputs.add(posix.join(...parts));
      }
      ts.forEachChild(node, visit);
    }
    visit(source);
  }

  function traceNpmScript(name: string): void {
    if (visitedScripts.has(name)) return;
    visitedScripts.add(name);
    const command = manifest.scripts[name];
    assert.equal(typeof command, 'string', `missing npm script ${name}`);
    for (const match of command.matchAll(/\bnpm run ([\w:-]+)/g)) {
      traceNpmScript(match[1]);
    }
    for (const match of command.matchAll(/\bnode ([\w./-]+\.mjs)/g)) {
      traceModule(posix.join('mcp-server', match[1]));
    }
  }

  traceNpmScript('prebuild');
  traceNpmScript('build');
  return inputs;
}

function missingBuildInputs(recipe: string): string[] {
  let workdir = '/';
  const copies: { source: string; destination: string; directory: boolean }[] = [];
  let foundBuild = false;
  for (const line of recipe.split(/\r?\n/)) {
    const parts = line.trim().split(/\s+/);
    if (parts[0] === 'WORKDIR') workdir = posix.resolve(workdir, parts[1]);
    if (parts[0] === 'COPY') {
      assert.equal(parts.length, 3, 'review changed COPY syntax in the standalone build contract');
      const [, source, target] = parts;
      const directory = source.includes('*') || statSync(diskPath(source)).isDirectory();
      let destination = posix.resolve(workdir, target);
      if (!directory && (target.endsWith('/') || target === '.')) {
        destination = posix.join(destination, posix.basename(source));
      }
      copies.push({ source, destination, directory });
    }
    if (line.trim() === 'RUN npm run build') {
      foundBuild = true;
      break;
    }
  }
  assert.ok(foundBuild, 'the standalone image must run its complete npm build lifecycle');

  return [...requiredBuildInputs()].filter((input) => !copies.some((copy) => {
    let relativePath: string;
    if (copy.source.includes('*')) {
      const pattern = posix.basename(copy.source)
        .split('*').map((part) => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('.*');
      if (
        posix.dirname(input) !== posix.dirname(copy.source) ||
        !new RegExp(`^${pattern}$`).test(posix.basename(input))
      ) return false;
      relativePath = posix.basename(input);
    } else if (input === copy.source) {
      relativePath = '';
    } else if (copy.directory && input.startsWith(`${copy.source}/`)) {
      relativePath = posix.relative(copy.source, input);
    } else {
      return false;
    }
    return posix.join(copy.destination, relativePath) === `/app/${input}`;
  }));
}

test('standalone Docker copies the complete traced MCP build inputs to their runtime paths', () => {
  const inputs = requiredBuildInputs();
  for (const input of [
    'scripts/prep-pricing-data.mjs',
    'public/pricing/regions',
    'src/data/serviceIconMapping.ts',
    'src/data/fabricIconCatalog.ts',
    'src/data/microsoftProductIconCatalog.ts',
    'Azure_Public_Service_Icons/Icons',
    'mcp-server/scripts/copy-build-assets.mjs',
  ]) {
    assert.ok(inputs.has(input), `the build dependency trace must include ${input}`);
  }
  assert.deepEqual(missingBuildInputs(dockerfile), []);
});

test('standalone input coverage detects either missing pricing dependency independently', () => {
  for (const input of ['scripts/prep-pricing-data.mjs', 'public/pricing/regions']) {
    const omitted = dockerfile.split(/\r?\n/)
      .filter((line) => !line.startsWith(`COPY ${input} `)).join('\n');
    assert.ok(missingBuildInputs(omitted).includes(input), `must detect missing ${input}`);
  }
});

test('Linux MCP CI builds the actual standalone image from the repository root', () => {
  const ci = readFileSync(diskPath('.github/workflows/ci.yml'), 'utf8');
  const job = ci.slice(ci.indexOf('\n  build-mcp:'));
  assert.match(job, /runs-on: ubuntu-latest/);
  assert.match(
    job,
    /- name: Build standalone MCP runtime image\s+run: \|\s+docker build \\\s+--file mcp-server\/Dockerfile \\\s+--tag azurediagarm-mcp-validation:"\$\{GITHUB_SHA\}" \./,
  );
});
