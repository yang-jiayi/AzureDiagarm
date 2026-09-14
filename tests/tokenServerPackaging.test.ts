import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { posix } from 'node:path';
import test from 'node:test';
import ts from 'typescript';

const root = new URL('../', import.meta.url);
const dockerfile = readFileSync(new URL('Dockerfile', root), 'utf8');

function runtimeModules(): Set<string> {
  const visited = new Set<string>();
  function trace(file: string) {
    if (visited.has(file)) return;
    assert.ok(file.startsWith('server/'), `Review the packaging of runtime input ${file}`);
    visited.add(file);
    if (!file.endsWith('.js')) return;
    const source = ts.createSourceFile(file, readFileSync(new URL(file, root), 'utf8'), ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
    function visit(node: ts.Node) {
      if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === 'require') {
        const argument = node.arguments[0];
        assert.ok(argument && ts.isStringLiteral(argument), `${file}: review a computed runtime require`);
        if (argument.text.startsWith('.')) {
          const dependency = posix.normalize(posix.join(posix.dirname(file), argument.text));
          trace(posix.extname(dependency) ? dependency : `${dependency}.js`);
        }
      }
      ts.forEachChild(node, visit);
    }
    visit(source);
  }
  trace('server/token-server.js');
  return visited;
}

function missingModules(recipe: string) {
  const start = recipe.lastIndexOf('WORKDIR /srv/token-server');
  const end = recipe.indexOf('WORKDIR /srv/mcp-server', start);
  assert.ok(start >= 0 && end > start, 'Review changed runtime stage boundaries');
  const copies = [...recipe.slice(start, end).matchAll(/^COPY ([^\r\n]+) \.\/\s*$/gm)]
    .flatMap(match => match[1].trim().split(/\s+/))
    .map(value => new RegExp(`^${value.split('*').map(part => part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('.*')}$`));
  return [...runtimeModules()].filter(file => !copies.some(copy => copy.test(file)));
}

test('the runtime image includes every transitively required token-server module', () => {
  const required = runtimeModules();
  assert.ok(required.has('server/ai-http.js'));
  assert.ok(required.has('server/ai-jobs.js'));
  assert.deepEqual(missingModules(dockerfile), []);
});

test('the packaging gate detects an omitted new runtime helper even if it exists in the build stage', () => {
  const incomplete = dockerfile.replace('COPY server/ai-http.js server/ai-jobs.js ./', '');
  assert.deepEqual(missingModules(incomplete).sort(), ['server/ai-http.js', 'server/ai-jobs.js']);
});
