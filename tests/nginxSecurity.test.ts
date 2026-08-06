// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import test from 'node:test';

const nginx = readFileSync(new URL('../nginx.conf', import.meta.url), 'utf8');
const policies = Array.from(
  nginx.matchAll(/add_header Content-Security-Policy "([^"]+)" always;/g),
  (match) => match[1],
);

test('nginx applies one consistent CSP to every response block with custom headers', () => {
  assert.equal(policies.length, 8);
  assert.equal(new Set(policies).size, 1);
});

test('the production CSP blocks executable and embedding fallbacks', () => {
  const [policy] = policies;
  assert.match(policy, /default-src 'self'/);
  assert.match(policy, /object-src 'none'/);
  assert.match(policy, /frame-ancestors 'none'/);
  assert.match(policy, /base-uri 'self'/);
  assert.match(policy, /script-src 'self' 'wasm-unsafe-eval'/);
  assert.doesNotMatch(policy, /script-src[^;]*'unsafe-eval'/);
});

test('the production CSP permits only required application integrations', () => {
  const [policy] = policies;
  assert.match(policy, /https:\/\/login\.microsoftonline\.com/);
  assert.match(policy, /https:\/\/management\.azure\.com/);
  assert.match(policy, /https:\/\/\*\.applicationinsights\.azure\.com/);
  assert.match(policy, /wss:\/\/\*\.speech\.microsoft\.com/);
  assert.match(policy, /worker-src 'self' blob:/);
});
