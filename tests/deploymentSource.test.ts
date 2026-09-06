import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
  verifyDeploymentSource,
  type DeploymentSource,
} from '../scripts/verify-deployment-source.mjs';

const currentCommit = 'a'.repeat(40);
const newerCommit = 'b'.repeat(40);
const source: DeploymentSource = {
  ref: 'refs/heads/main',
  repository: 'owner/repository',
  expectedCommit: currentCommit,
  checkedOutCommit: currentCommit,
  token: 'test-only-token',
};
const mainRef = (sha: unknown = currentCommit) => ({
  ref: 'refs/heads/main',
  object: { type: 'commit', sha },
});
const jsonResponse = (body: unknown) => new Response(JSON.stringify(body));

test('only the exact validated checkout at current main is deployable', async () => {
  let requests = 0;
  const actual = await verifyDeploymentSource(source, async (url, options) => {
    requests++;
    assert.equal(String(url), 'https://api.github.com/repos/owner/repository/git/ref/heads/main');
    const headers = new Headers(options?.headers);
    assert.equal(headers.get('Authorization'), `Bearer ${source.token}`);
    assert.equal(headers.get('User-Agent'), 'AzureDiagarm-release-source-guard');
    assert.equal(headers.get('Cache-Control'), 'no-cache');
    assert.equal(options?.cache, 'no-store');
    assert.equal(options?.redirect, 'error');
    assert.ok(options?.signal instanceof AbortSignal);
    return jsonResponse(mainRef());
  });
  assert.equal(actual, currentCommit);
  assert.equal(requests, 1);
});

test('a historical main candidate is rejected even when its checkout matches', async () => {
  await assert.rejects(
    verifyDeploymentSource(source, async () => jsonResponse(mainRef(newerCommit))),
    /no longer current main/,
  );
});

test('each guard invocation reads main again, detecting supersession during a build', async () => {
  let requests = 0;
  const fetchMain: typeof fetch = async () => jsonResponse(
    mainRef(requests++ === 0 ? currentCommit : newerCommit),
  );
  assert.equal(await verifyDeploymentSource(source, fetchMain), currentCommit);
  await assert.rejects(verifyDeploymentSource(source, fetchMain), /no longer current main/);
  assert.equal(requests, 2);
});

for (const ref of ['refs/heads/feature', 'refs/tags/main', 'main', '', undefined]) {
  test(`non-main or missing workflow ref is rejected before lookup: ${String(ref)}`, async () => {
    let requested = false;
    await assert.rejects(
      verifyDeploymentSource({ ...source, ref }, async () => {
        requested = true;
        return jsonResponse(mainRef());
      }),
      /requires refs\/heads\/main/,
    );
    assert.equal(requested, false);
  });
}

for (const [label, overrides, message] of [
  ['missing candidate', { expectedCommit: undefined }, /validated deployment commit/],
  ['abbreviated candidate', { expectedCommit: 'a'.repeat(12) }, /validated deployment commit/],
  ['malformed candidate', { expectedCommit: 'g'.repeat(40) }, /validated deployment commit/],
  ['different checkout', { checkedOutCommit: newerCommit }, /checkout does not match/],
  ['missing checkout', { checkedOutCommit: undefined }, /checkout does not match/],
  ['missing repository', { repository: undefined }, /repository is missing or malformed/],
  ['malformed repository', { repository: 'owner/repo?other' }, /repository is missing or malformed/],
  ['repository path traversal', { repository: 'owner/..' }, /repository is missing or malformed/],
  ['missing token', { token: undefined }, /read-only GitHub token is required/],
  ['blank token', { token: ' ' }, /read-only GitHub token is required/],
] satisfies [string, Partial<DeploymentSource>, RegExp][]) {
  test(`invalid deployment input fails closed before lookup: ${label}`, async () => {
    let requested = false;
    await assert.rejects(
      verifyDeploymentSource({ ...source, ...overrides }, async () => {
        requested = true;
        return jsonResponse(mainRef());
      }),
      message,
    );
    assert.equal(requested, false);
  });
}

for (const [ref, expectedCommit, message] of [
  ['refs/heads/feature', currentCommit, /requires refs\/heads\/main/],
  ['refs/heads/main', 'invalid', /validated deployment commit is missing or malformed/],
] satisfies [string, string, RegExp][]) {
  test(`the CLI exits unsuccessfully on refused input: ${ref}, ${expectedCommit}`, () => {
    const result = spawnSync(process.execPath, [
      fileURLToPath(new URL('../scripts/verify-deployment-source.mjs', import.meta.url)),
    ], {
      cwd: fileURLToPath(new URL('../', import.meta.url)),
      encoding: 'utf8',
      timeout: 30_000,
      env: {
        PATH: process.env.PATH,
        SystemRoot: process.env.SystemRoot,
        GITHUB_REF: ref,
        GITHUB_REPOSITORY: source.repository,
        EXPECTED_COMMIT: expectedCommit,
      },
    });
    assert.equal(result.error, undefined);
    assert.equal(result.status, 1);
    assert.match(result.stderr, message);
    assert.equal(result.stdout, '');
  });
}

for (const status of [401, 403, 404, 429, 500, 503]) {
  test(`GitHub HTTP ${status} fails closed without reporting response contents`, async () => {
    await assert.rejects(
      verifyDeploymentSource(
        source,
        async () => new Response('untrusted diagnostic body', { status }),
      ),
      new RegExp(`^Error: GitHub main lookup failed \\(HTTP ${status}\\); refusing deployment\\.$`),
    );
  });
}

test('GitHub network errors fail closed without exposing request credentials', async () => {
  await assert.rejects(
    verifyDeploymentSource(source, async () => {
      throw new Error(`network diagnostic containing ${source.token}`);
    }),
    /^Error: Unable to read current main from GitHub; refusing deployment\.$/,
  );
});

test('GitHub timeouts fail closed', async () => {
  await assert.rejects(
    verifyDeploymentSource(source, async () => {
      throw new DOMException('Request timed out', 'TimeoutError');
    }),
    /Unable to read current main/,
  );
});

test('invalid JSON from GitHub fails closed', async () => {
  await assert.rejects(
    verifyDeploymentSource(source, async () => new Response('not JSON')),
    /invalid main ref response/,
  );
});

for (const [label, body] of [
  ['null response', null],
  ['missing object', { ref: 'refs/heads/main' }],
  ['missing SHA', { ref: 'refs/heads/main', object: { type: 'commit' } }],
  ['abbreviated SHA', mainRef('a'.repeat(12))],
  ['non-hex SHA', mainRef('z'.repeat(40))],
  ['non-string SHA', mainRef(42)],
  ['other ref', { ...mainRef(), ref: 'refs/heads/other' }],
  ['non-commit object', { ref: 'refs/heads/main', object: { type: 'tag', sha: currentCommit } }],
] satisfies [string, unknown][]) {
  test(`malformed GitHub main response fails closed: ${label}`, async () => {
    await assert.rejects(
      verifyDeploymentSource(source, async () => jsonResponse(body)),
      /missing or malformed main commit/,
    );
  });
}

test('workflow guards Azure entry and the revision update before its rollback marker', () => {
  const workflow = readFileSync(
    new URL('../.github/workflows/azurediagarm-sync-deploy.yml', import.meta.url),
    'utf8',
  );
  const deploy = workflow.slice(workflow.indexOf('\n  deploy:'));
  const guard = 'node scripts/verify-deployment-source.mjs';
  const initialGuard = deploy.indexOf(guard);
  assert.ok(initialGuard >= 0);
  assert.ok(initialGuard < deploy.indexOf('uses: azure/login@'));
  assert.equal(deploy.split(guard).length - 1, 2);
  assert.match(
    deploy,
    /- name: Verify current main before Azure changes\s+id: deployment_source\s+env:\s+EXPECTED_COMMIT: \$\{\{ needs\.merge_validate\.outputs\.fork_commit \}\}\s+GITHUB_TOKEN: \$\{\{ github\.token \}\}\s+run: node scripts\/verify-deployment-source\.mjs/,
  );
  assert.match(
    deploy,
    /- name: Deploy the new Container Apps revision[\s\S]*?env:[\s\S]*?EXPECTED_COMMIT: \$\{\{ needs\.merge_validate\.outputs\.fork_commit \}\}\s+GITHUB_TOKEN: \$\{\{ github\.token \}\}/,
  );
  assert.match(
    deploy,
    /node scripts\/verify-deployment-source\.mjs\s+touch \/tmp\/azurediagarm-deployment-started\s+az containerapp update "\$\{update_args\[@\]\}"/,
  );
  assert.match(
    deploy,
    /- name: Send failure email\s+if: \$\{\{ failure\(\) && steps\.deployment_source\.outcome == 'success' \}\}/,
  );
});
