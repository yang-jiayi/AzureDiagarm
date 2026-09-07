import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync, existsSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
  verifyDeploymentSource,
  type DeploymentSource,
} from '../scripts/verify-deployment-source.mjs';
import {
  readApprovedBaseline,
  recordSelectedBaseline,
  upstreamRepository,
} from '../scripts/upstream-baseline.mjs';

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
    /node scripts\/verify-astra-deployment\.mjs\s+node scripts\/verify-deployment-source\.mjs\s+# Freeze[^\n]+\s+printf '%s\\0' "\$AZURE_OPENAI_ENDPOINT" "\$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA" "\$AZURE_OPENAI_RESOURCE_ID" "\$ALLOW_BYO_AI_ENDPOINTS" \\\s+> azurediagarm-verified-astra\.env\s+touch azurediagarm-deployment-started\s+az containerapp update "\$\{update_args\[@\]\}"/,
  );
  assert.match(
    deploy,
    /- name: Send failure email\s+if: \$\{\{ failure\(\) && steps\.deployment_source\.outcome == 'success' \}\}/,
  );
});

const workflow = readFileSync(
  new URL('../.github/workflows/azurediagarm-sync-deploy.yml', import.meta.url), 'utf8',
);
const repoRoot = fileURLToPath(new URL('../', import.meta.url));
const bash = process.platform === 'win32'
  ? join(process.env.ProgramFiles || 'C:\\Program Files', 'Git', 'bin', 'bash.exe')
  : 'bash';
const approvedCommit = '6b38cd6145afbe3332f08194a86769c4e169d026';
const verifiedAstraTuple = [
  'https://verified.openai.azure.com/',
  'verified-production-alias',
  '/subscriptions/11111111-1111-1111-1111-111111111111/resourceGroups/approved/providers/Microsoft.CognitiveServices/accounts/verified',
];

function stepSource(name: string) {
  const section = workflow.split(`      - name: ${name}\n`)[1]?.split('\n      - name:')[0];
  assert.ok(section, `missing step ${name}`);
  const run = section.split('        run: |\n')[1];
  assert.ok(run, `missing shell body for ${name}`);
  return run.replace(/^          /gm, '');
}

function shellFixture(run: (directory: string) => void) {
  const directory = join(repoRoot, `.test-release-${randomUUID()}`);
  mkdirSync(directory);
  try { run(directory); } finally { rmSync(directory, { recursive: true, force: true }); }
}

function localizeWorkflowStep(name: string) {
  return stepSource(name)
    .replaceAll('/tmp/azurediagarm-', '$FIXTURE/azurediagarm-')
    .replace('headers="$(mktemp)"', 'headers="$FIXTURE/headers"')
    .replace(/\$\{\{ steps\.rollback_baseline\.outputs\.revision \}\}/g, 'previous')
    .replace(/\$\{\{ vars\.AZURE_CONTAINER_APP \}\}/g, 'app')
    .replace(/\$\{\{ vars\.AZURE_APP_URL \}\}/g, 'https://private.example')
    .replace(/\$\{\{[^}]+\}\}/g, 'test');
}

function writeRollbackTuple(directory: string, byoFlag = 'true') {
  writeFileSync(join(directory, 'azurediagarm-verified-astra.env'), `${[...verifiedAstraTuple, byoFlag].join('\0')}\0`);
}

function writeCleanupHelper(directory: string) {
  mkdirSync(join(directory, 'scripts'));
  writeFileSync(join(directory, 'scripts', 'retired-ai-environment.mjs'),
    readFileSync(new URL('../scripts/retired-ai-environment.mjs', import.meta.url)));
}

for (const scenario of ['invalid-byo-policy', 'model-rejected', 'source-rejected', 'capture-failed', 'approved-byo-true', 'approved-byo-false']) {
  test(`the approved BYO policy, final Astra and current-main guards precede revision mutation: ${scenario}`, () => {
    const approved = scenario.startsWith('approved-');
    const byoFlag = scenario === 'invalid-byo-policy' ? 'TRUE' : scenario === 'approved-byo-false' ? 'false' : 'true';
    const boundary = stepSource('Deploy the new Container Apps revision').match(
      /^\[\[ "\$ALLOW_BYO_AI_ENDPOINTS"[\s\S]*?^az containerapp update "\$\{update_args\[@\]\}"$/m,
    )?.[0];
    assert.ok(boundary);
    shellFixture(directory => {
      const script = `
set -Eeuo pipefail
export AZURE_OPENAI_ENDPOINT='${verifiedAstraTuple[0]}'
export AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA='${verifiedAstraTuple[1]}'
export AZURE_OPENAI_RESOURCE_ID='${verifiedAstraTuple[2]}'
export ALLOW_BYO_AI_ENDPOINTS='${byoFlag}'
${scenario === 'capture-failed' ? 'printf() { return 74; }' : ''}
update_args=(--image approved-test-image)
node() {
  case "$1" in
    scripts/verify-astra-deployment.mjs)
      echo model >> calls
      ${scenario === 'model-rejected' ? 'return 1' : 'return 0'};;
    scripts/verify-deployment-source.mjs)
      echo source >> calls
      ${scenario === 'source-rejected' ? 'return 1' : 'return 0'};;
    *) return 90;;
  esac
}
az() {
  [[ "$*" == "containerapp update --image approved-test-image" ]] || return 91
  [[ -f azurediagarm-deployment-started ]] || return 92
  echo update >> calls
}
${boundary.replaceAll('/tmp/azurediagarm-', '$PWD/azurediagarm-')}
`;
      const result = spawnSync(bash, ['--noprofile', '--norc', '-s'], {
        cwd: directory, input: script, encoding: 'utf8', timeout: 30_000,
      });
      assert.equal(result.error, undefined);
      assert.equal(result.status, approved ? 0 : scenario === 'capture-failed' ? 74 : 1, result.stderr);
      const calls = existsSync(join(directory, 'calls')) ? readFileSync(join(directory, 'calls'), 'utf8').trim().split('\n') : [];
      assert.deepEqual(calls, scenario === 'invalid-byo-policy' ? [] : scenario === 'model-rejected' ? ['model']
        : approved ? ['model', 'source', 'update'] : ['model', 'source']);
      assert.equal(existsSync(join(directory, 'azurediagarm-deployment-started')), approved);
      if (approved) {
        assert.deepEqual(readFileSync(join(directory, 'azurediagarm-verified-astra.env'), 'utf8').split('\0'),
          [...verifiedAstraTuple, byoFlag, '']);
      } else {
        const rollback = spawnSync(bash, ['--noprofile', '--norc', '-s'], {
          cwd: directory, encoding: 'utf8', timeout: 30_000,
          input: `export FIXTURE="$PWD"\naz() { echo unexpected-cloud-call; return 91; }\nnode() { return 92; }\n`
            + localizeWorkflowStep('Roll back a failed production revision'),
        });
        assert.equal(rollback.status, 0, rollback.stderr);
        assert.match(rollback.stdout, /Deployment did not start/);
        assert.doesNotMatch(rollback.stdout, /unexpected-cloud-call/);
      }
    });
  });
}

test('approved provenance reads the committed record and verifies local commit ancestry without fetching', () => {
  const calls: string[] = [];
  const actual = readApprovedBaseline((...args) => {
    calls.push(args.join(' '));
    if (args[0] === 'show') return JSON.stringify({ repository: upstreamRepository, commit: approvedCommit });
    if (args[0] === 'cat-file') return 'commit';
    if (args[0] === 'merge-base') return '';
    throw new Error('Unexpected Git operation');
  });
  assert.equal(actual, approvedCommit);
  assert.deepEqual(calls, [
    'show HEAD:.github/upstream-baseline.json',
    `cat-file -t ${approvedCommit}`,
    `merge-base --is-ancestor ${approvedCommit} HEAD`,
  ]);
});

for (const record of [
  undefined, 'not JSON', '{}',
  JSON.stringify({ repository: upstreamRepository, commit: '1234' }),
  JSON.stringify({ repository: 'other/repository', commit: approvedCommit }),
]) {
  test(`missing or malformed approved provenance fails closed: ${record}`, () => {
    assert.throws(() => readApprovedBaseline(() => {
      if (record === undefined) throw new Error('missing committed file');
      return record;
    }), /baseline/);
  });
}

test('an approved SHA outside checkout ancestry is rejected', () => {
  assert.throws(() => readApprovedBaseline((...args) => {
    if (args[0] === 'show') return JSON.stringify({ repository: upstreamRepository, commit: approvedCommit });
    if (args[0] === 'cat-file') return 'commit';
    throw new Error('not an ancestor');
  }), /checkout ancestry/);
});

test('recording provenance requires the selected remote tip and successful merge selection', () => {
  for (const failure of ['', 'remote', 'tip', 'merge']) {
    const writes: string[] = [];
    const run = () => recordSelectedBaseline(newerCommit, (...args) => {
      if (args[0] === 'remote') return failure === 'remote' ? 'other' : upstreamRepository;
      if (args.includes('MERGE_HEAD')) return failure === 'merge' ? currentCommit : newerCommit;
      return failure === 'tip' ? currentCommit : newerCommit;
    }, (_path, content) => { writes.push(content); });
    if (failure) {
      assert.throws(run);
      assert.deepEqual(writes, []);
    } else {
      run();
      assert.deepEqual(JSON.parse(writes[0]), { repository: upstreamRepository, commit: newerCommit });
    }
  }
});

test('already-integrated upstream selection must be an ancestor before recording', () => {
  for (const integrated of [true, false]) {
    let written = false;
    const run = () => recordSelectedBaseline(newerCommit, (...args) => {
      if (args[0] === 'remote') return upstreamRepository;
      if (args.includes('MERGE_HEAD')) throw Object.assign(new Error('no pending merge'), { status: 1 });
      if (args[0] === 'merge-base') {
        if (!integrated) throw new Error('not integrated');
        return '';
      }
      return newerCommit;
    }, () => { written = true; });
    if (integrated) run(); else assert.throws(run);
    assert.equal(written, integrated);
  }
});

test('unexpected merge-state lookup failures cannot record provenance', () => {
  for (const status of [128, undefined]) {
    const failure = Object.assign(new Error('merge state could not be read'), { status });
    let written = false;
    assert.throws(() => recordSelectedBaseline(newerCommit, (...args) => {
      if (args[0] === 'remote') return upstreamRepository;
      if (args.includes('MERGE_HEAD')) throw failure;
      assert.notEqual(args[0], 'merge-base');
      return newerCommit;
    }, () => { written = true; }), error => error === failure);
    assert.equal(written, false);
  }
});

test('seed provenance matches the reviewed upstream merge parent in existing Git history', () => {
  const record = JSON.parse(readFileSync(new URL('../.github/upstream-baseline.json', import.meta.url), 'utf8'));
  assert.equal(record.repository, upstreamRepository);
  // Future synchronization advances the record; this historical check documents
  // the bootstrap source without freezing subsequent approved upstream updates.
  if (record.commit === approvedCommit) {
    const result = spawnSync('git', ['rev-parse', '09b2bfeeffa36849d310acb2e8a72901102418cd^2'], {
      cwd: repoRoot, encoding: 'utf8',
    });
    // Shallow CI checkouts need not contain the bootstrap merge.
    if (result.status === 0) assert.equal(result.stdout.trim(), approvedCommit);
  }
});

for (const scenario of [
  'push', 'sync-fetch-failure', 'sync-diff-failure', 'sync-protected',
  'sync-history-failure', 'sync-merge-failure', 'invalid-baseline', 'sync-success',
]) {
  test(`source selection stays offline for releases and fail-closed for sync: ${scenario}`, () => {
    shellFixture((directory) => {
      const prefix = `
export RUNNER_TEMP="$PWD" GITHUB_OUTPUT="$PWD/outputs"
export GITHUB_EVENT_NAME=${scenario === 'push' || scenario === 'invalid-baseline' ? 'push' : 'workflow_dispatch'}
export UPSTREAM_REPOSITORY='${upstreamRepository}'
node() {
  if [[ "$*" == *"--record"* ]]; then
    [[ -f merged ]] || return 90
    printf '%s' "$3" > recorded
  else
    ${scenario === 'invalid-baseline' ? 'return 1' : `echo '${approvedCommit}'`}
  fi
}
git() {
  if [[ "$1" == fetch ]]; then
    echo fetch >> calls
    ${scenario === 'push' || scenario === 'sync-fetch-failure' ? 'return 128' : 'return 0'}
  elif [[ "$1 $2" == "rev-parse HEAD" ]]; then
    if [[ -f committed ]]; then echo '${newerCommit}'; else echo '${currentCommit}'; fi
  elif [[ "$*" == *"upstream/main^{commit}"* ]]; then echo '${newerCommit}'
  elif [[ "$1 $2" == "merge-base --is-ancestor" ]]; then
    ${scenario === 'sync-history-failure' ? 'return 1' : 'return 0'}
  elif [[ "$1 $2" == "diff --name-only" ]]; then
    ${scenario === 'sync-diff-failure' ? 'return 99' : scenario === 'sync-protected' ? 'echo scripts/changed.sh' : 'return 0'}
  elif [[ "$1 $2" == "merge --no-ff" ]]; then
    ${scenario === 'sync-merge-failure' ? 'return 1' : ': > merged'}
  elif [[ "$1" == show ]]; then echo 100
  elif [[ "$1" == commit ]]; then : > committed
  elif [[ "$1" == log ]]; then echo 'approved upstream change'
  fi
}
`;
      const result = spawnSync(bash, ['--noprofile', '--norc', '-s'], {
        cwd: directory, encoding: 'utf8', timeout: 30_000,
        input: prefix + stepSource('Merge approved upstream source'),
      });
      assert.equal(result.error, undefined);
      if (scenario === 'push') {
        assert.equal(result.status, 0, result.stderr);
        assert.equal(existsSync(join(directory, 'calls')), false);
        assert.match(readFileSync(join(directory, 'outputs'), 'utf8'), new RegExp(`upstream_commit=${approvedCommit}`));
        assert.equal(existsSync(join(directory, 'recorded')), false);
      } else if (scenario === 'sync-success') {
        assert.equal(result.status, 0, result.stderr);
        assert.equal(readFileSync(join(directory, 'recorded'), 'utf8'), newerCommit);
        assert.ok(existsSync(join(directory, 'committed')));
      } else {
        assert.notEqual(result.status, 0);
        assert.equal(existsSync(join(directory, 'recorded')), false);
        assert.equal(existsSync(join(directory, 'committed')), false);
      }
    });
  });
}

test('bounded stalled verification fails and enters the unchanged rollback path', () => {
  const section = workflow.split('      - name: Verify the production endpoint\n')[1]
    .split('\n      - name: Roll back')[0];
  assert.match(section, /timeout-minutes: 20/);
  assert.match(workflow, /- name: Roll back a failed production revision\s+if: \$\{\{ failure\(\) && steps\.deployment\.outputs\.should_deploy == 'true' \}\}/);
  shellFixture((directory) => {
    writeFileSync(join(directory, 'verify.sh'), localizeWorkflowStep('Verify the production endpoint'));
    writeFileSync(join(directory, 'rollback.sh'), localizeWorkflowStep('Roll back a failed production revision'));
    writeRollbackTuple(directory);
    writeCleanupHelper(directory);
    const script = `
export FIXTURE="$PWD" GITHUB_RUN_ID=1 GITHUB_RUN_ATTEMPT=1 ACCESS_CONTROL_ENABLED=true
: > "$FIXTURE/azurediagarm-deployment-started"
seq() { echo 1; }
sleep() { :; }
az() {
  if [[ "$*" == *"auth show"* ]]; then echo true
  elif [[ "$*" == *"revision copy"* ]]; then echo restored > "$FIXTURE/rollback-called"
  elif [[ "$*" == *"latestReadyRevisionName"* ]]; then echo app--rb1-1
  elif [[ "$*" == *"revision show"* ]]; then echo 'Provisioned|Healthy|Running'
  elif [[ "$*" == *"ingress.fqdn"* ]]; then echo origin.example
  else echo test
  fi
}
curl() {
  local output= headers= url= connect= maximum=
  while (( $# )); do
    case "$1" in
      --output) output="$2"; shift;;
      --dump-header) headers="$2"; shift;;
      --connect-timeout) connect="$2"; shift;;
      --max-time) maximum="$2"; shift;;
      https:*) url="$1";;
    esac
    shift
  done
  if [[ "$url" == "https://private.example/healthz" ]]; then
    echo ok > "$output"
    printf 'x-robots-tag: noindex\\nreferrer-policy: same-origin\\n' > "$headers"
  elif [[ "$url" == "https://private.example/" ]]; then
    echo 'location: /.auth/login/aad' > "$headers"; printf 302
  elif [[ "$url" == *origin.example* ]]; then printf 403
  else
    [[ "$connect" == 15 && "$maximum" == 45 ]] || { echo unbounded >> "$FIXTURE/bad-timeout"; return 99; }
    echo "$url" >> "$FIXTURE/bounded-stalls"
    printf 000
    return 28
  fi
}
export -f seq sleep az curl
if bash verify.sh; then exit 91; else bash rollback.sh; fi
`;
    const result = spawnSync(bash, ['--noprofile', '--norc', '-s'], {
      cwd: directory, input: script, encoding: 'utf8', timeout: 30_000,
    });
    assert.equal(result.error, undefined);
    assert.equal(result.status, 0, result.stderr);
    assert.equal(existsSync(join(directory, 'bad-timeout')), false);
    const stalled = readFileSync(join(directory, 'bounded-stalls'), 'utf8');
    assert.match(stalled, /\/api\/openai/);
    assert.match(stalled, /\/mcp/);
    assert.equal(readFileSync(join(directory, 'rollback-called'), 'utf8').trim(), 'restored');
  });
});

for (const scenario of ['old-image-byo-enabled', 'old-image-byo-disabled', 'missing-tuple', 'invalid-tuple', 'invalid-byo-policy', 'cleanup-failed']) {
  test(`rollback preserves approved BYO policy, sole managed Astra and unrelated baseline settings: ${scenario}`, () => {
    const approvedByoFlag = scenario === 'old-image-byo-disabled' ? 'false' : 'true';
    const rollback = localizeWorkflowStep('Roll back a failed production revision');
    assert.doesNotMatch(rollback, /verify-astra-deployment|az rest|vars\.(?:AZURE_OPENAI|ALLOW_BYO)|--replace-env-vars|ALLOW_BYO_AI_ENDPOINTS=false/);
    shellFixture(directory => {
      writeFileSync(join(directory, 'azurediagarm-deployment-started'), '');
      writeCleanupHelper(directory);
      if (scenario !== 'missing-tuple') writeRollbackTuple(directory, approvedByoFlag);
      if (scenario === 'invalid-tuple') writeFileSync(join(directory, 'azurediagarm-verified-astra.env'), 'incomplete\0');
      if (scenario === 'invalid-byo-policy') writeRollbackTuple(directory, 'TRUE');
      const script = `
export FIXTURE="$PWD" GITHUB_RUN_ID=1 GITHUB_RUN_ATTEMPT=1
export AZURE_OPENAI_ENDPOINT=https://changed.openai.azure.com/
export AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=changed-alias AZURE_OPENAI_RESOURCE_ID=changed-resource
export ALLOW_BYO_AI_ENDPOINTS=${approvedByoFlag === 'true' ? 'false' : 'true'}
node() {
  [[ "$*" == "scripts/retired-ai-environment.mjs --rollback" ]] || return 98
  ${scenario === 'cleanup-failed' ? 'return 1' : 'command node "$@"'}
}
seq() { echo 1; }
sleep() { :; }
az() {
  if [[ "$1 $2 $3" == "containerapp revision copy" ]]; then
    printf '%s\\0' "$@" > copy-args
  elif [[ "$*" == *"latestReadyRevisionName"* ]]; then echo app--rb1-1
  elif [[ "$1 $2 $3" == "containerapp revision show" ]]; then echo 'Provisioned|Healthy|Running'
  elif [[ "$1 $2 $3" == "afd endpoint purge" ]]; then :
  else echo "unexpected Azure command" >&2; return 99
  fi
}
${rollback}
`;
      const result = spawnSync(bash, ['--noprofile', '--norc', '-s'], {
        cwd: directory, input: script, encoding: 'utf8', timeout: 30_000,
      });
      assert.equal(result.error, undefined);
      if (!scenario.startsWith('old-image-')) {
        assert.notEqual(result.status, 0);
        assert.equal(existsSync(join(directory, 'copy-args')), false);
        return;
      }
      assert.equal(result.status, 0, result.stderr);
      const args = readFileSync(join(directory, 'copy-args'), 'utf8').split('\0').slice(0, -1);
      const values = (name: string) => {
        const start = args.indexOf(name);
        assert.ok(start >= 0, `missing revision-copy flag ${name}`);
        const end = args.findIndex((value, index) => index > start && value.startsWith('--'));
        return args.slice(start + 1, end === -1 ? undefined : end);
      };
      assert.deepEqual(values('--from-revision'), ['previous']);
      assert.equal(args.includes('--image'), false);
      const set = values('--set-env-vars');
      const remove = values('--remove-env-vars');
      assert.deepEqual(set, [
        `AZURE_OPENAI_ENDPOINT=${verifiedAstraTuple[0]}`,
        `AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=${verifiedAstraTuple[1]}`,
        `AZURE_OPENAI_ALLOWED_DEPLOYMENTS=${verifiedAstraTuple[1]}`,
        `ALLOW_BYO_AI_ENDPOINTS=${approvedByoFlag}`,
      ]);
      assert.ok(!remove.includes('ALLOW_BYO_AI_ENDPOINTS'));
      assert.ok(remove.every(name => /^(?:VITE_)?AZURE_(?:OPENAI|FOUNDRY)_/.test(name)));
      const baseline = [
        { name: 'AZURE_OPENAI_ALLOWED_DEPLOYMENTS', value: 'gpt-5.6-sol,gpt-5.6-terra,gpt-6-astra' },
        { name: 'AZURE_OPENAI_ENDPOINT', value: 'https://old.openai.azure.com/' },
        { name: 'AZURE_OPENAI_DEPLOYMENT_GPT56SOL', value: 'gpt-5.6-sol' },
        { name: 'AZURE_OPENAI_API_KEY', secretRef: 'retired-openai-key' },
        { name: 'AZURE_FOUNDRY_ENDPOINT', value: 'https://old.services.ai.azure.com/' },
        { name: 'AZURE_FOUNDRY_API_KEY', secretRef: 'retired-foundry-key' },
        { name: 'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS', value: 'claude-opus-5' },
        { name: 'ALLOW_BYO_AI_ENDPOINTS', value: approvedByoFlag === 'true' ? 'false' : 'true' },
        { name: 'ACCESS_CONTROL_ENABLED', value: 'true' },
        { name: 'ACCESS_ADMIN_EMAIL', secretRef: 'unchanged-admin-reference' },
        { name: 'APP_PRIVATE_SETTING', secretRef: 'unchanged-private-reference' },
        { name: 'AZURE_CLIENT_ID', value: 'unchanged-managed-identity' },
        { name: 'AZURE_SPEECH_REGION', value: 'unchanged-speech-region' },
      ];
      const inherited = new Map<string, { name: string; value?: string; secretRef?: string }>(
        baseline.map(setting => [setting.name, { ...setting }]),
      );
      for (const name of remove) inherited.delete(name);
      for (const setting of set) {
        const separator = setting.indexOf('=');
        const name = setting.slice(0, separator);
        inherited.set(name, { name, value: setting.slice(separator + 1) });
      }
      for (const original of baseline.filter(setting => !/^(?:AZURE_OPENAI_|AZURE_FOUNDRY_|ALLOW_BYO_)/.test(setting.name))) {
        assert.deepEqual(inherited.get(original.name), original, 'non-AI baseline env/secretRefs must be preserved');
      }
      assert.equal(inherited.get('AZURE_OPENAI_ALLOWED_DEPLOYMENTS')?.value, verifiedAstraTuple[1]);
      assert.equal(inherited.get('ALLOW_BYO_AI_ENDPOINTS')?.value, approvedByoFlag);
      for (const name of ['AZURE_FOUNDRY_ENDPOINT', 'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS',
        'AZURE_FOUNDRY_API_KEY', 'AZURE_OPENAI_API_KEY', 'AZURE_OPENAI_DEPLOYMENT_GPT56SOL']) {
        assert.equal(inherited.has(name), false, `${name} must not remain executable on the old image`);
      }
    });
  });
}
