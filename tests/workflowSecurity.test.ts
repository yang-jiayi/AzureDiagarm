// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import test from 'node:test';

const workflow = readFileSync(
  new URL('../.github/workflows/azurediagarm-sync-deploy.yml', import.meta.url),
  'utf8',
);
const emailHelper = readFileSync(
  new URL('../scripts/send-acs-email.sh', import.meta.url),
  'utf8',
);
const dockerfile = readFileSync(new URL('../Dockerfile', import.meta.url), 'utf8');
const dockerignore = readFileSync(new URL('../.dockerignore', import.meta.url), 'utf8');
const feedbackServer = readFileSync(
  new URL('../server/token-server.js', import.meta.url),
  'utf8',
);
const resourcesBicep = readFileSync(
  new URL('../infra/resources.bicep', import.meta.url),
  'utf8',
);
const deployScript = readFileSync(
  new URL('../scripts/deploy_aca.sh', import.meta.url),
  'utf8',
);

test('deployment notifications use OIDC instead of a Communication Services secret', () => {
  assert.doesNotMatch(workflow, /AZURE_COMMUNICATION_CONNECTION_STRING/);
  assert.doesNotMatch(workflow, /az communication email send/);
  assert.equal(
    Array.from(workflow.matchAll(/bash scripts\/send-acs-email\.sh email-body\.txt/g)).length,
    3,
  );
  assert.equal(
    Array.from(workflow.matchAll(/AZURE_COMMUNICATION_ENDPOINT: \$\{\{ vars\.FEEDBACK_EMAIL_ENDPOINT \}\}/g))
      .length,
    3,
  );
  assert.equal(
    Array.from(
      workflow.matchAll(
        /uses: azure\/login@532459ea530d8321f2fb9bb10d1e0bcf23869a43/g,
      ),
    ).length,
    3,
  );
  assert.match(
    workflow,
    /- name: Refresh Azure sign-in for deployment notification[\s\S]*?if: \$\{\{ !cancelled\(\) && steps\.deployment\.outputs\.should_deploy == 'true' \}\}[\s\S]*?- name: Send detailed update email/,
  );
});

test('the notification helper acquires a scoped token and waits for delivery', () => {
  assert.match(
    emailHelper,
    /az account get-access-token[\s\S]*--resource "\$ACS_TOKEN_RESOURCE"/,
  );
  assert.match(emailHelper, /https:\/\/communication\.azure\.com/);
  assert.match(emailHelper, /\/emails:send\?api-version=\$ACS_EMAIL_API_VERSION/);
  assert.equal(
    Array.from(emailHelper.matchAll(/--header "Authorization: Bearer \$access_token"/g)).length,
    2,
  );
  assert.match(emailHelper, /Operation-Location/);
  assert.match(emailHelper, /Succeeded\)/);
});

test('validated upstream publication isolates the main bypass credential', () => {
  const validationIndex = workflow.indexOf(
    '- name: Validate merged source before publishing',
  );
  const stagingIndex = workflow.indexOf('stage_validated_merge:');
  const publisherIndex = workflow.indexOf('publish_validated_merge:');
  const deployKeyIndex = workflow.indexOf('secrets.MAIN_BRANCH_DEPLOY_KEY');

  assert.ok(validationIndex >= 0);
  assert.ok(stagingIndex > validationIndex);
  assert.ok(publisherIndex > stagingIndex);
  assert.ok(deployKeyIndex > publisherIndex);
  assert.equal(
    Array.from(workflow.matchAll(/secrets\.MAIN_BRANCH_DEPLOY_KEY/g)).length,
    1,
  );
  assert.match(
    workflow,
    /merge_validate:[\s\S]*?permissions:\s+contents: read[\s\S]*?artifact_name: \$\{\{ steps\.artifact\.outputs\.name \}\}[\s\S]*?details_file="\$RUNNER_TEMP\/update-details\.txt"[\s\S]*?git merge --no-ff --no-commit upstream\/main[\s\S]*?GIT_AUTHOR_DATE="@\$merge_epoch \+0000"[\s\S]*?\[\[ "\$actual_commit" == "\$expected_commit" \]\][\s\S]*?status --porcelain=v1 --untracked-files=all[\s\S]*?git bundle create[\s\S]*?actions\/upload-artifact@ea165f8d65b6e75b540449e92b4886f43607fa02[\s\S]*?name: \$\{\{ steps\.artifact\.outputs\.name \}\}/,
  );
  assert.match(
    workflow,
    /stage_validated_merge:[\s\S]*?permissions:\s+contents: write[\s\S]*?actions\/download-artifact@d3f86a106a0bac45b974a628896c90dbdf5c8093[\s\S]*?name: \$\{\{ needs\.merge_validate\.outputs\.artifact_name \}\}[\s\S]*?bundle_commit[\s\S]*?\[\[ "\$bundle_commit" == "\$EXPECTED_COMMIT" \]\][\s\S]*?\[\[ "\$current_main" == "\$BASE_COMMIT" \|\| "\$current_main" == "\$EXPECTED_COMMIT" \]\][\s\S]*?staged_branch="automation\/upstream-sync"[\s\S]*?--force-with-lease="\$staged_ref:\$existing_commit"[\s\S]*?"\$EXPECTED_COMMIT:\$staged_ref"/,
  );
  assert.match(
    workflow,
    /publish_validated_merge:[\s\S]*?ref: \$\{\{ needs\.merge_validate\.outputs\.fork_commit \}\}[\s\S]*?ssh-key: \$\{\{ secrets\.MAIN_BRANCH_DEPLOY_KEY \}\}[\s\S]*?if \[\[ "\$current_main" == "\$EXPECTED_COMMIT" \]\][\s\S]*?\[\[ "\$current_main" == "\$BASE_COMMIT" \]\][\s\S]*?--atomic[\s\S]*?"\$EXPECTED_COMMIT:refs\/heads\/main"/,
  );
  assert.match(
    workflow,
    /deploy:[\s\S]*?if: needs\.merge_validate\.outputs\.source_changed != 'true'/,
  );
  assert.doesNotMatch(workflow, /git push origin HEAD:main/);
  assert.doesNotMatch(
    workflow,
    /"https:\/\/x-access-token:\$\{GITHUB_TOKEN\}@github\.com\/\$\{GITHUB_REPOSITORY\}\.git"\s+HEAD:main/,
  );
});

test('follow-up contact stays disabled unless client and server opt in together', () => {
  assert.match(dockerfile, /ARG VITE_FEEDBACK_CONTACT_ENABLED=false/);
  assert.match(
    workflow,
    /VITE_FEEDBACK_CONTACT_ENABLED=\$\{\{ vars\.FEEDBACK_CONTACT_ENABLED \|\| 'false' \}\}/,
  );
  assert.match(
    workflow,
    /"FEEDBACK_CONTACT_ENABLED=\$\{\{ vars\.FEEDBACK_CONTACT_ENABLED \|\| 'false' \}\}"/,
  );
  assert.match(
    feedbackServer,
    /process\.env\.FEEDBACK_CONTACT_ENABLED === 'true'/,
  );
  assert.match(feedbackServer, /createArchivedFeedbackContact\(item\.contact\)/);
});

test('BYO AI remains server-gated and is wired into every production deployment path', () => {
  assert.match(feedbackServer, /process\.env\.ALLOW_BYO_AI_ENDPOINTS === 'true'/);
  assert.match(resourcesBicep, /param allowByoAIEndpoints bool = false/);
  assert.match(
    resourcesBicep,
    /\{ name: 'ALLOW_BYO_AI_ENDPOINTS', value: string\(allowByoAIEndpoints\) \}/,
  );
  assert.match(
    workflow,
    /ALLOW_BYO_AI_ENDPOINTS: \$\{\{ vars\.ALLOW_BYO_AI_ENDPOINTS \|\| 'false' \}\}/,
  );
  assert.match(workflow, /AZURE_OPENAI_ENDPOINT ALLOW_BYO_AI_ENDPOINTS AZURE_OPENAI_DEPLOYMENT_GPT56SOL/);
  assert.match(workflow, /"ALLOW_BYO_AI_ENDPOINTS=\$ALLOW_BYO_AI_ENDPOINTS"/);
  assert.match(
    workflow,
    /"ALLOW_BYO_AI_ENDPOINTS=\$\{\{ vars\.ALLOW_BYO_AI_ENDPOINTS \|\| 'false' \}\}"/,
  );
  assert.match(
    deployScript,
    /"ALLOW_BYO_AI_ENDPOINTS=\$\{ALLOW_BYO_AI_ENDPOINTS:-false\}"/,
  );
});

test('the container build installs the image parser safeguard before npm ci', () => {
  const patchCopyIndex = dockerfile.indexOf(
    'COPY scripts/patch-image-size.mjs ./scripts/patch-image-size.mjs',
  );
  const installIndex = dockerfile.indexOf('RUN npm ci');

  assert.ok(patchCopyIndex >= 0);
  assert.ok(installIndex > patchCopyIndex);
  assert.match(dockerignore, /^!scripts\/patch-image-size\.mjs$/m);
});

// The .dockerignore is allowlist-style ("ignore **, then re-include"), so a new
// cross-package import from mcp-server into the repo-root scripts/ folder is
// invisible locally and in CI but fails the image build with
// ERR_MODULE_NOT_FOUND. Derive the requirement instead of hardcoding it.
test('root scripts imported by the MCP build are re-included in the Docker context', () => {
  const scriptsDir = new URL('../mcp-server/scripts/', import.meta.url);
  const imported = new Set<string>();
  for (const entry of readdirSync(scriptsDir)) {
    if (!entry.endsWith('.mjs')) continue;
    const source = readFileSync(new URL(entry, scriptsDir), 'utf8');
    for (const match of source.matchAll(/from\s+'\.\.\/\.\.\/(scripts\/[\w.-]+)'/g)) {
      imported.add(match[1]);
    }
  }

  assert.ok(imported.size > 0, 'expected at least one cross-package build import');
  for (const path of imported) {
    assert.match(
      dockerignore,
      new RegExp(`^!${path.replace(/[.]/g, '\\.')}$`, 'm'),
      `${path} is imported by an mcp-server build script but excluded from the Docker context`,
    );
  }
});
