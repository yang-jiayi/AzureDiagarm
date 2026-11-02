// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import test from 'node:test';
import { verifyDeploymentTarget } from '../scripts/verify-deployment-target.mjs';

const workflow = readFileSync(
  new URL('../.github/workflows/azurediagarm-sync-deploy.yml', import.meta.url),
  'utf8',
);
const ciWorkflow = readFileSync(new URL('../.github/workflows/ci.yml', import.meta.url), 'utf8');
const codeqlWorkflow = readFileSync(
  new URL('../.github/workflows/codeql.yml', import.meta.url),
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
    Array.from(workflow.matchAll(/uses: azure\/login@[0-9a-f]{40}/g)).length,
    3,
  );
  assert.match(
    workflow,
    /- name: Refresh Azure sign-in for deployment notification[\s\S]*?if: \$\{\{ !cancelled\(\) && steps\.deployment\.outputs\.should_deploy == 'true' \}\}[\s\S]*?- name: Send detailed update email/,
  );
});

test('production administrator configuration is masked and not interpolated into shell code', () => {
  assert.doesNotMatch(workflow, /vars\.ACCESS_ADMIN_EMAIL/);
  assert.equal(
    workflow.split('ACCESS_ADMIN_EMAIL: ${{ secrets.ACCESS_ADMIN_EMAIL }}').length - 1,
    4,
  );
  assert.match(
    workflow,
    /- name: Deploy the new Container Apps revision[\s\S]*?env:\s+ACCESS_ADMIN_EMAIL: \$\{\{ secrets\.ACCESS_ADMIN_EMAIL \}\}/,
  );
  assert.match(workflow, /"ACCESS_ADMIN_EMAIL=\$ACCESS_ADMIN_EMAIL"/);
  assert.doesNotMatch(workflow, /"ACCESS_ADMIN_EMAIL=\$\{\{/);
});

test('Speech STS token exchange uses the resource-scoped Speech Contributor role', () => {
  assert.match(feedbackServer, /SPEECH_STS_ENDPOINT\}\/sts\/v1\.0\/issueToken/);
  assert.match(
    resourcesBicep,
    /resource speechContributorRole 'Microsoft\.Authorization\/roleAssignments@2022-04-01' = if \(deploySpeech\)/,
  );
  assert.match(resourcesBicep, /scope: speech[\s\S]*?0e75ca1e-0464-4b4d-8b93-68208a576181/);
  assert.match(resourcesBicep, /name: '8813c46e-1a44-5c14-9459-c10475883e01'/);
  assert.doesNotMatch(resourcesBicep, /guid\(deploySpeech \? speech\.id : resourceToken, appIdentity\.id, 'speechuser'\)/);
  assert.doesNotMatch(resourcesBicep, /f2dc8367-1007-4938-bd23-fe263f013447/);
});

// Third-party actions run with repository credentials, so a moved tag is a
// supply-chain takeover. Three assertions above happened to spell out SHAs,
// which meant every Dependabot bump failed CI and left the *other* `uses:`
// lines unguarded. Enforce the property itself, once, across every workflow.
test('every workflow action is pinned to a full commit SHA', () => {
  const workflowsDir = new URL('../.github/workflows/', import.meta.url);
  const files = readdirSync(workflowsDir).filter(
    (entry) => entry.endsWith('.yml') || entry.endsWith('.yaml'),
  );

  assert.ok(files.length > 0, 'expected at least one workflow file');

  const references: string[] = [];
  for (const file of files) {
    const source = readFileSync(new URL(file, workflowsDir), 'utf8');
    for (const match of source.matchAll(/^\s*uses:\s*(\S+)/gm)) {
      const reference = match[1];
      // Local composite actions and reusable workflows in this repo are
      // covered by this repo's own review, not by a third-party tag.
      if (reference.startsWith('./')) continue;
      references.push(`${file}: ${reference}`);
      assert.match(
        reference,
        /^[\w.-]+\/[\w.-]+(\/[\w.-]+)*@[0-9a-f]{40}$/,
        `${file} uses "${reference}", which is not pinned to a full 40-character commit SHA`,
      );
    }
  }

  assert.ok(references.length > 0, 'expected at least one third-party action');
});

test('deployment uses and records the image digest produced by the pushed build', () => {
  assert.match(workflow, /deploy:[\s\S]*?permissions:\s+contents: read\s+id-token: write/);
  assert.match(
    workflow,
    /\n  deploy:\n[\s\S]*?environment: production\n\s+permissions:[\s\S]*?id-token: write/,
  );
  assert.match(
    workflow,
    /\n  notify_validation_failure:\n[\s\S]*?environment: production\n\s+permissions:[\s\S]*?id-token: write/,
  );
  assert.match(
    workflow,
    /- name: Build and push application image[\s\S]*?id: build_image[\s\S]*?push: true[\s\S]*?tags: \$\{\{ steps\.image\.outputs\.image \}\}/,
  );
  assert.match(
    workflow,
    /- name: Record immutable release image[\s\S]*?IMAGE_DIGEST: \$\{\{ steps\.build_image\.outputs\.digest \}\}[\s\S]*?\[\[ "\$IMAGE_DIGEST" =~ \^sha256:\[0-9a-f\]\{64\}\$ \]\][\s\S]*?immutable_ref=\$immutable_image/,
  );
  assert.match(
    workflow,
    /- name: Upload release image metadata[\s\S]*?actions\/upload-artifact@[0-9a-f]{40}[\s\S]*?name: release-image-\$\{\{ github\.run_id \}\}[\s\S]*?retention-days: 90/,
  );
  assert.match(workflow, /--image "\$\{\{ steps\.release_image\.outputs\.immutable_ref \}\}"/);
  assert.match(workflow, /\[\[ "\$actual_image" == "\$\{\{ steps\.release_image\.outputs\.immutable_ref \}\}" \]\]/);
  assert.match(workflow, /Image: \$\{\{ steps\.release_image\.outputs\.immutable_ref \}\}/);
});

test('Azure OIDC login follows an exact new-tenant target check', () => {
  const expected = {
    AZURE_TENANT_ID: '02310481-6752-4d53-8fc4-f597a8e047f4',
    AZURE_SUBSCRIPTION_ID: '0df5f427-7359-447e-9b60-fa67c5588b8f',
    AZURE_RESOURCE_GROUP: 'AzureDiagarm_rg',
    AZURE_CONTAINER_APP: 'azurediagarm-app',
    AZURE_CLIENT_ID: 'synthetic-client-id',
  };
  assert.doesNotThrow(() => verifyDeploymentTarget(expected));
  for (const [name, wrongValue] of [
    ['AZURE_TENANT_ID', 'old-tenant'],
    ['AZURE_SUBSCRIPTION_ID', 'old-subscription'],
    ['AZURE_RESOURCE_GROUP', 'old-resource-group'],
    ['AZURE_CONTAINER_APP', 'old-container-app'],
  ]) {
    assert.throws(
      () => verifyDeploymentTarget({ ...expected, [name]: wrongValue }),
      new RegExp(name),
    );
  }
  assert.throws(
    () => verifyDeploymentTarget({ ...expected, AZURE_CLIENT_ID: '' }),
    /AZURE_CLIENT_ID/,
  );
  const notificationJob = workflow.split('  notify_validation_failure:\n')[1]?.split('  deploy:\n')[0];
  const deploymentJob = workflow.split('  deploy:\n')[1];
  assert.ok(notificationJob);
  assert.ok(deploymentJob);
  const guard = 'run: node scripts/verify-deployment-target.mjs';
  const login = 'uses: azure/login@';
  assert.ok(notificationJob!.indexOf(guard) >= 0);
  assert.ok(notificationJob!.indexOf(login) > notificationJob!.indexOf(guard));
  assert.ok(deploymentJob!.indexOf(guard) >= 0);
  assert.ok(deploymentJob!.indexOf(login) > deploymentJob!.indexOf(guard));
});

test('CI reports full root advisories while retaining the production high-severity gate', () => {
  assert.match(ciWorkflow, /actions\/dependency-review-action@[0-9a-f]{40}/);
  assert.match(ciWorkflow, /fail-on-severity: high[\s\S]*?fail-on-scopes: runtime/);
  assert.match(ciWorkflow, /npm run audit:production/);
  assert.match(ciWorkflow, /npm --prefix server audit --omit=dev --audit-level=high/);
  assert.match(ciWorkflow, /- name: Audit all root dependencies[\s\S]*?npm audit --audit-level=critical/);
});

test('CodeQL produces an artifact without requiring repository code-scanning permissions', () => {
  assert.match(codeqlWorkflow, /languages: javascript-typescript/);
  assert.match(codeqlWorkflow, /build-mode: none/);
  assert.match(codeqlWorkflow, /upload: never/);
  assert.match(codeqlWorkflow, /actions\/upload-artifact@[0-9a-f]{40}/);
  assert.match(codeqlWorkflow, /permissions:\s+contents: read/);
  assert.doesNotMatch(codeqlWorkflow, /security-events:\s*write/);
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
    /merge_validate:[\s\S]*?permissions:\s+contents: read[\s\S]*?artifact_name: \$\{\{ steps\.artifact\.outputs\.name \}\}[\s\S]*?details_file="\$RUNNER_TEMP\/update-details\.txt"[\s\S]*?git merge --no-ff --no-commit upstream\/main[\s\S]*?GIT_AUTHOR_DATE="@\$merge_epoch \+0000"[\s\S]*?\[\[ "\$actual_commit" == "\$expected_commit" \]\][\s\S]*?status --porcelain=v1 --untracked-files=all[\s\S]*?git bundle create[\s\S]*?actions\/upload-artifact@[0-9a-f]{40}[\s\S]*?name: \$\{\{ steps\.artifact\.outputs\.name \}\}/,
  );
  assert.match(
    workflow,
    /stage_validated_merge:[\s\S]*?permissions:\s+contents: write[\s\S]*?actions\/download-artifact@[0-9a-f]{40}[\s\S]*?name: \$\{\{ needs\.merge_validate\.outputs\.artifact_name \}\}[\s\S]*?bundle_commit[\s\S]*?\[\[ "\$bundle_commit" == "\$EXPECTED_COMMIT" \]\][\s\S]*?\[\[ "\$current_main" == "\$BASE_COMMIT" \|\| "\$current_main" == "\$EXPECTED_COMMIT" \]\][\s\S]*?staged_branch="automation\/upstream-sync"[\s\S]*?--force-with-lease="\$staged_ref:\$existing_commit"[\s\S]*?"\$EXPECTED_COMMIT:\$staged_ref"/,
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

test('custom AI is an explicit administrator opt-in alongside Astra; retired managed providers stay removed', () => {
  assert.match(feedbackServer, /allowByoAIEndpoints: deployment\.allowByoAIEndpoints/);
  assert.match(feedbackServer, /runtimeAstraConfiguration\(deployment\.astra, deployment\.allowByoAIEndpoints\)/);
  assert.match(resourcesBicep, /param allowByoAIEndpoints bool = false/);
  assert.match(resourcesBicep, /name: 'ALLOW_BYO_AI_ENDPOINTS', value: string\(allowByoAIEndpoints\)/);
  assert.doesNotMatch(resourcesBicep, /azureFoundryEndpoint/);
  assert.doesNotMatch(workflow, /vars\.AZURE_FOUNDRY/);
  assert.match(workflow, /ALLOW_BYO_AI_ENDPOINTS: \$\{\{ vars\.ALLOW_BYO_AI_ENDPOINTS \}\}/);
  assert.doesNotMatch(workflow, /ALLOW_BYO_AI_ENDPOINTS:\s*\$\{\{\s*vars\.ALLOW_BYO_AI_ENDPOINTS\s*\|\|/);
  assert.match(workflow, /node scripts\/verify-deployment-preflight\.mjs runtime-config/);
  assert.match(workflow, /ALLOW_BYO_AI_ENDPOINTS: \$\{\{ steps\.desired_config\.outputs\.allow_byo_ai_endpoints \}\}/);
  assert.match(workflow, /"ALLOW_BYO_AI_ENDPOINTS=\$ALLOW_BYO_AI_ENDPOINTS"/);
  assert.match(workflow, /echo "allow_byo_ai_endpoints=\$ALLOW_BYO_AI_ENDPOINTS" >> "\$GITHUB_OUTPUT"/);
  assert.match(workflow, /node scripts\/retired-ai-environment\.mjs/);
  assert.match(deployScript, /scripts\/retired-ai-environment\.mjs/);
  assert.match(workflow, /AZURE_OPENAI_ENDPOINT AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA AZURE_OPENAI_RESOURCE_ID/);
});

test('container model build arguments and environment contain only the approved Astra alias', () => {
  const allowed = ['VITE_AZURE_OPENAI_ENDPOINT', 'VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA'];
  for (const instruction of ['ARG', 'ENV']) {
    const declared = Array.from(dockerfile.matchAll(
      new RegExp(`^${instruction} (VITE_(?:AZURE_OPENAI|AZURE_FOUNDRY)_[A-Z0-9_]+)`, 'gm'),
    ), match => match[1]);
    assert.deepEqual(declared, allowed, `${instruction} must not reintroduce another model or credentials`);
  }
  assert.match(dockerfile, /COPY server\/astra-policy\.js \.\//);
});

test('the maintained cloud script checks actual Astra identity before build and immediately before update', () => {
  const modelChecks = Array.from(deployScript.matchAll(/^node .*verify-astra-deployment\.mjs"$/gm));
  assert.equal(modelChecks.length, 2);
  assert.ok(modelChecks[0].index! < deployScript.indexOf('\naz acr build '));
  assert.match(deployScript, /node "\$SOURCE_DIR\/scripts\/verify-astra-deployment\.mjs"\r?\naz containerapp update/);
  assert.match(deployScript, /export AZURE_OPENAI_ALLOWED_DEPLOYMENTS="\$\{AZURE_OPENAI_ALLOWED_DEPLOYMENTS:-\$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA\}"/);
  assert.match(deployScript, /retired_ai_env="\$\(node [^\n]+\)"\r?\n\[\[ -n "\$retired_ai_env" \]\]\r?\nmapfile -t REMOVE_ENV_VARS/);
  assert.doesNotMatch(deployScript, /AZURE_OPENAI_DEPLOYMENT_(?:GPT5|DEEPSEEK|GROK|KIMI|MISTRAL)|AZURE_FOUNDRY/);
  assert.match(deployScript, /export ALLOW_BYO_AI_ENDPOINTS="\$\{ALLOW_BYO_AI_ENDPOINTS:-false\}"/);
  assert.match(deployScript, /"ALLOW_BYO_AI_ENDPOINTS=\$ALLOW_BYO_AI_ENDPOINTS"/);
});

test('the retired bundled MCP experiment cannot provision a legacy inference application', () => {
  const source = readFileSync(new URL('../scripts/deploy-mcp-instance.sh', import.meta.url), 'utf8');
  assert.match(source, /standalone mcp-server\/Dockerfile/);
  assert.match(source, /^exit 1\s*$/m);
  assert.doesNotMatch(source, /\baz\s|docker\s|AZURE_OPENAI_DEPLOYMENT_|AZURE_FOUNDRY|ALLOW_BYO_AI_ENDPOINTS/);
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
  // Compare against exact allowlist lines rather than building a RegExp from a
  // path, which would need every metacharacter escaped to stay correct.
  const allowlist = new Set(
    dockerignore.split('\n').map((line) => line.trim()).filter(Boolean),
  );
  for (const path of imported) {
    assert.ok(
      allowlist.has(`!${path}`),
      `${path} is imported by an mcp-server build script but excluded from the Docker context`,
    );
  }
});
