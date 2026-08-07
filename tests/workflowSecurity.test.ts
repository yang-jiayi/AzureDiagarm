// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
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
const feedbackServer = readFileSync(
  new URL('../server/token-server.js', import.meta.url),
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
