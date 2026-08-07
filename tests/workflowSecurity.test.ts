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
