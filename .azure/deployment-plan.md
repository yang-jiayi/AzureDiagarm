# Azure Deployment Plan

> **Status:** Ready for Validation

Generated: 2026-09-05

## 1. Current Release

Publish the completed AzureDiagarm workspace, Office export, privacy, and runtime
safeguard improvements to `yang-jiayi/AzureDiagarm`, then update the existing
production application. The user explicitly requested GitHub publication and
production deployment. The user also explicitly requested switching the actual
application inference model, including its picker and persisted selections, from
GPT-5.6 to GPT-6 Astra. Reuse the existing application, subscription, region, and
OpenAI account; do not simply relabel an old deployment.

## 2. Scope and Existing Architecture

Modify the existing React/Vite and Node application. Reuse its Container App,
registry, Front Door, authentication, access-list store, Cosmos account, and
GitHub OIDC identity. Do not deploy the separate Product Analytics application
described in the historical plan below.

## 3. Release Recipe

`recipe.type: azcli`

Use Azure CLI through the existing GitHub Actions pipeline, not the retired azd
hooks. The user's explicit instruction to publish and deploy approves this
existing-application release; source integration and validation remain required.

Use the existing `.github/workflows/azurediagarm-sync-deploy.yml` push-to-main
release path. A push deploys the customization commit without merging unrelated
upstream updates. Do not use manual upstream synchronization for this release.

The initial local checkout was 78 commits behind the existing production branch.
Feature work is preserved in checkpoint `5de4fa2` on
`release/workspace-quality-20260905`; integrate `origin/main` at `c885477` before
release. Preserve its newer Office, cloud document, accessibility, authentication,
and deployment capabilities rather than replacing them with the older checkout.
Publish the identical integration tree as release snapshot `bc42a8d` on
`release/astra-workspace-20260905`, PR
[63](https://github.com/yang-jiayi/AzureDiagarm/pull/63). The original local
checkpoint and integration commits remain preserved.

The only new Azure resource is the genuine Astra model deployment, represented
by `infra/gpt6-astra.bicep` under an existing OpenAI account. Validate this actual
deployment template and preview it at the production resource-group scope.
Do not validate/deploy the legacy subscription bootstrap as a substitute for
the current application's release. Existing bootstrap templates receive static
compilation checks only. Container build validation runs on the established
Docker-capable CI runner, with all application gates required before rollout.

## 4. Production Context

The existing production application is `azurediagarm-app` in `AzureDiagarm_rg`,
West US, at `https://azurediagarm.mssql.biz`. Subscription and identity values
remain in the existing GitHub variables rather than being copied into source.
Read-only preflight confirmed the current revision is healthy, origin isolation
and platform authentication hold, and the runtime identity has registry pull and
account-scoped Table data access. Rollback image, digest, and revision are recorded
in the private session state.

The existing OpenAI account now contains genuine `gpt-6-astra`, version
`2026-09-03`, in West US. The usage-based GlobalStandard deployment reserves
50K TPM / 50 RPM of the verified available quota, uses `Microsoft.DefaultV2`,
and retains the existing default-version upgrade policy. Existing GPT-5.6
deployments remain intact for rollback. Wire
`VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=gpt-6-astra` through the frontend build
and runtime allowlist; migrate previous managed-model selections only after
the new deployment is configured. Preserve separate BYO settings.

The existing `ACCESS_CONTROL_ENABLED=false` configuration must become `true`
for the new public-mode safeguards. The configured application administrator
matches the operator and remains allowed. The application email allowlist is
currently empty: Entra group assignment alone does not authorize other users
through the application's allowlist. Treat this access-policy transition
explicitly rather than claiming all prior users retain access.
The unchanged administrator value is now stored in the like-named Actions
secret, and the workflow consumes that secret instead of a repository variable
to prevent public deployment logs from exposing the address.

The existing deployment identity is correctly federated to this repository's
immutable-ID main-branch subject. The initial legacy-subject comparison was
incorrect; the existing production workflow has already deployed successfully
using this credential. No federation or branch-protection changes are needed.

## 5. Preparation

Artifacts and the existing release path are prepared. The editor integration,
hook dependencies, canonical style contracts, theme/contrast repairs, and static
gates now pass. Focused browser checks cover recovery, modal keyboard safety,
cancelled AI requests, validation metadata, and atomic import into a new cloud
document. Cloud hydration now records its normalized baseline without rewriting
the source document. The initial Linux container/core and CodeQL gates passed.
Older cloud test arrangements now explicitly edit metadata before expecting a
write. Focus-mode recovery waits for a committed draft. All 85 functional browser
cases and the patched runtime/container checks now pass on Linux. All four
workflow visual changes have been reviewed against stable Linux artifacts.
Newly reported dependency advisories have been patched. Fresh complete CI with
the reviewed baselines remains pending; this status does not authorize deployment.

- [x] Identify repository, existing workflow, and application boundaries.
- [x] Preserve the unrelated historical deployment plan below.
- [x] Fetch the current protected production branch and preserve local work in a release checkpoint.
- [x] Restore dependencies from the merged lockfiles and pass the production dependency audit.
- [x] Resolve and validate the current-main integration without duplicate editor systems.
- [x] Resolve production subscription, region, revision, and endpoint.
- [x] Verify platform authentication, origin protection, runtime Table permissions, and data-retention impact.
- [x] Reconcile application access-control and deployment federation settings.
- [x] Verify Astra catalog availability and quota, and provision the actual model.
- [x] Verify Astra inference capabilities through the authorized runtime identity.
- [x] Complete and verify build/runtime model configuration and persisted-settings migration.
- [x] Confirm local source is current with the remote branch.
- [ ] Complete release validation and record evidence.
- [x] Core Azure validation: CLI/authentication, compile the actual model template, ARM validation, and what-if.
- [ ] Build the final container using a Docker-capable builder and inspect its context.
- [x] Review applicable Azure policy constraints and unchanged runtime role scopes; record existing account-level audit findings separately.
- [ ] Pass complete application, server, MCP, browser, and required GitHub checks.

## 6. Deployment Steps

Commit the requested source changes with the required coauthor trailer, publish
the release branch, and open a pull request against protected `main`. Complete
the required app, MCP, browser, and CodeQL checks and use the repository's existing
permitted pull-request merge path. Do not change branch rules or rewrite remote
history. Monitor the production workflow triggered by the resulting main commit.
Confirm the deployed commit, active revision, health, protected API behavior, and
expected frontend assets.

## 7. Validation Proof

| Check | Result | Timestamp |
| --- | --- | --- |
| Merged dependency restore and production audit | Pass; existing image parser safeguards verified | 2026-09-05 |
| Read-only production security and health | Healthy, public health 200, protected routes 401, direct origin 403 | 2026-09-05 12:54 UTC |
| Runtime data permissions | Existing account-scoped Table Data Contributor permits budget-table operations | 2026-09-05 12:54 UTC |
| Existing feedback retention impact | Fully paginated authorized metadata read found zero feedback rows | 2026-09-05 12:52:55 UTC |
| Astra model provisioning | Succeeded; actual `gpt-6-astra` v2026-09-03, GlobalStandard 50K TPM, default content filter | 2026-09-05 |
| Astra authenticated inference | Existing runtime identity returned model `gpt-6-astra`; Responses v1, JSON output, 32K application cap accepted | 2026-09-05 |
| Astra image and reasoning support | Fresh synthetic 64px red image identified as red; none/low/medium/high/xhigh/max accepted and echoed by actual model | 2026-09-05 |
| Production deployment trust | Existing immutable-ID repository/main OIDC subject matches; prior workflow run 32490277791 authenticated and deployed | 2026-09-05 |
| Administrator access | Runtime/GitHub configured administrator matches operator and remains allowed; other users need explicit application authorization | 2026-09-05 |
| Full Office export corpus | 167 scenarios, 333 checks, zero issues; unchanged golden thresholds | 2026-09-05 |
| Astra GitHub configuration | `AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=gpt-6-astra`, following existing repository variable naming | 2026-09-05 |
| Required public access flag | GitHub `ACCESS_CONTROL_ENABLED=true` prepared; the existing live revision is unchanged until release | 2026-09-05 |
| Core Azure CLI validation | Official `validate-deployment.ps1` helper passed CLI, authentication, Bicep compilation, `az deployment group validate`, and resource-group what-if for `infra/gpt6-astra.bicep` | 2026-09-05 |
| Structured infrastructure preview | No resource creation or deletion. Astra has one Modify entry for omitted service-reported `properties.currentCapacity`; all other resources are Ignore. The helper's textual Delete count describes property lines, not deleted resources. No template application was performed | 2026-09-05 |
| Container build context | Allowlist-style `.dockerignore` excludes credentials and unrelated artifacts; root, server, and MCP lockfiles are present. Complete Linux image build and API/MCP readiness smoke passed in required Vite job 101373981082, CI run 33991330053, for release snapshot `bc42a8d`, without Azure credentials or registry publication. Rebuild after dependency patches is required | 2026-09-05 20:56 UTC |
| Static model RBAC review | Astra template references an existing account and introduces no identities or roles. Existing `infra/openai-role.bicep` scopes Cognitive Services OpenAI User to that account, matching runtime inference operations. No role widening is required | 2026-09-05 |
| Applicable policy constraints | Inherited deny initiative and actual definitions reviewed. OpenAI policy denies `ProvisionedManaged`, whereas Astra uses `GlobalStandard`; classic-resource, VM, AKS, SQL, and HSM constraints do not match this change. Actual model ARM validation passed without a policy denial | 2026-09-05 |
| Final static application gates | `npm run lint`, `npm run typecheck:scripts`, and all 807 root unit tests pass on the published repaired source, including the bounded AI queue, cloud hydration/import fixes, and administrator log-privacy contract | 2026-09-05 |
| Remaining core scripts | Workflow contracts, ARM extraction, layout preservation, icon library, icon workspace, validation freshness, and service-name normalization all pass | 2026-09-05 |
| Genuine Astra production build | `npm run build` passes with the verified OpenAI endpoint and `VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=gpt-6-astra` | 2026-09-05 |
| Focused browser repairs | Theme/contrast, explicit draft recovery, service-inspector focus, AI cancellation/review, score-zero metadata, modal keyboard safety, and atomic AI import checks pass. The import preserves authoritative IDs/pricing and makes zero writes to the old source document | 2026-09-05 |
| AI comparison admission | Shared cancellable budget queue passes 111 focused units and 85 AI/UI checks, including 43 comparison cases; concurrency limits and cancellation assertions remain intact | 2026-09-05 |
| Cloud browser arrangements | Explicit author edits replace 18 incidental hydration-write setups. All 19 focused cases and three replacement-race runs pass. The complete Windows run reaches 84/85; the remaining unchanged keyboard/access-dialog case encountered delayed Vite loading and subsequently passed both keyboard/WCAG control runs. Fresh Linux CI remains required | 2026-09-05 |
| Focus-mode persistence | Linux traces reloaded before confirmed persistence. The test now waits for `Saved on this device`, retaining every recovery/Escape/focus assertion. Focus-mode and recent-work recovery pass all six repeated local runs | 2026-09-05 |
| Workflow visual review | All four baselines use exact reviewed Linux captures, with byte-identical attempts and unchanged dimensions. Differences reflect approved status-bar separation, shadows, and upper-border positioning. Dark/mobile/forced-colors comparisons are 0 differing pixels against both attempts; forced-color current-step text retains 11.31:1 contrast. The 1% threshold and failure semantics remain unchanged | 2026-09-05 |
| Initial required CI | Vite/core/container and all three CodeQL analyses passed in runs 33991330053/33991329211. Browser and MCP audit failures block merge until corrected | 2026-09-05 |
| Repaired required CI | Runs 33995733963/33995732452 pass Vite/core, patched MCP, complete image/startup/readiness, all 85 functional browser cases, and all CodeQL analyses. Only the now-reviewed three visual baselines prevented the full browser job from completing; fresh complete CI remains required | 2026-09-05 |
| Runtime dependency patches | Patched MCP `fast-uri` to 3.1.7 and `qs` to 6.16.0. API Express 4/body-parser restrict `qs` to the vulnerable 6.15 line, so a compatible `qs` override selects 6.16.0 without upgrading Express's major version. Both production audits now report zero vulnerabilities; patched MCP 65 and API 120 tests pass, with generated MCP assets unchanged | 2026-09-05 |
| Deployment log privacy | Created the `ACCESS_ADMIN_EMAIL` Actions secret from the exact existing administrator value through stdin, without a trailing newline or command-line disclosure. Workflow references use the secret and a quoted environment variable; the administrator and runtime value remain unchanged | 2026-09-05 |
| Final source review | Read-only review reports no significant findings in normalized cloud baselines, canonical serialization, and their save/copy/conflict interactions. Final lint and script type checks pass; nine deployment workflow security contracts pass | 2026-09-05 |
| Remote source currency | Refreshed `origin/main` remains `c885477c799a35a044a73ac05b230aeab7160f95`; no unresolved index entries or whitespace errors | 2026-09-05 |
| Refreshed live preflight | Existing healthy revision and rollback image remain unchanged. Astra is Succeeded at the verified model/version and capacity. Without following authentication redirects, public health returns 200, protected root/API/MCP return 401, and the direct origin returns 403. Browser-style requests redirect to Microsoft sign-in, not anonymous API JSON | 2026-09-05 |

Policy Insights reports no evaluated rows for the existing Container App; an
empty result is not proof of estate-wide compliance. The existing OpenAI account
has two pre-existing audit-only findings: Private Link is absent and network
access is unrestricted. Both policies target the account, not the child model
deployment. This release neither replaces nor updates the account's networking.
Do not silently change its shared connectivity or create policy exemptions to
clear those findings; they require a separately planned network migration.
No subscription-wide compliance claim is made.

Integrated-source static validation and the Astra build are complete; full
browser and container/required CI validation remain pending. Historical evidence
below does not authorize this release. The operator's direct AI data-plane call lacks the
required inference permission; model checks instead used the existing authorized
runtime identity, without granting new roles or changing network controls.

## 8. Rollback and Data Safety

Retain the previous healthy image and revision. Do not weaken authentication or
origin restrictions to make deployment succeed. Determine the effect of feedback
retention on existing records before enabling any irreversible deletion.
The preflight found no existing feedback records to delete. New feedback remains
subject to the explicitly documented retention policy.

---

# Historical Product Analytics Deployment Plan

> **Historical status:** Deployed

Generated: 2026-07-20

## 1. Project Overview

**Goal:** Build and prepare a private AADB Product Analytics web application that turns Application Insights telemetry and Cosmos DB feedback into actionable product-maintenance insights.

**Path:** Add Components

## 2. Requirements

| Attribute | Value |
| --- | --- |
| Classification | Production internal application |
| Scale | Small |
| Budget | Balanced |
| Subscription | Existing AADB subscription |
| Location | East US 2 |
| Availability | `minReplicas: 1`, `maxReplicas: 3` |

The application must use Microsoft Entra ID authentication, assignment-required access, managed identity, least-privilege Azure RBAC, HTTPS-only ingress, and server-side access to Azure Monitor Logs. Raw KQL, Azure credentials, prompts, architecture content, and access tokens must not be exposed to the browser.

## 3. Components Detected

| Component | Type | Technology | Path |
| --- | --- | --- | --- |
| Analytics dashboard | Frontend | React, TypeScript, Vite | `NEW-WEB-APP/src` |
| Analytics API | API | Node.js, Express, TypeScript | `NEW-WEB-APP/server` |
| Query registry | Data access | Azure Monitor Query Logs SDK, KQL | `NEW-WEB-APP/server/analytics` |
| Product telemetry | Shared contract | Application Insights custom events | `src/services/telemetryService.ts` |
| Existing analytics | Query reference | Azure Workbook JSON | `scripts/workbook-content.json` |

## 4. Recipe Selection

**Selected:** Script-driven Bicep and Azure CLI

**Rationale:** The analytics application must not use azd or project templates. Explicit shell scripts build the image with ACR Tasks, provision resources with resource-group-scoped Bicep, update the Container App, and verify health. The root `azure.yaml` is now an explicit guard that blocks the retired generic azd deployment path. The application reuses the existing Container Apps environment, ACR, and source Log Analytics workspace.

## 5. Architecture

**Stack:** Single Azure Container App hosting a static React dashboard and Node.js API.

| Component | Azure Service | SKU |
| --- | --- | --- |
| Analytics web/API container | Azure Container Apps | Consumption, 0.5 vCPU / 1 GiB, min 1, max 3 |
| Container image | Existing Azure Container Registry | Existing Basic registry |
| Runtime identity | User-assigned managed identity | Dedicated analytics identity |
| Source telemetry | Existing workspace-based Application Insights | Existing |
| Query plane | Existing Log Analytics workspace | Existing PerGB2018 |
| Analytics observability | Application Insights | Workspace-based, existing workspace |
| Authentication | Container Apps built-in auth with Microsoft Entra ID | Single tenant, assignment required |

The API uses `DefaultAzureCredential` and `LogsQueryClient`. Its managed identity receives Log Analytics Reader on the source workspace. KQL is stored in a named, typed server-side query registry. API routes validate filters, cap time ranges and rows, batch compatible queries, and cache aggregates briefly. Optional AI recommendations consume only pre-aggregated analytics summaries and are disabled unless an Azure OpenAI endpoint and deployment are configured.

## 6. Provisioning Limit Checklist

This change reuses the existing Container Apps managed environment, ACR, Log Analytics workspace, and network. It creates one Container App, one user-assigned identity, one Application Insights component, and three role assignments. These resource types do not consume regional compute-family quota. The existing Container Apps environment supports additional apps, and the requested steady-state replica count is one.

| Resource Type | Number to Deploy | Total After Deployment | Limit/Quota | Notes |
| --- | --- | --- | --- | --- |
| `Microsoft.App/containerApps` | 1 | Existing count + 1 | Environment/platform limit, not vCPU family quota | Reuses existing East US 2 environment; 1 minimum replica |
| `Microsoft.ManagedIdentity/userAssignedIdentities` | 1 | Existing count + 1 | Subscription resource limit | Dedicated least-privilege identity |
| `Microsoft.Insights/components` | 1 | Existing count + 1 | Subscription resource limit | Uses existing workspace |
| `Microsoft.Authorization/roleAssignments` | 3 | Existing count + 3 | 4,000 per subscription | ACR pull, Log Analytics Reader, optional Cosmos read |

**Status:** All planned resources are within documented platform limits; no scarce SKU or regional vCPU quota is requested. Live subscription preflight remains required before deployment.

## 7. Execution Checklist

### Phase 1: Planning

- [x] Analyze workspace
- [x] Gather requirements
- [x] Confirm existing deployment location from repository deployment configuration
- [x] Prepare resource inventory
- [x] Scan codebase and existing workbook
- [x] Select recipe
- [x] Plan architecture
- [x] User approved the architecture and requested full implementation

### Phase 2: Execution

- [x] Generate application scaffold and dependencies
- [x] Implement telemetry contract and typed query registry
- [x] Implement secured analytics API, caching, and health endpoints
- [x] Implement dashboard, filters, drilldowns, and recommendation views
- [x] Add telemetry schema and API tests
- [x] Generate Docker and Bicep deployment artifacts
- [x] Add script-driven provision, deploy, and verification commands
- [x] Run local build, lint, tests, and container validation
- [x] Set status to `Ready for Validation`

### Phase 3: Validation

- [x] Invoke Azure validation workflow
- [x] Validate Bicep and deployment prerequisites
- [x] Record validation proof

### Phase 4: Deployment

- [x] Confirm live Azure subscription and Entra application parameters
- [x] Deploy with Azure deployment workflow
- [x] Verify authenticated endpoint and telemetry queries
- [x] Set status to `Deployed`

## 8. Validation Proof

| Check | Command Run | Result | Timestamp |
| --- | --- | --- | --- |
| Lint | `npm run lint` | Pass, zero warnings | 2026-07-20 |
| TypeScript and Vite | `npm run build` | Pass | 2026-07-20 |
| API contracts | `npm test` | Pass, 4 tests | 2026-07-20 |
| Shell syntax | `bash -n scripts/*.sh` | Pass | 2026-07-20 |
| Bicep compilation | `az bicep build --file infra/main.bicep` | Pass | 2026-07-20 |
| Production image | `docker build -t aadb-product-analytics:local .` | Pass | 2026-07-20 |
| Runtime health | `GET /api/health` in production container | Pass, healthy | 2026-07-20 |
| Browser workflow | Playwright navigation and responsive checks | Pass, 8 views and named controls | 2026-07-20 |
| Azure context | `az account show` and read-only resource resolution | Pass, ARTURO-MngEnvMCAP094150 / East US 2 | 2026-07-20 |
| ARM group validation | `az deployment group validate -g azure-diagrams-rg -f infra/main.bicep ...` | Pass, provisioning state Succeeded | 2026-07-20 |
| ARM what-if | `az deployment group what-if -g azure-diagrams-rg -f infra/main.bicep ...` | Pass, 6 creates / 15 ignores / 0 modifies / 0 deletes | 2026-07-20 |
| Production deployment | `scripts/01-provision.sh` and `scripts/02-deploy.sh` | Pass, authenticated Container App deployed | 2026-07-20 |
| Authentication callback | Entra browser login with ID token issuance enabled | Pass | 2026-07-20 |
| AADB telemetry source | In-container `GET /api/analytics/overview?range=90d` | Pass, 1,152 users / 1,601 sessions / 10,892 events | 2026-07-20 |
| Decision intelligence | In-container `GET /api/analytics/insights?range=90d` | Pass, funnel/models/findings/reliability/cohorts/actions populated | 2026-07-20 |
| Cosmos feedback | Authenticated Feedback view using `diagrams-db/feedback` | Pass | 2026-07-20 |

## 9. Files to Generate

| File | Purpose | Status |
| --- | --- | --- |
| `.azure/deployment-plan.md` | Deployment source of truth | Complete |
| `NEW-WEB-APP/package.json` | Application scripts and dependencies | Complete |
| `NEW-WEB-APP/src/` | Product analytics dashboard | Complete |
| `NEW-WEB-APP/server/` | Analytics API and KQL registry | Complete |
| `NEW-WEB-APP/Dockerfile` | Production image | Complete |
| `NEW-WEB-APP/infra/` | Resource-group-scoped Bicep | Complete |
| `NEW-WEB-APP/scripts/` | Azure CLI setup, deployment, and verification | Complete |

## 10. Next Steps

1. Use the deployed application at `https://aadb-usage-analytics.thankfulbeach-7e8f01bc.eastus2.azurecontainerapps.io`.
2. Grant additional maintainers with `scripts/04-grant-access.sh <user-upn>`.
3. Deploy future application revisions with `scripts/02-deploy.sh`.
