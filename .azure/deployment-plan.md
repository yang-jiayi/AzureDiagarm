# Azure Deployment Plan

> **Status:** Approved

Generated: 2026-09-07

## 1. Current Release

The user explicitly resumed GitHub publication and production deployment on
2026-09-07. Publish the completed local application changes from
`fix/collapsible-icon-palette-20260906`, based on published main
`a72e702f4dce07781b3a68a246fa6359fbe58d42`.

Include the compact, Draw.io-inspired editor, preserved authored styling,
Office/web/PNG paint and legend repairs, managed GPT-6 Astra-only catalog,
enhanced explicit BYO profiles and routing, and completed graph, reliability,
pricing-maintenance and release safeguards. The separately reported twelve
baseline export findings remain unfixed; this release does not claim to
resolve them.

## 2. Scope and Existing Architecture

Update the existing application image and its approved runtime settings only.
Reuse the existing Container App, environment, registry, Front Door, managed
identity, storage and OpenAI account. Preserve public-mode authentication,
application access control, shared Table budget accounting and origin isolation.
MCP and Azure Import remain disabled.

The existing GitHub and application `ALLOW_BYO_AI_ENDPOINTS` values are both
`true`; preserve that approved policy. Managed requests must use the verified
genuine Astra deployment only, while explicitly selected BYO connections retain
their own credentials, approved endpoints and model selection. Retired managed
model environment settings must not revive old models during release or rollback.

No new Azure resources, model/capacity/SKU changes, RBAC grants, network changes,
data migrations, stored-data deletions or Speech authorization changes are
included. The required Davis Dragon narration remains blocked separately.

## 3. Release Recipe

`recipe.type: azcli`

Use the existing checked-PR-to-main path and the current push-triggered
`.github/workflows/azurediagarm-sync-deploy.yml`. Do not dispatch upstream
synchronization or rerun an old workflow as a deployment shortcut. The exact
validated source, protected-main checks and current-main guards must agree.

The unchanged group-scoped `infra/gpt6-astra.bicep` is used only for current
Azure consistency validation and what-if. Do not apply that template or the
historical subscription bootstrap for this application-only release.

## 4. Production Context

Keep the existing `azurediagarm-app` in `AzureDiagarm_rg`, West US, at
`https://azurediagarm.mssql.biz`, using the existing GitHub OIDC subscription
and identity. Read-only preflight confirmed that the CLI subscription matches
the repository release target.

The current healthy baseline is main `a72e702f4dce07781b3a68a246fa6359fbe58d42`,
revision `azurediagarm-app--g34024872580-1`, with 100% latest-revision traffic and
1-2 replicas. Its image is
`sqlserverevoacr.azurecr.io/azurediagarm/app:u71ef7e82e354-ca72e702f4dce-20260906094914`.
Retain the workflow's rollback receipt and guarded rollback behavior.

## 5. Preparation and Validation Steps

- [x] Confirm explicit publication authorization and preserve existing work.
- [x] Confirm the same current main, target application, subscription and region.
- [x] Preserve BYO policy, managed Astra, authentication and budget boundaries.
- [ ] All validation checks pass.
  - [x] Core AZCLI validation: CLI/auth, Bicep compile, ARM validate and what-if.
  - [ ] Docker build/context and current-source standalone/combined runtime gates.
  - [x] Applicable Azure policy validation; no policy changes or exemptions.
  - [x] Static role review; no additional grants are needed by this release.
  - [x] Complete local source/server/MCP validation and release-guard review.
  - [ ] Required protected-main Linux build/browser/CodeQL checks.
- [ ] Record final evidence and validated source before merging for deployment.

## 6. Deployment Steps

Commit only the reviewed release inputs, create a checked PR, and merge only
after current-source gates pass. Monitor the resulting main-triggered workflow,
then confirm the deployed source/image/revision, readiness, authentication
boundaries, genuine Astra binding and preserved BYO policy.

No branch protections, required checks, identity boundaries or deployment guards
may be weakened to complete this release.

## 7. Validation Proof

Current-release validation is in progress. The earlier records below are
historical and do not authorize this candidate. Final command results, exact
source and GitHub run references will be recorded here by azure-validate.

The completed local UI/export integration already includes the normal build,
125 targeted units, 27 normal-build browser cases, four separately configured
managed-Astra browser cases, fourteen offline HTML cases, and native PowerPoint
paint/legend evidence. Those targeted results are not a substitute for the
complete current-source and protected-main release gates.

| Check | Current result | Evidence |
| --- | --- | --- |
| Existing target and rollback baseline | Current CLI subscription agrees with GitHub release variables. Existing application is Succeeded at the recorded PR #66 image/revision, 100% latest traffic, 1-2 replicas. Health returns 200; unauthenticated root returns the expected protected 401 | Read-only `az account show`, allowlisted `az containerapp show`, selected non-secret GitHub variables and public HTTP probes on 2026-09-07 |
| Existing AI and application policy | Genuine `gpt-6-astra` v2026-09-03, GlobalStandard 50 is Succeeded. Both configured BYO flags are already true; application access control, public mode and Table budget remain enabled. MCP/Azure Import remain disabled | Read-only deployment and application projections; no setting or role changes |
| Bootstrap syntax | Updated `infra/main.bicep` and its modules compile. This bootstrap is not applied | `az bicep build --file infra/main.bicep` with output outside the worktree |
| Core AZCLI consistency | All five official helper checks pass against a hash-matched copy of `infra/gpt6-astra.bicep` with the live account and capacity. Structured what-if has 15 Ignore and one Modify for omitted service-reported `properties.currentCapacity`; no resource creates/deletes. The helper's text counters also count property deltas and are not resource-deletion evidence | `validate-deployment.ps1 -Scope group -ResourceGroup AzureDiagarm_rg`; separate structured `az deployment group what-if --no-pretty-print`; no template apply |
| Applicable Azure policy | Subscription and inherited management-group assignments retrieved with Azure MCP. Actual deny conditions cover unrelated classic/VM/VMSS/AKS/SQL/HSM/Sentinel resources or ProvisionedManaged OpenAI capacity, not this app-only update or unchanged GlobalStandard model | Azure MCP `policy_assignment_list` and read-only definition/set-definition reads; no exemptions or governance changes |
| Existing runtime roles | Registry-scoped AcrPull, account-scoped Cognitive Services OpenAI User and storage-scoped Blob/Table data contributors are already assigned to the runtime identity. Existing ancillary roles are unchanged | Read-only role assignment projection; static source review and post-deployment comparison remain required |
| Static role review | Independent release review confirmed the existing runtime identity, resource-scoped AcrPull, OpenAI User and Blob/Table data roles cover this release. No grants or scope expansion are indicated | `infra/resources.bicep`, `infra/openai-role.bicep`, `infra/main.bicep`; existing live roles match |
| Complete local source gates | All 1,192 application units, 245 server cases and 71 MCP cases pass; full lint, production audit gates and MCP build pass with unchanged generated inputs. The final application build and script typecheck pass | Existing npm/tsx commands; final full application units used `--test-concurrency=1` after a Windows Node 24 transport failure under concurrent workload |
| BYO responsive correction | Full units exposed one noncanonical 520px media query. It now uses the existing 640px compact breakpoint. Both form grids switch correctly at 640/641px, without horizontal overflow or WCAG violations; the four existing managed/BYO policy cases also pass | `tests/breakpoints.test.ts` and the actual browser boundary regression; no assertion or timeout weakening |
| Release model-version correction | Independent release review identified a missing version comparison. The guard now rejects missing, different, empty, null or numeric versions and requires `2026-09-03`, with a regression tying that constant to the model template. The corrected helper also passes against live ARM account/deployment metadata | Ten deployment-security cases and 82 source/workflow/rollback cases pass; scoped lint passes; no model or Azure resource modification |
| Explicit account reference | Added the missing non-secret `AZURE_OPENAI_RESOURCE_ID` GitHub variable for the already verified production account. Existing endpoint, deployment alias and BYO policy are unchanged | Exact read-back matches the verified account; no credentials, role grants or alternative target introduced |

## 8. Rollback and Limitations

Retain the current image/revision and use only this release's guarded rollback
path if needed. Never restore retired managed deployments or weaken BYO,
authentication or budget policy as a workaround.

Deployment success does not establish recovery of the earlier upstream Astra
inference incident. Browser AI fixtures are mocked, native Visio remains
unavailable, and the twelve export-assessment findings remain explicitly open.

---

# Historical PR #66 Validation Record

> **Status:** Validated

Generated: 2026-09-06

## 1. Current Release

Publish the prioritized application repairs in
[PR #66](https://github.com/yang-jiayi/AzureDiagarm/pull/66) from
`fix/astra-generation-rate-limits-20260906`, then update the existing production
AzureDiagarm application. The user requested production/GitHub publication,
reliable generation, a substantially more compact workspace, and prioritized
quality improvements using multiple assessment owners.

This is an application release, not another model migration. Genuine GPT-6 Astra
is already deployed. Do not fall back to GPT-5.6, lower selected generation
quality, refund unknown usage, or change authentication to conceal a failure.

## 2. Scope and Existing Architecture

The changes cover compact/responsive controls, accessible palette and recovery
actions, bounded AI retry/admission/error handling, graph and Office export
fidelity, explicit incomplete-IaC comparisons, measured editor performance, and
release/standalone-MCP packaging safeguards.

Reuse the existing Container App, registry, Front Door, managed identity,
storage, OpenAI account and protected-main workflow. No new Azure resources,
model/SKU/capacity changes, role grants, network changes, data migrations or
retention-policy changes are part of this release. MCP stays disabled in
production. The unrelated historical Product Analytics plan below is not in
scope.

## 3. Release Recipe

`recipe.type: azcli`

Use the existing `.github/workflows/azurediagarm-sync-deploy.yml` push-to-main
release path after a reviewed, fully checked PR. Do not use manual upstream
synchronization or a historical workflow rerun as a deployment shortcut.
The new guard requires the allowed main ref, validated checkout and current
remote main to agree before Azure access and again before the deployment marker.
It does not retrofit guards into immutable pre-fix workflow versions.

Azure validation uses the existing group-scoped model template for unchanged
infrastructure consistency checks; it must not be applied during this app-only
release. The actual changed application and both container recipes require
their own complete current-source gates. An unchanged infrastructure preview
alone is not application validation.

## 4. Production Context

Target the existing `azurediagarm-app` in `AzureDiagarm_rg`, West US, at
`https://azurediagarm.mssql.biz`, using the same configured subscription and
GitHub OIDC identity as the successful PR #63 rollout. Production currently
serves main commit `1521996ba15ca4cbf94e5d186db586090be74e5d`, revision
`azurediagarm-app--g34002727786-1`. Preserve that healthy application revision
as this release's rollback baseline.

Keep public deployment mode, Easy Auth, application access control, shared Table
budget accounting and origin isolation enabled. Existing administrator access
must remain intact; non-admin users still need application allowlist entries.
Secrets and operator identity values must not be copied into this plan.

## 5. Preparation and Validation Steps

The user-approved application release is validated against source commit
`95968aff1777e501756888f54797d497929e6acf`. Complete Linux application, browser,
standalone/combined-runtime and required CodeQL checks now pass, together with
the recorded native Office and Azure preflight evidence. The subsequent
documentation-only validation record must satisfy the same protected-branch
checks before merge. Historical records below do not authorize this candidate.

- [x] Preserve existing production and local work; publish only completed owned checkpoints.
- [x] Complete scoped security, maintenance, graph, design and performance assessments.
- [x] Publish server, release/packaging, graph/IaC and compact-UI checkpoints in draft PR #66.
- [x] All validation checks pass.
  - [x] Core AZCLI validation: authenticated CLI, compile the unchanged scoped template, ARM validation and structured what-if.
  - [x] Docker build/context: current-source standalone MCP and complete application images, generated inputs and runtime readiness.
  - [x] Azure policy validation: unchanged SKU/identity/network scope remains compatible with applicable constraints.
  - [x] Static role verification: existing account-scoped roles still match application operations; no widening is needed.
  - [x] Complete the populated JA/768x720 toolbar correction without raising the existing 220px bound.
  - [x] Complete real production Worker/CSP/parity/cancellation proof and controlled performance measurements.
  - [x] Complete AI client signal compatibility and final retry/cancellation/partial-output browser coverage.
  - [x] Generate current Office fixtures and exercise real PowerPoint opening, text bounds and connector movement.
  - [x] Pass the final integrated source, browser, runtime and all six required GitHub checks.

## 6. Deployment Steps

Once every current application and infrastructure preflight gate passes, record
the exact evidence below and set this release to Validated through azure-validate.
Then invoke azure-deploy, guard the exact PR head and current main, and use the
existing permitted PR merge path. Monitor the new main-triggered production
workflow and confirm its actual commit, image, healthy revision, traffic,
protected endpoints and fresh frontend/Worker assets.

Do not weaken checks, branch rules, identity boundaries or rollback protection.
Save in-progress user work before requesting a browser reload.

## 7. Validation Proof

Evidence timestamps on 2026-09-06 (UTC): Azure preflight log finalized at
08:17:16, native PowerPoint log at 08:37:56, and the successful final source CI
run completed at 09:07:37.

| Check | Current result | Evidence |
| --- | --- | --- |
| Server diagnostics/cooldown checkpoint | Passed scoped proxy, budget and rate-limit regressions; no accounting/auth changes | `586e550` |
| Maintenance checkpoint | Passed Linux core, standalone/complete runtime image builds and browser safeguards | `131ccc3`, CI `34012863530` |
| Graph/IaC checkpoint | Passed integrated Linux source, Office/core, browser, MCP and runtime gates | `fd2f765`, CI `34015994681` |
| Initial compact UI checkpoint | Main critical/visual cases and Vite/MCP/runtime passed, but one populated header failed at 238px vs 220px; corrected and confirmed in the next row | `ff4f138`, CI `34016824587` |
| Populated UI correction | Dated pricing metadata shares the existing 44px row, recovering 47.219px of Japanese width. All 36 unchanged Linux combinations now pass, with maximum header 194.375px instead of the failing 238px. No App JSX, control-size, snapshot or 220px-threshold changes | `cd1af5c`; integrated candidate `c466b11`, CI `34022669317`, browser job `101457987210` |
| AI client handoff | Signal compatibility and synthetic correlation fixtures are integrated; all 133 targeted units and the complete 93-case browser suite pass. All three Both cancellation cases run without AbortSignal.any, proving active request cancellation and no late output/fan-out | `npx tsx --test` over the six affected AI unit files; `npm run test:ai-ui`; scoped ESLint |
| AI fixture integration | Full browser coverage exposed two obsolete assumptions: completed manifest scopes need not stay linked to cancellation, and a reasoning-capable Astra report includes its effort suffix. Fixtures now assert exact active-output cancellation, no extra dispatches, and the real reasoning-qualified report key | Seven focused regressions and then all 93 browser cases pass; production comparison/report behavior is unchanged |
| Performance handoff | Production/CSP and cold-development browser cases each pass 6/6; strict ten-fixture Node/browser parity and all 18 warm output hashes match; source types/lint and 34 focused cases pass | Published checkpoint `c5d1194`; all 12 source hashes match the original controlled private performance report |
| Performance measurements | Paired synthetic desktop drag frame-gap p95 improves 70.9 to 50.7 ms at 80 nodes and 203.2 to 135.8 ms at 250 nodes. Warm Worker samples have no main-thread tasks of at least 50 ms. Some layout wall times increase; no small-diagram, 60fps or field-INP claim | Both run orders, four valid 24-transform runs per size/mode; same UI and unchanged layout output |
| Core Azure consistency | Official helper passes installed/authenticated CLI, Bicep compilation, group ARM validation and what-if. A structured preview has 15 Ignore and one Modify, only omitted service-reported `properties.currentCapacity`; zero resource creates/deletes and no template apply | `validate-deployment.ps1 -Scope group -ResourceGroup AzureDiagarm_rg`, using a hash-matched private copy of `infra/gpt6-astra.bicep`, current subscription and account/capacity parameter file; separate `az deployment group what-if --no-pretty-print` |
| Actual Astra and runtime preflight | Genuine `gpt-6-astra` v2026-09-03 remains Succeeded at GlobalStandard 50. Current application revision is Running/Succeeded at 1-2 replicas with public/Easy Auth/access control/Table budget enabled and MCP/Azure import disabled | Read-only `az account show`, `az cognitiveservices account deployment show` and an allowlisted `az containerapp show` projection |
| Applicable Azure policy | Effective assignments and actual deny definitions reviewed. ProvisionedManaged and unrelated classic/VM/AKS/SQL/HSM restrictions do not match this release. Actual model ARM validation passes; no exemptions or assignment changes | Azure MCP `policy_assignment_list`; native Azure CLI policy definition/set-definition reads |
| Static role verification | Existing templates scope Cognitive Services OpenAI User to its account, Blob/Table data contributor roles to storage and AcrPull to the registry. This app-only release changes no principal, assignment or required data operation | `infra/openai-role.bicep`, `infra/resources.bicep`; live role confirmation remains a post-deployment step |
| Fresh browser/native Office | Seven fixture files plus three actual UI exports generated from current source. PowerPoint opens/renders all six presentations (16 slides), fits 372 native text blocks, and retains the three role-like-ID groups and connector movement | `npm run test:exports:browser` with an isolated artifact directory, then `npm run test:exports:desktop`; representative rendered slides visually reviewed |
| Current source build/types | Fresh integrated Vite/application build, script/test type check and final scoped lint pass | `npm run build`, `npm run typecheck:scripts`, scoped `npx eslint` |
| Final security follow-up | No reportable vulnerabilities in reviewed Worker/Vite/App performance and small AI signal-compatibility changes | Existing read-only security specialist; no complete-security guarantee |
| Integrated Linux browser checks | All 100 critical/visual/Worker cases, 36 populated workspace combinations, 93 AI/UI cases, 16 inspector/review cases and modal safeguards pass | Candidate `c466b11`, CI `34022669317`; all three CodeQL analyses and standalone MCP image also pass |
| Provenance fixture integration | The first run stopped at 930/932 units because its transport mock omitted the new budget-classifier export. The fixture now re-exports the real classifier rather than replacing retry semantics; all 67 affected cases, scoped lint and the complete Linux rerun pass | Follow-up `95968af`; CI `34023258712` succeeds |
| Final integrated Linux/runtime checks | All application/core/Office/workspace/server gates, standalone MCP image, complete runtime image and runtime readiness pass. All 100 critical/visual/Worker and supplemental browser safeguards pass | CI `34023258712`, exact source `95968aff1777e501756888f54797d497929e6acf`; Vite job `101459614086`, MCP job `101459614107`, browser job `101459614133` |
| Final required GitHub checks | All six protected-main contexts are successful and completed on that same source head: Vite, MCP, critical browser flows and actions/JavaScript-TypeScript/Python CodeQL | Fresh `gh pr view 66` and `gh run view 34023258712` results; existing named-user PR-only bypass confirmed without changing rules |
| Existing-app pre-deployment RBAC | The live Container App and registry use the same existing user-assigned identity; its registry-scoped AcrPull assignment is present. Reuse the existing environment and image-deployment workflow, not an AZD/placeholder provisioning flow | Read-only `az containerapp show` and `az role assignment list`; no identity, role or environment changes |

Native Visio is not installed on this machine. Browser VSDX/package/relationship
coverage is not represented as native Visio execution.

The existing OpenAI account still has two audit-only findings: Private Link is
absent and network access is unrestricted. Fresh Policy Insights results and
the actual built-in definitions confirm both target the unchanged account, not
this application revision or child-template consistency check. Do not change
shared connectivity or create an exemption for this release. A separate network
migration is needed to address them; no estate-wide compliance claim is made.

### Material provider limitation

Meaningful max-quality Astra generation failed independently of application
orchestration. A bounded same-model canary did not complete and was removed;
its quota was restored. The final streaming hypothesis remains untested because
runtime-console setup and the alternate existing operator credential could not
establish an authorized inference control. No streaming/configuration switch
is justified. Private diagnostic/support evidence is retained outside Git.

This release can improve application safeguards and usability without claiming
that the upstream generation incident has recovered. Any remaining provider
blocker must be explicit in the production handoff; tiny readiness responses,
HTTP 200 headers or incomplete output are not proof of architecture generation.

## 8. Rollback and Data Safety

Retain the current healthy PR #63 application image/revision before rollout.
Use the established guarded rollback path only for this deployment's failure;
do not rerun an old pre-guard workflow or alter model/auth/budget policy as a
shortcut. No stored architecture or feedback records are intentionally deleted
or rewritten by this release.

---

# Historical PR #63 Validation Record

> **Historical status:** Validated

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
reviewed visual baselines also pass. The supplemental workspace harness now
exercises genuine AI regeneration rather than file import; this exposed and
fixed canonical service-alias identity handling. All local supplemental browser
safeguards and 810 unit tests pass. Linux additionally exposed a two-pixel
tablet export-header overflow; tablet-only row spacing now reserves additional
font-metric headroom without changing controls or clipping content. Linux now
passes the complete critical/visual and workspace harnesses. Its remaining AI/UI
failure resolved a mocked response before an external prop edit had committed;
that fixture now explicitly commits those edits while retaining asynchronous
application callbacks. The complete Linux gate now passes: all 86 critical and
visual cases, the workspace harness, all 85 AI/UI cases, all 16 inspector/review
cases, and modal focus. App, MCP, full-container readiness, and all three CodeQL
checks also pass on `1c84d16535dd374be526a89275a8588638cb5ce7`. The official
Azure checks and structured infrastructure preview were rerun successfully.
The validation-record commit must also satisfy the unchanged protected-branch
checks before the permitted PR merge triggers production deployment.

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
- [x] Complete release validation and record evidence.
- [x] Core Azure validation: CLI/authentication, compile the actual model template, ARM validation, and what-if.
- [x] Build the final container using a Docker-capable builder and inspect its context.
- [x] Review applicable Azure policy constraints and unchanged runtime role scopes; record existing account-level audit findings separately.
- [x] Pass complete application, server, MCP, browser, and required GitHub checks.

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
| Final static application gates | `npm run lint`, `npm run typecheck:scripts`, and all 810 root unit tests pass, including canonical service aliases, the bounded AI queue, cloud hydration/import fixes, and administrator log-privacy contract | 2026-09-05 |
| Remaining core scripts | Workflow contracts, ARM extraction, layout preservation, icon library, icon workspace, validation freshness, and service-name normalization all pass | 2026-09-05 |
| Genuine Astra production build | `npm run build` passes with the verified OpenAI endpoint and `VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=gpt-6-astra` | 2026-09-05 |
| Focused browser repairs | Theme/contrast, explicit draft recovery, service-inspector focus, AI cancellation/review, score-zero metadata, modal keyboard safety, and atomic AI import checks pass. The import preserves authoritative IDs/pricing and makes zero writes to the old source document | 2026-09-05 |
| AI comparison admission | Shared cancellable budget queue passes 111 focused units and 85 AI/UI checks, including 43 comparison cases; concurrency limits and cancellation assertions remain intact | 2026-09-05 |
| Cloud browser arrangements | Explicit author edits replace 18 incidental hydration-write setups. All 19 focused cases and three replacement-race runs pass. The complete Windows run reaches 84/85; the remaining unchanged keyboard/access-dialog case encountered delayed Vite loading and subsequently passed both keyboard/WCAG control runs. Fresh Linux CI remains required | 2026-09-05 |
| Focus-mode persistence | Linux traces reloaded before confirmed persistence. The test now waits for `Saved on this device`, retaining every recovery/Escape/focus assertion. Focus-mode and recent-work recovery pass all six repeated local runs | 2026-09-05 |
| Workflow visual review | All four baselines use exact reviewed Linux captures, with byte-identical attempts and unchanged dimensions. Differences reflect approved status-bar separation, shadows, and upper-border positioning. Dark/mobile/forced-colors comparisons are 0 differing pixels against both attempts; forced-color current-step text retains 11.31:1 contrast. The 1% threshold and failure semantics remain unchanged | 2026-09-05 |
| Initial required CI | Vite/core/container and all three CodeQL analyses passed in runs 33991330053/33991329211. Browser and MCP audit failures block merge until corrected | 2026-09-05 |
| Repaired required CI | Runs 33995733963/33995732452 pass Vite/core, patched MCP, complete image/startup/readiness, all 85 functional browser cases, and all CodeQL analyses. Only the now-reviewed three visual baselines prevented the full browser job from completing; fresh complete CI remains required | 2026-09-05 |
| Complete critical/visual browser gate | Run 33996721727 passes all 86 critical and visual cases. Its subsequent workspace harness failed because it used a new-document import as an in-place regeneration proxy; the harness now invokes explicitly mocked genuine Astra generation and proves a real snapshot transaction abort before retry | 2026-09-05 |
| Regeneration service identities | Two focused tests reproduced lost manual instance IDs and falsely unique aliases. Proposal reconciliation now resolves catalog service aliases before uniqueness checks, preserves custom labels/pricing/connections, and leaves group names and explicit import identities unchanged. All focused and full unit cases pass | 2026-09-05 |
| Supplemental browser safeguards | Current `test:workspace:browser`, all 85 AI/UI cases, all 16 inspector/review cases, and modal-focus safeguards pass. Regeneration proves retained node/workflow IDs, exactly three genuine Astra requests, snapshot-failure retry without duplicate generation, undo, pricing, and review persistence | 2026-09-05 |
| Tablet header headroom | Linux reported 221.96875px against the existing 220px bound for JA/EN populated export headers at 768px. A tablet-only row-gap adjustment preserves all 42 control dimensions, passes all 36 local responsive-header cases and both WCAG audits, and passes 17 canonical/contrast cases. Run 34000522676 confirms the full Linux workspace harness passes with maximum header height 213.1875px. Desktop/mobile styles and every threshold remain unchanged | 2026-09-06 |
| Committed AI baseline fixture | Run 34000522676 passes all 86 critical/visual cases and the workspace harness, but its AI/UI suite reaches 84/85 because an external `root.render` edit had not committed before the mocked response resolved. Only the two edited-baseline arrangements now use explicit `flushSync`; ordinary application callbacks and asynchronous race/cancellation behavior remain unchanged. `npx tsx --test tests\aiGenerationUI.browser.ts` passes all 85 cases; focused ESLint and `npm run typecheck:scripts` pass. Complete Linux confirmation remains required | 2026-09-06 |
| Final complete Linux application gates | CI run 34001474402 succeeds on `1c84d16535dd374be526a89275a8588638cb5ce7`: Vite/core/Office/server tests, patched MCP audit/build/tests, complete runtime image and API/MCP readiness, all 86 critical/visual cases, workspace browser safeguards, all 85 AI/UI cases, all 16 inspector/review cases, and modal focus. Browser job 101400994008 completes every supplemental command; maximum responsive header height is 213.1875px. No timeout, screenshot threshold, or correctness assertion was weakened | 2026-09-06 00:40 UTC |
| Final required CodeQL checks | Run 34001472961 succeeds on the same release head for actions, JavaScript/TypeScript, and Python. All six required checks are successful; the existing PR-only owner authorization remains available without changing branch protections | 2026-09-06 |
| Runtime dependency patches | Patched MCP `fast-uri` to 3.1.7 and `qs` to 6.16.0. API Express 4/body-parser restrict `qs` to the vulnerable 6.15 line, so a compatible `qs` override selects 6.16.0 without upgrading Express's major version. Both production audits now report zero vulnerabilities; patched MCP 65 and API 120 tests pass, with generated MCP assets unchanged | 2026-09-05 |
| Deployment log privacy | Created the `ACCESS_ADMIN_EMAIL` Actions secret from the exact existing administrator value through stdin, without a trailing newline or command-line disclosure. Workflow references use the secret and a quoted environment variable; the administrator and runtime value remain unchanged | 2026-09-05 |
| Final source review | Read-only review reports no significant findings in normalized cloud baselines, canonical serialization, and their save/copy/conflict interactions. Final lint and script type checks pass; nine deployment workflow security contracts pass | 2026-09-05 |
| Remote source currency | Refreshed `origin/main` remains `c885477c799a35a044a73ac05b230aeab7160f95`; no unresolved index entries or whitespace errors | 2026-09-05 |
| Refreshed live preflight | Existing healthy revision and rollback image remain unchanged. Astra is Succeeded at the verified model/version and capacity. Without following authentication redirects, public health returns 200, protected root/API/MCP return 401, and the direct origin returns 403. Browser-style requests redirect to Microsoft sign-in, not anonymous API JSON | 2026-09-05 |
| Final official Azure revalidation | `validate-deployment.ps1 -Scope group -ResourceGroup AzureDiagarm_rg -Template infra\gpt6-astra.bicep` with the existing subscription and non-secret account/capacity parameter file again passes CLI/authentication, Bicep compilation, ARM validation, and what-if. A fresh `az deployment group what-if --no-pretty-print` structured preview confirms 15 Ignore and one Modify, with only service-reported `properties.currentCapacity` omitted; zero resource creates/deletes. Generated compilation output was removed | 2026-09-06 |

Policy Insights reports no evaluated rows for the existing Container App; an
empty result is not proof of estate-wide compliance. The existing OpenAI account
has two pre-existing audit-only findings: Private Link is absent and network
access is unrestricted. Both policies target the account, not the child model
deployment. This release neither replaces nor updates the account's networking.
Do not silently change its shared connectivity or create policy exemptions to
clear those findings; they require a separately planned network migration.
No subscription-wide compliance claim is made.

Integrated-source, genuine Astra, browser, complete-container, required GitHub,
and official Azure validation are complete. Evidence covers the application at
`1c84d16535dd374be526a89275a8588638cb5ce7`; the subsequent documentation-only
validation record must finish the same required checks before merge. Historical
evidence below does not authorize this release. The operator's direct AI
data-plane call lacks the required inference permission; model checks instead
used the existing authorized runtime identity, without granting new roles or
changing network controls.

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
