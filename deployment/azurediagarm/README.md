# AzureDiagarm production deployment

The production site is deployed to Azure Container Apps and exposed only through Azure Front Door Standard with WAF.

## Cost controls

- Azure Front Door uses the Standard tier, matching the `SQLServerEvo_rg` reference architecture.
- The candidate workflow targets the Consumption workload profile with a minimum of 1 and a maximum of 2 replicas. The read-only release preflight observed the current production maximum at 3; the candidate now fails closed rather than silently reducing it. A reviewer must explicitly resolve this difference before an ordinary release can proceed.
- Images use the existing `sqlserverevoacr` Basic registry instead of creating another paid registry.
- Managed AI uses genuine GPT-6 Astra. The frontend deployment, server `AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA`, and singleton managed-deployment allowlist must agree. The observed production setting is `ALLOW_BYO_AI_ENDPOINTS=false`; the candidate requires the explicit setting and compares it with the running app instead of defaulting or silently enabling BYO.
- Upstream merging requires manual dispatch on `main`; deployments run automatically for validated changes pushed to `main`.
- Azure Communication Services sends deployment-result and upstream-validation-failure notifications. A deployment rejected by the initial source guard stops before Azure sign-in and does not send a deployment email.

## Update flow

1. **Ordinary release:** merge a reviewed pull request into `main`. The push starts `.github/workflows/azurediagarm-sync-deploy.yml`, validates the source, builds and pushes a uniquely tagged image, deploys a Container Apps revision, verifies production, and emails the result. Do not use manual workflow dispatch for an ordinary release.
2. **Upstream synchronization:** manually dispatch the workflow on `main` only when intentionally reviewing synchronization with `Arturo-Quiroga-MSFT/azure-architecture-diagram-builder`. Protected-file checks and validation still apply. A changed, validated merge is published through the guarded publication jobs; its new push to `main` starts deployment. Do not dispatch from a feature or release-record branch.
3. **Source freshness:** before Azure changes, the deployment job requires `refs/heads/main`, checks that its checkout is the exact validated commit, and reads the current `main` SHA from GitHub. It rejects a mismatch, malformed response, or unavailable API. It reads `main` again immediately before the revision update, before creating the deployment-started marker. `force_deploy` does not bypass either check or the upstream guard.

### New-tenant release prerequisites and review differences

The workflow source is pinned to tenant `02310481-6752-4d53-8fc4-f597a8e047f4`, subscription `0df5f427-7359-447e-9b60-fa67c5588b8f`, resource group `AzureDiagarm_rg`, and Container App `azurediagarm-app`. It fails before Azure sign-in if those GitHub variables do not identify that exact target.

The production OIDC jobs name the GitHub Environment `production`, making the expected Azure federated-credential subject `repo:yang-jiayi@28970082/AzureDiagarm@1301512171:environment:production`. The Environment, its required-reviewer/branch protections, the corresponding Azure federated credential, and the Azure roles have **not** been created or approved. Adding `environment: production` to workflow source does not prove those controls exist; do not treat the candidate as releasable until those prerequisites are reviewed and configured.

The candidate validates exactly one access-list store. For the preflighted Tables-backed target, set `AZURE_TABLES_ACCESS_ENDPOINT` and leave `AZURE_ACCESS_KEY_VAULT_RESOURCE_ID` empty; both-present and neither-present configurations fail closed because the server prefers Key Vault when its resource ID is set. The candidate also requires explicit `MCP_ENABLED=true`, `ALLOW_BYO_AI_ENDPOINTS=false`, `AI_BUDGET_STORE=table`, the budget limit/concurrency values, and the budget Tables endpoint. These settings must match the running app; omitted values or runtime drift block before Azure changes. Budget values were not recorded in the source preflight and must not be replaced with workflow defaults.

Review the candidate's deliberate runtime differences against the read-only production preflight before release:

- Maximum replicas change from 3 to 2. The workflow rejects a current maximum above the candidate value, so this release is blocked until reviewers explicitly resolve the scale reduction.
- Startup/readiness probes use `/readyz` instead of the observed `/healthz` probe path; liveness remains `/healthz`.
- Container Apps ingress is restricted to the Azure Front Door backend service tag.
- Front Door origin timeout changes from 30 to 240 seconds, and the health-probe interval changes from 100 to 120 seconds.
- The deployment adds the diagnostic setting `azurediagarm-local-logs`.

These are not pre-approved merely because they are documented. The candidate keeps the existing Easy Auth audience check and Defender-for-Storage PUT behavior unchanged; their one-time prerequisites remain pending separate review/authorization.

Ordinary releases use the committed `.github/upstream-baseline.json` rather than fetching upstream. The record must identify the approved repository and a full commit in the checkout's ancestry. Keep the record, `scripts/upstream-baseline.mjs`, and the workflow in the same change; an absent or invalid committed record fails closed. Guarded synchronization requires the fetched upstream history to retain the approved baseline and advances the record only after protected-file review and successful merge selection.

Set `AZURE_OPENAI_RESOURCE_ID` to the full account resource ID for the approved
endpoint. Before image publication and revision update, model preflight reads
that account and deployment through ARM, verifies endpoint ownership, and
requires the actual model to be `OpenAI/gpt-6-astra` with successful provisioning.
Lookup failures or a differently named model behind the alias block release.
This does not perform inference, grant roles, or delete old Azure deployments.

### Safe retries and rollbacks

- Compare the run's commit with current `main` before using GitHub's **Re-run jobs**. Retry only a current-main run whose workflow version already contains both source checks. A transient GitHub lookup failure must be resolved, not bypassed.
- If `main` has advanced, do not retry the superseded run. Use the release associated with the new current-main push; corrections belong in a reviewed pull request into `main`.
- **Historical-run limitation:** GitHub reruns use the original immutable workflow version. Adding the guard to today's files does not retrofit pre-fix runs. Never rerun a pre-guard workflow as a recovery shortcut, even if its original checks were green. Do not change branch protections, bypass rights, OIDC federation, or authentication to work around a refusal.
- The existing automatic rollback remains separate: if a revision update actually started and subsequently fails verification, the workflow copies the previously ready revision and checks its health. A source check rejected before the deployment-started marker must not trigger a revision rollback.
- Rollback preserves the pre-update verified Astra endpoint, singleton managed deployment, and approved BYO enablement policy. It must not restore an older image's stale model allowlist or reread mutable configuration during recovery. Unrelated baseline settings and secret references stay intact; no new model lookup is required during recovery.
- The production verification step is bounded to 20 minutes. OpenAI and MCP boundary probes use 15-second connection and 45-second total request limits so stalled responses can fail the check and reach rollback.
- For an intentional rollback, use the approved recovery procedure or a reviewed revert released through current `main`, not a historical workflow rerun. Manual dispatch remains an upstream-sync operation, not a rollback command.

## Local image testing

The [local Docker recipe](../../README.md#docker-deployment-local-only) explicitly selects `APP_DEPLOYMENT_MODE=local`, disables the local access-list gate, uses an in-memory development budget, and publishes only on loopback. This is not a public deployment configuration. Public deployments retain mandatory Entra authentication, the app access list, shared durable state, and Front Door origin isolation.

The standalone MCP image is also built from the repository root (`docker build -f mcp-server/Dockerfile -t aadb-mcp .`). Its prebuild imports the shared pricing expansion module and reads the public pricing dataset, Microsoft catalogs, and official icons. Linux CI builds this actual image in the existing MCP job; source-tree tests alone do not validate Docker's copied inputs.

## Security layers

- The wildcard `*.mssql.biz` certificate is read from `westuskvl` by the Front Door managed identity and tracks the latest Key Vault secret version.
- WAF runs in Prevention mode with rate limiting and known AI crawler `User-Agent` blocking.
- The origin validates `X-Azure-FDID`, preventing direct Container Apps access from bypassing WAF.
- Application responses include anti-indexing headers and `robots.txt`; these controls discourage compliant crawlers while WAF handles known automated clients.
- Managed Astra requests are proxied server-side with the app's managed identity; no managed credential is embedded in the browser bundle. BYO requests use only the supplied user's key, never the app identity or its server-side key. BYO keys remain in tab memory and request transit, not saved profiles, diagrams, telemetry, or deployment settings.
- Diagram blobs and shared rate-limit tables use one keyless storage account with shared-key access, public blobs, and public-network access disabled. A Network Security Perimeter is the intentional ingress control.
- Defender for Storage is explicitly overridden at the account: sensitive-data discovery remains enabled, while on-upload malware scanning is disabled because its Event Grid integration is incompatible with the perimeter. The deployment workflow enforces this state to prevent recurring authorization failures.
- The resource-scoped `AzureDiagarm-Storage-NSP` exemption records the NSP as the mitigating control for applicable Private Link and VNet-rule audits in the policy initiative actually assigned to the target. The workflow queries inherited assignments, then accepts only `Azure_Security_Baseline` at the tenant management-group scope with the expected initiative ID; child-scope assignments are not substitutes. It validates the exemption's ARM resource ID for the exact storage-account scope, its assignment binding, `Mitigated` category, future expiry, and every applicable reference present in that initiative. It does not skip a missing exemption or require references absent from the initiative. The read-only preflight found two of the six previously assumed reference candidates in the current initiative. The exemption currently expires on 2027-08-04, so the design must be reviewed. Do not add a Private Endpoint until Container Apps has a verified VNet and private-DNS path.
