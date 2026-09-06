# AzureDiagarm production deployment

The production site is deployed to Azure Container Apps and exposed only through Azure Front Door Standard with WAF.

## Cost controls

- Azure Front Door uses the Standard tier, matching the `SQLServerEvo_rg` reference architecture.
- The Container App uses the Consumption workload profile with a minimum of 1 and a maximum of 2 replicas; the production workflow does not enable scale-to-zero.
- Images use the existing `sqlserverevoacr` Basic registry instead of creating another paid registry.
- AI generation defaults to genuine GPT-6 Astra. Additional model deployments are available only when their frontend configuration and server-side provider allowlists agree; do not substitute GPT-5.4 Mini to recover an Astra incident.
- Upstream merging requires manual dispatch on `main`; deployments run automatically for validated changes pushed to `main`.
- Azure Communication Services sends deployment-result and upstream-validation-failure notifications. A deployment rejected by the initial source guard stops before Azure sign-in and does not send a deployment email.

## Update flow

1. **Ordinary release:** merge a reviewed pull request into `main`. The push starts `.github/workflows/azurediagarm-sync-deploy.yml`, validates the source, builds and pushes a uniquely tagged image, deploys a Container Apps revision, verifies production, and emails the result. Do not use manual workflow dispatch for an ordinary release.
2. **Upstream synchronization:** manually dispatch the workflow on `main` only when intentionally reviewing synchronization with `Arturo-Quiroga-MSFT/azure-architecture-diagram-builder`. Protected-file checks and validation still apply. A changed, validated merge is published through the guarded publication jobs; its new push to `main` starts deployment. Do not dispatch from a feature or release-record branch.
3. **Source freshness:** before Azure changes, the deployment job requires `refs/heads/main`, checks that its checkout is the exact validated commit, and reads the current `main` SHA from GitHub. It rejects a mismatch, malformed response, or unavailable API. It reads `main` again immediately before the revision update, before creating the deployment-started marker. `force_deploy` does not bypass either check or the upstream guard.

### Safe retries and rollbacks

- Compare the run's commit with current `main` before using GitHub's **Re-run jobs**. Retry only a current-main run whose workflow version already contains both source checks. A transient GitHub lookup failure must be resolved, not bypassed.
- If `main` has advanced, do not retry the superseded run. Use the release associated with the new current-main push; corrections belong in a reviewed pull request into `main`.
- **Historical-run limitation:** GitHub reruns use the original immutable workflow version. Adding the guard to today's files does not retrofit pre-fix runs. Never rerun a pre-guard workflow as a recovery shortcut, even if its original checks were green. Do not change branch protections, bypass rights, OIDC federation, or authentication to work around a refusal.
- The existing automatic rollback remains separate: if a revision update actually started and subsequently fails verification, the workflow copies the previously ready revision and checks its health. A source check rejected before the deployment-started marker must not trigger a revision rollback.
- For an intentional rollback, use the approved recovery procedure or a reviewed revert released through current `main`, not a historical workflow rerun. Manual dispatch remains an upstream-sync operation, not a rollback command.

## Local image testing

The [local Docker recipe](../../README.md#docker-deployment-local-only) explicitly selects `APP_DEPLOYMENT_MODE=local`, disables the local access-list gate, uses an in-memory development budget, and publishes only on loopback. This is not a public deployment configuration. Public deployments retain mandatory Entra authentication, the app access list, shared durable state, and Front Door origin isolation.

The standalone MCP image is also built from the repository root (`docker build -f mcp-server/Dockerfile -t aadb-mcp .`). Its prebuild imports the shared pricing expansion module and reads the public pricing dataset, Microsoft catalogs, and official icons. Linux CI builds this actual image in the existing MCP job; source-tree tests alone do not validate Docker's copied inputs.

## Security layers

- The wildcard `*.mssql.biz` certificate is read from `westuskvl` by the Front Door managed identity and tracks the latest Key Vault secret version.
- WAF runs in Prevention mode with rate limiting and known AI crawler `User-Agent` blocking.
- The origin validates `X-Azure-FDID`, preventing direct Container Apps access from bypassing WAF.
- Application responses include anti-indexing headers and `robots.txt`; these controls discourage compliant crawlers while WAF handles known automated clients.
- Azure OpenAI requests are proxied server-side and authorized with the Container App's user-assigned managed identity; no API key is embedded in the browser bundle.
- Diagram blobs and shared rate-limit tables use one keyless storage account with shared-key access, public blobs, and public-network access disabled. A Network Security Perimeter is the intentional ingress control.
- Defender for Storage is explicitly overridden at the account: sensitive-data discovery remains enabled, while on-upload malware scanning is disabled because its Event Grid integration is incompatible with the perimeter. The deployment workflow enforces this state to prevent recurring authorization failures.
- The resource-scoped `AzureDiagarm-Storage-NSP` exemption records the NSP as the mitigating control for the six built-in Private Link and VNet-rule audits. It expires on 2027-08-04 so the design must be reviewed; the workflow verifies the exemption and its scope. Do not add a Private Endpoint until Container Apps has a verified VNet and private-DNS path.
