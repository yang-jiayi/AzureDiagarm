# AzureDiagarm production deployment

The production site is deployed to Azure Container Apps and exposed only through Azure Front Door Standard with WAF.

## Cost controls

- Azure Front Door uses the Standard tier, matching the `SQLServerEvo_rg` reference architecture.
- The Container App uses the Consumption workload profile and can scale to zero.
- Images use the existing `sqlserverevoacr` Basic registry instead of creating another paid registry.
- AI generation uses a single pay-as-you-go `gpt-5.4-mini` GlobalStandard deployment with low reasoning effort by default.
- Upstream checks run only when manually dispatched; deployments still run automatically for validated changes pushed to `main`.
- Azure Communication Services sends mail only after a deployment succeeds or fails.

## Update flow

`.github/workflows/azurediagarm-sync-deploy.yml` deploys validated changes pushed to `main`. Upstream synchronization with `Arturo-Quiroga-MSFT/azure-architecture-diagram-builder` now starts only through manual workflow dispatch; after validation it pushes the merge, builds a unique image tag, deploys a new Container Apps revision, purges Front Door, verifies the production URL, and emails the update details.

## Security layers

- The wildcard `*.mssql.biz` certificate is read from `westuskvl` by the Front Door managed identity and tracks the latest Key Vault secret version.
- WAF runs in Prevention mode with rate limiting and known AI crawler `User-Agent` blocking.
- The origin validates `X-Azure-FDID`, preventing direct Container Apps access from bypassing WAF.
- Application responses include anti-indexing headers and `robots.txt`; these controls discourage compliant crawlers while WAF handles known automated clients.
- Azure OpenAI requests are proxied server-side and authorized with the Container App's user-assigned managed identity; no API key is embedded in the browser bundle.
