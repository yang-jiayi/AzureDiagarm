# AzureDiagarm application access and authentication

> Last updated: July 16, 2026

## Production configuration

| Item | Value |
|---|---|
| Application URL | `https://azurediagarm.mssql.biz` |
| Azure subscription | `f2c0fe9a-0171-42ed-803d-3e78322545a1` |
| Resource group | `AzureDiagarm_rg` |
| Container App | `azurediagarm-app` |
| Microsoft Entra tenant | `376417b8-4dea-4ba0-b980-ac5323856cbd` |
| Entra application | `AzureDiagarm-Production` |
| Application (client) ID | `5cd8361b-e235-493b-95a2-c2e8f444c3a2` |
| Permanent administrator | `yangjiayi@msft.jp` |
| Access store | `azurediagarm-access-kv` |
| Speech resource | `azurediagarmspeech` (`westus2`, S0) |
| Easy Auth action | `RedirectToLoginPage` |
| Whitelist enforcement | Enabled |

## Security model

1. Azure Front Door and WAF are the public entry point.
2. Azure Container Apps Easy Auth requires Microsoft Entra ID authentication.
3. The Entra application is single-tenant (`AzureADMyOrg`).
4. Nginx sends an internal authorization subrequest for every page, static asset,
   API, and MCP request.
5. The Node server reads the trusted individual principal headers injected by
   Container Apps, with `X-MS-CLIENT-PRINCIPAL` claim parsing as a fallback.
6. `yangjiayi@msft.jp` is always allowed and is the only account that receives
   access-list management APIs and UI.
7. Other users are allowed only when their normalized email address is active in
   the application whitelist.

External requests cannot set the `X-MS-CLIENT-PRINCIPAL-*` headers. Container
Apps removes external values and injects the authenticated principal headers.

## Front Door requirements

- `forwardProxy.convention` is `Standard`, so Easy Auth uses the
  `X-Forwarded-Host` value and returns users to `azurediagarm.mssql.biz`.
- The registered callback is:
  `https://azurediagarm.mssql.biz/.auth/login/aad/callback`.
- Front Door caching is disabled on every route. Authentication responses and
  protected static assets must never be stored in a shared cache.
- The origin response timeout is 240 seconds for GPT-5.6 Sol generation.
- `/healthz` is the only unauthenticated application path and is used by the
  Front Door origin health probe.
- Container Apps ingress allows only the current IPv4 ranges from the
  `AzureFrontDoor.Backend` service tag and Azure platform infrastructure
  addresses. Nginx also requires this profile's `X-Azure-FDID`, so another
  Front Door profile cannot use the shared backend address space to bypass WAF.
- The deployment workflow refreshes the origin allowlist from Azure service
  tags and verifies that direct origin access still returns HTTP 403.
- The GitHub OIDC service principal `azurediagarm-github-deploy` has the custom
  `AzureDiagarm Service Tag Reader` role at the target subscription. That role
  grants only `Microsoft.Resources/subscriptions/locations/read` and
  `Microsoft.Network/locations/serviceTags/read`, which are required to refresh
  the allowlist without granting subscription-wide Reader access.
- Front Door access, health-probe, WAF, and metric diagnostics are sent to the
  `azurediagarm-logs` Log Analytics workspace.

## Easy Auth CSRF and referrer policy

Keep the response header `Referrer-Policy: same-origin`.

Container Apps Easy Auth performs CSRF validation for unsafe,
cookie-authenticated requests such as `POST /api/openai`. A stricter
`no-referrer` policy removes the same-origin `Referer` header and causes Easy
Auth to reject the request before it reaches Nginx or the OpenAI proxy. The
corresponding platform log contains HTTP 403, substatus 60, and
`Cross-site request forgery detected`.

`same-origin` preserves the referrer required for same-origin API calls while
still withholding it from external sites.

## Whitelist persistence

The whitelist uses Azure Resource Manager child resources under the dedicated
Key Vault:

```text
/subscriptions/f2c0fe9a-0171-42ed-803d-3e78322545a1
  /resourceGroups/AzureDiagarm_rg
  /providers/Microsoft.KeyVault/vaults/azurediagarm-access-kv/secrets/<email-hash>
```

The Key Vault has Public Network access disabled. The app does not use the Key
Vault data endpoint. It reads and writes only the secret resource metadata
through `management.azure.com`, which remains available through the Azure
control plane.

The managed identity has the custom role
`AzureDiagarm Access Whitelist Manager` at this vault only. The role grants:

- `Microsoft.KeyVault/vaults/read`
- `Microsoft.KeyVault/vaults/secrets/read`
- `Microsoft.KeyVault/vaults/secrets/write`

It does not grant access to secret values in the Key Vault data plane, other
vaults, or other Azure resources.

The Avatar Presenter uses keyless authentication to the dedicated Speech
resource. The Container Apps managed identity has only the **Cognitive Services
Speech User** role on that resource, and local key authentication is disabled.

## Administrator workflow

1. Sign in to `https://azurediagarm.mssql.biz` as `yangjiayi@msft.jp`.
2. Select **Access** / **アクセス管理** in the header.
3. Enter an email address and select **Add** / **追加**.
4. To revoke access, select the delete button next to the email address.

The permanent administrator cannot be removed. The management button is not
rendered for other users, and the server independently rejects non-admin API
requests with HTTP 403.

## Sign out

Use:

```text
https://azurediagarm.mssql.biz/.auth/logout?post_logout_redirect_uri=%2F
```

## Operational checks

Inspect the authentication policy without exposing credentials:

```powershell
az containerapp auth show `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --resource-group AzureDiagarm_rg `
  --name azurediagarm-app `
  --query "{platform:platform,globalValidation:globalValidation,httpSettings:httpSettings,aad:identityProviders.azureActiveDirectory}"
```

Inspect the access-store role:

```powershell
az role assignment list `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --assignee-object-id 5ca2361f-57c4-4840-82c6-8d71fb41c00f `
  --scope /subscriptions/f2c0fe9a-0171-42ed-803d-3e78322545a1/resourceGroups/AzureDiagarm_rg/providers/Microsoft.KeyVault/vaults/azurediagarm-access-kv `
  --output table
```

Inspect origin isolation and Front Door diagnostics:

```powershell
az role assignment list `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --assignee-object-id 5ef2bf3e-436a-46c4-929a-b31c50ef2881 `
  --role "AzureDiagarm Service Tag Reader" `
  --scope /subscriptions/f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --output table

az containerapp ingress access-restriction list `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --resource-group AzureDiagarm_rg `
  --name azurediagarm-app `
  --output table

$frontDoorId = az afd profile show `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --resource-group AzureDiagarm_rg `
  --profile-name azurediagarm-fd `
  --query id `
  --output tsv

az monitor diagnostic-settings show `
  --name azurediagarm-local-logs `
  --resource $frontDoorId
```

Useful Log Analytics queries:

```kusto
AzureDiagnostics
| where TimeGenerated > ago(24h)
| where ResourceProvider == "MICROSOFT.CDN"
| where Category in ("FrontDoorAccessLog", "FrontDoorWebApplicationFirewallLog")
| extend RequestUri = tostring(column_ifexists("requestUri_s", "")),
         Status = tostring(column_ifexists("httpStatusCode_s", "")),
         Action = tostring(column_ifexists("action_s", "")),
         Details = tostring(column_ifexists("details_msg_s", ""))
| project TimeGenerated, Category, RequestUri, Status, Action, Details
| order by TimeGenerated desc
```

```kusto
union isfuzzy=true ContainerAppSystemLogs_CL, ContainerAppConsoleLogs_CL
| where TimeGenerated > ago(24h)
| where Log_s has "Cross-site request forgery" or Log_s has "SubStatus: 60"
| project TimeGenerated, ContainerAppName_s, RevisionName_s, Log_s
| order by TimeGenerated desc
```

## OpenAI proxy diagnostics

`/api/openai` returns a structured error envelope with a safe `code`,
`requestId`, and, when available, `upstreamStatus`, `upstreamCode`, and
`upstreamRequestId`. The proxy logs request identifiers and status metadata,
but never logs prompts, access tokens, or upstream response bodies.

Common codes include:

- `credential_acquisition_failed`
- `azure_openai_authentication_failed`
- `azure_openai_rate_limited`
- `azure_openai_timeout`
- `azure_openai_unavailable`
- `application_authentication_required`
- `application_request_rejected`
- `edge_request_blocked`

Use the displayed request ID to correlate browser errors with Container Apps
logs. Do not treat every HTTP 401 or 403 as a managed identity failure: Easy
Auth, the application whitelist, WAF, and Azure OpenAI RBAC are separate
enforcement layers.

## Credential rotation

The Easy Auth client credential named
`AzureDiagarm Container Apps Easy Auth` expires on July 16, 2028.

Rotate it before expiry:

1. Append a new Entra application credential.
2. Immediately pass the new value to
   `az containerapp auth microsoft update --client-secret`.
3. Confirm sign-in through the custom domain.
4. Delete the old credential by key ID.

Never print, email, or commit the credential value.

## Emergency recovery

If an authentication configuration error prevents administrator access:

```powershell
az containerapp auth update `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --resource-group AzureDiagarm_rg `
  --name azurediagarm-app `
  --action AllowAnonymous `
  --yes

az containerapp update `
  --subscription f2c0fe9a-0171-42ed-803d-3e78322545a1 `
  --resource-group AzureDiagarm_rg `
  --name azurediagarm-app `
  --set-env-vars ACCESS_CONTROL_ENABLED=false `
  --output none
```

Restore `RedirectToLoginPage` and `ACCESS_CONTROL_ENABLED=true` immediately
after correcting the configuration.
