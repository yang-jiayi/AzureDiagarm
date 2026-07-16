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

## Security model

1. Azure Front Door and WAF are the public entry point.
2. Azure Container Apps Easy Auth requires Microsoft Entra ID authentication.
3. The Entra application is single-tenant (`AzureADMyOrg`).
4. Nginx sends an internal authorization subrequest for every page, static asset,
   API, and MCP request.
5. The Node server reads the trusted `X-MS-CLIENT-PRINCIPAL-NAME` and
   `X-MS-CLIENT-PRINCIPAL-ID` headers injected by Container Apps.
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
