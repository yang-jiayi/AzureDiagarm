# AzureDiagarm application access and authentication

This public runbook documents the production security model without publishing
tenant IDs, subscription IDs, object IDs, administrator addresses, or private
resource names. Keep environment-specific values in GitHub configuration and
the approved operational inventory.

## Production security model

1. Azure Front Door and WAF are the only public entry point.
2. Azure Container Apps ingress rejects direct-origin requests.
3. Container Apps Easy Auth requires Microsoft Entra ID authentication when
   access control is enabled.
4. Enterprise Application assignment limits access to approved security groups.
5. Guest accounts receive no Azure RBAC roles, Microsoft Fabric permissions, or
   unrelated Enterprise Application assignments.
6. Nginx authorizes page, static asset, API, and MCP requests consistently.
7. The Node server trusts only Easy Auth principal headers injected by the
   platform and parses the aggregate principal claim as a fallback.
8. Production workloads use managed identity instead of committed credentials.

The public source repository does not grant access to the Azure subscription or
the deployed application.

## Configuration inventory

Store the following values outside source control:

| Item | Example placeholder |
|---|---|
| Azure subscription | `<subscription-id>` |
| Resource group | `<resource-group>` |
| Container App | `<container-app-name>` |
| Microsoft Entra tenant | `<tenant-id>` |
| Entra application client ID | `<application-client-id>` |
| Guest access group object ID | `<guest-group-object-id>` |
| Administrator group object ID | `<administrator-group-object-id>` |
| Deployment managed identity object ID | `<deployment-principal-object-id>` |
| Runtime managed identity object ID | `<runtime-principal-object-id>` |
| Access Key Vault resource ID | `<access-key-vault-resource-id>` |
| Front Door profile | `<front-door-profile>` |
| Log Analytics workspace | `<log-analytics-workspace>` |

Non-secret deployment values belong in GitHub Actions variables. Credentials,
tokens, certificates, and private keys belong in GitHub Actions secrets or an
Azure-managed identity flow.

## Front Door requirements

- Use `forwardProxy.convention=Standard` so Easy Auth honors
  `X-Forwarded-Host`.
- Register the custom-domain Easy Auth callback.
- Disable Front Door caching for authenticated pages, static assets, and APIs.
- Permit Container Apps ingress only from the current
  `AzureFrontDoor.Backend` service-tag ranges and Azure platform probes.
- Require the expected `X-Azure-FDID` value at Nginx.
- Keep `/healthz` limited to health checks and return no sensitive data.
- Send Front Door access, WAF, health, and metric diagnostics to Log Analytics.
- Verify that direct and spoofed-origin requests return HTTP 403 after every
  deployment.

The deployment workflow refreshes the origin allowlist, verifies TLS and WAF
settings, deploys a new revision, checks its probes, and rolls back on failure.

## Easy Auth and browser security

Keep `Referrer-Policy: same-origin`. Easy Auth performs CSRF validation for
unsafe cookie-authenticated requests, and a stricter `no-referrer` policy can
remove the same-origin `Referer` header required by the platform.

Also retain:

- a restrictive Content Security Policy;
- `X-Content-Type-Options: nosniff`;
- a least-privilege Permissions Policy;
- HTTPS-only redirects and HSTS; and
- private/no-store caching for authenticated responses.

## Entra group access

Use separate security groups for application users and administrators.

1. Invite an external user through Microsoft Entra ID.
2. Add the user to the application access group.
3. Assign the access and administrator groups to the Enterprise Application.
4. Keep **Assignment required** enabled.
5. Do not assign users directly unless an approved exception requires it.
6. Remove group membership to revoke access.
7. Do not grant the application access group Azure RBAC, Fabric workspace,
   licenses, or unrelated application assignments.

If Conditional Access limits guest access to Azure management or Fabric, exclude
the administrator group through the approved identity-governance process. Keep
the actual group and user object IDs in the restricted operational inventory.

The following Microsoft application IDs are public platform identifiers, not
tenant credentials:

- Microsoft Azure Management: `797f4846-ba00-4fd7-ba43-dac1f8f63013`
- Power BI Service: `00000009-0000-0000-c000-000000000000`

## Legacy access store

The optional legacy email access store uses Azure Resource Manager metadata
under a dedicated Key Vault:

```text
/subscriptions/<subscription-id>
  /resourceGroups/<resource-group>
  /providers/Microsoft.KeyVault/vaults/<access-key-vault>/secrets/<email-hash>
```

The runtime identity should receive only the custom metadata permissions needed
to read and update application-owned secret resources. It must not receive Key
Vault data-plane access to secret values or broad subscription roles.

Disable the store with `ACCESS_CONTROL_ENABLED=false` when Entra assignment is
the authoritative access model.

## Operational checks

Set local shell variables from the restricted inventory before running checks:

```powershell
$subscriptionId = "<subscription-id>"
$resourceGroup = "<resource-group>"
$containerApp = "<container-app-name>"
$accessVaultScope = "<access-key-vault-resource-id>"
$runtimePrincipalId = "<runtime-principal-object-id>"
$deploymentPrincipalId = "<deployment-principal-object-id>"
$frontDoorProfile = "<front-door-profile>"
```

Inspect authentication without exposing credential values:

```powershell
az containerapp auth show `
  --subscription $subscriptionId `
  --resource-group $resourceGroup `
  --name $containerApp `
  --query "{platform:platform,globalValidation:globalValidation,httpSettings:httpSettings,aad:identityProviders.azureActiveDirectory}"
```

Inspect least-privilege role assignments:

```powershell
az role assignment list `
  --subscription $subscriptionId `
  --assignee-object-id $runtimePrincipalId `
  --scope $accessVaultScope `
  --output table

az role assignment list `
  --subscription $subscriptionId `
  --assignee-object-id $deploymentPrincipalId `
  --scope "/subscriptions/$subscriptionId" `
  --output table
```

Inspect origin isolation:

```powershell
az containerapp ingress access-restriction list `
  --subscription $subscriptionId `
  --resource-group $resourceGroup `
  --name $containerApp `
  --output table
```

Never print secret values, access tokens, Easy Auth client credentials, or
private feedback content while collecting diagnostics.

## Credential rotation

Track Easy Auth and other credential expiration dates in the restricted
operational inventory. Rotate before expiry:

1. Create the replacement credential.
2. Update the Container App authentication configuration.
3. Confirm sign-in through the custom domain.
4. Delete the previous credential.

Prefer workload identity federation and managed identity wherever the platform
supports them.

## Emergency recovery

If an authentication configuration error blocks administrators, use an
authenticated Azure operator session and the placeholders above:

```powershell
az containerapp auth update `
  --subscription $subscriptionId `
  --resource-group $resourceGroup `
  --name $containerApp `
  --action AllowAnonymous `
  --yes

az containerapp update `
  --subscription $subscriptionId `
  --resource-group $resourceGroup `
  --name $containerApp `
  --set-env-vars ACCESS_CONTROL_ENABLED=false `
  --output none
```

Restrict network access before using this break-glass path. Restore
`RedirectToLoginPage` immediately after correcting the configuration, then
re-run the production security verification workflow.
