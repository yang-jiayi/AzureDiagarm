# Import from Azure

`Import from Azure` signs an approved AzureDiagarm user in with Microsoft Entra ID, queries Azure Resource Graph with that user's delegated token, and maps a selected Resource Group into a diagram.

## Security model

- **Read-only delegated access.** The browser requests the Azure Service Management `user_impersonation` scope. Azure RBAC still determines which subscriptions and Resource Groups the user can read.
- **No production server identity exposure.** `AZURE_IMPORT_ENABLED` remains unset in the hosted application, so the Container App managed identity cannot be used through the import API.
- **Browser-session token cache.** MSAL stores tokens in `sessionStorage`; tokens are sent directly to `management.azure.com`, not to the AzureDiagarm server.
- **Private application access.** Container Apps Easy Auth and the Key Vault-backed application whitelist continue to protect the complete site.
- **No imported Azure inventory is persisted by the server.** Resource data is held in browser memory while the diagram is built.

## Production registration

| Setting | Value |
|---|---|
| Application | `AzureDiagarm-Production` |
| Client ID | `5cd8361b-e235-493b-95a2-c2e8f444c3a2` |
| Tenant ID | `376417b8-4dea-4ba0-b980-ac5323856cbd` |
| Authority | `https://login.microsoftonline.com/376417b8-4dea-4ba0-b980-ac5323856cbd` |
| SPA redirect URI | `https://azurediagarm.mssql.biz` |
| ARM scope | `https://management.azure.com/user_impersonation` |

The app registration must include the delegated Azure Service Management `user_impersonation` permission. Grant tenant admin consent if tenant policy does not allow the approved user to consent during first sign-in.

## Using the feature

1. Open `https://azurediagarm.mssql.biz` and complete the existing application sign-in.
2. Select **Import from Azure**.
3. Select **Sign in to Azure** and approve the read-only delegated permission if prompted.
4. Choose a Subscription and Resource Group.
5. Select **Import resource group**.

The diagram maps supported top-level resources deterministically. Unsupported Azure resource types are reported through the import coverage summary rather than sent to an AI model.

## Local or self-hosted server mode

When no `VITE_AZURE_AD_CLIENT_ID` is built into the client, the feature can use the co-located server identity. This mode is disabled by default and should only be enabled in an isolated environment:

```text
AZURE_IMPORT_ENABLED=true
```

Grant that identity only Reader access to the minimum required scope. Do not enable server mode in the hosted AzureDiagarm production deployment.
