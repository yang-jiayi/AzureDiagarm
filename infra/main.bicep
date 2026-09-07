// main.bicep — subscription-scoped entry point for azd
// Provisions the resource group and all resources for the
// Microsoft Product Architecture Diagram Builder.
targetScope = 'subscription'

// ── Environment ────────────────────────────────────────────────────────────────
@minLength(1)
@maxLength(64)
@description('Name of the azd environment (used to derive resource names).')
param environmentName string

@minLength(1)
@description('Primary Azure region for all resources.')
param location string

// ── Managed GPT-6 Astra account ───────────────────────────────────────────────
@description('Your Azure OpenAI endpoint URL.')
param azureOpenAiEndpoint string = ''

@description('Full resource ID of the Azure OpenAI account for managed-identity RBAC.')
param azureOpenAiResourceId string = ''

@description('Azure Communication Services endpoint used to deliver feedback email.')
param feedbackEmailEndpoint string = ''

@description('Verified Azure Communication Services sender address.')
param feedbackEmailSender string = ''

@description('Recipient address for feedback submissions.')
param feedbackEmailRecipient string = ''

@description('Allow explicitly consented follow-up contact through feedback email delivery.')
param feedbackContactEnabled bool = false

@description('Optional Azure Table Storage endpoint for feedback and shared rate limiting. Diagram Storage is used when this is empty.')
param azureTablesEndpoint string = ''

@description('Azure Table Storage table name for feedback.')
param azureTablesFeedbackTable string = 'feedback'

@description('Existing Front Door identifier; public application ingress must be restricted to that Front Door.')
param frontDoorId string = ''
@description('HTTPS public application origin, without a path, query, or fragment.')
param publicAppUrl string = ''
@description('Administrator of the existing Easy Auth-backed application access list.')
param accessAdminEmail string = ''
param accessKeyVaultResourceId string = ''
param accessTablesEndpoint string = ''
@allowed(['cosmos', 'table'])
param aiBudgetStore string = 'table'
param aiBudgetTablesEndpoint string = ''
param aiBudgetTable string = 'aibudgets'

@description('Set true only after verifying preconfigured single-tenant Easy Auth and Front Door origin isolation.')
param easyAuthVerified bool = false
@minValue(1)
param aiDailyTokenBudget int = 250000
@minValue(1)
param aiMaxConcurrentRequests int = 2
@minValue(1)
param feedbackRetentionDays int = 30
@description('Explicitly approve retention cleanup of existing pre-expiry feedback after reviewing its impact.')
param feedbackLegacyRetentionEnabled bool = false

@description('Actual GPT-6 Astra deployment name. Configure only after provisioning the genuine gpt-6-astra model.')
param openAiDeploymentGpt6Astra string = ''

@description('Explicit administrator opt-in for user-key Azure OpenAI or official OpenAI connections. Does not authorize additional managed models.')
param allowByoAIEndpoints bool = false

// ── Avatar presenter (Speech) ──────────────────────────────────────────────────
@description('Provision an Azure Speech resource for the avatar presenter feature.')
param deploySpeech bool = true

@description('Azure region for the Speech resource (must support Avatar API: westus2, eastus2, etc.).')
param speechRegion string = 'westus2'

// ── Diagram persistence (Cosmos DB) ───────────────────────────────────────────
@description('Provision an Azure Cosmos DB account for saving diagrams across sessions.')
param deployCosmos bool = false

@description('Provision low-cost Azure Blob Storage for authenticated diagram autosave, versions, comments, and share links.')
param deployDiagramStorage bool = true

@description('Azure region for zone-redundant diagram storage. This may differ from the app region when ZRS is unavailable there.')
param diagramStorageLocation string = 'westus2'

// ── MCP server (decoupled Container App) ──────────────────────────────────────
@secure()
@description('Optional bearer token required on the MCP /mcp endpoint. Empty keeps MCP external ingress disabled.')
param mcpAuthToken string = ''

// ── Internals ──────────────────────────────────────────────────────────────────
var abbrs = loadJsonContent('./abbreviations.json')
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = { 'azd-env-name': environmentName }

// ── Resource group ─────────────────────────────────────────────────────────────
resource rg 'Microsoft.Resources/resourceGroups@2022-09-01' = {
  name: 'rg-${environmentName}'
  location: location
  tags: tags
}

// ── All resources ──────────────────────────────────────────────────────────────
module resources './resources.bicep' = {
  name: 'resources'
  scope: rg
  params: {
    location: location
    tags: tags
    abbrs: abbrs
    resourceToken: resourceToken
    deploySpeech: deploySpeech
    speechRegion: speechRegion
    deployCosmos: deployCosmos
    deployDiagramStorage: deployDiagramStorage
    diagramStorageLocation: diagramStorageLocation
    mcpAuthToken: mcpAuthToken
    azureOpenAiEndpoint: azureOpenAiEndpoint
    azureOpenAiDeploymentGpt6Astra: openAiDeploymentGpt6Astra
    allowByoAIEndpoints: allowByoAIEndpoints
    feedbackEmailEndpoint: feedbackEmailEndpoint
    feedbackEmailSender: feedbackEmailSender
    feedbackEmailRecipient: feedbackEmailRecipient
    feedbackContactEnabled: feedbackContactEnabled
    azureTablesEndpoint: azureTablesEndpoint
    azureTablesFeedbackTable: azureTablesFeedbackTable
    frontDoorId: frontDoorId
    publicAppUrl: publicAppUrl
    accessAdminEmail: accessAdminEmail
    accessKeyVaultResourceId: accessKeyVaultResourceId
    accessTablesEndpoint: accessTablesEndpoint
    aiBudgetStore: aiBudgetStore
    aiBudgetTablesEndpoint: aiBudgetTablesEndpoint
    aiBudgetTable: aiBudgetTable
    easyAuthVerified: easyAuthVerified
    aiDailyTokenBudget: aiDailyTokenBudget
    aiMaxConcurrentRequests: aiMaxConcurrentRequests
    feedbackRetentionDays: feedbackRetentionDays
    feedbackLegacyRetentionEnabled: feedbackLegacyRetentionEnabled
  }
}

var openAiResourceParts = split(azureOpenAiResourceId, '/')
module openAiRole './openai-role.bicep' = if (!empty(azureOpenAiResourceId)) {
  name: 'openai-role'
  scope: resourceGroup(openAiResourceParts[2], openAiResourceParts[4])
  params: {
    accountName: openAiResourceParts[8]
    principalId: resources.outputs.appIdentityPrincipalId
  }
}

// ── Outputs captured by azd ────────────────────────────────────────────────────
output AZURE_LOCATION string = location
output AZURE_TENANT_ID string = tenant().tenantId
output AZURE_RESOURCE_GROUP string = rg.name

// Container registry — azd uses this to push the built image
output AZURE_CONTAINER_REGISTRY_ENDPOINT string = resources.outputs.registryLoginServer
output AZURE_CONTAINER_REGISTRY_NAME string = resources.outputs.registryName

// Container app — azd locates it by the azd-service-name tag, but the name is
// also emitted here for reference.
output SERVICE_APP_NAME string = resources.outputs.containerAppName
output SERVICE_APP_IDENTITY_PRINCIPAL_ID string = resources.outputs.appIdentityPrincipalId

// App URL
output SERVICE_APP_URL string = 'https://${resources.outputs.containerAppFqdn}'

// MCP server (decoupled) — azd locates it by the azd-service-name: mcp tag.
output SERVICE_MCP_NAME string = resources.outputs.mcpAppName
output SERVICE_MCP_URL string = 'https://${resources.outputs.mcpAppFqdn}'
output MCP_ENDPOINT string = 'https://${resources.outputs.mcpAppFqdn}/mcp'

// Azure OpenAI — non-secret values available for explicitly configured builds.
output AZURE_OPENAI_ENDPOINT string = azureOpenAiEndpoint
output AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA string = openAiDeploymentGpt6Astra
output AZURE_OPENAI_RESOURCE_ID string = azureOpenAiResourceId
output AZURE_OPENAI_ALLOWED_DEPLOYMENTS string = openAiDeploymentGpt6Astra
output ALLOW_BYO_AI_ENDPOINTS string = string(allowByoAIEndpoints)

// Speech
output AZURE_SPEECH_REGION string = resources.outputs.speechRegionOut
output AZURE_SPEECH_RESOURCE_ID string = resources.outputs.speechResourceId

// Cosmos DB (empty strings when deployCosmos = false)
output AZURE_COSMOS_ENDPOINT string = resources.outputs.cosmosEndpoint
output COSMOS_DATABASE_ID string = resources.outputs.cosmosDatabaseId
output COSMOS_CONTAINER_ID string = resources.outputs.cosmosContainerId
output COSMOS_FEEDBACK_CONTAINER_ID string = resources.outputs.cosmosFeedbackContainerId

// Authenticated diagram persistence (empty strings when deployDiagramStorage = false)
output AZURE_BLOB_ENDPOINT string = resources.outputs.diagramStorageEndpoint
output AZURE_BLOB_DIAGRAMS_CONTAINER string = resources.outputs.diagramStorageContainer
output AZURE_TABLES_ENDPOINT string = resources.outputs.tableStorageEndpoint

// App Insights — available to explicitly configured builds.
output APPLICATIONINSIGHTS_CONNECTION_STRING string = resources.outputs.appInsightsConnectionString
