// resources.bicep — all resources deployed into the resource group
targetScope = 'resourceGroup'

param location string
param tags object
param abbrs object
param resourceToken string

// Feature flags
param deploySpeech bool
param speechRegion string
param deployCosmos bool
param deployDiagramStorage bool
param diagramStorageLocation string

@secure()
@description('Optional bearer token required on the decoupled MCP server /mcp endpoint. Empty keeps external ingress disabled.')
param mcpAuthToken string = ''

// Azure OpenAI (passed through to container app env; not provisioned here)
param azureOpenAiEndpoint string
param azureOpenAiAllowedDeployments string = ''
@secure()
param azureOpenAiApiKey string
param azureFoundryEndpoint string = ''
param azureFoundryAllowedDeployments string = ''
@secure()
param azureFoundryApiKey string = ''
param feedbackEmailEndpoint string = ''
param feedbackEmailSender string = ''
param feedbackEmailRecipient string = ''
param azureTablesEndpoint string = ''
param azureTablesFeedbackTable string = 'feedback'
param frontDoorId string = ''

// ── Log Analytics ──────────────────────────────────────────────────────────────
resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: '${abbrs.operationalInsightsWorkspaces}${resourceToken}'
  location: location
  tags: tags
  properties: {
    sku: { name: 'PerGB2018' }
    retentionInDays: 30
  }
}

// ── Application Insights ───────────────────────────────────────────────────────
resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${abbrs.insightsComponents}${resourceToken}'
  location: location
  tags: tags
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalytics.id
  }
}

// ── Container Registry ─────────────────────────────────────────────────────────
resource acr 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  name: '${abbrs.containerRegistryRegistries}${resourceToken}'
  location: location
  tags: tags
  sku: { name: 'Basic' }
  properties: {
    adminUserEnabled: false
  }
}

// ── User-assigned Managed Identity ────────────────────────────────────────────
resource appIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${abbrs.managedIdentityUserAssignedIdentities}app-${resourceToken}'
  location: location
  tags: tags
}

// ── AcrPull role → managed identity ──────────────────────────────────────────
resource acrPullRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(acr.id, appIdentity.id, 'acrpull')
  scope: acr
  properties: {
    // AcrPull built-in role
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '7f951dda-4ed3-4680-a7ca-43fe172d538d'
    )
    principalId: appIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Container Apps Environment ────────────────────────────────────────────────
resource caEnv 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: '${abbrs.appManagedEnvironments}${resourceToken}'
  location: location
  tags: tags
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalytics.properties.customerId
        sharedKey: logAnalytics.listKeys().primarySharedKey
      }
    }
  }
}

// ── Speech (optional) ─────────────────────────────────────────────────────────
resource speech 'Microsoft.CognitiveServices/accounts@2023-05-01' = if (deploySpeech) {
  name: '${abbrs.cognitiveServicesSpeech}${resourceToken}'
  // Speech Avatar API is only available in select regions; use the caller-supplied region.
  location: speechRegion
  tags: tags
  kind: 'SpeechServices'
  sku: { name: 'S0' }
  properties: {
    customSubDomainName: '${abbrs.cognitiveServicesSpeech}${resourceToken}'
    disableLocalAuth: true
    publicNetworkAccess: 'Enabled'
  }
}

// Cognitive Services Speech User role → managed identity
resource speechUserRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (deploySpeech) {
  name: guid(deploySpeech ? speech.id : resourceToken, appIdentity.id, 'speechuser')
  scope: speech
  properties: {
    // Cognitive Services Speech User built-in role
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'f2dc8367-1007-4938-bd23-fe263f013447'
    )
    principalId: appIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Cosmos DB (optional) ──────────────────────────────────────────────────────
var cosmosDatabaseId = 'diagrams'
var cosmosContainerId = 'diagrams'
var cosmosFeedbackContainerId = 'feedback'

resource cosmos 'Microsoft.DocumentDB/databaseAccounts@2024-02-15-preview' = if (deployCosmos) {
  name: '${abbrs.documentDBDatabaseAccounts}${resourceToken}'
  location: location
  tags: tags
  kind: 'GlobalDocumentDB'
  properties: {
    databaseAccountOfferType: 'Standard'
    consistencyPolicy: { defaultConsistencyLevel: 'Session' }
    locations: [
      { locationName: location, failoverPriority: 0, isZoneRedundant: false }
    ]
    enableFreeTier: true
  }
}

resource cosmosDb 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases@2024-02-15-preview' = if (deployCosmos) {
  parent: cosmos
  name: cosmosDatabaseId
  properties: {
    resource: { id: cosmosDatabaseId }
  }
}

resource cosmosContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2024-02-15-preview' = if (deployCosmos) {
  parent: cosmosDb
  name: cosmosContainerId
  properties: {
    resource: {
      id: cosmosContainerId
      partitionKey: { paths: ['/id'], kind: 'Hash' }
    }
  }
}

// Dedicated container for in-app user feedback (append-only, low read volume).
resource cosmosFeedbackContainer 'Microsoft.DocumentDB/databaseAccounts/sqlDatabases/containers@2024-02-15-preview' = if (deployCosmos) {
  parent: cosmosDb
  name: cosmosFeedbackContainerId
  properties: {
    resource: {
      id: cosmosFeedbackContainerId
      partitionKey: { paths: ['/id'], kind: 'Hash' }
    }
  }
}

// Cosmos DB Built-in Data Contributor → managed identity
resource cosmosRole 'Microsoft.DocumentDB/databaseAccounts/sqlRoleAssignments@2024-02-15-preview' = if (deployCosmos) {
  parent: cosmos
  name: guid(deployCosmos ? cosmos.id : resourceToken, appIdentity.id, 'cosmoscontributor')
  properties: {
    roleDefinitionId: '${deployCosmos ? cosmos.id : ''}/sqlRoleDefinitions/00000000-0000-0000-0000-000000000002'
    principalId: appIdentity.properties.principalId
    scope: deployCosmos ? '${cosmos.id}/dbs/${cosmosDatabaseId}/colls/${cosmosFeedbackContainerId}' : ''
  }
}

// ── Authenticated persistence and shared counters (Blob/Table Storage) ────────
var diagramStorageContainerName = 'diagrams'
var diagramRateLimitTableName = 'ratelimit'
var diagramStoragePerimeterName = 'azurediagarm-storage-perimeter'
var diagramStoragePerimeterProfileName = 'diagram-storage-profile'

// The tenant governance policy disables ordinary public storage endpoints.
// NSP keeps the HTTPS endpoint usable by authenticated resources in this
// subscription while enforcing a network boundary ahead of Storage RBAC.
resource diagramStoragePerimeter 'Microsoft.Network/networkSecurityPerimeters@2024-07-01' = if (deployDiagramStorage) {
  name: diagramStoragePerimeterName
  location: diagramStorageLocation
  tags: tags
  properties: {}
}

resource diagramStoragePerimeterProfile 'Microsoft.Network/networkSecurityPerimeters/profiles@2024-07-01' = if (deployDiagramStorage) {
  parent: diagramStoragePerimeter
  name: diagramStoragePerimeterProfileName
  properties: {}
}

resource diagramStorageSubscriptionRule 'Microsoft.Network/networkSecurityPerimeters/profiles/accessRules@2024-07-01' = if (deployDiagramStorage) {
  parent: diagramStoragePerimeterProfile
  name: 'allow-azurediagarm-subscription'
  properties: {
    direction: 'Inbound'
    subscriptions: [
      {
        id: subscription().id
      }
    ]
  }
}

resource diagramStorage 'Microsoft.Storage/storageAccounts@2025-06-01' = if (deployDiagramStorage) {
  name: take('stg${resourceToken}', 24)
  location: diagramStorageLocation
  tags: tags
  kind: 'StorageV2'
  sku: {
    name: 'Standard_ZRS'
  }
  properties: {
    accessTier: 'Hot'
    allowBlobPublicAccess: false
    allowCrossTenantReplication: false
    allowSharedKeyAccess: false
    defaultToOAuthAuthentication: true
    minimumTlsVersion: 'TLS1_2'
    publicNetworkAccess: 'SecuredByPerimeter'
    supportsHttpsTrafficOnly: true
  }
}

resource diagramBlobService 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' = if (deployDiagramStorage) {
  parent: diagramStorage
  name: 'default'
  properties: {
    isVersioningEnabled: true
    deleteRetentionPolicy: {
      enabled: true
      days: 30
    }
    containerDeleteRetentionPolicy: {
      enabled: true
      days: 30
    }
  }
}

resource diagramBlobContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = if (deployDiagramStorage) {
  parent: diagramBlobService
  name: diagramStorageContainerName
  properties: {
    publicAccess: 'None'
  }
}

resource diagramTableService 'Microsoft.Storage/storageAccounts/tableServices@2023-05-01' = if (deployDiagramStorage) {
  parent: diagramStorage
  name: 'default'
}

resource diagramRateLimitTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = if (deployDiagramStorage) {
  parent: diagramTableService
  name: diagramRateLimitTableName
}

resource diagramFeedbackTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = if (deployDiagramStorage && azureTablesFeedbackTable != diagramRateLimitTableName) {
  parent: diagramTableService
  name: azureTablesFeedbackTable
}

resource diagramStorageRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (deployDiagramStorage) {
  name: guid(deployDiagramStorage ? diagramStorage.id : resourceToken, appIdentity.id, 'diagram-blob-contributor')
  scope: diagramStorage
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
    )
    principalId: appIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

resource diagramTableStorageRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (deployDiagramStorage) {
  name: guid(deployDiagramStorage ? diagramStorage.id : resourceToken, appIdentity.id, 'diagram-table-contributor')
  scope: diagramStorage
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
    )
    principalId: appIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

resource diagramStoragePerimeterAssociation 'Microsoft.Network/networkSecurityPerimeters/resourceAssociations@2024-07-01' = if (deployDiagramStorage) {
  parent: diagramStoragePerimeter
  name: 'diagram-storage-association'
  properties: {
    accessMode: 'Enforced'
    privateLinkResource: {
      id: diagramStorage.id
    }
    profile: {
      id: diagramStoragePerimeterProfile.id
    }
  }
}

var effectiveAzureTablesEndpoint = !empty(azureTablesEndpoint)
  ? azureTablesEndpoint
  : (deployDiagramStorage ? diagramStorage!.properties.primaryEndpoints.table : '')
var appHealthProbeHttpGet = union(
  {
    path: '/healthz'
    port: 80
    scheme: 'HTTP'
  },
  empty(frontDoorId)
    ? {}
    : {
        httpHeaders: [
          {
            name: 'X-Azure-FDID'
            value: frontDoorId
          }
        ]
      }
)

// ── Container App ─────────────────────────────────────────────────────────────
// azd locates the service by the 'azd-service-name' tag value matching
// the service key in azure.yaml ('app').
resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: '${abbrs.appContainerApps}diagram-builder-${resourceToken}'
  location: location
  tags: union(tags, { 'azd-service-name': 'app' })
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: { '${appIdentity.id}': {} }
  }
  properties: {
    managedEnvironmentId: caEnv.id
    configuration: {
      secrets: concat(
        empty(azureOpenAiApiKey)
          ? []
          : [
              { name: 'azure-openai-api-key', value: azureOpenAiApiKey }
            ],
        empty(azureFoundryApiKey)
          ? []
          : [
              { name: 'azure-foundry-api-key', value: azureFoundryApiKey }
            ]
      )
      ingress: {
        external: true
        targetPort: 80
        transport: 'auto'
      }
      registries: [
        {
          server: acr.properties.loginServer
          identity: appIdentity.id
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'app'
          // Placeholder image replaced by 'azd deploy'
          image: 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'
          resources: { cpu: json('0.5'), memory: '1.0Gi' }
          probes: [
            {
              type: 'Startup'
              httpGet: appHealthProbeHttpGet
              initialDelaySeconds: 1
              periodSeconds: 5
              timeoutSeconds: 3
              failureThreshold: 30
            }
            {
              type: 'Liveness'
              httpGet: appHealthProbeHttpGet
              initialDelaySeconds: 15
              periodSeconds: 15
              timeoutSeconds: 5
              failureThreshold: 3
            }
            {
              type: 'Readiness'
              httpGet: appHealthProbeHttpGet
              initialDelaySeconds: 3
              periodSeconds: 5
              timeoutSeconds: 3
              failureThreshold: 3
              successThreshold: 1
            }
          ]
          env: concat([
            // Identity — lets DefaultAzureCredential pick up the managed identity
            { name: 'AZURE_CLIENT_ID', value: appIdentity.properties.clientId }
            { name: 'AZURE_OPENAI_ENDPOINT', value: azureOpenAiEndpoint }
            { name: 'AZURE_OPENAI_ALLOWED_DEPLOYMENTS', value: azureOpenAiAllowedDeployments }
            { name: 'AZURE_FOUNDRY_ENDPOINT', value: azureFoundryEndpoint }
            { name: 'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS', value: azureFoundryAllowedDeployments }
            { name: 'OPENAI_RATE_LIMIT_PER_HOUR', value: '120' }
            { name: 'FEEDBACK_EMAIL_ENDPOINT', value: feedbackEmailEndpoint }
            { name: 'FEEDBACK_EMAIL_SENDER', value: feedbackEmailSender }
            { name: 'FEEDBACK_EMAIL_RECIPIENT', value: feedbackEmailRecipient }
            { name: 'AZURE_TABLES_ENDPOINT', value: effectiveAzureTablesEndpoint }
            { name: 'AZURE_TABLES_FEEDBACK_TABLE', value: azureTablesFeedbackTable }
            { name: 'AZURE_TABLES_RATE_LIMIT_TABLE', value: diagramRateLimitTableName }
            // Speech
            { name: 'AZURE_SPEECH_REGION', value: deploySpeech ? speech!.location : '' }
            { name: 'AZURE_SPEECH_RESOURCE_ID', value: deploySpeech ? speech!.id : '' }
            // Cosmos DB
            { name: 'AZURE_COSMOS_ENDPOINT', value: deployCosmos ? cosmos!.properties.documentEndpoint : '' }
            { name: 'COSMOS_DATABASE_ID', value: deployCosmos ? cosmosDatabaseId : '' }
            { name: 'COSMOS_CONTAINER_ID', value: deployCosmos ? cosmosContainerId : '' }
            { name: 'COSMOS_FEEDBACK_CONTAINER_ID', value: deployCosmos ? cosmosFeedbackContainerId : '' }
            // Authenticated cloud diagram persistence
            { name: 'AZURE_BLOB_ENDPOINT', value: deployDiagramStorage ? diagramStorage!.properties.primaryEndpoints.blob : '' }
            { name: 'AZURE_BLOB_DIAGRAMS_CONTAINER', value: deployDiagramStorage ? diagramStorageContainerName : '' }
            { name: 'MCP_ENABLED', value: 'true' }
            { name: 'MCP_HTTP_STATELESS', value: 'true' }
            // Public URL (self-referential — set after first deploy if needed)
            {
              name: 'PUBLIC_URL'
              value: 'https://${abbrs.appContainerApps}diagram-builder-${resourceToken}.${caEnv.properties.defaultDomain}'
            }
          ],
          concat(
            empty(azureOpenAiApiKey)
              ? []
              : [
                  { name: 'AZURE_OPENAI_API_KEY', secretRef: 'azure-openai-api-key' }
                ],
            empty(azureFoundryApiKey)
              ? []
              : [
                  { name: 'AZURE_FOUNDRY_API_KEY', secretRef: 'azure-foundry-api-key' }
                ]
          ))
        }
      ]
      scale: {
        minReplicas: 0
        maxReplicas: 3
      }
    }
  }
}

// ── MCP Container App (decoupled, own FQDN) ──────────────────────────────────
// The Azure Architecture Diagram Builder MCP server as its own Container App,
// so agent traffic scales, releases, and fails independently of the web UI.
// azd locates it by the 'azd-service-name: mcp' tag (matches azure.yaml).
var mcpExternalEnabled = !empty(mcpAuthToken)
var mcpBaseEnv = [
  { name: 'AZURE_CLIENT_ID', value: appIdentity.properties.clientId }
  { name: 'MCP_HTTP_HOST', value: mcpExternalEnabled ? '0.0.0.0' : '127.0.0.1' }
  { name: 'MCP_HTTP_PORT', value: '3030' }
  { name: 'MCP_HTTP_PATH', value: '/mcp' }
  { name: 'MCP_SESSION_MAX', value: '100' }
  { name: 'MCP_SESSION_IDLE_SECONDS', value: '1800' }
  { name: 'MCP_SESSION_TTL_SECONDS', value: '7200' }
  { name: 'MCP_SESSION_GC_SECONDS', value: '60' }
]
var mcpAuthEnv = empty(mcpAuthToken)
  ? []
  : [ { name: 'MCP_AUTH_TOKEN', secretRef: 'mcp-auth-token' } ]
var mcpSecrets = empty(mcpAuthToken)
  ? []
  : [ { name: 'mcp-auth-token', value: mcpAuthToken } ]

resource mcpApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: '${abbrs.appContainerApps}diagram-mcp-${resourceToken}'
  location: location
  tags: union(tags, { 'azd-service-name': 'mcp' })
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: { '${appIdentity.id}': {} }
  }
  properties: {
    managedEnvironmentId: caEnv.id
    configuration: {
      secrets: mcpSecrets
      ingress: {
        external: mcpExternalEnabled
        targetPort: 3030
        transport: 'auto'
      }
      registries: [
        {
          server: acr.properties.loginServer
          identity: appIdentity.id
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'mcp'
          // Placeholder image replaced by 'azd deploy mcp'.
          image: 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'
          resources: { cpu: json('0.5'), memory: '1.0Gi' }
          probes: [
            {
              type: 'Startup'
              httpGet: { path: '/healthz', port: 3030, scheme: 'HTTP' }
              initialDelaySeconds: 1
              periodSeconds: 5
              timeoutSeconds: 3
              failureThreshold: 30
            }
            {
              type: 'Liveness'
              httpGet: { path: '/healthz', port: 3030, scheme: 'HTTP' }
              initialDelaySeconds: 10
              periodSeconds: 15
              timeoutSeconds: 5
              failureThreshold: 3
            }
            {
              type: 'Readiness'
              httpGet: { path: '/healthz', port: 3030, scheme: 'HTTP' }
              initialDelaySeconds: 2
              periodSeconds: 5
              timeoutSeconds: 3
              failureThreshold: 3
              successThreshold: 1
            }
          ]
          env: concat(mcpBaseEnv, mcpAuthEnv)
        }
      ]
      scale: {
        minReplicas: 0
        // Streamable HTTP sessions are process-local, so a single replica is
        // required until the transport has a shared session backend.
        maxReplicas: 1
      }
    }
  }
}

// ── Outputs ────────────────────────────────────────────────────────────────────
output registryLoginServer string = acr.properties.loginServer
output registryName string = acr.name
output appIdentityPrincipalId string = appIdentity.properties.principalId
output containerAppName string = containerApp.name
output containerAppFqdn string = containerApp.properties.configuration.ingress.fqdn
output mcpAppName string = mcpApp.name
output mcpAppFqdn string = mcpApp.properties.configuration.ingress.fqdn
output speechRegionOut string = deploySpeech ? speech!.location : ''
output speechResourceId string = deploySpeech ? speech!.id : ''
output cosmosEndpoint string = deployCosmos ? cosmos!.properties.documentEndpoint : ''
output cosmosDatabaseId string = deployCosmos ? cosmosDatabaseId : ''
output cosmosContainerId string = deployCosmos ? cosmosContainerId : ''
output cosmosFeedbackContainerId string = deployCosmos ? cosmosFeedbackContainerId : ''
output diagramStorageEndpoint string = deployDiagramStorage ? diagramStorage!.properties.primaryEndpoints.blob : ''
output diagramStorageContainer string = deployDiagramStorage ? diagramStorageContainerName : ''
output tableStorageEndpoint string = effectiveAzureTablesEndpoint
output appInsightsConnectionString string = appInsights.properties.ConnectionString
