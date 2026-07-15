targetScope = 'resourceGroup'

param accountName string
param principalId string

resource openAi 'Microsoft.CognitiveServices/accounts@2023-05-01' existing = {
  name: accountName
}

resource openAiUserRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(openAi.id, principalId, 'openaiuser')
  scope: openAi
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      '5e0bd9bd-7b93-4f28-af87-19fc36ad61bd'
    )
    principalId: principalId
    principalType: 'ServicePrincipal'
  }
}
