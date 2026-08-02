targetScope = 'resourceGroup'

param accountName string
param principalId string

resource foundry 'Microsoft.CognitiveServices/accounts@2023-05-01' existing = {
  name: accountName
}

resource cognitiveServicesUserRole 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(foundry.id, principalId, 'cognitiveservicesuser')
  scope: foundry
  properties: {
    roleDefinitionId: subscriptionResourceId(
      'Microsoft.Authorization/roleDefinitions',
      'a97b65f3-24c7-4388-baec-2e87135dc908'
    )
    principalId: principalId
    principalType: 'ServicePrincipal'
  }
}
