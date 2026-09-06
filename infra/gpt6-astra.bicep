targetScope = 'resourceGroup'

@minLength(2)
@description('Name of the existing Azure OpenAI account. The account, network controls, and role assignments are not changed.')
param openAiAccountName string

@minValue(1)
@description('GlobalStandard capacity units. Confirm regional capacity and subscription quota before changing this value.')
param capacity int = 50

resource account 'Microsoft.CognitiveServices/accounts@2025-06-01' existing = {
  name: openAiAccountName
}

resource astra 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: account
  name: 'gpt-6-astra'
  sku: {
    name: 'GlobalStandard'
    capacity: capacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'gpt-6-astra'
      version: '2026-09-03'
    }
    raiPolicyName: 'Microsoft.DefaultV2'
    versionUpgradeOption: 'OnceNewDefaultVersionAvailable'
  }
}

output deploymentName string = astra.name
