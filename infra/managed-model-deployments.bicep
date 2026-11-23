// Explicit opt-in template for the two managed deployments.
// Apply only after confirming model availability, regional capacity, quota,
// pricing, approval and the impact of creating these new deployments.
// This template does not update, rename, or delete existing deployments.
targetScope = 'resourceGroup'

@minLength(2)
@description('Existing Azure OpenAI account. The account itself is not changed.')
param openAiAccountName string

@minValue(1)
@description('GlobalStandard capacity units for GPT-6.2 Astra. Confirm quota and capacity before deployment.')
param gpt62AstraCapacity int

@minValue(1)
@description('GlobalStandard capacity units for GPT-6.2 Sol. Confirm quota and capacity before deployment.')
param gpt62SolCapacity int

resource openAiAccount 'Microsoft.CognitiveServices/accounts@2025-06-01' existing = {
  name: openAiAccountName
}

resource gpt62Astra 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: openAiAccount
  name: 'gpt-6.2-astra'
  sku: {
    name: 'GlobalStandard'
    capacity: gpt62AstraCapacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'gpt-6.2-astra'
      version: '2026-11-17'
    }
    raiPolicyName: 'Microsoft.DefaultV2'
    versionUpgradeOption: 'NoAutoUpgrade'
  }
}

resource gpt62Sol 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: openAiAccount
  name: 'gpt-6.2-sol'
  sku: {
    name: 'GlobalStandard'
    capacity: gpt62SolCapacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'gpt-6.2-sol'
      version: '2026-11-17'
    }
    raiPolicyName: 'Microsoft.DefaultV2'
    versionUpgradeOption: 'NoAutoUpgrade'
  }
}

output gpt62AstraDeploymentId string = gpt62Astra.id
output gpt62SolDeploymentId string = gpt62Sol.id
