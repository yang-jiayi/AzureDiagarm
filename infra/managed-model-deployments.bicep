// Explicit opt-in template for the two managed deployments.
// Apply only after confirming model availability, regional capacity, quota,
// pricing, approval and the impact of creating these new deployments.
// This template does not update, rename, or delete the existing Astra deployment.
targetScope = 'resourceGroup'

@minLength(2)
@description('Existing Azure OpenAI account. The account itself is not changed.')
param openAiAccountName string

@minLength(2)
@description('Existing Azure AI Services account hosting the Foundry model. The account itself is not changed.')
param foundryAccountName string

@minValue(1)
@description('GlobalStandard capacity units for GPT-6.1 Sol. Confirm quota and capacity before deployment.')
param gpt61SolCapacity int

@minValue(1)
@description('GlobalStandard capacity units for Claude Opus 5.5. Confirm quota and capacity before deployment.')
param claudeOpus55Capacity int

resource openAiAccount 'Microsoft.CognitiveServices/accounts@2025-06-01' existing = {
  name: openAiAccountName
}

resource foundryAccount 'Microsoft.CognitiveServices/accounts@2025-06-01' existing = {
  name: foundryAccountName
}

resource gpt61Sol 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: openAiAccount
  name: 'gpt-6.1-sol'
  sku: {
    name: 'GlobalStandard'
    capacity: gpt61SolCapacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'gpt-6.1-sol'
      version: '2026-09-29'
    }
    raiPolicyName: 'Microsoft.DefaultV2'
    versionUpgradeOption: 'NoAutoUpgrade'
  }
}

resource claudeOpus55 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: foundryAccount
  name: 'claude-opus-5-5'
  sku: {
    name: 'GlobalStandard'
    capacity: claudeOpus55Capacity
  }
  properties: {
    model: {
      format: 'Anthropic'
      name: 'claude-opus-5-5'
      version: '2'
    }
    raiPolicyName: 'Microsoft.DefaultV2'
    versionUpgradeOption: 'NoAutoUpgrade'
  }
}

output gpt61SolDeploymentId string = gpt61Sol.id
output claudeOpus55DeploymentId string = claudeOpus55.id
