// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const DEPLOYMENT_NAME_RE = /^[A-Za-z0-9._-]{1,64}$/;
const AZURE_HOST_SUFFIXES = [
  '.openai.azure.com', '.openai.azure.us', '.openai.azure.cn',
  '.cognitiveservices.azure.com', '.cognitiveservices.azure.us', '.cognitiveservices.azure.cn',
  '.services.ai.azure.com', '.services.ai.azure.us', '.services.ai.azure.cn',
];

function normalizeHttpsOrigin(value, allowedHost) {
  let url;
  // Check the original authority too: URL normalizes away :443, empty query
  // delimiters, backslashes and other noncanonical origin spellings.
  if (typeof value !== 'string' || !/^https:\/\/[A-Za-z0-9.-]+\/?$/.test(value)) {
    throw new Error('An allowed HTTPS origin is required.');
  }
  try { url = new URL(value); } catch { throw new Error('An allowed HTTPS origin is required.'); }
  if (url.hostname.length > 253 || !url.hostname.split('.').every(label => (
    /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/.test(label)
  )) || !allowedHost(url.hostname)) throw new Error('An allowed HTTPS origin is required.');
  return `${url.origin}/`;
}

function normalizeAzureOpenAIEndpoint(value) {
  try {
    return normalizeHttpsOrigin(value, hostname => AZURE_HOST_SUFFIXES.some(suffix => (
      hostname.endsWith(suffix) && hostname.length > suffix.length
    )));
  } catch { throw new Error('An Azure OpenAI HTTPS origin is required.'); }
}

function byoEndpointsEnabled(env = process.env) {
  const flag = env.ALLOW_BYO_AI_ENDPOINTS;
  if (flag !== undefined && flag !== '' && flag !== 'true' && flag !== 'false') {
    throw new Error('ALLOW_BYO_AI_ENDPOINTS must be true or false.');
  }
  return flag === 'true';
}

function managedModelsConfiguration(env = process.env) {
  const definitions = [
    { key: 'astra', model: 'gpt-6.2-astra', deploymentKey: 'AZURE_OPENAI_DEPLOYMENT_GPT62ASTRA' },
    { key: 'sol', model: 'gpt-6.2-sol', deploymentKey: 'AZURE_OPENAI_DEPLOYMENT_GPT62SOL' },
  ];
  const endpoint = env.AZURE_OPENAI_ENDPOINT || '';
  const allowed = (env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS || '').split(',').map(value => value.trim()).filter(Boolean);
  const deployments = definitions.map(({ key, model, deploymentKey }) => ({
    key, model, apiFormat: 'responses', deploymentKey, deployment: env[deploymentKey] || '',
  }));
  const noManagedSettings = deployments.every(item => !item.deployment)
    && !endpoint && !allowed.length && !env.AZURE_OPENAI_RESOURCE_ID && !env.AZURE_OPENAI_API_KEY;
  for (const name of [
    'AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA', 'AZURE_OPENAI_DEPLOYMENT_GPT61SOL',
    'AZURE_FOUNDRY_ENDPOINT', 'AZURE_FOUNDRY_RESOURCE_ID',
    'AZURE_FOUNDRY_DEPLOYMENT_OPUS55', 'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS',
    'AZURE_FOUNDRY_API_KEY',
  ]) {
    if (env[name]) throw new Error(`Remove retired managed-model setting ${name} before starting the new application.`);
  }
  if (noManagedSettings) {
    return {
      astra: { model: 'gpt-6.2-astra', apiFormat: 'responses', configured: false, deployment: null, endpoint: null },
      sol: { model: 'gpt-6.2-sol', apiFormat: 'responses', configured: false, deployment: null, endpoint: null },
      configured: false,
    };
  }
  const hasBothDeployments = deployments.every(item => DEPLOYMENT_NAME_RE.test(item.deployment));
  const exactAllowlist = allowed.length === 2 && new Set(allowed).size === 2
    && deployments.every(item => allowed.includes(item.deployment));
  if (!hasBothDeployments || deployments[0].deployment === deployments[1].deployment || !exactAllowlist) {
    throw new Error('GPT-6.2 Astra and GPT-6.2 Sol require distinct aliases and an exact two-deployment AZURE_OPENAI_ALLOWED_DEPLOYMENTS allowlist.');
  }
  const normalizedEndpoint = normalizeAzureOpenAIEndpoint(endpoint);
  const routes = Object.fromEntries(deployments.map(({ key, model, apiFormat, deployment }) => [
    key, { model, apiFormat, configured: true, deployment, endpoint: normalizedEndpoint },
  ]));
  return { ...routes, configured: true };
}

function runtimeManagedModelsConfiguration(config, allowByoAIEndpoints = false) {
  return {
    features: { bringYourOwnAI: allowByoAIEndpoints === true },
    ai: {
      defaultModel: 'gpt-6.2-sol', configured: config.configured,
      models: [config.astra, config.sol].map(({ model, apiFormat, deployment, configured }) => ({ model, apiFormat, deployment, configured })),
    },
  };
}

module.exports = {
  managedModelsConfiguration, runtimeManagedModelsConfiguration, byoEndpointsEnabled,
  normalizeAzureOpenAIEndpoint, normalizeHttpsOrigin, DEPLOYMENT_NAME_RE,
};
