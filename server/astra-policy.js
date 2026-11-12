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

function normalizeFoundryEndpoint(value) {
  return normalizeHttpsOrigin(value, hostname => hostname.endsWith('.services.ai.azure.com')
    && hostname.length > '.services.ai.azure.com'.length);
}

function managedModelsConfiguration(env = process.env) {
  const route = (model, apiFormat, prefix, deploymentKey, normalize) => {
    const deployment = env[deploymentKey] || '';
    const endpoint = env[`${prefix}_ENDPOINT`];
    const allowed = (env[`${prefix}_ALLOWED_DEPLOYMENTS`] || '').split(',').map(value => value.trim()).filter(Boolean);
    if (!deployment && !endpoint && !allowed.length && !env[`${prefix}_RESOURCE_ID`] && !env[`${prefix}_API_KEY`]) {
      return { model, apiFormat, configured: false, deployment: null, endpoint: null };
    }
    if (!DEPLOYMENT_NAME_RE.test(deployment) || deployment === 'gpt-6-astra' || allowed.length !== 1 || allowed[0] !== deployment) {
      throw new Error(`${model} requires ${deploymentKey} and its identical singleton ${prefix}_ALLOWED_DEPLOYMENTS.`);
    }
    return { model, apiFormat, configured: true, deployment, endpoint: normalize(endpoint) };
  };
  if (env.AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA) throw new Error('Remove the retired AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA configuration before starting the new application.');
  if (env.AZURE_FOUNDRY_API_KEY) throw new Error('Managed Claude requires Microsoft Entra authentication, not AZURE_FOUNDRY_API_KEY.');
  const sol = route('gpt-6.1-sol', 'responses', 'AZURE_OPENAI', 'AZURE_OPENAI_DEPLOYMENT_GPT61SOL', normalizeAzureOpenAIEndpoint);
  const opus = route('claude-opus-5-5', 'anthropic-messages', 'AZURE_FOUNDRY', 'AZURE_FOUNDRY_DEPLOYMENT_OPUS55', normalizeFoundryEndpoint);
  if (sol.configured && opus.configured && sol.deployment === opus.deployment) throw new Error('Managed model deployment aliases must be distinct.');
  return { sol, opus, configured: sol.configured && opus.configured };
}

function runtimeManagedModelsConfiguration(config, allowByoAIEndpoints = false) {
  return {
    features: { bringYourOwnAI: allowByoAIEndpoints === true },
    ai: {
      defaultModel: 'gpt-6.1-sol', configured: config.configured,
      models: [config.sol, config.opus].map(({ model, apiFormat, deployment, configured }) => ({ model, apiFormat, deployment, configured })),
    },
  };
}

module.exports = {
  managedModelsConfiguration, runtimeManagedModelsConfiguration, byoEndpointsEnabled,
  normalizeAzureOpenAIEndpoint, normalizeFoundryEndpoint, normalizeHttpsOrigin, DEPLOYMENT_NAME_RE,
};
