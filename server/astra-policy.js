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

function astraConfiguration(env = process.env) {
  const deployment = env.AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA || '';
  const allowed = (env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS || '').split(',').map(value => value.trim()).filter(Boolean);
  if (!deployment && !env.AZURE_OPENAI_ENDPOINT && !env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS
    && !env.AZURE_OPENAI_API_KEY && !env.AZURE_OPENAI_RESOURCE_ID) {
    return { configured: false, deployment: null, endpoint: null };
  }
  if (!DEPLOYMENT_NAME_RE.test(deployment) || allowed.length !== 1 || allowed[0] !== deployment) {
    throw new Error('GPT-6 Astra requires AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA and the identical singleton AZURE_OPENAI_ALLOWED_DEPLOYMENTS.');
  }
  return { configured: true, deployment, endpoint: normalizeAzureOpenAIEndpoint(env.AZURE_OPENAI_ENDPOINT) };
}

function runtimeAstraConfiguration(config, allowByoAIEndpoints = false) {
  return {
    features: { bringYourOwnAI: allowByoAIEndpoints === true },
    ai: {
      model: 'gpt-6-astra', apiFormat: 'responses',
      deployment: config.deployment, configured: config.configured,
    },
  };
}

module.exports = {
  astraConfiguration, runtimeAstraConfiguration, byoEndpointsEnabled,
  normalizeAzureOpenAIEndpoint, normalizeHttpsOrigin, DEPLOYMENT_NAME_RE,
};
