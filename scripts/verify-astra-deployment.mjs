import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import policy from '../server/astra-policy.js';

export const APPROVED_ASTRA_MODEL_VERSION = '2026-09-03';

export function verifyAstraDeployment(env, read = (url) => JSON.parse(execFileSync('az', [
  'rest', '--method', 'get', '--url', url, '--output', 'json', '--only-show-errors',
], { encoding: 'utf8', timeout: 30_000, stdio: ['ignore', 'pipe', 'pipe'] }))) {
  const id = env.AZURE_OPENAI_RESOURCE_ID;
  const alias = env.AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA;
  if (!/^\/subscriptions\/[a-f0-9-]{36}\/resourceGroups\/[A-Za-z0-9._()-]+\/providers\/Microsoft\.CognitiveServices\/accounts\/[A-Za-z0-9-]+$/i.test(id || '')
    || !policy.DEPLOYMENT_NAME_RE.test(alias || '')) {
    throw new Error('A full AZURE_OPENAI_RESOURCE_ID and explicit AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA are required for model verification.');
  }
  const endpoint = policy.normalizeAzureOpenAIEndpoint(env.AZURE_OPENAI_ENDPOINT);
  const accountUrl = `https://management.azure.com${id}?api-version=2024-10-01`;
  const account = read(accountUrl);
  const accountEndpoints = [account?.properties?.endpoint, ...Object.values(account?.properties?.endpoints || {})];
  if (account?.id?.toLowerCase() !== id.toLowerCase()
    || !accountEndpoints.some(value => {
      try { return policy.normalizeAzureOpenAIEndpoint(value) === endpoint; } catch { return false; }
    })) {
    throw new Error('The configured endpoint does not belong to the verified Azure OpenAI account.');
  }
  const deploymentId = `${id}/deployments/${alias}`;
  const deployment = read(`https://management.azure.com${deploymentId}?api-version=2024-10-01`);
  if (deployment?.id?.toLowerCase() !== deploymentId.toLowerCase()
    || deployment?.properties?.provisioningState !== 'Succeeded'
    || deployment?.properties?.model?.format !== 'OpenAI'
    || deployment?.properties?.model?.name !== 'gpt-6-astra'
    || deployment?.properties?.model?.version !== APPROVED_ASTRA_MODEL_VERSION) {
    throw new Error(`The approved deployment alias must resolve to a successfully provisioned OpenAI gpt-6-astra model, version ${APPROVED_ASTRA_MODEL_VERSION}. No fallback is permitted.`);
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    verifyAstraDeployment(process.env);
    console.log(`Verified the configured deployment is genuine GPT-6 Astra, version ${APPROVED_ASTRA_MODEL_VERSION}.`);
  } catch {
    console.error('GPT-6 Astra model identity verification failed. Check the resource ID, endpoint, deployment model/version and read permissions; no deployment is permitted.');
    process.exitCode = 1;
  }
}
