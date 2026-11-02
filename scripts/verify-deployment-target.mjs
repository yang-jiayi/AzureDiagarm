import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const approvedTarget = Object.freeze({
  AZURE_TENANT_ID: '02310481-6752-4d53-8fc4-f597a8e047f4',
  AZURE_SUBSCRIPTION_ID: '0df5f427-7359-447e-9b60-fa67c5588b8f',
  AZURE_RESOURCE_GROUP: 'AzureDiagarm_rg',
  AZURE_CONTAINER_APP: 'azurediagarm-app',
});

export function verifyDeploymentTarget(env = process.env) {
  const mismatches = Object.entries(approvedTarget)
    .filter(([name, value]) => env[name] !== value)
    .map(([name]) => name);
  if (!env.AZURE_CLIENT_ID?.trim()) mismatches.push('AZURE_CLIENT_ID');
  if (mismatches.length) {
    throw new Error(`Deployment target check failed for: ${mismatches.join(', ')}`);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  try {
    verifyDeploymentTarget();
    console.info('Approved deployment target verified.');
  } catch {
    console.error('::error::Deployment target variables do not match the approved new-tenant target.');
    process.exitCode = 1;
  }
}
