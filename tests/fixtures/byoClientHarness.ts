import { build } from 'esbuild';
import { createRequire } from 'node:module';

export const BYO_STORAGE_KEY = 'azure-diagrams-byo-ai-settings';
export const MANAGED_TEST_ENV = {
  VITE_AZURE_OPENAI_ENDPOINT: 'https://offline.openai.azure.com/',
  VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'actual-astra-production-alias',
};

type Client = typeof import('../../src/stores/byoAISettingsStore')
  & typeof import('../../src/services/aiModelRuntime')
  & typeof import('../../src/services/runtimeConfig')
  & typeof import('../../src/services/apiHelper')
  & typeof import('../../src/services/byoAIConnection')
  & Pick<typeof import('../../src/stores/modelSettingsStore'), 'updateModelSettings' | 'updateFeatureOverride' | 'MODEL_CONFIG'>
  & Pick<typeof import('../../src/services/azureOpenAI'),
    'callAzureOpenAI' | 'generateArchitectureWithAI' | 'generateArchitectureFromIaC' | 'analyzeArchitectureDiagramImage' | 'generateFollowUpSuggestions'>
  & Pick<typeof import('../../src/services/blueprintArchitectureAI'), 'generateBlueprintArchitectureWithAI'>
  & Pick<typeof import('../../src/services/referenceArchitectureAI'), 'generateReferenceArchitectureWithAI'>
  & Pick<typeof import('../../src/services/componentManifestAI'), 'generateComponentManifest'>
  & Pick<typeof import('../../src/services/architectureValidator'), 'validateArchitecture' | 'formatValidationReport'>
  & Pick<typeof import('../../src/services/deploymentGuideGenerator'),
    'generateDeploymentGuide' | 'downloadDeploymentGuide' | 'downloadAllBicepTemplates'>
  & Pick<typeof import('../../src/utils/modelNaming'), 'getModelSuffix' | 'generateModelFilename'>
  & { getTestModelUsage(): Array<Record<string, unknown>> };

const bundles = new Map<string, Promise<string>>();
export async function loadBYOClient(
  stored?: unknown,
  env: Record<string, string> = MANAGED_TEST_ENV,
  entries = new Map<string, string>(),
): Promise<{ client: Client; entries: Map<string, string> }> {
  if (stored !== undefined) entries.set(BYO_STORAGE_KEY, JSON.stringify(stored));
  const key = JSON.stringify(env);
  if (!bundles.has(key)) bundles.set(key, build({
    stdin: {
      contents: `
        export * from './src/stores/byoAISettingsStore';
        export * from './src/services/aiModelRuntime';
        export * from './src/services/runtimeConfig';
        export * from './src/services/apiHelper';
        export * from './src/services/byoAIConnection';
        export { updateModelSettings, updateFeatureOverride, MODEL_CONFIG } from './src/stores/modelSettingsStore';
        export { callAzureOpenAI, generateArchitectureWithAI, generateArchitectureFromIaC,
          analyzeArchitectureDiagramImage, generateFollowUpSuggestions } from './src/services/azureOpenAI';
        export { generateBlueprintArchitectureWithAI } from './src/services/blueprintArchitectureAI';
        export { generateReferenceArchitectureWithAI } from './src/services/referenceArchitectureAI';
        export { generateComponentManifest } from './src/services/componentManifestAI';
        export { validateArchitecture, formatValidationReport } from './src/services/architectureValidator';
        export { generateDeploymentGuide, downloadDeploymentGuide, downloadAllBicepTemplates } from './src/services/deploymentGuideGenerator';
        export { getModelSuffix, generateModelFilename } from './src/utils/modelNaming';
        export { getTestModelUsage } from './src/services/telemetryService';
      `,
      resolveDir: process.cwd(), loader: 'ts',
    },
    bundle: true, write: false, platform: 'node', format: 'cjs', logLevel: 'silent',
    define: { 'import.meta.env': JSON.stringify(env) },
    plugins: [{
      name: 'offline-telemetry',
      setup(builder) {
        builder.onResolve({ filter: /telemetryService$/ }, () => ({ path: 'telemetry', namespace: 'offline-test' }));
        builder.onLoad({ filter: /.*/, namespace: 'offline-test' }, () => ({
          contents: `const events = []; export const trackAIModelUsage = event => events.push(event);
            export const getTestModelUsage = () => structuredClone(events);`, loader: 'js',
        }));
      },
    }],
  }).then(result => result.outputFiles[0].text));
  const module = { exports: {} };
  new Function('module', 'exports', 'require', 'localStorage', await bundles.get(key)!)(
    module, module.exports, createRequire(import.meta.url), {
      getItem: (name: string) => entries.get(name) ?? null,
      setItem: (name: string, value: string) => { entries.set(name, value); },
    },
  );
  return { client: module.exports as Client, entries };
}

export function jsonResponse(value: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(value), { status, headers: { 'Content-Type': 'application/json', ...headers } });
}

export const capabilityResponse = (enabled = true) => jsonResponse({ features: { bringYourOwnAI: enabled } });
export const testResponse = (apiFormat = 'responses') => jsonResponse(apiFormat === 'responses'
  ? { status: 'completed', output_text: '{"status":"ok"}' }
  : { choices: [{ finish_reason: 'stop', message: { content: '{"status":"ok"}' } }] });
export const settle = () => new Promise<void>(resolve => setImmediate(resolve));

export function deferred<T>(): { promise: Promise<T>; resolve(value: T): void; reject(error: unknown): void } {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

export async function verifiedProfile(client: Client, profile: Parameters<Client['upsertBYOAIProfile']>[0], key = 'sk-offline-profile-key') {
  const saved = client.upsertBYOAIProfile(profile);
  client.setBYOAIApiKey(saved.id, key);
  await client.testBYOAIConnection(saved.id);
  return saved;
}
