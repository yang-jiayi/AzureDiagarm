import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { chromium, type Browser, type Page } from 'playwright';

let browser: Browser;
let script: string;
const modelSettingsScripts = new Map<string, string>();
before(async () => {
  const mocks: Record<string, string> = {
    modelSettingsStore: `
      const settings = {model:'test-model', reasoningEffort:'none'};
      export const useModelSettings = () => [window.h.settings ?? settings];
      export const getModelSettings = () => settings;
      export const getModelSettingsForFeature = () => window.h.settings ?? settings;
      export const getAvailableModels = () => window.h.availableModels ??
        (['validation','models'].includes(window.h.surface) ? ['test-model','second-model'] : ['test-model']);
      export const MODEL_CONFIG = {
        'test-model': {displayName:'Test', apiFormat:'responses', isReasoning:false},
        'second-model': {displayName:'Second', apiFormat:'responses', isReasoning:false},
        'gpt-6-astra': {displayName:'GPT-6 Astra', apiFormat:'responses', isReasoning:true, maxCompletionTokens:32000},
        'gpt-5.6-sol': {displayName:'GPT-5.6 Sol', apiFormat:'responses', isReasoning:false}
      };
      export const FEATURE_CONFIG = {architectureGeneration:{displayName:'Topology',recommendedModel:'test-model'},blueprint:{displayName:'Blueprint',recommendedModel:'test-model'}};
      export const isModelAvailable = () => true;
      export const updateFeatureOverride = () => {};
      export const getSupportedReasoningEfforts = () => ['none','max'];
      export const getReasoningEffortLabel = x => x;
      export const getCommonSupportedReasoningEfforts = () => ['none'];
      export const normalizeReasoningEffort = (_, value) => value;
    `,
    byoAISettingsStore: `
      export const useBYOAISettings = () => ({settings:{enabled:false},verified:false});
      export const getBYOAIProviderLabel = provider => provider;
    `,
    runtimeConfig: `export const useRuntimeConfig = () => ({status:'ready',bringYourOwnAI:false});`,
    LanguageContext: `export const useLanguage = () => ({language:window.h.language ?? 'en',t:x=>x,translate:x=>x});`,
    safeStorage: `
      export const readLocalStorage = key => key==='aiGenerator.mode' ? window.h.mode : null;
      export const readBooleanPreference = (key, fallback) => key==='aiGenerator.bothInParallel' ? window.h.parallel : fallback;
      export const writeLocalStorage = () => {};
    `,
    telemetryService: `
      export const trackImageImport = () => {};
      export const trackValidationCompared = event => window.h.telemetry.push({kind:'compared',...event});
      export const trackValidationCritiqueRanked = event => window.h.telemetry.push({kind:'critique',...event});
      export const trackValidationFindings = event => window.h.telemetry.push({kind:'findings',...event});
    `,
    aiBudgetService: `
      export const getAIBudget = signal => {
        const h = window.h;
        h.budgetReads.push({signal});
        if(h.budgetFailure) return Promise.reject(new Error(h.budgetFailure));
        const snapshot = () => ({concurrentLimit:h.budgetLimit, concurrentRequests:h.inFlight+h.budgetExternal});
        if(h.deferBudget) return new Promise(resolve => h.budgetPending.push({signal,resolve:()=>resolve(snapshot())}));
        return Promise.resolve(snapshot());
      };
    `,
    ImageUploader: `export default function ImageUploader(){return null}`,
    azureOpenAI: `
      export const isAzureOpenAIConfigured = () => true;
      export const isManagedAIConfigured = () => true;
      export const generateValidationCritique = (...args) => window.h.generate('critique', args[2]?.signal, args[1], args[2]);
      export const generateCritique = (...args) => window.h.generate('critique', args[2]?.signal, args[1], args[2]);
      export const throwIfGenerationAborted = signal => {if(signal?.aborted) throw new DOMException('Cancelled','AbortError');};
      export const generateArchitectureWithAI = (...args) => window.h.generate('topology', args[4]?.signal || args[4] || args[1]?.signal, args[0], args[1]);
      export const generateFollowUpSuggestions = input => {window.h.followups++; window.h.followupInputs.push(input); return Promise.resolve(['Add monitoring']);};
      export const analyzeArchitectureDiagramImage = () => Promise.resolve({description:'image'});
    `,
    referenceArchitectureAI: `export const generateReferenceArchitectureWithAI = (prompt, settings) => window.h.generate('reference', settings.signal, prompt);`,
    blueprintArchitectureAI: `export const generateBlueprintArchitectureWithAI = (prompt, settings) => window.h.generate('blueprint', settings.signal, prompt, settings);`,
    componentManifestAI: `export const generateComponentManifest = (prompt, settings) => window.h.generate('manifest', settings.signal, prompt, settings);`,
    architectureValidator: `export const validateArchitecture = (...args) =>
      window.h.generate('validation', args[4]?.signal, args[3], args[4]);`,
    avatarPresenter: `export class AvatarPresenter { disconnect(){} }`,
    exportReferencePng: `export const exportReferenceArchitectureAsPng = async () => {window.h.exports++;};`,
    exportBlueprintPng: `export const exportBlueprintArchitectureAsPng = async () => {window.h.exports++;};`,
  };
  const result = await build({
    stdin: {
      contents: `
        import React from 'react';
        import {createRoot} from 'react-dom/client';
        import {flushSync} from 'react-dom';
        import Generator from './src/components/AIArchitectureGenerator';
        import Chat from './src/components/ArchitectureChatPanel';
        import Review from './src/components/AIChangeReview';
        import CompareValidation from './src/components/CompareValidationModal';
        import CompareModels from './src/components/CompareModelsModal';
        import {buildDiagramChanges} from './src/services/diagramChanges';
        import {runWithRateLimitRetry} from './src/services/aiRetry';
        import {createOpenAIProxyError} from './src/services/apiHelper';
        const h = window.h;
        const root = createRoot(document.getElementById('root'));
        h.dispatch = (kind, signal, prompt, override) => {
          h.calls.push({kind,signal,prompt,override});
          const comparison = ['models','validation'].includes(h.surface) || (h.surface==='generator' && h.mode==='both');
          if(comparison && (h.inFlight+h.budgetExternal >= h.budgetLimit || h.contentionRemaining > 0)) {
            if(h.inFlight+h.budgetExternal >= h.budgetLimit) h.dispatchViolations++;
            if(h.contentionRemaining > 0) {
              h.contentionRemaining--;
              if(h.occupyOnContention) h.budgetExternal=h.budgetLimit;
            }
            return Promise.reject(Object.assign(new Error('Capacity is occupied.'), {code:'ai_concurrency_limit',status:429}));
          }
          if(comparison) { h.inFlight++; h.peakRequests=Math.max(h.peakRequests,h.inFlight); }
          let released = false;
          const release = () => { if(comparison && !released) {released=true;h.inFlight--;} };
          signal?.addEventListener('abort',release,{once:true});
          const metrics={totalTokens:1,completionTokens:1,promptTokens:0,elapsedTimeMs:1};
          const value = kind==='validation' ? {pillars:[],overallScore:80,quickWins:[],summary:'Review',metrics}
            : kind==='critique' ? {content:'## Recommendation\\n**Test** is recommended.',metrics}
            : {services:[],connections:[],groups:[],components:[],zones:[],metrics};
          let request;
          if(h.rateLimits[kind] > 0) {
            h.rateLimits[kind]--;
            request=Promise.reject(Object.assign(new Error('The AI provider is rate-limiting requests.'), {
              code:'azure_openai_rate_limited',status:429,retryAfterMs:h.retryAfterMs,
            }));
          }
          else if(kind==='validation' || h.deferKinds.includes(kind)) request = new Promise((resolve,reject) => {
            const cleanup = () => {
              h.pending=h.pending.filter(item=>item!==entry);
              h.validationPending=h.validationPending.filter(finish=>finish!==entry.resolve);
            };
            const entry={kind,override,resolve:()=>{cleanup();resolve(value);},
              reject:message=>{cleanup();reject(new Error(message));}};
            h.pending.push(entry);
            if(kind==='validation') h.validationPending.push(entry.resolve);
          });
          else if(h.failKinds.includes(kind)) request=Promise.reject(h.failureCodes[kind] ? createOpenAIProxyError({
            ok:false,status:h.failureStatuses[kind] ?? 429,data:null,
            error:{
              source:h.failureSources[kind] ?? 'azure_openai',code:h.failureCodes[kind],
              requestId:'test-'+kind+'-failure',retryAfterMs:h.retryAfterMs,
            },
          }) : new Error(kind+' generation failed'));
          else request=Promise.resolve(value);
          return request.finally(()=>{release();signal?.removeEventListener('abort',release);});
        };
        h.generate = (kind, signal, prompt, override) => h.useRatePolicy
          ? runWithRateLimitRetry(()=>h.dispatch(kind,signal,prompt,override),{signal,onRetryWait:override?.onRetryWait})
          : h.dispatch(kind,signal,prompt,override);
        const apply = (...args) => {
          h.applies.push(args);
          if(h.deferReview) return new Promise(resolve => {
            h.finishReview=resolve;
            args.at(-1)?.addEventListener('abort',()=>resolve(false),{once:true});
          });
          if(h.accepted && h.acceptedNodes) {h.nodes=h.acceptedNodes; h.render();}
          return Promise.resolve(h.legacyVoid ? undefined : h.accepted);
        };
        h.render = () => {
          const currentArchitecture={nodes:h.nodes || [],edges:[],architectureName:'Test',revision:h.revision};
          if(h.surface==='chat') root.render(<Chat isOpen={h.open} onClose={()=>{h.open=false;h.render()}}
            currentArchitecture={currentArchitecture} diagramKey={h.diagramKey} onApply={apply}/>);
          else if(h.surface==='validation') root.render(<CompareValidation isOpen={h.open}
            onClose={()=>{h.open=false;h.render()}} diagramFingerprint={h.diagramFingerprint}
            services={[{name:'Web App',type:'App Service',category:'app services'}]} connections={[]}
            onApply={(...args)=>{h.applies.push(args);return h.accepted;}}/>);
          else if(h.surface==='models') root.render(<CompareModels isOpen={h.open}
            onClose={()=>{h.open=false;h.render()}} onApply={apply}/>);
          else if(h.surface==='review') {
            const n=id=>({id,position:{x:0,y:0},data:{label:id}});
            h.changeSet ||= buildDiagramChanges({nodes:[n('a')],edges:[]},{nodes:[n('a'),n('b')],edges:[{id:'e',source:'a',target:'b'}]});
            root.render(<Review changeSet={h.changeSet} error={h.reviewError} onApply={graph=>{
              h.reviewAttempts++;
              if(h.snapshotError) {h.reviewError='Snapshot failed; your diagram was not changed.'; h.render(); return;}
              h.reviewError=''; h.reviewed.push(graph); h.render();
            }} onCancel={()=>h.cancelled++}/>);
          } else root.render(<Generator onGenerate={apply} currentArchitecture={currentArchitecture}
            onReferenceArchitecture={()=>h.references++} onBlueprintArchitecture={()=>h.blueprints++}/>);
        };
        h.commitRender = () => flushSync(h.render);
        h.unmount = () => root.unmount();
        h.render();
      `, resolveDir: process.cwd(), loader: 'tsx',
    },
    bundle: true, write: false, platform: 'browser', format: 'iife', logLevel: 'silent',
    define: { 'process.env.NODE_ENV': '"test"', 'import.meta.env': '{}' },
    plugins: [{
      name: 'component-boundaries',
      setup(build) {
        build.onResolve({ filter: /\.css$/ }, args => ({ path: args.path, namespace: 'empty-css' }));
        build.onLoad({ filter: /.*/, namespace: 'empty-css' }, () => ({ contents: '' }));
        build.onResolve({ filter: /\/(modelSettingsStore|byoAISettingsStore|runtimeConfig|LanguageContext|safeStorage|telemetryService|aiBudgetService|ImageUploader|azureOpenAI|referenceArchitectureAI|blueprintArchitectureAI|componentManifestAI|architectureValidator|avatarPresenter|exportReferencePng|exportBlueprintPng)$/ },
          args => ({ path: args.path.split('/').at(-1)!, namespace: 'component-mock' }));
        build.onLoad({ filter: /.*/, namespace: 'component-mock' }, args => ({ contents: mocks[args.path], loader: 'js' }));
      },
    }],
  });
  script = result.outputFiles[0].text;
  const legacyDeployments = {
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL: 'gpt-5.6-sol',
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA: 'gpt-5.6-terra',
    VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA: 'gpt-5.6-luna',
  };
  const environments: Record<string, Record<string, string>> = {
    astra: { ...legacyDeployments, VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'gpt-6-astra' },
    legacy: legacyDeployments,
    none: {},
  };
  await Promise.all(Object.entries(environments).map(async ([profile, environment]) => {
    const settingsBundle = await build({
      stdin: {
        contents: `
          import React from 'react';
          import {createRoot} from 'react-dom/client';
          import Popover from './src/components/ModelSettingsPopover';
          import Selector from './src/components/ModelSelector';
          import Badge from './src/components/ModelBadge';
          import {LanguageProvider, useLanguage} from './src/i18n/LanguageContext';
          import {getModelSettings, getModelSettingsForFeature, getRecommendedModelSettings,
            updateModelSettings} from './src/stores/modelSettingsStore';
          const h = window.h;
          if(h.initialSettings) updateModelSettings(h.initialSettings);
          h.getModelSettings = getModelSettings;
          h.getModelSettingsForFeature = getModelSettingsForFeature;
          h.getRecommendedModelSettings = getRecommendedModelSettings;
          const root = createRoot(document.getElementById('root'));
          function Harness() {
            const language = useLanguage();
            h.setLanguage = language.setLanguage;
            h.translate = language.translate;
            if(h.surface === 'popover') return <Popover isOpen={h.open}
              onToggle={()=>{h.open=!h.open;h.render();}}
              onOpenBYOSettings={()=>h.byoSettingsOpened++}/>;
            if(h.surface === 'badge') return <Badge modelName={h.modelName} elapsedTimeMs={1200}/>;
            return <Selector compact={h.surface === 'compact'}/>;
          }
          h.render = () => root.render(<LanguageProvider><Harness/></LanguageProvider>);
          h.render();
        `,
        resolveDir: process.cwd(), loader: 'tsx',
      },
      bundle: true, write: false, platform: 'browser', format: 'iife', logLevel: 'silent',
      define: { 'process.env.NODE_ENV': '"test"', 'import.meta.env': JSON.stringify(environment) },
      plugins: [{
        name: 'model-settings-environment',
        setup(build) {
          build.onResolve({ filter: /\.css$/ }, args => ({ path: args.path, namespace: 'empty-css' }));
          build.onLoad({ filter: /.*/, namespace: 'empty-css' }, () => ({ contents: '' }));
          build.onResolve({ filter: /\/(byoAISettingsStore|runtimeConfig)$/ },
            args => ({ path: args.path.split('/').at(-1)!, namespace: 'model-settings-mock' }));
          build.onLoad({ filter: /.*/, namespace: 'model-settings-mock' }, args => ({
            contents: args.path === 'runtimeConfig'
              ? `export const useRuntimeConfig = () => ({status:'ready',bringYourOwnAI:true});`
              : `export const useBYOAISettings = () => window.h.byo;
                 export const getBYOAIModelLabel = () => window.h.byo.settings.model;`,
            loader: 'js',
          }));
        },
      }],
    });
    modelSettingsScripts.set(profile, settingsBundle.outputFiles[0].text);
  }));
  browser = await chromium.launch({ headless: true });
});
after(async () => { await browser?.close(); });

async function setup(t: { after: (fn: () => Promise<void>) => void }, config: Record<string, unknown> = {}): Promise<Page> {
  const page = await browser.newPage({ locale: 'en-US' });
  t.after(() => page.close());
  await page.setContent(config.surface === 'chat'
    ? '<html><body><div class="app"><header class="app-header"><button id="outside">Outside</button></header><main class="workspace" tabindex="0" data-modal-focus-fallback>Canvas</main><div id="root"></div></div></body></html>'
    : '<html><body><button id="outside">Outside</button><div id="root"></div></body></html>');
  await page.evaluate(config => {
    if (config.abortSignalAnyUnavailable) {
      Object.defineProperty(AbortSignal, 'any', { value: undefined, configurable: true });
    }
    (window as any).h = {
      surface: 'generator', mode: 'topology', parallel: true, revision: 7, accepted: false,
      open: true, diagramKey: 'diagram-a', diagramFingerprint: 'fingerprint-a',
      deferKinds: [], failKinds: [], deferReview: false, calls: [], pending: [], applies: [], validationPending: [],
      followups: 0, followupInputs: [], blueprints: 0, references: 0, exports: 0, reviewed: [], cancelled: 0,
      reviewAttempts: 0, reviewError: '', snapshotError: false,
      budgetLimit: 2, budgetExternal: 0, budgetReads: [], budgetFailure: null, deferBudget: false, budgetPending: [],
      inFlight: 0, peakRequests: 0, dispatchViolations: 0, contentionRemaining: 0, telemetry: [],
      rateLimits: {}, failureCodes: {}, failureSources: {}, failureStatuses: {},
      useRatePolicy: false, retryAfterMs: 60_000, ...config,
    };
  }, config);
  await page.addScriptTag({ content: script });
  return page;
}
async function setupModelSettings(
  t: { after: (fn: () => Promise<void>) => void },
  profile: string,
  config: Record<string, unknown> = {},
): Promise<Page> {
  const page = await browser.newPage({ locale: 'en-US' });
  t.after(() => page.close());
  await page.route('http://model-settings.test/**', route => route.fulfill({
    contentType: 'text/html',
    body: '<html><body><div id="root"></div></body></html>',
  }));
  await page.goto('http://model-settings.test/');
  await page.evaluate(config => {
    (window as any).h = {
      surface: 'popover', open: true, byoSettingsOpened: 0,
      byo: { settings: { enabled: false }, verified: false, connectionState: 'unverified' },
      ...config,
    };
  }, config);
  await page.addScriptTag({ content: modelSettingsScripts.get(profile)! });
  await page.locator('#root > *').first().waitFor();
  return page;
}
async function openModelGuidance(page: Page, surface: string) {
  await page.locator(surface === 'popover'
    ? '.msp-advanced-settings summary'
    : surface === 'compact' ? '.advanced-toggle-compact' : '.advanced-toggle').click();
}

async function generate(page: Page) {
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  await page.locator('#architecture-description').fill('Add a secure web application');
  await page.getByRole('button', { name: 'Continue to output', exact: true }).click();
  await page.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
}
async function brief(page: Page) {
  return await page.locator('#architecture-description').count()
    ? page.locator('#architecture-description').inputValue()
    : page.locator('.generator-output-summary > p').innerText();
}
async function finish(page: Page, kind: string, error?: string) {
  await page.waitForFunction(kind => (window as any).h.pending.some((item: any) => item.kind === kind), kind);
  await page.evaluate(({ kind, error }) => {
    const h = (window as any).h;
    const pending = h.pending.splice(h.pending.findIndex((item: any) => item.kind === kind), 1)[0];
    if (error) pending.reject(error);
    else pending.resolve();
  }, { kind, error });
}

for (const [mode, parallel] of [['topology', true], ['both', true], ['both', false]] as const) {
  test(`${mode}/${parallel ? 'parallel' : 'sequential'} rejected review keeps prompt and produces no success or blueprint side effects`, async t => {
    const page = await setup(t, { mode, parallel });
    await generate(page);
    await page.getByRole('button', { name: 'Retry generation' }).waitFor();
    assert.equal(await brief(page), 'Add a secure web application');
    const state = await page.evaluate(() => {
      const h = (window as any).h;
      return { applies: h.applies.length, revision: h.applies[0][4], blueprints: h.blueprints, exports: h.exports };
    });
    assert.deepEqual(state, { applies: 1, revision: 7, blueprints: 0, exports: 0 });
    assert.equal(await page.locator('.generator-success-panel').count(), 0);
    assert.equal(await page.getByRole('dialog').count(), 1);
  });
}

test('generator captures revision before async work; cancelled late results never apply and retry works', async t => {
  const page = await setup(t, { deferKinds: ['topology'] });
  await generate(page);
  await page.evaluate(() => { (window as any).h.revision = 8; (window as any).h.render(); });
  await page.getByRole('button', { name: 'Cancel request' }).click();
  assert.equal(await page.evaluate(() => (window as any).h.calls[0].signal.aborted), true);
  await finish(page, 'topology');
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  await page.getByRole('button', { name: 'Retry generation' }).click();
  await page.evaluate(() => { (window as any).h.revision = 9; (window as any).h.render(); });
  await finish(page, 'topology');
  await page.waitForFunction(() => (window as any).h.applies.length === 1);
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][4]), 8);
});

test('a cancelled old response cannot finish a newer generator request', async t => {
  const page = await setup(t, { deferKinds: ['topology'] });
  await generate(page);
  await page.getByRole('button', { name: 'Cancel request' }).click();
  await page.getByRole('button', { name: 'Retry generation' }).click();
  await page.waitForFunction(() => (window as any).h.pending.length === 2);
  await finish(page, 'topology');
  assert.equal(await page.getByRole('button', { name: 'Generating...' }).isDisabled(), true);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  await finish(page, 'topology');
  await page.getByRole('button', { name: 'Retry generation' }).waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1);
});

test('cancel during parent review aborts the callback signal and retains prompt without success', async t => {
  const page = await setup(t, { deferReview: true });
  await generate(page);
  await page.waitForFunction(() => (window as any).h.applies.length === 1);
  await page.getByRole('button', { name: 'Cancel request' }).click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][5].aborted), true);
  assert.equal(await brief(page), 'Add a secure web application');
  assert.equal(await page.locator('.generator-success-panel').count(), 0);
});

test('Both-mode manifest cancellation without AbortSignal.any does not start fallback providers', async t => {
  const page = await setup(t, { mode: 'both', deferKinds: ['manifest'], abortSignalAnyUnavailable: true });
  await generate(page);
  await page.getByRole('button', { name: 'Cancel request' }).click();
  await finish(page, 'manifest');
  await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 30)));
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => call.kind)), ['manifest']);
  assert.equal(await page.evaluate(() => (window as any).h.calls[0].signal.aborted), true);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
});

for (const parallel of [true, false]) {
  test(`Both-mode ${parallel ? 'parallel' : 'sequential'} cancellation suppresses both late outputs`, async t => {
    const page = await setup(t, {
      mode: 'both', parallel, deferKinds: ['topology', 'blueprint'], abortSignalAnyUnavailable: true,
    });
    const outputKinds = parallel ? ['topology', 'blueprint'] : ['topology'];
    await generate(page);
    await page.waitForFunction(count => (window as any).h.pending.length === count, outputKinds.length);
    await page.getByRole('button', { name: 'Cancel request' }).click();
    await finish(page, 'topology');
    if (parallel) await finish(page, 'blueprint');
    await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 30)));
    assert.equal(await page.evaluate(() => (window as any).h.applies.length + (window as any).h.blueprints + (window as any).h.exports), 0);
    const calls = await page.evaluate(() => (window as any).h.calls.map(
      (call: { kind: string; signal: AbortSignal }) => ({ kind: call.kind, aborted: call.signal.aborted }),
    ));
    assert.equal(calls[0].kind, 'manifest');
    // The completed manifest has left its admission scope; all pending outputs must abort.
    assert.deepEqual(calls.slice(1), outputKinds.map(kind => ({ kind, aborted: true })));
  });
}

const astraMaxBoth = {
  mode: 'both', parallel: true, accepted: true, useRatePolicy: true,
  settings: { model: 'gpt-6-astra', reasoningEffort: 'max' },
  availableModels: ['gpt-6-astra'],
};

test('Both/parallel/Astra MAX waits for provider capacity without replaying a completed topology', async t => {
  const page = await setup(t, { ...astraMaxBoth, rateLimits: { blueprint: 1 }, deferKinds: ['topology'] });
  await page.clock.install();
  await generate(page);
  await page.locator('.generator-retry-wait').waitFor();
  assert.match(await page.locator('.generator-retry-wait').innerText(), /attempt 2\/3.*same model, reasoning, and output limit/);
  await finish(page, 'topology');
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => call.kind)), [
    'manifest', 'topology', 'blueprint',
  ]);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  await page.clock.fastForward(61_000);
  await page.waitForFunction(() => (window as any).h.blueprints === 1);
  const state = await page.evaluate(() => {
    const h = (window as any).h;
    return {
      calls: h.calls.map((call: any) => call.kind),
      settings: h.calls.map((call: any) => [call.override.model, call.override.reasoningEffort]),
      applies: h.applies.length, blueprints: h.blueprints, violations: h.dispatchViolations,
    };
  });
  assert.deepEqual(state, {
    calls: ['manifest', 'topology', 'blueprint', 'blueprint'],
    settings: Array.from({ length: 4 }, () => ['gpt-6-astra', 'max']),
    applies: 1, blueprints: 1, violations: 0,
  });
});

test('Both manifest rate-limit exhaustion stays bounded and never fans out into fallback requests', async t => {
  const page = await setup(t, { ...astraMaxBoth, rateLimits: { manifest: 3 } });
  await page.clock.install();
  await generate(page);
  await page.locator('.generator-retry-wait').filter({ hasText: 'attempt 2/3' }).waitFor();
  await page.clock.fastForward(61_000);
  await page.locator('.generator-retry-wait').filter({ hasText: 'attempt 3/3' }).waitFor();
  await page.clock.fastForward(61_000);
  await page.getByRole('alert').waitFor();
  assert.match(await page.getByRole('alert').innerText(), /Wait at least 60s.*never lower the model, reasoning, or output limit.*deployment capacity/);
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    return { calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints };
  }), { calls: ['manifest', 'manifest', 'manifest'], applies: 0, blueprints: 0 });
});

for (const [code, status, source, message] of [
  ['ai_daily_budget_exceeded', 429, 'budget', /daily AI budget.*midnight UTC/],
  ['azure_openai_unavailable', 500, 'azure_openai', /internal server error.*not automatically retried/],
] as const) {
  test(`Both manifest ${code} is terminal without fallback or misleading provider-throttle guidance`, async t => {
    const page = await setup(t, {
      ...astraMaxBoth, failKinds: ['manifest'], failureCodes: { manifest: code },
      failureSources: { manifest: source }, failureStatuses: { manifest: status }, retryAfterMs: 1000,
    });
    await page.clock.install();
    await generate(page);
    await page.getByRole('alert').waitFor();
    assert.match(await page.getByRole('alert').innerText(), message);
    assert.match(await page.getByRole('alert').innerText(), /test-manifest-failure/);
    assert.doesNotMatch(await page.getByRole('alert').innerText(), /provider.*rate.limit|deployment capacity/);
    await page.clock.fastForward(121_000);
    assert.deepEqual(await page.evaluate(() => {
      const h = (window as any).h;
      return { calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints };
    }), { calls: ['manifest'], applies: 0, blueprints: 0 });
  });
}

test('Both partial daily-budget exhaustion preserves the completed output and waits for an explicit missing-output retry', async t => {
  const page = await setup(t, {
    ...astraMaxBoth, failKinds: ['blueprint'], failureCodes: { blueprint: 'ai_daily_budget_exceeded' },
    failureSources: { blueprint: 'budget' }, retryAfterMs: 1000,
  });
  await page.clock.install();
  await generate(page);
  await page.getByRole('button', { name: 'Retry missing output' }).waitFor();
  const warning = page.locator('.generator-output-summary .azd-callout--warning');
  assert.match(await warning.innerText(), /daily AI budget.*midnight UTC.*After the daily budget resets/);
  assert.doesNotMatch(await warning.innerText(), /lowering reasoning|faster model|deployment capacity/);
  await page.clock.fastForward(121_000);
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    return { calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints };
  }), { calls: ['manifest', 'topology', 'blueprint'], applies: 1, blueprints: 0 });
  await page.evaluate(() => { (window as any).h.failKinds = []; });
  await page.getByRole('button', { name: 'Retry missing output' }).click();
  await page.waitForFunction(() => (window as any).h.blueprints === 1);
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => call.kind)), [
    'manifest', 'topology', 'blueprint', 'blueprint',
  ]);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1);
});

test('Japanese Both mode explains an excessive provider cooldown without silent downgrade or fan-out', async t => {
  const page = await setup(t, { ...astraMaxBoth, language: 'ja', rateLimits: { manifest: 1 }, retryAfterMs: 180_000 });
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  await page.locator('#architecture-description').fill('もっともセキュアの構成で、FabricのE2EのArchitecture図を作成してください。');
  await page.getByRole('button', { name: '出力設定へ進む', exact: true }).click();
  await page.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
  await page.getByRole('alert').waitFor();
  assert.match(await page.getByRole('alert').innerText(), /180秒.*モデル、推論強度、出力上限を下げません/);
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => call.kind)), ['manifest']);
});

test('cancelling a rate-limited missing-output retry keeps accepted work and retries only the missing blueprint', async t => {
  const page = await setup(t, {
    ...astraMaxBoth, failKinds: ['blueprint'], failureCodes: { blueprint: 'azure_openai_rate_limited' },
    retryAfterMs: 180_000,
    acceptedNodes: [{ id: 'accepted', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Accepted topology' } }],
  });
  await page.clock.install();
  await generate(page);
  await page.getByRole('button', { name: 'Retry missing output' }).waitFor();
  assert.doesNotMatch(await page.locator('.generator-output-summary .azd-callout--warning').innerText(), /lowering reasoning|faster model/);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.failKinds = [];
    h.rateLimits.blueprint = 1;
    h.retryAfterMs = 60_000;
  });
  await page.getByRole('button', { name: 'Retry missing output' }).click();
  await page.locator('.generator-retry-wait').waitFor();
  await page.getByRole('button', { name: 'Cancel request' }).click();
  await page.getByRole('button', { name: 'Retry missing output' }).waitFor();
  await page.clock.fastForward(121_000);
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    return { calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints };
  }), { calls: ['manifest', 'topology', 'blueprint', 'blueprint'], applies: 1, blueprints: 0 });
  assert.equal(await page.locator('.generator-retry-wait').count(), 0);
  await page.getByRole('button', { name: 'Retry missing output' }).click();
  await page.waitForFunction(() => (window as any).h.blueprints === 1);
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    return {
      calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints,
      samePrompt: h.calls.filter((call: any) => call.kind === 'blueprint').every((call: any) => call.prompt === h.calls[0].prompt),
      sameSettings: h.calls.every((call: any) => call.override.model === 'gpt-6-astra' && call.override.reasoningEffort === 'max'),
    };
  }), {
    calls: ['manifest', 'topology', 'blueprint', 'blueprint', 'blueprint'],
    applies: 1, blueprints: 1, samePrompt: true, sameSettings: true,
  });
});

test('Both/parallel respects server cap one while retaining output completed before a budget outage', async t => {
  const page = await setup(t, { ...astraMaxBoth, budgetLimit: 1, deferKinds: ['topology'] });
  await generate(page);
  await page.waitForFunction(() => (window as any).h.pending.some((item: any) => item.kind === 'topology'));
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => call.kind)), ['manifest', 'topology']);
  await page.evaluate(() => { (window as any).h.budgetFailure = 'Budget unavailable'; });
  await finish(page, 'topology');
  await page.getByRole('button', { name: 'Retry missing output' }).waitFor();
  assert.match(await page.locator('.generator-output-summary .azd-callout--warning').innerText(), /budget could not be checked/);
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1);
  await page.evaluate(() => { (window as any).h.budgetFailure = null; });
  await page.getByRole('button', { name: 'Retry missing output' }).click();
  await page.waitForFunction(() => (window as any).h.blueprints === 1);
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    return {
      calls: h.calls.map((call: any) => call.kind), applies: h.applies.length,
      peak: h.peakRequests, violations: h.dispatchViolations,
    };
  }), { calls: ['manifest', 'topology', 'blueprint'], applies: 1, peak: 1, violations: 0 });
});

test('blueprint-only cancellation suppresses export and reference updates', async t => {
  const page = await setup(t, { mode: 'blueprint', deferKinds: ['blueprint'] });
  await generate(page);
  await page.getByRole('button', { name: 'Cancel request' }).click();
  await finish(page, 'blueprint');
  await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 30)));
  assert.equal(await page.evaluate(() => (window as any).h.blueprints + (window as any).h.exports), 0);
});

test('generator unmount aborts an in-flight provider and suppresses late callbacks', async t => {
  const page = await setup(t, { deferKinds: ['topology'] });
  await generate(page);
  await page.evaluate(() => (window as any).h.unmount());
  assert.equal(await page.evaluate(() => (window as any).h.calls[0].signal.aborted), true);
  await finish(page, 'topology');
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
});

test('Both accepted generation preserves legacy void callbacks and publishes its blueprint', async t => {
  const page = await setup(t, { mode: 'both', legacyVoid: true });
  await generate(page);
  await page.getByText('✓ Diagram created — review it before validation').waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1);
  assert.equal(await page.evaluate(() => (window as any).h.blueprints), 1);
  assert.equal(await page.evaluate(() => (window as any).h.exports), 1);
  await page.getByRole('button', { name: '1. Brief', exact: true }).click();
  assert.equal(await page.locator('#architecture-description').inputValue(), '');
});

for (const missing of ['topology', 'blueprint']) {
  test(`Both retry preserves the completed output and regenerates only ${missing}`, async t => {
    const page = await setup(t, {
      mode: 'both', accepted: true, failKinds: [missing],
      acceptedNodes: [{ id: 'sql', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'SQL' } }],
    });
    await generate(page);
    await page.getByRole('button', { name: 'Retry missing output' }).waitFor();
    assert.equal(await brief(page), 'Add a secure web application');
    await page.evaluate(() => { (window as any).h.failKinds = []; });
    await page.getByRole('button', { name: 'Retry missing output' }).click();
    await page.getByText('✓ Diagram created — review it before validation').waitFor();
    assert.deepEqual(await page.evaluate(() => {
      const h = (window as any).h;
      return { calls: h.calls.map((call: any) => call.kind), applies: h.applies.length, blueprints: h.blueprints };
    }), { calls: ['manifest', 'topology', 'blueprint', missing], applies: 1, blueprints: 1 });
    assert.equal(await page.evaluate(() => {
      const calls = (window as any).h.calls;
      return calls.at(-1).prompt === calls[1].prompt;
    }), true, 'retry uses the captured prompt, not the already-modified canvas');
  });
}

test('generator rejects an edited baseline but ignores selection and measurement changes', async t => {
  const node = { id: 'app', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'App' } };
  const page = await setup(t, { nodes: [node], deferKinds: ['topology'] });
  await generate(page);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.nodes = h.nodes.map((node: any) => ({ ...node, selected: true, width: 180, height: 80 }));
    h.commitRender();
  });
  await finish(page, 'topology');
  await page.getByRole('button', { name: 'Retry generation' }).waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1);
  await page.getByRole('button', { name: 'Retry generation' }).click();
  await page.evaluate(() => {
    const h = (window as any).h;
    h.nodes = h.nodes.map((node: any) => ({ ...node, data: { ...node.data, label: 'Manual edit' } }));
    h.commitRender();
  });
  await finish(page, 'topology');
  await page.getByRole('alert').filter({ hasText: 'diagram changed' }).waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 1, 'stale response must not reach apply');
  assert.equal(await brief(page), 'Add a secure web application');
});

test('chat rejected review retains prompt and suppresses summary and follow-ups', async t => {
  const page = await setup(t, { surface: 'chat', deferKinds: ['topology'] });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.evaluate(() => { (window as any).h.revision = 8; (window as any).h.render(); });
  await finish(page, 'topology');
  await page.getByRole('button', { name: 'Retry request' }).waitFor();
  assert.equal(await page.locator('textarea').inputValue(), 'Add SQL');
  assert.equal(await page.locator('.arch-chat-msg-assistant').count(), 0);
  assert.equal(await page.evaluate(() => (window as any).h.followups), 0);
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][3]), 7);
});

test('chat cancellation aborts the request; late results cannot apply', async t => {
  const page = await setup(t, { surface: 'chat', deferKinds: ['topology'] });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.getByRole('button', { name: 'Cancel request' }).click();
  assert.equal(await page.evaluate(() => (window as any).h.calls[0].signal.aborted), true);
  await finish(page, 'topology');
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.locator('textarea').inputValue(), 'Add SQL');
});

test('chat summaries and follow-ups describe only the accepted graph', async t => {
  const page = await setup(t, {
    surface: 'chat', accepted: true,
    acceptedNodes: [{ id: 'sql', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'SQL' } }],
  });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.locator('.arch-chat-msg-assistant').filter({ hasText: 'Added SQL.' }).waitFor();
  assert.equal(await page.locator('textarea').inputValue(), '');
  assert.deepEqual(await page.evaluate(() => (window as any).h.followupInputs[0].services), ['SQL']);
});

test('chat cancellation during pending review and unmount reaches the parent callback', async t => {
  const page = await setup(t, { surface: 'chat', deferReview: true });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.applies.length === 1);
  await page.evaluate(() => (window as any).h.unmount());
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][4].aborted), true);
  assert.equal(await page.evaluate(() => (window as any).h.followups), 0);
});

test('desktop chat stays nonmodal and allows focus outside the panel', async t => {
  const page = await setup(t, { surface: 'chat' });
  await page.setViewportSize({ width: 1600, height: 1000 });
  await page.getByRole('complementary', { name: 'Architecture chat' }).waitFor();
  assert.equal(await page.getByRole('dialog').count(), 0);
  assert.equal(await page.locator('.arch-chat-backdrop').count(), 0);
  assert.equal(await page.locator('.arch-chat-panel').getAttribute('data-modal'), 'false');
  await page.locator('#outside').focus();
  await page.waitForTimeout(160);
  assert.equal(await page.locator('#outside').evaluate(element => element === document.activeElement), true);
  assert.equal(await page.locator('#root').getAttribute('aria-hidden'), null);
});

for (const [width, dismissal] of [[1180, 'backdrop'], [900, 'Escape'], [390, 'backdrop']] as const) {
  test(`compact chat at ${width}px contains focus and ${dismissal} dismissal cancels pending work`, async t => {
    const page = await setup(t, { surface: 'chat', open: false, deferKinds: ['topology'] });
    await page.setViewportSize({ width, height: 800 });
    await page.locator('#outside').focus();
    await page.evaluate(() => { (window as any).h.open = true; (window as any).h.render(); });
    const dialog = page.getByRole('dialog', { name: 'Architecture chat' });
    await dialog.waitFor();
    assert.equal(await dialog.getAttribute('aria-modal'), 'true');
    assert.equal(await dialog.getAttribute('data-modal'), 'true');
    await page.locator('.arch-chat-backdrop').waitFor();
    await page.waitForFunction(() => document.querySelector('.arch-chat-panel')?.contains(document.activeElement));
    assert.equal(await page.locator('.app-header').getAttribute('inert'), '');
    assert.equal(await page.locator('.workspace').getAttribute('inert'), '');
    await page.locator('textarea').fill('Add SQL');
    await page.getByRole('button', { name: 'Send', exact: true }).click();
    await page.waitForFunction(() => (window as any).h.calls.length === 1);
    await page.keyboard.press('Tab');
    assert.equal(await dialog.evaluate(element => element.contains(document.activeElement)), true);
    await page.keyboard.press('Shift+Tab');
    assert.equal(await dialog.evaluate(element => element.contains(document.activeElement)), true);
    if (dismissal === 'Escape') await page.keyboard.press('Escape');
    else await page.locator('.arch-chat-backdrop').click();
    await dialog.waitFor({ state: 'hidden' });
    await page.waitForFunction(() => document.activeElement?.id === 'outside');
    assert.equal(await page.locator('.app-header').getAttribute('inert'), null);
    assert.equal(await page.locator('.workspace').getAttribute('inert'), null);
    assert.equal(await page.evaluate(() => (window as any).h.calls[0].signal.aborted), true);
    await finish(page, 'topology');
    assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
    assert.equal(await page.evaluate(() => (window as any).h.followups), 0);
  });
}

test('switching diagrams cancels chat work and cannot populate the new thread', async t => {
  const page = await setup(t, { surface: 'chat', deferKinds: ['topology'] });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.evaluate(() => { (window as any).h.diagramKey = 'diagram-b'; (window as any).h.render(); });
  await page.waitForFunction(() => (window as any).h.calls[0].signal.aborted);
  await finish(page, 'topology');
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.locator('.arch-chat-msg').count(), 0);
  assert.equal(await page.locator('textarea').inputValue(), '');
});

test('chat preserves manual edits made while generation runs', async t => {
  const page = await setup(t, { surface: 'chat', deferKinds: ['topology'] });
  await page.locator('textarea').fill('Add SQL');
  await page.getByRole('button', { name: 'Send', exact: true }).click();
  await page.evaluate(() => {
    const h = (window as any).h;
    h.nodes = [{ id: 'manual', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Manual edit' } }];
    h.render();
  });
  await finish(page, 'topology');
  await page.locator('.arch-chat-msg-error').filter({ hasText: 'diagram changed' }).waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.evaluate(() => (window as any).h.followups), 0);
  assert.equal(await page.locator('textarea').inputValue(), 'Add SQL');
});

const comparisonModels = ['test-model', 'second-model', 'gpt-6-astra', 'gpt-5.6-sol'];

async function startComparison(page: Page, surface: string) {
  if (surface === 'models') await page.locator('textarea').fill('Add a secure web application');
  await page.getByRole('button', {
    name: surface === 'models' ? 'Compare 4 Models' : 'Compare Validation Across 4 Models', exact: true,
  }).click();
}

for (const surface of ['models', 'validation']) {
  const kind = surface === 'models' ? 'topology' : 'validation';
  const comparisonConfig = { surface, availableModels: comparisonModels, deferKinds: [kind, 'critique'] };

  for (const limit of [1, 2]) {
    test(`${surface} comparison queue dispatches four models at cap ${limit}, preserves progress, and continues after an error`, async t => {
      const page = await setup(t, { ...comparisonConfig, budgetLimit: limit });
      await startComparison(page, surface);
      await page.waitForFunction(limit => (window as any).h.calls.length === limit, limit);
      assert.equal(await page.locator('.compare-result-card.running').count(), limit);
      assert.equal(await page.locator('.compare-result-card.pending').count(), 4 - limit);
      for (let index = 0; index < 4; index += 1) {
        await finish(page, kind, index === 1 ? 'The second model failed.' : undefined);
      }
      await page.locator('.compare-progress').waitFor({ state: 'hidden' });
      assert.equal(await page.locator('.compare-result-card.success').count(), 3);
      assert.equal(await page.locator('.compare-result-card.error').count(), 1);
      assert.equal(await page.locator('.compare-result-card.pending, .compare-result-card.running').count(), 0);
      assert.match(await page.locator('.compare-result-error-msg').innerText(), /second model failed/);
      const state = await page.evaluate(() => {
        const h = (window as any).h;
        return { models: h.calls.map((call: any) => call.override.model),
          managed: h.calls.every((call: any) => call.override.forceManaged === true && !!call.signal),
          peak: h.peakRequests, violations: h.dispatchViolations, telemetry: h.telemetry };
      });
      assert.deepEqual(state.models, comparisonModels);
      assert.equal(state.managed, true);
      assert.equal(state.peak, limit);
      assert.equal(state.violations, 0);
      if (surface === 'validation') {
        assert.equal(state.telemetry.filter((event: any) => event.kind === 'compared').length, 1);
        assert.equal(state.telemetry.find((event: any) => event.kind === 'compared').modelCount, 3);
      }
    });
  }

  for (const [name, configuration, message] of [
    ['unavailable budget', { budgetFailure: 'Budget service offline.' }, /budget could not be checked/],
    ['zero cap', { budgetLimit: 0 }, /budget is invalid/],
    ['fractional cap', { budgetLimit: 1.5 }, /budget is invalid/],
    ['negative occupancy', { budgetExternal: -1 }, /budget is invalid/],
  ] as const) {
    test(`${surface} comparison queue surfaces ${name} and dispatches no models`, async t => {
      const page = await setup(t, { ...comparisonConfig, ...configuration });
      await startComparison(page, surface);
      await page.getByRole('alert').waitFor();
      assert.match(await page.getByRole('alert').innerText(), message);
      assert.equal(await page.locator('.compare-result-card.error').count(), 4);
      assert.equal(await page.evaluate(() => (window as any).h.calls.length), 0);
      assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), 0);
    });
  }

  test(`${surface} comparison queue preserves completed results if its budget endpoint fails later`, async t => {
    const page = await setup(t, { ...comparisonConfig, budgetLimit: 1 });
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.calls.length === 1);
    await page.evaluate(() => { (window as any).h.budgetFailure = 'Budget service offline.'; });
    await finish(page, kind);
    await page.getByRole('alert').waitFor();
    assert.match(await page.getByRole('alert').innerText(), /budget could not be checked/);
    assert.equal(await page.locator('.compare-result-card.success').count(), 1);
    assert.equal(await page.locator('.compare-result-card.error').count(), 3);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 1);
    assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), 0);
  });

  test(`${surface} comparison queue waits on occupied capacity and cancels without further polling`, async t => {
    const page = await setup(t, { ...comparisonConfig, budgetLimit: 1, budgetExternal: 1 });
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.budgetReads.length >= 1);
    assert.equal(await page.locator('.compare-result-card.pending').count(), 4);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 0);
    await page.getByRole('button', { name: 'Cancel comparison', exact: true }).click();
    const reads = await page.evaluate(() => (window as any).h.budgetReads.length);
    await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 1200)));
    assert.equal(await page.evaluate(() => (window as any).h.budgetReads.length), reads);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 0);
    assert.equal(await page.locator('.compare-result-card.error').count(), 4);
    assert.match(await page.getByRole('alert').innerText(), /Comparison cancelled/);
  });

  test(`${surface} comparison queue aborts a pending budget lookup on close and ignores its late response`, async t => {
    const page = await setup(t, { ...comparisonConfig, deferBudget: true });
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.budgetPending.length === 1);
    await page.getByRole('button', { name: 'Close', exact: true }).click();
    assert.equal(await page.evaluate(() => (window as any).h.budgetPending[0].signal.aborted), true);
    await page.evaluate(() => (window as any).h.budgetPending[0].resolve());
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 0);
    assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), 0);
  });

  for (const dismiss of ['close', 'escape', 'backdrop', 'prop', 'unmount']) {
    test(`${surface} comparison queue ${dismiss} aborts running models and suppresses queued and late results`, async t => {
      const page = await setup(t, comparisonConfig);
      await startComparison(page, surface);
      await page.waitForFunction(() => (window as any).h.calls.length === 2);
      if (dismiss === 'close') await page.getByRole('button', { name: 'Close', exact: true }).click();
      else if (dismiss === 'escape') await page.keyboard.press('Escape');
      else if (dismiss === 'backdrop') await page.locator('.modal-overlay').dispatchEvent('click');
      else if (dismiss === 'unmount') await page.evaluate(() => (window as any).h.unmount());
      else await page.evaluate(() => { const h = (window as any).h; h.open = false; h.render(); });
      await page.getByRole('dialog').waitFor({ state: 'hidden' });
      assert.equal(await page.evaluate(() => (window as any).h.calls.every((call: any) => call.signal?.aborted)), true);
      await page.evaluate(() => (window as any).h.pending.splice(0).forEach((item: any) => item.resolve()));
      assert.equal(await page.evaluate(() => (window as any).h.calls.length), 2);
      assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), 0);
    });
  }

  test(`${surface} comparison queue cancellation keeps partial success but never reports the cancelled batch as complete`, async t => {
    const page = await setup(t, comparisonConfig);
    await startComparison(page, surface);
    await finish(page, kind);
    await page.waitForFunction(() => (window as any).h.calls.length === 3);
    await page.getByRole('button', { name: 'Cancel comparison', exact: true }).click();
    await page.evaluate(() => (window as any).h.pending.splice(0).forEach((item: any) => item.resolve()));
    assert.equal(await page.locator('.compare-result-card.success').count(), 1);
    assert.equal(await page.locator('.compare-result-card.error').count(), 3);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 3);
    assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), 0);
    assert.match(await page.getByRole('alert').innerText(), /Completed results were kept/);
  });

  test(`${surface} comparison queue ignores an old response after cancellation and a fresh comparison`, async t => {
    const page = await setup(t, { ...comparisonConfig, budgetLimit: 1 });
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.calls.length === 1);
    await page.getByRole('button', { name: 'Cancel comparison', exact: true }).click();
    await page.getByRole('button', { name: 'New Comparison', exact: true }).click();
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.calls.length === 2);
    await finish(page, kind);
    assert.equal(await page.locator('.compare-result-card.success').count(), 0);
    assert.equal(await page.locator('.compare-result-card.running').count(), 1);
    assert.equal(await page.locator('.compare-result-card.pending').count(), 3);
    for (let index = 0; index < 4; index += 1) await finish(page, kind);
    await page.locator('.compare-progress').waitFor({ state: 'hidden' });
    assert.equal(await page.locator('.compare-result-card.success').count(), 4);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 5);
    if (surface === 'validation') {
      assert.equal(await page.evaluate(() => (window as any).h.telemetry.filter((event: any) => event.kind === 'compared').length), 1);
    }
  });

  test(`${surface} comparison critique uses the same budget, then cancellation blocks late critique and telemetry`, async t => {
    const page = await setup(t, comparisonConfig);
    await startComparison(page, surface);
    for (let index = 0; index < 4; index += 1) await finish(page, kind);
    await page.locator('.compare-progress').waitFor({ state: 'hidden' });
    const telemetryCount = await page.evaluate(() => (window as any).h.telemetry.length);
    await page.evaluate(() => { (window as any).h.budgetExternal = 2; });
    await page.getByRole('button', { name: 'Generate AI Critique', exact: true }).click();
    await page.getByRole('button', { name: 'Cancel critique', exact: true }).waitFor();
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 4);
    await page.evaluate(() => { (window as any).h.budgetExternal = 0; });
    await page.waitForFunction(() => (window as any).h.calls.length === 5);
    assert.equal(await page.evaluate(() => (window as any).h.calls[4].override.forceManaged), true);
    await page.getByRole('button', { name: 'Cancel critique', exact: true }).click();
    assert.equal(await page.evaluate(() => (window as any).h.calls[4].signal.aborted), true);
    await finish(page, 'critique');
    assert.equal(await page.getByRole('button', { name: 'Save Critique', exact: true }).count(), 0);
    assert.match(await page.locator('.compare-critique-error').innerText(), /Critique cancelled/);
    assert.equal(await page.evaluate(() => (window as any).h.telemetry.length), telemetryCount);
  });

  test(`${surface} comparison queue requeues a capacity race without downgrading a model or reporting failure`, async t => {
    const page = await setup(t, {
      ...comparisonConfig, budgetLimit: 1, contentionRemaining: 1, occupyOnContention: true,
    });
    await startComparison(page, surface);
    await page.waitForFunction(() => (window as any).h.calls.length === 1);
    assert.equal(await page.locator('.compare-result-card.error').count(), 0);
    assert.equal(await page.locator('.compare-result-card.pending').count(), 4);
    await page.evaluate(() => { (window as any).h.budgetExternal = 0; });
    for (let index = 0; index < 4; index += 1) await finish(page, kind);
    await page.locator('.compare-progress').waitFor({ state: 'hidden' });
    assert.equal(await page.locator('.compare-result-card.success').count(), 4);
    const models = await page.evaluate(() => (window as any).h.calls.map((call: any) => call.override.model));
    assert.deepEqual(models, ['test-model', 'second-model', 'gpt-6-astra', 'gpt-5.6-sol', 'test-model']);
    assert.equal(await page.evaluate(() => (window as any).h.dispatchViolations), 0);
  });
}

test('validation comparison queue captures fingerprints before capacity waits and never restamps later queued reviews', async t => {
  const page = await setup(t, {
    surface: 'validation', availableModels: comparisonModels, budgetLimit: 1, budgetExternal: 1,
  });
  await startComparison(page, 'validation');
  await page.waitForFunction(() => (window as any).h.budgetReads.length >= 1);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.diagramFingerprint = 'fingerprint-b';
    h.budgetExternal = 0;
    h.render();
  });
  for (let index = 0; index < 4; index += 1) await finish(page, 'validation');
  await page.locator('.compare-progress').waitFor({ state: 'hidden' });
  await page.evaluate(() => {
    const h = (window as any).h;
    h.diagramFingerprint = 'fingerprint-c';
    h.render();
  });
  for (let index = 0; index < 4; index += 1) {
    await page.getByRole('button', { name: 'Use This Validation', exact: true }).nth(index).click();
  }
  assert.deepEqual(await page.evaluate(() => (window as any).h.applies.map((args: any[]) => args[1])),
    Array(4).fill('fingerprint-a'));
  assert.equal(await page.getByRole('dialog').count(), 1);
  assert.match(await page.getByRole('alert').innerText(), /could not be applied/);
});

test('validation comparisons preserve their request fingerprint and keep rejected reviews open', async t => {
  const page = await setup(t, { surface: 'validation' });
  await page.getByRole('button', { name: 'Compare Validation Across 2 Models' }).click();
  await page.waitForFunction(() => (window as any).h.validationPending.length === 2);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.diagramFingerprint = 'fingerprint-b';
    h.render();
    h.validationPending.splice(0).forEach((finish: () => void) => finish());
  });
  await page.getByRole('button', { name: 'Use This Validation' }).first().click();
  await page.getByRole('alert').filter({ hasText: 'could not be applied' }).waitFor();
  assert.equal(await page.getByRole('dialog', { name: 'Compare Validation', exact: true }).count(), 1);
  assert.equal(await page.evaluate(() => (window as any).h.applies[0][1]), 'fingerprint-a');
  await page.evaluate(() => { (window as any).h.accepted = true; });
  await page.getByRole('button', { name: 'Use This Validation' }).first().click();
  assert.equal(await page.evaluate(() => (window as any).h.applies[1][1]), 'fingerprint-a');
  await page.waitForFunction(() => !(window as any).h.open);
});

test('model comparison keeps rejected reviews open and preserves source model metadata', async t => {
  const page = await setup(t, { surface: 'models' });
  await page.locator('textarea').fill('Add a secure web application');
  await page.getByRole('button', { name: 'Compare 2 Models', exact: true }).click();
  await page.getByRole('button', { name: 'Use This Architecture' }).first().click();
  await page.waitForFunction(() => (window as any).h.applies.length === 1);
  assert.equal(await page.getByRole('dialog').count(), 1);
  assert.deepEqual(await page.evaluate(() => (window as any).h.applies[0].slice(1)),
    ['Add a secure web application', 'test-model', 'none']);
  await page.evaluate(() => { (window as any).h.accepted = true; });
  await page.getByRole('button', { name: 'Use This Architecture' }).first().click();
  await page.waitForFunction(() => !(window as any).h.open);
});

for (const surface of ['models', 'validation']) {
  for (const astraAvailable of [true, false]) {
    test(`${surface} comparison ${astraAvailable ? 'prefers and labels Astra' : 'preserves the pre-Astra fallback'} using configured deployments`, async t => {
      const page = await setup(t, {
        surface,
        availableModels: astraAvailable ? ['gpt-5.6-sol', 'gpt-6-astra'] : ['test-model', 'gpt-5.6-sol'],
      });
      assert.equal(await page.locator('.compare-model-chip.selected').count(), 2);
      if (surface === 'models') {
        await page.locator('textarea').fill('Add a secure web application');
        await page.getByRole('button', { name: 'Compare 2 Models', exact: true }).click();
      } else {
        await page.getByRole('button', { name: 'Compare Validation Across 2 Models', exact: true }).click();
        await page.waitForFunction(() => (window as any).h.validationPending.length === 2);
        await page.evaluate(() => (window as any).h.validationPending.splice(0).forEach((finish: () => void) => finish()));
      }
      await page.locator('.compare-critique-model-select').waitFor();
      assert.equal(await page.locator('.compare-critique-model-select').inputValue(),
        astraAvailable ? 'gpt-6-astra' : 'gpt-5.6-sol');
      await page.evaluate(() => {
        HTMLAnchorElement.prototype.click = function () {
          (window as any).h.savedReport = JSON.parse(decodeURIComponent(this.href.split(',')[1]));
        };
      });
      await page.getByRole('button', { name: 'Save JSON', exact: true }).click();
      const results = await page.evaluate(() => (window as any).h.savedReport.results);
      if (astraAvailable) {
        assert.equal(results['gpt6astra-none'].model, 'gpt-6-astra');
        assert.equal(results['gpt6astra-none'].displayName, 'GPT-6 Astra');
        assert.equal(results.gpt6astra, undefined);
      } else {
        assert.equal(results.gpt6astra, undefined);
        assert.equal(results['gpt6astra-none'], undefined);
        assert.equal(results.gpt56sol.model, 'gpt-5.6-sol');
      }
    });
  }
}

test('review defaults to all, blocks invalid subsets, and applies exactly the chosen graph', async t => {
  const page = await setup(t, { surface: 'review' });
  const checkboxes = page.getByRole('checkbox');
  assert.equal(await checkboxes.count(), 2);
  assert.equal(await checkboxes.nth(0).isChecked(), true);
  assert.equal(await checkboxes.nth(1).isChecked(), true);
  await checkboxes.nth(0).uncheck();
  assert.equal(await page.getByRole('button', { name: 'Apply selected changes' }).isDisabled(), true);
  await page.getByRole('alert').waitFor();
  await checkboxes.nth(0).check();
  await checkboxes.nth(1).uncheck();
  await page.getByRole('button', { name: 'Apply selected changes' }).click();
  assert.deepEqual(await page.evaluate(() => {
    const graph = (window as any).h.reviewed[0];
    return { nodes: graph.nodes.map((node: any) => node.id), edges: graph.edges.length };
  }), { nodes: ['a', 'b'], edges: 0 });
  assert.equal(await page.getByRole('dialog', { name: 'Review AI changes' }).count(), 1,
    'a void onApply callback must not close the parent-controlled review');
});

test('review retains selection after parent snapshot failure and can retry the same proposal', async t => {
  const page = await setup(t, { surface: 'review', snapshotError: true });
  const dialog = page.getByRole('dialog', { name: 'Review AI changes' });
  await dialog.getByRole('checkbox').nth(1).uncheck();
  await dialog.getByRole('button', { name: 'Apply selected changes' }).click();
  await dialog.getByRole('alert').filter({ hasText: 'Snapshot failed' }).waitFor();
  assert.equal(await dialog.count(), 1);
  assert.equal(await page.evaluate(() => (window as any).h.reviewed.length), 0);
  assert.equal(await dialog.getByRole('checkbox').nth(0).isChecked(), true);
  assert.equal(await dialog.getByRole('checkbox').nth(1).isChecked(), false);
  await page.evaluate(() => { (window as any).h.snapshotError = false; });
  await dialog.getByRole('button', { name: 'Apply selected changes' }).click();
  await page.waitForFunction(() => (window as any).h.reviewed.length === 1);
  assert.equal(await page.evaluate(() => (window as any).h.reviewAttempts), 2);
  assert.deepEqual(await page.evaluate(() => {
    const graph = (window as any).h.reviewed[0];
    return { nodes: graph.nodes.map((node: any) => node.id), edges: graph.edges.length };
  }), { nodes: ['a', 'b'], edges: 0 });
  assert.equal(await dialog.count(), 1);
});

test('review keyboard dismissal cancels without applying a graph', async t => {
  const page = await setup(t, { surface: 'review' });
  await page.getByRole('dialog').waitFor();
  await page.keyboard.press('Escape');
  assert.equal(await page.evaluate(() => (window as any).h.cancelled), 1);
  assert.equal(await page.evaluate(() => (window as any).h.reviewed.length), 0);
});

const featureNames = [
  'Architecture Generation', 'Architecture Validation', 'Deployment Guide & Bicep', 'Blueprint Diagrams',
];

for (const profile of ['astra', 'legacy', 'none']) {
  for (const surface of ['popover', 'compact', 'selector']) {
    test(`${surface} guidance resolves the ${profile} deployment portfolio and applies exactly that profile`, async t => {
      const page = await setupModelSettings(t, profile, {
        surface,
        initialSettings: profile === 'none' ? undefined : {
          model: 'gpt-5.6-terra', reasoningEffort: 'high',
          featureOverrides: { blueprint: { model: 'gpt-5.6-sol', reasoningEffort: 'medium' } },
        },
      });
      await openModelGuidance(page, surface);
      const apply = page.locator(surface === 'popover' ? '.msp-portfolio-btn'
        : surface === 'compact' ? '.recommended-portfolio-compact' : '.recommended-portfolio');
      const recommendation = await page.evaluate(() => (window as any).h.getRecommendedModelSettings());
      const summary = surface === 'popover' ? page.locator('.msp-portfolio-copy > span')
        : surface === 'compact' ? page.locator('.compact-advanced-footer') : undefined;

      if (profile === 'none') {
        assert.equal(await apply.isDisabled(), true);
        assert.doesNotMatch(await page.locator('#root').innerText(), /GPT-6 Astra|GPT-5\.6/);
        assert.match(summary ? await summary.innerText() : await page.locator('.feature-recommended').first().innerText(),
          /No managed models are configured\./);
        return;
      }

      const models = profile === 'astra'
        ? Array<string>(4).fill('GPT-6 Astra')
        : ['GPT-5.6 Sol', 'GPT-5.6 Terra', 'GPT-5.6 Terra', 'GPT-5.6 Luna'];
      if (summary) {
        const text = await summary.innerText();
        if (profile === 'astra') assert.equal(text, 'All features: GPT-6 Astra (Low)');
        else featureNames.forEach((name, index) => assert.ok(text.includes(`${name}: ${models[index]} (Low)`)));
      }
      if (surface === 'compact') {
        assert.deepEqual(await page.locator('.help-defaults > span').allTextContents(),
          featureNames.map((name, index) => `• ${name}: ${models[index]} (Low)`));
      } else if (surface === 'selector') {
        assert.deepEqual(await page.locator('.feature-recommended').allTextContents(),
          models.map(model => `Recommended: ${model} (Low)`));
      }

      if (profile === 'legacy') {
        assert.doesNotMatch(await page.locator('#root').innerText(), /GPT-6 Astra/);
      } else if (surface !== 'compact') {
        const cards = page.locator(surface === 'popover' ? '.msp-model-btn' : '.model-button');
        const role = surface === 'popover' ? '.msp-model-role' : 'small';
        assert.equal(await cards.filter({ hasText: 'GPT-6 Astra' }).locator(role).innerText(), 'Recommended for all features');
        for (const model of ['GPT-5.6 Sol', 'GPT-5.6 Terra', 'GPT-5.6 Luna']) {
          assert.equal(await cards.filter({ hasText: model }).locator(role).innerText(), 'Alternative model');
        }
      }
      assert.notDeepEqual(await page.evaluate(() => (window as any).h.getModelSettings()), recommendation);
      await apply.click();
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettings()), recommendation);
      if (surface === 'popover') {
        assert.deepEqual(await page.locator('.msp-feature-effective').allTextContents(), models.map(model => `${model} (Low)`));
      }
    });
  }
}

for (const connectionState of ['verified', 'key-required']) {
  test(`Astra recommendations do not replace a ${connectionState} BYO selection or its paused managed settings`, async t => {
    const initialSettings = { model: 'gpt-5.6-terra', reasoningEffort: 'high', featureOverrides: {} };
    const byo = {
      settings: { enabled: true, model: 'customer-owned-model', reasoningEffort: 'medium' },
      verified: connectionState === 'verified', connectionState,
    };
    const page = await setupModelSettings(t, 'astra', { initialSettings, byo });
    await openModelGuidance(page, 'popover');
    assert.equal(await page.locator('.model-popover-label').innerText(), 'Custom: customer-owned-model');
    assert.equal(await page.locator('.msp-portfolio-btn').isDisabled(), true);
    for (const control of await page.locator('.msp-model-btn, .msp-reasoning-btn, .msp-feature-select').all()) {
      assert.equal(await control.isDisabled(), true);
    }
    await page.locator('.msp-portfolio-btn').evaluate((button: HTMLButtonElement) => button.click());
    assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettings()), initialSettings);
    await page.getByRole('button', { name: connectionState === 'verified' ? 'Configure' : 'Enter key', exact: true }).click();
    assert.deepEqual(await page.evaluate(() => {
      const h = (window as any).h;
      return { byo: h.byo, opened: h.byoSettingsOpened, settings: h.getModelSettings() };
    }), { byo, opened: 1, settings: initialSettings });
  });
}

test('Astra model guidance and vision errors have real Japanese translations without legacy advice', async t => {
  const page = await setupModelSettings(t, 'astra');
  await openModelGuidance(page, 'popover');
  await page.evaluate(() => (window as any).h.setLanguage('ja'));
  await page.getByText('すべての機能: GPT-6 Astra (低)', { exact: true }).waitFor();
  const astra = page.locator('.msp-model-btn').filter({ hasText: 'GPT-6 Astra' });
  assert.equal(await astra.locator('.msp-model-role').innerText(), 'すべての機能で推奨');
  assert.match((await astra.getAttribute('title'))!, /アーキテクチャ設計/);
  assert.equal(await page.locator('.msp-model-btn').filter({ hasText: 'GPT-5.6 Terra' })
    .locator('.msp-model-role').innerText(), '代替モデル');
  const message = await page.evaluate(() => (window as any).h.translate(
    'The selected model may not support image analysis. Choose a vision-capable model in AI settings. Request ID: vision-request',
  ));
  assert.match(message, /画像分析/);
  assert.match(message, /vision-request/);
  assert.doesNotMatch(message, /Try using|GPT-/);
});

for (const modelName of ['GPT-6 Astra (Low)', 'GPT-5.6 Sol (Medium)']) {
  test(`model badges preserve ${modelName} provenance even when Astra is configured`, async t => {
    const page = await setupModelSettings(t, 'astra', { surface: 'badge', modelName });
    assert.equal(await page.locator('.model-generation-badge-text strong').innerText(), modelName);
    assert.match(await page.locator('.model-generation-badge-time').innerText(), /1\.2s/);
  });
}
