import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { build } from 'esbuild';
import { chromium, type Browser, type Page } from 'playwright';

let browser: Browser;
let script: string;
const modelSettingsScripts = new Map<string, string>();
before(async () => {
  const mocks: Record<string, string> = {
    modelSettingsStore: `
      const settings = {model:'gpt-6-astra', reasoningEffort:'none'};
      export const useModelSettings = () => [window.h.settings ?? settings];
      export const getModelSettings = () => settings;
      export const getModelSettingsForFeature = feature => ({
        ...(window.h.settings ?? settings), ...window.h.settings?.featureOverrides?.[feature],
      });
      export const getAvailableModels = () => window.h.configured === false ? [] : ['gpt-6-astra'];
      export const MODEL_CONFIG = {
        'gpt-6-astra': {displayName:'GPT-6 Astra', apiFormat:'responses', isReasoning:true, maxCompletionTokens:32000}
      };
      export const FEATURE_CONFIG = {architectureGeneration:{displayName:'Topology',recommendedModel:'gpt-6-astra'},blueprint:{displayName:'Blueprint',recommendedModel:'gpt-6-astra'}};
      export class AIModelConfigurationError extends Error {
        constructor(code, message) { super(message); this.name='AIModelConfigurationError'; this.code=code; this.source='client'; this.retryable=false; }
      }
      export const isModelAvailable = () => window.h.configured !== false;
      export const getDeploymentName = () => {
        if(!isModelAvailable()) throw new AIModelConfigurationError('astra_not_configured','GPT-6 Astra is not configured.');
        return 'test-gpt-6-astra';
      };
      export const updateFeatureOverride = () => {};
      export const getSupportedReasoningEfforts = () => ['none','max'];
      export const getReasoningEffortLabel = x => x;
      export const isReasoningEffort = value => ['none','minimal','low','medium','high','xhigh','max'].includes(value);
    `,
    LanguageContext: `export const useLanguage = () => ({language:window.h.language ?? 'en',t:x=>x,translate:x=>x});`,
    byoAISettingsStore: `
      export const getBYOAISettings = () => window.h.byo ?? {profiles:[],activeProfileId:null};
      export const useBYOAISettings = () => ({settings:getBYOAISettings()});
      export const getBYOAIConnectionState = id => window.h.connectionStates?.[id] ?? {status:'key-required',hasApiKey:false,verified:false,revision:0};
      export const selectBYOAIProfile = id => {window.h.byo={...getBYOAISettings(),activeProfileId:id};window.h.render();};
      export const normalizeBYOAIEndpoint = (provider, value) => value;
      export const validateBYOAIProfile = profile => ({valid:true,profile});
    `,
    runtimeConfig: `
      export const getRuntimeConfigSnapshot = () => ({status:'ready',bringYourOwnAI:true});
      export const useRuntimeConfig = getRuntimeConfigSnapshot;
      export const isBYOAIEnabledOnServer = () => true;
      export const loadRuntimeConfig = async () => getRuntimeConfigSnapshot();
      export const runtimeConfigCancellationError = () => new DOMException('Cancelled','AbortError');
      export const awaitWithAISignal = (work, signal) => signal.aborted
        ? Promise.reject(runtimeConfigCancellationError()) : work;
    `,
    aiModelRuntime: `
      export const getEffectiveAIModelInfo = feature => {
        const settings={model:'gpt-6-astra',reasoningEffort:'none',...window.h.settings,
          ...window.h.settings?.featureOverrides?.[feature]};
        return window.h.connectionInfo ?? {source:'managed',model:'gpt-6-astra',displayName:'GPT-6 Astra',
          reasoningEffort:settings.reasoningEffort,apiFormat:'responses',isReasoning:true,
          supportsVision:true,maxCompletionTokens:32000,ready:window.h.configured!==false,
          ...(window.h.configured===false?{code:'astra_not_configured'}:{})};
      };
      export const captureRuntimeModelOverride = feature => {
        const info=getEffectiveAIModelInfo(feature);
        if(!info.ready) throw new Error('The selected AI connection is not ready.');
        return {model:'gpt-6-astra',reasoningEffort:info.reasoningEffort,
          connection:Object.freeze({...info,feature,revision:window.h.connectionRevision??0})};
      };
      export const assertCapturedAIConnectionCurrent = () => {};
    `,
    safeStorage: `
      export const readLocalStorage = key => key==='aiGenerator.mode' ? window.h.mode : null;
      export const readBooleanPreference = (key, fallback) => key==='aiGenerator.bothInParallel' ? window.h.parallel : fallback;
      export const writeLocalStorage = () => {};
    `,
    telemetryService: `
      export const trackImageImport = () => {};
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
      export const isAzureOpenAIConfigured = () => window.h.configured !== false;
      export const throwIfGenerationAborted = signal => {if(signal?.aborted) throw new DOMException('Cancelled','AbortError');};
      export const generateArchitectureWithAI = (...args) => window.h.generate('topology', args[4]?.signal || args[4] || args[1]?.signal, args[0], args[1]);
      export const generateFollowUpSuggestions = input => {window.h.followups++; window.h.followupInputs.push(input); return Promise.resolve(['Add monitoring']);};
      export const analyzeArchitectureDiagramImage = () => Promise.resolve({description:'image'});
    `,
    referenceArchitectureAI: `export const generateReferenceArchitectureWithAI = (prompt, settings) => window.h.generate('reference', settings.signal, prompt);`,
    blueprintArchitectureAI: `export const generateBlueprintArchitectureWithAI = (prompt, settings) => window.h.generate('blueprint', settings.signal, prompt, settings);`,
    componentManifestAI: `export const generateComponentManifest = (prompt, settings) => window.h.generate('manifest', settings.signal, prompt, settings);`,
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
        import {buildDiagramChanges} from './src/services/diagramChanges';
        import {runWithRateLimitRetry} from './src/services/aiRetry';
        import {createOpenAIProxyError} from './src/services/apiHelper';
        const h = window.h;
        const root = createRoot(document.getElementById('root'));
        h.dispatch = (kind, signal, prompt, override) => {
          h.calls.push({kind,signal,prompt,override});
          const budgetedBoth = h.surface==='generator' && h.mode==='both';
          if(budgetedBoth && (h.inFlight+h.budgetExternal >= h.budgetLimit || h.contentionRemaining > 0)) {
            if(h.inFlight+h.budgetExternal >= h.budgetLimit) h.dispatchViolations++;
            if(h.contentionRemaining > 0) {
              h.contentionRemaining--;
              if(h.occupyOnContention) h.budgetExternal=h.budgetLimit;
            }
            return Promise.reject(Object.assign(new Error('Capacity is occupied.'), {code:'ai_concurrency_limit',status:429}));
          }
          if(budgetedBoth) { h.inFlight++; h.peakRequests=Math.max(h.peakRequests,h.inFlight); }
          let released = false;
          const release = () => { if(budgetedBoth && !released) {released=true;h.inFlight--;} };
          signal?.addEventListener('abort',release,{once:true});
          const metrics={totalTokens:1,completionTokens:1,promptTokens:0,elapsedTimeMs:1};
          const value = {services:[],connections:[],groups:[],components:[],zones:[],metrics};
          let request;
          if(h.rateLimits[kind] > 0) {
            h.rateLimits[kind]--;
            request=Promise.reject(Object.assign(new Error('The AI provider is rate-limiting requests.'), {
              code:'azure_openai_rate_limited',status:429,retryAfterMs:h.retryAfterMs,
            }));
          }
          else if(h.deferKinds.includes(kind)) request = new Promise((resolve,reject) => {
            const cleanup = () => {
              h.pending=h.pending.filter(item=>item!==entry);
            };
            const entry={kind,override,resolve:()=>{cleanup();resolve(value);},
              reject:message=>{cleanup();reject(new Error(message));}};
            h.pending.push(entry);
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
        build.onResolve({ filter: /\/(modelSettingsStore|byoAISettingsStore|runtimeConfig|aiModelRuntime|LanguageContext|safeStorage|telemetryService|aiBudgetService|ImageUploader|azureOpenAI|referenceArchitectureAI|blueprintArchitectureAI|componentManifestAI|exportReferencePng|exportBlueprintPng)$/ },
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
    astra: { ...legacyDeployments, VITE_AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com', VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA: 'gpt-6-astra' },
    legacy: { ...legacyDeployments, VITE_AZURE_OPENAI_ENDPOINT: 'https://example.openai.azure.com' },
    none: {},
  };
  await Promise.all(Object.entries(environments).map(async ([profile, environment]) => {
    const settingsBundle = await build({
      stdin: {
        contents: `
          import React from 'react';
          import {createRoot} from 'react-dom/client';
          import {flushSync} from 'react-dom';
          import Popover from './src/components/ModelSettingsPopover';
          import Connections from './src/components/BYOAISettingsDialog';
          import Validation from './src/components/ValidationModal';
          import Badge from './src/components/ModelBadge';
          import {LanguageProvider, useLanguage} from './src/i18n/LanguageContext';
          import {getModelSettings, getModelSettingsForFeature, getRecommendedModelSettings,
            updateModelSettings} from './src/stores/modelSettingsStore';
          import {getBYOAISettings, getBYOAIConnectionState, reloadBYOAISettings,
            upsertBYOAIProfile, selectBYOAIProfile} from './src/stores/byoAISettingsStore';
          import {getEffectiveAIModelInfo, captureRuntimeModelOverride} from './src/services/aiModelRuntime';
          import {loadRuntimeConfig} from './src/services/runtimeConfig';
          const h = window.h;
          if(h.surface==='validation-download') {
            h.downloads=[];
            const blobs=new Map();
            const originalCreate=URL.createObjectURL.bind(URL);
            URL.createObjectURL=blob=>{const url=originalCreate(blob);blobs.set(url,blob);return url;};
            HTMLAnchorElement.prototype.click=function(){
              const entry={filename:this.download,href:this.href,text:null};
              h.downloads.push(entry);
              const blob=blobs.get(this.href);
              if(blob) void blob.text().then(text=>{entry.text=text;});
            };
          }
          const originalFetch=window.fetch.bind(window);
          h.requests=[]; h.pendingTests=[];
          window.fetch=async (input,init) => {
            const url=String(input);
            if(url==='/api/runtime-config') return new Response(JSON.stringify({
              features:{bringYourOwnAI:h.allowBYO!==false}
            }),{headers:{'content-type':'application/json'}});
            if(url!=='/api/openai') return originalFetch(input,init);
            const payload=JSON.parse(init.body);
            h.requests.push({url,model:payload.deployment,body:payload.body,
              hasKey:!!payload.byo?.apiKey,
              provider:payload.byo?.provider,endpoint:payload.byo?.endpoint});
            const result=() => {
              if(h.testOutcome==='failed') return new Response(JSON.stringify({error:{
                code:'byo_authentication_failed',source:'bring_your_own',requestId:'byo-ui-test-123',
                message:'Rejected secret sk-ui-private-test-key https://secret.invalid/?key=sk-ui-private-test-key',
              }}),{status:401,headers:{'content-type':'application/json'}});
              const content=JSON.stringify({status:'ok'});
              const body=payload.apiFormat==='chat-completions'
                ? {choices:[{finish_reason:'stop',message:{content}}],usage:{prompt_tokens:1,completion_tokens:1,total_tokens:2}}
                : {status:h.testOutcome==='incomplete'?'incomplete':'completed',
                  output:[{type:'message',status:'completed',content:[{type:'output_text',text:content}]}],
                  usage:{input_tokens:1,output_tokens:1,total_tokens:2}};
              return new Response(JSON.stringify(body),{headers:{'content-type':'application/json'}});
            };
            if(h.deferTest) return new Promise(resolve=>h.pendingTests.push({signal:init.signal,resolve:()=>resolve(result())}));
            return result();
          };
          if(h.initialSettings) updateModelSettings(h.initialSettings);
          h.getModelSettings = getModelSettings;
          h.getModelSettingsForFeature = getModelSettingsForFeature;
          h.getRecommendedModelSettings = getRecommendedModelSettings;
          h.getBYOAISettings = getBYOAISettings;
          h.getBYOAIConnectionState = getBYOAIConnectionState;
          h.reloadBYOAISettings = reloadBYOAISettings;
          h.upsertBYOAIProfile = upsertBYOAIProfile;
          h.selectBYOAIProfile = selectBYOAIProfile;
          h.getEffectiveAIModelInfo = getEffectiveAIModelInfo;
          h.captureRuntimeModelOverride = captureRuntimeModelOverride;
          h.loadRuntimeConfig = loadRuntimeConfig;
          const root = createRoot(document.getElementById('root'));
          function Harness() {
            const language = useLanguage();
            h.setLanguage = language.setLanguage;
            h.translate = language.translate;
            if(h.surface === 'badge') return <Badge modelName={h.modelName} elapsedTimeMs={1200}/>;
            if(h.surface === 'validation-download') return <Validation validation={h.validation}
              isOpen={h.open} onClose={()=>{h.open=false;h.render();}}/>;
            if(h.surface === 'connections') return <>
              <button id="connections-opener" onClick={()=>{h.open=true;h.render();}}>Open AI connections</button>
              <Connections isOpen={h.open} onClose={()=>{h.open=false;h.render();}}
                returnFocusTarget={document.getElementById('connections-opener')}/>
            </>;
            return <Popover isOpen={h.open} onToggle={()=>{h.open=!h.open;h.render();}}
              onConfigureConnections={()=>{h.surface='connections';h.open=true;h.render();}}/>;
          }
          h.render = () => root.render(<LanguageProvider><Harness/></LanguageProvider>);
          h.commitRender = () => flushSync(h.render);
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
      open: true, diagramKey: 'diagram-a',
      deferKinds: [], failKinds: [], deferReview: false, calls: [], pending: [], applies: [],
      followups: 0, followupInputs: [], blueprints: 0, references: 0, exports: 0, reviewed: [], cancelled: 0,
      reviewAttempts: 0, reviewError: '', snapshotError: false,
      budgetLimit: 2, budgetExternal: 0, budgetReads: [], budgetFailure: null, deferBudget: false, budgetPending: [],
      inFlight: 0, peakRequests: 0, dispatchViolations: 0, contentionRemaining: 0,
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
  const page = await browser.newPage({ locale: 'en-US', hasTouch: Boolean(config.touch) });
  t.after(() => page.close());
  await page.route('https://model-settings.test/**', route => route.fulfill({
    contentType: 'text/html',
    body: '<html><body><div id="root"></div></body></html>',
  }));
  await page.goto('https://model-settings.test/');
  await page.evaluate(config => {
    (window as any).h = {
      surface: 'popover', open: true,
      ...config,
    };
    if (config.savedSettings) localStorage.setItem('azure-diagrams-model-settings', JSON.stringify(config.savedSettings));
    if (config.savedBYO) localStorage.setItem('azure-diagrams-byo-ai-settings', JSON.stringify(config.savedBYO));
  }, config);
  await page.addScriptTag({ content: modelSettingsScripts.get(profile)! });
  await page.locator('#root > *').first().waitFor();
  return page;
}
async function openModelGuidance(page: Page) {
  await page.locator('.msp-advanced-settings summary').click();
}

async function addComponentStyles(page: Page, component: 'AIArchitectureGenerator' | 'ArchitectureChatPanel' | 'AIChangeReview' | 'ModelSettingsPopover' | 'BYOAISettingsDialog') {
  const styles = [
    '../src/styles/design-tokens.css', '../src/styles/tokens.css', '../src/index.css',
    '../src/styles/control-primitives.css', '../src/styles/modal-primitives.css',
    '../src/components/AIConnectionSelector.css', '../src/components/ModelSettingsPopover.css',
    `../src/components/${component}.css`, '../src/styles/surfaces.css',
  ];
  await page.addStyleTag({ content: styles.map(path => readFileSync(new URL(path, import.meta.url), 'utf8')).join('\n') });
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
  await page.waitForFunction(() => (window as any).h.pending.some((request: { kind: string }) => request.kind === 'topology'));
  await page.evaluate(() => {
    const h = (window as any).h;
    h.nodes = [{ id: 'manual', type: 'azureNode', position: { x: 0, y: 0 }, data: { label: 'Manual edit' } }];
    // Commit the simulated edit before the reply settles in this same browser task.
    h.commitRender();
    h.pending.find((request: { kind: string }) => request.kind === 'topology').resolve();
  });
  await page.locator('.arch-chat-msg-error').filter({ hasText: 'diagram changed' }).waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.applies.length), 0);
  assert.equal(await page.evaluate(() => (window as any).h.followups), 0);
  assert.equal(await page.locator('textarea').inputValue(), 'Add SQL');
});

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

for (const language of ['en', 'ja']) {
  test(`review exposes supported before and after values without internal metadata in ${language}`, async t => {
    const node = (id: string, label: string) => ({ id, position: { x: 0, y: 0 }, data: { label } });
    const groups = [node('old-group', 'Original network'), node('new-group', 'Replacement network')];
    const targets = [node('old-db', 'Orders database'), node('new-db', 'Audit database')];
    const before = {
      ...node('web', 'Customer portal'), parentNode: 'old-group',
      data: { label: 'Customer portal', description: 'Original description', tags: ['Original tag'],
        pricing: { quantity: 1, estimatedCost: 10, region: 'eastus2' }, internalCache: 'HIDDEN_OLD' },
      style: { color: 'red' },
    };
    const after = {
      ...before, parentNode: 'new-group',
      data: { ...before.data, description: 'Replacement description', tags: ['Replacement tag'],
        pricing: { quantity: 2, estimatedCost: 15, region: 'japaneast' }, internalCache: 'HIDDEN_NEW' },
      style: { color: 'blue' },
    };
    const oldEdge = { id: 'connection', source: 'web', target: 'old-db', label: 'Write records' };
    const newEdge = { ...oldEdge, target: 'new-db' };
    const page = await setup(t, {
      surface: 'review', language,
      changeSet: {
        before: { nodes: [...groups, ...targets, before], edges: [oldEdge] },
        proposed: { nodes: [...groups, ...targets, after], edges: [newEdge] },
        changes: [
          { id: 'node:change:web', entity: 'node', entityId: 'web', kind: 'change', label: 'Customer portal',
            before, after, fields: ['data', 'parentNode', 'style'], costDelta: 20 },
          { id: 'edge:change:connection', entity: 'edge', entityId: 'connection', kind: 'change', label: 'Write records',
            before: oldEdge, after: newEdge, fields: ['target'] },
        ],
      },
    });
    const dialog = page.getByRole('dialog');
    await page.setViewportSize({ width: 390, height: 844 });
    await addComponentStyles(page, 'AIChangeReview');
    await page.evaluate(() => {
      const h = (window as any).h;
      for (const entity of [h.changeSet.changes[0].before, h.changeSet.changes[0].after]) {
        entity.data.onEdit = window.getSelection;
      }
      h.render();
    });
    const details = dialog.locator('.ai-review-details').first();
    await details.locator('summary').focus();
    await page.keyboard.press('Enter');
    assert.equal(await details.getAttribute('open'), '', 'Enter must open review details without changing selection');
    const content = await details.innerText();
    for (const value of ['Original description', 'Replacement description', 'Original tag', 'Replacement tag',
      'Original network', 'Replacement network', 'eastus2', 'japaneast', 'red', 'blue']) {
      assert.ok(content.includes(value), `review must expose ${value}`);
    }
    for (const label of language === 'ja' ? ['変更前', '変更後', '説明', '数量'] : ['Before', 'After', 'Description', 'Quantity']) {
      assert.ok(content.includes(label), `review detail labels must be localized: ${label}`);
    }
    assert.doesNotMatch(content, /HIDDEN_|internalCache/);
    assert.equal(await dialog.getByRole('checkbox').first().isChecked(), true, 'opening details must not toggle selection');
    await dialog.locator('.ai-review-details').last().locator('summary').focus();
    await page.keyboard.press('Space');
    assert.equal(await dialog.locator('.ai-review-details').last().getAttribute('open'), '', 'Space must open connection details');
    assert.equal(await dialog.getByRole('checkbox').last().isChecked(), true, 'opening connection details must preserve selection');
    const connection = await dialog.locator('.ai-review-details').last().innerText();
    assert.ok(connection.includes('Orders database') && connection.includes('Audit database'));
    const bounds = await dialog.locator('.ai-review-list').evaluate(element => ({
      width: element.clientWidth, scrollWidth: element.scrollWidth,
      detailCount: element.querySelectorAll('.ai-review-detail-row').length,
    }));
    assert.ok(bounds.scrollWidth <= bounds.width + 1, 'expanded review values must fit at phone width');
    t.diagnostic(`${language} 390px expanded review: clientWidth=${bounds.width}, scrollWidth=${bounds.scrollWidth}, ${bounds.detailCount} localized before/after rows; Enter/Space disclosures preserve selection`);
    await dialog.getByRole('button', { name: language === 'ja' ? '選択した変更を適用' : 'Apply selected changes' }).click();
    assert.equal(await page.evaluate(() => (window as any).h.reviewed[0].nodes.find((item: any) => item.id === 'web').data.description),
      'Replacement description');
  });
}

const savedLegacySettings = {
  version: 3, model: 'gpt-5.6-terra', reasoningEffort: 'high',
  featureOverrides: { blueprint: { model: 'gpt-5.6-sol', reasoningEffort: 'medium' } },
};
const savedBYO = {
  version: 1, enabled: true, provider: 'openai', model: 'customer-owned-model',
  baseUrl: 'https://api.openai.com/v1', apiFormat: 'responses', reasoningEffort: 'max',
};

for (const profile of ['astra', 'legacy', 'none']) {
    test(`managed Astra popover ${profile} migrates legacy reasoning and preserves selected BYO profiles`, async t => {
      const page = await setupModelSettings(t, profile, {
        savedSettings: savedLegacySettings, savedBYO,
      });
      await openModelGuidance(page);
      assert.match(await page.locator('#root').innerText(), /GPT-6 Astra/);
      assert.doesNotMatch(await page.locator('#root').innerText(), /GPT-5|alternative model|Compare Models|Compare Validation/i);
      assert.match(await page.locator('#root').innerText(), /customer-owned-model/);
      assert.match(await page.locator('.ai-connection-current').innerText(), /Key required/);
      assert.equal(await page.locator('.msp-model-btn, .msp-byo-card, .model-select').count(), 0);
      assert.equal(await page.locator('.astra-reasoning-settings').getByRole('combobox').count(), 4, 'managed per-feature reasoning remains intact');
      assert.equal(await page.getByRole('combobox', { name: 'AI connection', exact: true }).count(), 1);
      assert.deepEqual(await page.locator('.astra-reasoning-settings').getByRole('combobox').last().locator('option').evaluateAll(options => (
        options.map(option => (option as HTMLOptionElement).value)
      )), ['default', 'none', 'low', 'medium', 'high', 'xhigh', 'max']);
      const migrated = {
        model: 'gpt-6-astra', reasoningEffort: 'high',
        featureOverrides: { blueprint: { model: 'gpt-6-astra', reasoningEffort: 'medium' } },
      };
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettings()), migrated);
      assert.deepEqual(await page.evaluate(() => JSON.parse(localStorage.getItem('azure-diagrams-model-settings')!)), {
        version: 4, astraOnlyVersion: 1, ...migrated,
      });
      const recommended = page.getByRole('button', { name: 'Use recommended reasoning', exact: true });
      if (profile !== 'astra') {
        assert.match(await page.getByRole('alert').innerText(), /GPT-6 Astra is not configured.*administrator/);
        for (const control of await page.locator('.astra-reasoning-settings button, .astra-reasoning-settings select').all()) {
          assert.equal(await control.isDisabled(), true);
        }
        return;
      }
      assert.equal(await page.getByRole('alert').count(), 0);
      await page.getByRole('group', { name: 'Default reasoning', exact: true })
        .getByRole('button', { name: 'Max', exact: true }).click();
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettingsForFeature('validation')),
        { model: 'gpt-6-astra', reasoningEffort: 'max' });
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettingsForFeature('blueprint')),
        { model: 'gpt-6-astra', reasoningEffort: 'medium' });
      await page.getByRole('combobox', { name: 'Blueprint Diagrams - Reasoning effort', exact: true }).selectOption('default');
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettingsForFeature('blueprint')),
        { model: 'gpt-6-astra', reasoningEffort: 'max' });
      await recommended.click();
      assert.deepEqual(await page.evaluate(() => (window as any).h.getModelSettings()), {
        model: 'gpt-6-astra', reasoningEffort: 'low', featureOverrides: {},
      });
    });
}

for (const language of ['en', 'ja']) {
  test(`Astra-only reasoning controls have feature-specific names and bounded touch layout in ${language}`, async t => {
    const page = await setupModelSettings(t, 'astra', { touch: true });
    await page.setViewportSize({ width: 390, height: 844 });
    await addComponentStyles(page, 'ModelSettingsPopover');
    await page.evaluate(language => (window as any).h.setLanguage(language), language);
    const summary = page.locator('.msp-advanced-settings summary');
    await summary.focus();
    await page.keyboard.press('Enter');
    assert.equal(await page.locator('.msp-advanced-settings').getAttribute('open'), '');
    const names: string[] = [];
    const rows = page.locator('.msp-feature-row');
    assert.equal(await rows.count(), 4);
    assert.deepEqual(await rows.locator('.msp-feature-name').allTextContents(), language === 'ja'
      ? ['アーキテクチャの生成', 'アーキテクチャの検証', 'デプロイガイドとBicep', 'ブループリント図']
      : ['Architecture Generation', 'Architecture Validation', 'Deployment Guide & Bicep', 'Blueprint Diagrams']);
    for (const row of await rows.all()) {
      const feature = await row.locator('.msp-feature-name').innerText();
      const name = `${feature} - ${language === 'ja' ? '推論強度' : 'Reasoning effort'}`;
      const select = row.getByRole('combobox', { name, exact: true });
      await select.waitFor();
      assert.ok((await select.boundingBox())!.height >= 44, name);
      names.push(name);
    }
    for (const width of [320, 390]) {
      await page.setViewportSize({ width, height: 844 });
      const bounds = await page.locator('.astra-reasoning-settings').evaluate(element => ({
        width: element.clientWidth, scrollWidth: element.scrollWidth,
        controlsFit: Array.from(element.querySelectorAll('select, button')).every(control => (
          control.getBoundingClientRect().right <= element.getBoundingClientRect().right + 1
        )),
      }));
      assert.ok(bounds.scrollWidth <= bounds.width + 1 && bounds.controlsFit,
        `Astra reasoning choices must not scroll horizontally at ${width}px`);
      t.diagnostic(`${language} ${width}px Astra reasoning: clientWidth=${bounds.width}, scrollWidth=${bounds.scrollWidth}, controls contained=${bounds.controlsFit}`);
    }
    t.diagnostic(`${language} accessible reasoning names: ${names.join('; ')}`);
  });

  test(`Astra-only ${language} missing setup blocks generator and chat without custom-provider advice`, async t => {
    const page = await setup(t, { configured: false, language });
    await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
    await page.locator('#architecture-description').fill('Add a secure web application');
    await page.getByRole('button', { name: language === 'ja' ? '出力設定へ進む' : 'Continue to output', exact: true }).click();
    await page.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
    const error = await page.getByRole('alert').innerText();
    assert.match(error, /GPT-6 Astra/);
    assert.match(error, language === 'ja' ? /管理者/ : /administrator/);
    assert.doesNotMatch(error, /custom|fallback|GPT-5/i);
    assert.equal(await page.evaluate(() => (window as any).h.calls.length), 0);
    const chat = await setup(t, { surface: 'chat', configured: false, language });
    assert.match(await chat.locator('.arch-chat-warning').innerText(), /GPT-6 Astra/);
    assert.equal(await chat.getByRole('button', { name: 'Send', exact: true }).isDisabled(), true);
    assert.equal(await chat.evaluate(() => (window as any).h.calls.length), 0);
  });
}

test('Astra-only Both output keeps independent reasoning and never offers a model picker', async t => {
  const page = await setup(t, {
    mode: 'both', accepted: true,
    settings: { model: 'gpt-6-astra', reasoningEffort: 'max',
      featureOverrides: { blueprint: { model: 'gpt-6-astra', reasoningEffort: 'none' } } },
  });
  await page.getByRole('button', { name: 'Generate Diagram', exact: true }).click();
  await page.locator('#architecture-description').fill('Add a secure web application');
  await page.getByRole('button', { name: 'Continue to output', exact: true }).click();
  assert.equal(await page.getByRole('combobox', { name: /Select AI model/ }).count(), 0);
  assert.equal(await page.getByRole('combobox', { name: /Select reasoning effort/ }).count(), 2);
  await page.getByRole('button', { name: 'Generate Architecture', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.blueprints === 1);
  assert.deepEqual(await page.evaluate(() => (window as any).h.calls.map((call: any) => ({
    kind: call.kind, model: call.override.model, effort: call.override.reasoningEffort,
  }))), [
    { kind: 'manifest', model: 'gpt-6-astra', effort: 'max' },
    { kind: 'topology', model: 'gpt-6-astra', effort: 'max' },
    { kind: 'blueprint', model: 'gpt-6-astra', effort: 'none' },
  ]);
});

for (const modelName of ['GPT-6 Astra (Low)', 'GPT-5.6 Sol (Medium)', 'Legacy customer endpoint / retired-model']) {
  test(`model badges preserve ${modelName} provenance even when Astra is configured`, async t => {
    const page = await setupModelSettings(t, 'astra', { surface: 'badge', modelName });
    assert.equal(await page.locator('.model-generation-badge-text strong').innerText(), modelName);
    assert.match(await page.locator('.model-generation-badge-time').innerText(), /1\.2s/);
  });
}

async function saveConnectionProfile(page: Page, name = 'Design team', language = 'en', enterKey = true) {
  const dialog = page.getByRole('dialog', { name: language === 'ja' ? 'AI 接続' : 'AI connections', exact: true });
  await dialog.getByLabel(language === 'ja' ? 'プロファイル名' : 'Profile name', { exact: true }).fill(name);
  await dialog.getByLabel(language === 'ja' ? 'プロバイダー' : 'Provider', { exact: true }).selectOption('openai');
  await dialog.getByLabel(language === 'ja' ? 'モデル / デプロイ名' : 'Model / deployment', { exact: true }).fill('customer-model');
  await dialog.getByRole('button', { name: language === 'ja' ? 'プロファイルを保存' : 'Save profile', exact: true }).click();
  if (enterKey) {
    await dialog.getByLabel(language === 'ja' ? 'API キー（タブのメモリ内のみ）' : 'API key (tab memory only)', { exact: true })
      .fill('sk-ui-private-test-key');
  }
  return dialog;
}

test('BYO actual dialog saves named profiles without activation and supports explicit switch-before-delete CRUD', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections' });
  const dialog = await saveConnectionProfile(page);
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null);
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null,
    'a successful test must never silently select BYO');
  await dialog.getByRole('button', { name: 'Use this profile', exact: true }).click();
  assert.equal(await dialog.getByRole('button', { name: 'Delete profile', exact: true }).isDisabled(), true);
  assert.match(await dialog.innerText(), /Explicitly switch to managed Astra/);
  await dialog.getByRole('button', { name: 'Use managed Astra', exact: true }).click();
  await dialog.getByRole('button', { name: 'Add profile', exact: true }).click();
  await saveConnectionProfile(page, 'Review team');
  assert.equal(await dialog.locator('.byo-ai-profile-list > button').count(), 2);
  await dialog.getByLabel('Profile name', { exact: true }).fill('Review production');
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  assert.match(await dialog.locator('.byo-ai-profile-list').innerText(), /Review production/);
  await dialog.getByRole('button', { name: 'Delete profile', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 2);
  await dialog.getByRole('button', { name: 'Confirm delete', exact: true }).click();
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 1);
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null);
});

test('BYO legacy reload preserves the chosen profile but never persisted keys or verification', async t => {
  const page = await setupModelSettings(t, 'astra', {
    surface: 'connections', savedBYO: { ...savedBYO, apiKey: 'sk-legacy-key-must-disappear', verified: true,
      nested: { apiKey: 'sk-nested-key-must-disappear' } },
  });
  const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
  await dialog.locator('.byo-ai-status--key-required').waitFor();
  assert.equal(await dialog.getByLabel('API key (tab memory only)', { exact: true }).inputValue(), '');
  const initial = await page.evaluate(() => {
    const h = (window as any).h;
    return { settings: h.getBYOAISettings(), stored: localStorage.getItem('azure-diagrams-byo-ai-settings') };
  });
  assert.ok(initial.settings.activeProfileId);
  assert.doesNotMatch(initial.stored!, /apiKey|verified|secret|sk-legacy|sk-nested|nested/);
  assert.equal(await dialog.getByRole('button', { name: 'Profile selected', exact: true }).isDisabled(), true);
  await dialog.getByLabel('API key (tab memory only)', { exact: true }).fill('sk-ui-private-test-key');
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.getEffectiveAIModelInfo('architectureGeneration').ready), true);
  await page.reload();
  await page.evaluate(() => { (window as any).h = { surface: 'connections', open: true }; });
  await page.addScriptTag({ content: modelSettingsScripts.get('astra')! });
  await page.locator('.byo-ai-status--key-required').waitFor();
  assert.equal(await page.getByLabel('API key (tab memory only)', { exact: true }).inputValue(), '');
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), initial.settings.activeProfileId);
  assert.equal(await page.evaluate(() => (window as any).h.getEffectiveAIModelInfo('architectureGeneration').ready), false);
  assert.doesNotMatch(await page.locator('body').innerText(), /sk-ui-private-test-key|sk-legacy|sk-nested/);
  const storage = await page.evaluate(() => ({ local: { ...localStorage }, session: { ...sessionStorage } }));
  assert.doesNotMatch(JSON.stringify(storage), /sk-ui-private-test-key|sk-legacy|sk-nested|apiKey|verified/);
});

test('BYO imported credential-bearing endpoints never appear in saved profiles or editor fields', async t => {
  const page = await setupModelSettings(t, 'astra', {
    surface: 'connections', savedBYO: { version: 3, activeProfileId: 'credential-profile', profiles: [{
      id: 'credential-profile', name: 'Imported team', provider: 'azure-openai',
      endpoint: 'https://user:sk-import-secret@resource.openai.azure.com/?api-key=sk-query-secret',
      model: 'team-model', apiFormat: 'responses', reasoningEffort: 'none',
      isReasoning: false, supportsVision: false, maxCompletionTokens: 2048,
    }] },
  });
  const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
  const state = await page.evaluate(() => {
    const h = (window as any).h;
    return { settings: h.getBYOAISettings(), stored: localStorage.getItem('azure-diagrams-byo-ai-settings'),
      effective: h.getEffectiveAIModelInfo('architectureGeneration') };
  });
  assert.equal(state.settings.activeProfileId, 'credential-profile');
  assert.equal(state.effective.source, 'bring-your-own');
  assert.equal(state.effective.ready, false);
  assert.doesNotMatch(JSON.stringify(state.settings) + state.stored, /sk-import-secret|sk-query-secret|user:/);
  assert.doesNotMatch(await dialog.getByLabel('Endpoint origin', { exact: true }).inputValue(), /sk-import-secret|sk-query-secret|user:/);
  assert.doesNotMatch(await dialog.innerText(), /sk-import-secret|sk-query-secret|user:/);
});

for (const language of ['en', 'ja']) {
  test(`BYO ${language} storage failures preserve add, activation, switch, and deletion state with safe guidance`, async t => {
    const page = await setupModelSettings(t, 'astra', { surface: 'connections' });
    await page.evaluate(language => (window as any).h.setLanguage(language), language);
    await page.evaluate(() => {
      const original = Storage.prototype.setItem;
      (window as any).h.failProfileStorage = true;
      Storage.prototype.setItem = function (key: string, value: string) {
        if (key === 'azure-diagrams-byo-ai-settings' && (window as any).h.failProfileStorage) {
          throw new DOMException('Unsafe quota details sk-storage-private-test-key', 'QuotaExceededError');
        }
        original.call(this, key, value);
      };
    });
    const dialog = await saveConnectionProfile(page, 'Transactional team', language, false);
    const managerError = dialog.locator('.byo-ai-test-section .byo-ai-error[role="alert"]');
    await managerError.waitFor();
    assert.match(await managerError.innerText(), language === 'ja' ? /保存|ストレージ/ : /save|storage/i);
    assert.doesNotMatch(await managerError.innerText(), /sk-storage-private-test-key|Unsafe quota/);
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 0);
    await page.evaluate(() => { (window as any).h.failProfileStorage = false; });
    await dialog.getByRole('button', { name: language === 'ja' ? 'プロファイルを保存' : 'Save profile', exact: true }).click();
    await dialog.getByLabel(language === 'ja' ? 'API キー（タブのメモリ内のみ）' : 'API key (tab memory only)', { exact: true }).fill('sk-ui-private-test-key');
    await dialog.getByRole('button', { name: language === 'ja' ? '接続をテスト' : 'Test connection', exact: true }).click();
    await dialog.locator('.byo-ai-status--verified').waitFor();
    await page.evaluate(() => { (window as any).h.failProfileStorage = true; });
    await dialog.getByRole('button', { name: language === 'ja' ? 'このプロファイルを使用' : 'Use this profile', exact: true }).click();
    await managerError.waitFor();
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null);
    await page.evaluate(() => { (window as any).h.failProfileStorage = false; });
    await dialog.getByRole('button', { name: language === 'ja' ? 'このプロファイルを使用' : 'Use this profile', exact: true }).click();
    const selectedId = await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId);
    assert.ok(selectedId);
    await page.evaluate(() => { (window as any).h.failProfileStorage = true; });
    await dialog.getByRole('button', { name: language === 'ja' ? '管理対象の Astra を使用' : 'Use managed Astra', exact: true }).click();
    await dialog.locator('.ai-connection-error').waitFor();
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), selectedId);
    await page.evaluate(() => { (window as any).h.failProfileStorage = false; });
    await dialog.getByRole('button', { name: language === 'ja' ? '管理対象の Astra を使用' : 'Use managed Astra', exact: true }).click();
    await page.evaluate(() => { (window as any).h.failProfileStorage = true; });
    await dialog.getByRole('button', { name: language === 'ja' ? 'プロファイルを削除' : 'Delete profile', exact: true }).click();
    await dialog.getByRole('button', { name: language === 'ja' ? '削除を確定' : 'Confirm delete', exact: true }).click();
    await managerError.waitFor();
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 1);
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null);
    assert.doesNotMatch(await dialog.innerText(), /sk-storage-private-test-key|Unsafe quota/);
    await page.evaluate(() => { (window as any).h.failProfileStorage = false; });
    await dialog.getByRole('button', { name: language === 'ja' ? '削除を確定' : 'Confirm delete', exact: true }).click();
    assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 0);
  });
}

test('BYO explicit test is cancellable while busy and stale ignored-abort replies cannot verify', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', deferTest: true });
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.pendingTests.length === 1);
  assert.equal(await dialog.getByRole('button', { name: 'Close AI connections', exact: true }).isDisabled(), true);
  await page.keyboard.press('Escape');
  assert.equal(await dialog.isVisible(), true);
  await dialog.getByRole('button', { name: 'Cancel test', exact: true }).click();
  await dialog.locator('.byo-ai-status--unverified').waitFor();
  await page.evaluate(() => (window as any).h.pendingTests.shift().resolve());
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  assert.equal(await page.evaluate(() => {
    const h = (window as any).h;
    return h.getBYOAIConnectionState(h.getBYOAISettings().profiles[0].id).verified;
  }), false);

  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.pendingTests.length === 1);
  await dialog.getByLabel('Model / deployment', { exact: true }).fill('updated-deployment');
  await page.evaluate(() => (window as any).h.pendingTests.shift().resolve());
  await page.waitForFunction(() => {
    const h = (window as any).h;
    return h.getBYOAIConnectionState(h.getBYOAISettings().profiles[0].id).status !== 'testing';
  });
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  await page.evaluate(() => { (window as any).h.deferTest = false; });
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.requests.at(-1).model), 'updated-deployment');
});

test('BYO profile capability and key edits revoke verification without silently selecting managed Astra', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections' });
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  await dialog.getByRole('button', { name: 'Use this profile', exact: true }).click();
  const selectedId = await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId);
  await dialog.getByLabel('Supports image / vision input', { exact: true }).check();
  assert.equal(await page.evaluate(id => (window as any).h.getBYOAIConnectionState(id).verified, selectedId), false);
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), selectedId);
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  await dialog.getByLabel('API key (tab memory only)', { exact: true }).fill('short');
  await dialog.locator('.byo-ai-status--key-required').waitFor();
  assert.equal(await dialog.getByRole('button', { name: 'Test connection', exact: true }).isDisabled(), true);
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), selectedId);
  assert.equal(await page.evaluate(() => (window as any).h.getEffectiveAIModelInfo('blueprint').ready), false);
});

test('BYO friendly-name edits preserve an in-progress test and verified captured connection', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', deferTest: true });
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.pendingTests.length === 1);
  await dialog.getByLabel('Profile name', { exact: true }).fill('Design review team');
  assert.equal(await dialog.getByRole('button', { name: 'Cancel test', exact: true }).isVisible(), true);
  await page.evaluate(() => (window as any).h.pendingTests.shift().resolve());
  await dialog.locator('.byo-ai-status--verified').waitFor();
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  await dialog.getByRole('button', { name: 'Use this profile', exact: true }).click();
  const revision = await page.evaluate(() => {
    const h = (window as any).h;
    h.beforeRenameCapture = h.captureRuntimeModelOverride('architectureGeneration');
    return h.getBYOAIConnectionState(h.getBYOAISettings().activeProfileId).revision;
  });
  await dialog.getByLabel('Profile name', { exact: true }).fill('Design final team');
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  assert.deepEqual(await page.evaluate(() => {
    const h = (window as any).h;
    const state = h.getBYOAIConnectionState(h.getBYOAISettings().activeProfileId);
    const captured = h.captureRuntimeModelOverride('architectureGeneration', h.beforeRenameCapture);
    return { revision: state.revision, verified: state.verified, requestCount: h.requests.length,
      capturedName: captured.connection.displayName, currentName: h.getBYOAISettings().profiles[0].name };
  }), { revision, verified: true, requestCount: 1,
    capturedName: 'BYO OpenAI · Design review team · customer-model', currentName: 'Design final team' });
});

test('BYO external profile revision changes during a test cannot mark the edited connection verified', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', deferTest: true });
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await page.waitForFunction(() => (window as any).h.pendingTests.length === 1);
  await page.evaluate(() => {
    const h = (window as any).h;
    h.upsertBYOAIProfile({ ...h.getBYOAISettings().profiles[0], maxCompletionTokens: 512 });
    h.pendingTests.shift().resolve();
  });
  await page.waitForFunction(() => {
    const h = (window as any).h;
    return h.getBYOAIConnectionState(h.getBYOAISettings().profiles[0].id).status !== 'testing';
  });
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  assert.equal(await page.evaluate(() => {
    const h = (window as any).h;
    return h.getBYOAIConnectionState(h.getBYOAISettings().profiles[0].id).verified;
  }), false);
});

test('BYO disabled server policy permits editing but blocks testing and activation until explicitly allowed', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', allowBYO: false });
  const dialog = await saveConnectionProfile(page);
  await dialog.locator('.byo-ai-status--admin-disabled').waitFor();
  assert.equal(await dialog.getByRole('button', { name: 'Save profile', exact: true }).isEnabled(), true);
  assert.equal(await dialog.getByRole('button', { name: 'Test connection', exact: true }).isDisabled(), true);
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  assert.equal(await page.evaluate(() => (window as any).h.requests.length), 0);
  await page.evaluate(() => { (window as any).h.allowBYO = true; });
  await dialog.getByRole('button', { name: 'Refresh server policy', exact: true }).click();
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId), null);
  await dialog.getByRole('button', { name: 'Use this profile', exact: true }).click();
  await page.evaluate(async () => { const h = (window as any).h; h.allowBYO = false; await h.loadRuntimeConfig(true); });
  await dialog.locator('.byo-ai-status--admin-disabled').waitFor();
  assert.ok(await page.evaluate(() => (window as any).h.getBYOAISettings().activeProfileId));
  assert.equal(await page.evaluate(() => (window as any).h.getEffectiveAIModelInfo('validation').ready), false);
});

test('BYO test failure exposes safe guidance and request ID but never key or upstream message text', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', testOutcome: 'failed' });
  const logs: string[] = [];
  page.on('console', message => logs.push(message.text()));
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--failed').waitFor();
  assert.match(await dialog.getByRole('alert').innerText(), /API key.*Re-enter.*Request ID: byo-ui-test-123/);
  assert.doesNotMatch(await page.locator('body').innerText(), /sk-ui-private-test-key|secret\.invalid/);
  assert.doesNotMatch(page.url(), /sk-ui-private-test-key/);
  assert.doesNotMatch(logs.join('\n'), /sk-ui-private-test-key|secret\.invalid/);
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
});

test('BYO incomplete test responses never enable activation and Chat Completions remains an explicit choice', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections', testOutcome: 'incomplete' });
  const dialog = await saveConnectionProfile(page);
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--failed').waitFor();
  assert.equal(await dialog.getByRole('button', { name: 'Use this profile', exact: true }).isDisabled(), true);
  await dialog.getByLabel('API format', { exact: true }).selectOption('chat-completions');
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  await page.evaluate(() => { (window as any).h.testOutcome = 'success'; });
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.ok(await page.evaluate(() => (window as any).h.requests.at(-1).body.messages));
});

test('BYO Azure profile rejects arbitrary hosts and submits the explicitly chosen capabilities and bounded output', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections' });
  const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
  await dialog.getByLabel('Profile name', { exact: true }).fill('Azure engineering');
  await dialog.getByLabel('Endpoint origin', { exact: true }).fill('https://untrusted.example.com');
  await dialog.getByLabel('Model / deployment', { exact: true }).fill('engineering-deployment');
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  assert.match(await dialog.getByRole('alert').innerText(), /Azure resource HTTPS origin/);
  assert.equal(await page.evaluate(() => (window as any).h.getBYOAISettings().profiles.length), 0);
  await dialog.getByLabel('Endpoint origin', { exact: true }).fill('https://engineering.openai.azure.com');
  await dialog.getByLabel('Supports reasoning', { exact: true }).check();
  await dialog.getByLabel('Profile reasoning effort', { exact: true }).selectOption('high');
  await dialog.getByLabel('Supports image / vision input', { exact: true }).check();
  await dialog.getByLabel('Maximum output tokens', { exact: true }).fill('1024');
  await dialog.getByRole('button', { name: 'Save profile', exact: true }).click();
  await dialog.getByLabel('API key (tab memory only)', { exact: true }).fill('sk-ui-private-test-key');
  await dialog.getByRole('button', { name: 'Test connection', exact: true }).click();
  await dialog.locator('.byo-ai-status--verified').waitFor();
  assert.deepEqual(await page.evaluate(() => {
    const request = (window as any).h.requests.at(-1);
    return { provider: request.provider, endpoint: request.endpoint, reasoning: request.body.reasoning.effort,
      output: request.body.max_output_tokens };
  }), { provider: 'azure-openai', endpoint: 'https://engineering.openai.azure.com', reasoning: 'high', output: 1024 });
});

test('BYO ten-profile limit is visible and deleting an inactive profile restores Add profile', async t => {
  const page = await setupModelSettings(t, 'astra', { surface: 'connections' });
  const dialog = page.getByRole('dialog', { name: 'AI connections', exact: true });
  await page.evaluate(() => {
    const h = (window as any).h;
    for (let i = 0; i < 10; i++) h.upsertBYOAIProfile({
      id: 'profile-' + i, name: 'Team ' + i, provider: 'openai', endpoint: 'https://api.openai.com',
      model: 'team-model', apiFormat: 'responses', isReasoning: false, supportsVision: false,
      reasoningEffort: 'none', maxCompletionTokens: 2048,
    });
  });
  await page.waitForFunction(() => document.querySelectorAll('.byo-ai-profile-list > button').length === 10);
  assert.equal(await dialog.getByRole('button', { name: 'Add profile', exact: true }).isDisabled(), true);
  await dialog.locator('.byo-ai-profile-list > button').first().click();
  await dialog.getByRole('button', { name: 'Delete profile', exact: true }).click();
  await dialog.getByRole('button', { name: 'Confirm delete', exact: true }).click();
  assert.equal(await dialog.getByRole('button', { name: 'Add profile', exact: true }).isEnabled(), true);
});

for (const language of ['en', 'ja']) {
  test(`BYO actual ${language} dialog keeps 44px touch targets, 320/390px fit, key visibility, and focus restoration`, async t => {
    const page = await setupModelSettings(t, 'astra', { surface: 'connections', open: false, touch: true });
    await addComponentStyles(page, 'BYOAISettingsDialog');
    await page.setViewportSize({ width: 390, height: 844 });
    await page.evaluate(language => (window as any).h.setLanguage(language), language);
    const opener = page.getByRole('button', { name: 'Open AI connections', exact: true });
    await opener.click();
    const dialog = await saveConnectionProfile(page, language === 'ja' ? '設計チーム' : 'Design team', language);
    const key = dialog.getByLabel(language === 'ja' ? 'API キー（タブのメモリ内のみ）' : 'API key (tab memory only)', { exact: true });
    assert.equal(await key.getAttribute('type'), 'password');
    await dialog.getByRole('button', { name: language === 'ja' ? 'API キーを表示' : 'Show API key', exact: true }).click();
    assert.equal(await key.getAttribute('type'), 'text');
    assert.doesNotMatch(await dialog.innerText(), /sk-ui-private-test-key/);
    await dialog.getByRole('button', { name: language === 'ja' ? 'API キーを非表示' : 'Hide API key', exact: true }).click();
    for (const width of [320, 390]) {
      await page.setViewportSize({ width, height: 844 });
      const bounds = await dialog.evaluate(element => ({
        width: element.clientWidth, scrollWidth: element.scrollWidth,
        bodyWidth: element.querySelector('.byo-ai-dialog-body')!.clientWidth,
        bodyScroll: element.querySelector('.byo-ai-dialog-body')!.scrollWidth,
        controlsFit: Array.from(element.querySelectorAll('button,select,input:not([type="checkbox"])'))
          .filter(control => (control as HTMLElement).offsetParent !== null)
          .every(control => control.getBoundingClientRect().height >= 44
            && control.getBoundingClientRect().right <= element.getBoundingClientRect().right + 1),
      }));
      assert.ok(bounds.scrollWidth <= bounds.width + 1 && bounds.bodyScroll <= bounds.bodyWidth + 1 && bounds.controlsFit,
        JSON.stringify({ language, viewportWidth: width, ...bounds }));
    }
    const done = dialog.getByRole('button', { name: language === 'ja' ? '完了' : 'Done', exact: true });
    const close = dialog.getByRole('button', { name: language === 'ja' ? 'AI 接続を閉じる' : 'Close AI connections', exact: true });
    await done.focus();
    await page.keyboard.press('Tab');
    assert.equal(await close.evaluate(element => element === document.activeElement), true);
    await page.keyboard.press('Shift+Tab');
    assert.equal(await done.evaluate(element => element === document.activeElement), true);
    await page.keyboard.press('Escape');
    await dialog.waitFor({ state: 'hidden' });
    assert.equal(await opener.evaluate(element => element === document.activeElement), true);
  });
}

for (const surface of ['generator', 'chat']) {
  test(`${surface} retains submitted BYO model provenance when connection settings change during a request`, async t => {
    const page = await setup(t, {
      surface, accepted: true, deferKinds: ['topology'],
      connectionInfo: { source: 'bring-your-own', profileId: 'team', model: 'first-deployment',
        displayName: 'BYO OpenAI · Design team · first-deployment', ready: true,
        reasoningEffort: 'low', apiFormat: 'responses', isReasoning: true, supportsVision: true, maxCompletionTokens: 8000 },
    });
    if (surface === 'generator') await generate(page);
    else {
      await page.locator('.arch-chat-input').fill('Add a secure web application');
      await page.getByRole('button', { name: 'Send', exact: true }).click();
    }

    await page.waitForFunction(() => (window as any).h.calls.length === 1);
    await page.evaluate(() => {
      const h = (window as any).h;
      h.connectionInfo = { ...h.connectionInfo, model: 'new-deployment', displayName: 'BYO OpenAI · Edited team · new-deployment' };
      h.commitRender();
    });
    assert.equal(await page.evaluate(() => (window as any).h.calls[0].override.connection.model), 'first-deployment');
    await finish(page, 'topology');
    if (surface === 'generator') {
      await page.locator('.generator-success-panel').waitFor();
      assert.match(await page.locator('.ai-submitted-model').innerText(), /first-deployment/);
      assert.doesNotMatch(await page.locator('.ai-submitted-model').innerText(), /new-deployment/);
    } else {
      const completedProvenance = page.locator('.arch-chat-bubble:not(.arch-chat-bubble-pending) .arch-chat-provenance');
      await completedProvenance.waitFor();
      await page.waitForFunction(() => (window as any).h.followupInputs.length === 1);
      assert.match(await completedProvenance.innerText(), /first-deployment/);
      assert.equal(await page.evaluate(() => (window as any).h.followupInputs[0].modelOverride.connection.model), 'first-deployment');
    }
  });
}

for (const [name, provenance, suffix] of [
  ['BYO', { source: 'bring-your-own', model: 'BYO OpenAI · Original team · original-model',
    deployment: 'original-custom-deployment', reasoningEffort: 'high' }, '-byo-original-custom-deployment-high'],
  ['managed', { source: 'managed', model: 'GPT-6 Astra', reasoningEffort: 'medium' }, '-gpt6astra-medium'],
  ['unattributed', undefined, ''],
] as const) {
  test(`validation ${name} download uses stored artifact provenance and matching image links, never current settings`, async t => {
    const page = await setupModelSettings(t, 'astra', {
      surface: 'validation-download', savedBYO,
      validation: {
        timestamp: new Date(123).toISOString(), overallScore: 75, summary: 'Stored assessment',
        pillars: [], quickWins: [],
        diagramImageDataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/mvIAAAAASUVORK5CYII=',
        ...(provenance ? { metrics: { ...provenance, elapsedTimeMs: 1,
          promptTokens: 1, completionTokens: 1, totalTokens: 2 } } : {}),
      },
    });
    const dialog = page.getByRole('dialog', { name: '🔍 Architecture Validation', exact: true });
    await dialog.getByRole('button', { name: 'Download Report', exact: true }).click();
    await page.waitForFunction(() => (window as any).h.downloads.length === 2 && (window as any).h.downloads[0].text !== null);
    const downloads = await page.evaluate(() => (window as any).h.downloads);
    assert.deepEqual(downloads.map((item: any) => item.filename), [
      `architecture-validation-123${suffix}.md`, `architecture-validation-diagram-123${suffix}.png`,
    ]);
    assert.match(downloads[0].text, new RegExp(`\\]\\(\\./architecture-validation-diagram-123${suffix}\\.png\\)`));
    assert.doesNotMatch(downloads.map((item: any) => item.filename).join(' '), /customer-owned-model|Migrated AI connection/);
    assert.equal(await page.evaluate(() => (window as any).h.requests.length), 0, 'downloading a saved report never calls AI');
  });
}
