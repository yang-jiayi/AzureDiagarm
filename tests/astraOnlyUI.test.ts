import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';

test('managed Astra retirement keeps comparison and legacy model selector modules removed', () => {
  for (const component of ['CompareModelsModal', 'CompareValidationModal', 'ModelSelector']) {
    for (const extension of ['tsx', 'css']) {
      assert.equal(existsSync(new URL(`../src/components/${component}.${extension}`, import.meta.url)), false,
        `${component}.${extension} must not be retained solely for retired UI tests`);
    }
  }
});

test('App lazily exposes BYO profiles without reviving retired comparison preview plumbing', () => {
  const app = readFileSync(new URL('../src/App.tsx', import.meta.url), 'utf8');
  assert.doesNotMatch(app, /CompareModelsModal|CompareValidationModal|ModelSelector/);
  assert.doesNotMatch(app, /isCompareModelsOpen|isCompareValidationOpen|batchPreview|isCapturingBatch/);
  assert.doesNotMatch(app, /projectValidationComparisonInput/);
  assert.match(app, /const BYOAISettingsDialog = lazyWhenOpen\(\(\) => import/);
  assert.match(app, /configure-byo-ai/);
  assert.match(app, /onOpenAIConnections=\{openAIConnections\}/);
  assert.match(app, /<ModelSettingsPopover/);
  assert.match(app, /<AIChangeReview/);
  assert.match(app, /<ArchitectureChatPanel/);
});

test('BYO UI delegates verification and key ownership to core and keeps explicit activation', () => {
  const dialog = readFileSync(new URL('../src/components/BYOAISettingsDialog.tsx', import.meta.url), 'utf8');
  assert.match(dialog, /testBYOAIConnection\(saved\.id, \{ signal: controller\.signal \}\)/);
  assert.match(dialog, /selectBYOAIProfile\(saved\.id\)/);
  assert.match(dialog, /state !== 'verified'/);
  assert.match(dialog, /invalidateBYOAIProfile\(saved\.id\)/);
  assert.match(dialog, /if \(saved && key !== 'name'\)/);
  assert.match(dialog, /disabled=\{active \|\| testing\}/);
  assert.match(dialog, /type=\{showKey \? 'text' : 'password'\}/);
  assert.doesNotMatch(dialog, /localStorage|sessionStorage|verifiedSignature|setVerified|apiKey:\s*keyValue|console\./);
  assert.doesNotMatch(dialog, /byoAIConnectionSession|readBYOAIConnectionSecret|finishBYOAIConnectionTest/);
  assert.match(dialog, /closeOnEscape=\{!testing\}/);
  assert.match(dialog, /Cancel test/);
});

test('AI entrypoints capture safe connection metadata before asynchronous work', () => {
  const app = readFileSync(new URL('../src/App.tsx', import.meta.url), 'utf8');
  const generator = readFileSync(new URL('../src/components/AIArchitectureGenerator.tsx', import.meta.url), 'utf8');
  const chat = readFileSync(new URL('../src/components/ArchitectureChatPanel.tsx', import.meta.url), 'utf8');
  assert.match(app, /captureAIConnection\('validation'\)/);
  assert.match(app, /captureAIConnection\('deploymentGuide'\)/);
  assert.match(app, /generateArchitectureFromIaC\([\s\S]*?\}, language, \{ modelOverride \}\)/);
  assert.match(generator, /captureRuntimeModelOverride\('blueprint'\)/);
  assert.match(generator, /analyzeArchitectureDiagramImage\([\s\S]*?signal: controller\.signal, modelOverride/);
  assert.match(chat, /generateArchitectureWithAI\(prompt, modelOverride,/);
  assert.match(chat, /send\(best\[0\], before, controller, modelOverride\)/);
  assert.match(chat, /submittedModel: submittedName/);
  assert.doesNotMatch(generator, /JSON\.stringify\(modelSettings|console\.log.*effective model/);
});

test('UI download filenames use artifact metrics rather than the active connection', () => {
  const app = readFileSync(new URL('../src/App.tsx', import.meta.url), 'utf8');
  const validation = readFileSync(new URL('../src/components/ValidationModal.tsx', import.meta.url), 'utf8');
  const diagramFilenames = [...app.matchAll(/generateModelFilename\('azure-diagram(?:-workflow|-animated)?', '[a-z]+', undefined, generatedWithModel\?\.metrics\)/g)];
  assert.equal(diagramFilenames.length, 7, 'all diagram download paths must use captured diagram metadata');
  assert.match(app, /model: architecture\.metrics\.model/);
  assert.match(app, /source: architecture\.metrics\.source/);
  assert.match(app, /deployment: architecture\.metrics\.deployment/);
  assert.match(validation, /generateModelFilename\('architecture-validation', 'md', ts, validation\.metrics\)/);
  assert.match(validation, /generateModelFilename\('architecture-validation-diagram', 'png', ts, validation\.metrics\)/);
});

test('retained AI UI harness has no comparison projection scaffolding and keeps deterministic lifecycle updates', () => {
  const harness = readFileSync(new URL('./aiGenerationUI.browser.ts', import.meta.url), 'utf8');
  assert.doesNotMatch(harness, /projectValidationComparisonInput|projectValidationInput|lastValidationProjection|initialHiddenProjection/);
  assert.doesNotMatch(harness, /CompareValidationModal|CompareModelsModal|diagramFingerprint|validationInputs|validationPending/);
  assert.match(harness, /h\.commitRender = \(\) => flushSync\(h\.render\)/);
  assert.match(harness, /generator rejects an edited baseline but ignores selection and measurement changes/);
  assert.match(harness, /Both retry preserves the completed output and regenerates only/);
  assert.match(harness, /review retains selection after parent snapshot failure/);
});
