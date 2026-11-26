import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import igniteManifest from '../Azure_Public_Service_Icons/ignite-manifest.json';
import { isCapacityConsumed, resolveServiceIconMapping, SERVICE_ICON_MAP } from '../src/data/serviceIconMapping';

const root = path.resolve(import.meta.dirname, '..');
const iconRoot = path.join(root, 'Azure_Public_Service_Icons', 'Icons');
const sourceLock = JSON.parse(readFileSync(path.join(root, 'scripts', 'ignite-icon-sources.json'), 'utf8'));
const sha256 = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

test('Ignite IQ fallback symbols are separate, locally authored, and hash-pinned', () => {
  assert.equal(igniteManifest.files.length, 4);
  assert.match(igniteManifest.assetPolicy, /locally authored/i);
  assert.match(igniteManifest.assetPolicy, /not official Microsoft artwork/i);
  assert.deepEqual(
    igniteManifest.files.map(file => file.serviceName),
    ['Microsoft IQ', 'Work IQ', 'Fabric IQ', 'Web IQ'],
  );
  assert.equal(new Set(igniteManifest.files.map(file => file.path)).size, 4);

  for (const file of igniteManifest.files) {
    assert.equal(file.kind, 'unofficial-concept-substitute');
    assert.equal(file.path, sourceLock.icons.find(entry => entry.serviceName === file.serviceName)?.path);
    assert.equal(
      sha256(readFileSync(path.join(iconRoot, file.path))),
      file.sha256,
      file.path,
    );
    const mapping = resolveServiceIconMapping(file.serviceName);
    assert.ok(mapping, file.serviceName);
    assert.equal(mapping.mapping.iconFile, path.basename(file.path, '.svg'));
    assert.equal(mapping.mapping.hasPricingData, false, file.serviceName);
    assert.equal(isCapacityConsumed(file.serviceName), false, file.serviceName);
    assert.equal(SERVICE_ICON_MAP[file.serviceName]?.costRange, undefined, file.serviceName);
  }

  assert.equal(resolveServiceIconMapping('Microsoft IQ')?.mapping.iconFile, 'microsoft-iq-context-symbol');
  assert.equal(resolveServiceIconMapping('Fabric IQ')?.mapping.iconFile, 'fabric-iq-context-symbol');
  assert.equal(resolveServiceIconMapping('Foundry IQ')?.mapping.iconFile, '038470497-icon-service-Azure-AI-Foundry-IQ');
  assert.equal(resolveServiceIconMapping('Foundry IQ')?.mapping.category, 'ai + machine learning');
  assert.equal(resolveServiceIconMapping('Foundry IQ')?.mapping.hasPricingData, false);
  assert.equal(isCapacityConsumed('Foundry IQ'), false);
  assert.equal(SERVICE_ICON_MAP['Foundry IQ']?.costRange, undefined);
  assert.equal(resolveServiceIconMapping('Azure AI Search')?.mapping.iconFile, '10044-icon-service-Cognitive-Search');
  assert.equal(resolveServiceIconMapping('Work IQ')?.mapping.iconFile, 'work-iq-context-symbol');
  assert.equal(resolveServiceIconMapping('Web IQ')?.mapping.iconFile, 'web-iq-context-symbol');
  assert.notEqual(
    resolveServiceIconMapping('Microsoft Foundry')?.mapping.iconFile,
    resolveServiceIconMapping('Foundry IQ')?.mapping.iconFile,
  );
});
