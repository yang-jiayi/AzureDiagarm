// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import {
  MICROSOFT_PRODUCT_ICON_CATALOG,
  MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION,
  getMicrosoftProductIconByFileName,
} from '../src/data/microsoftProductIconCatalog';
import { SERVICE_ICON_MAP, resolveServiceIconMapping } from '../src/data/serviceIconMapping';
import { classifyIconPaletteCategory, iconPaletteCategories } from '../src/data/iconCatalog';

const here = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(here, '..');
const iconRoot = path.join(repoRoot, 'Azure_Public_Service_Icons', 'Icons');

function readManifest() {
  return JSON.parse(
    readFileSync(
      path.join(repoRoot, 'Azure_Public_Service_Icons', 'microsoft-product-manifest.json'),
      'utf8',
    ),
  ) as {
    packageVersion: string;
    iconCount: number;
    files: Array<{ path: string; serviceName: string; family: string; sha256: string }>;
  };
}

test('every Power Platform, Dynamics 365, and Microsoft 365 icon exists on disk', () => {
  assert.equal(MICROSOFT_PRODUCT_ICON_CATALOG.length, 229);
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const iconPath = path.join(iconRoot, definition.category, `${definition.fileName}.svg`);
    assert.ok(
      existsSync(iconPath),
      `${definition.serviceName} is missing its icon at ${definition.category}/${definition.fileName}.svg`,
    );
    const svg = readFileSync(iconPath, 'utf8');
    assert.ok(svg.includes('<svg'), `${definition.fileName}.svg is not an SVG document`);
  }
});

test('all three official families are represented', () => {
  const byFamily = new Map<string, number>();
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    byFamily.set(definition.family, (byFamily.get(definition.family) ?? 0) + 1);
  }
  assert.equal(byFamily.get('power-platform'), 8);
  assert.equal(byFamily.get('dynamics-365'), 16);
  assert.equal(byFamily.get('microsoft-365'), 205);

  // The products the palette must never be missing again.
  const required = [
    'Microsoft Copilot Studio',
    'Microsoft Power Apps',
    'Microsoft Power Automate',
    'Microsoft Power Pages',
    'Microsoft Dataverse',
    'Microsoft Power Platform',
    'Dynamics 365',
    'Dynamics 365 Sales',
    'Dynamics 365 Customer Service',
    'Dynamics 365 Business Central',
  ];
  const serviceNames = new Set(MICROSOFT_PRODUCT_ICON_CATALOG.map(entry => entry.serviceName));
  for (const name of required) {
    assert.ok(serviceNames.has(name), `${name} is missing from the catalog`);
  }
});

test('Microsoft 365 symbols are searchable by the workload they stand for', () => {
  const symbols = MICROSOFT_PRODUCT_ICON_CATALOG.filter(
    entry => entry.family === 'microsoft-365' && entry.kind === 'symbol',
  );

  // The official package ships concept symbols, not product logos, so the only
  // way a user finds "the Teams shape" is through the aliases.
  const searchable = new Map<string, string>();
  for (const symbol of symbols) {
    assert.equal(symbol.category, 'microsoft 365');
    assert.equal(symbol.includeInServiceMap, false);
    assert.match(symbol.fileName, /^m365-(?!app-)[a-z0-9-]+$/);
    for (const alias of symbol.aliases) {
      if (!searchable.has(alias.toLocaleLowerCase())) {
        searchable.set(alias.toLocaleLowerCase(), symbol.fileName);
      }
    }
  }

  for (const workload of [
    'Microsoft Teams',
    'SharePoint',
    'OneDrive',
    'Exchange Online',
    'Outlook',
    'Microsoft Purview',
    'Microsoft Defender XDR',
    'Microsoft Intune',
    'Microsoft Entra ID',
    'Microsoft Copilot',
    'Microsoft Planner',
    'Microsoft OneNote',
    'Microsoft Stream',
    'Viva Learning',
    'Viva Connections',
    'Viva Engage',
    'Windows 365',
    'Microsoft 365 Admin Center',
    'Microsoft Search',
    'Microsoft Graph API',
  ]) {
    assert.ok(
      searchable.has(workload.toLocaleLowerCase()),
      `no Microsoft 365 symbol is reachable by searching "${workload}"`,
    );
  }
});

test('catalog service names and file names are unique and lookup-safe', () => {
  const serviceNames = new Set<string>();
  const fileNames = new Set<string>();
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const service = definition.serviceName.toLocaleLowerCase();
    const file = definition.fileName.toLocaleLowerCase();
    assert.ok(!serviceNames.has(service), `duplicate service name: ${definition.serviceName}`);
    assert.ok(!fileNames.has(file), `duplicate file name: ${definition.fileName}`);
    serviceNames.add(service);
    fileNames.add(file);

    assert.match(definition.fileName, /^[a-z0-9-]+$/);
    assert.equal(getMicrosoftProductIconByFileName(definition.fileName), definition);
  }
});

test('every mapped catalog service and alias resolves to its own icon', () => {
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    if (!definition.includeInServiceMap) continue;
    for (const name of [definition.serviceName, definition.displayName, ...definition.aliases]) {
      const resolved = resolveServiceIconMapping(name);
      assert.ok(resolved, `"${name}" does not resolve to any icon`);
      assert.equal(
        resolved.mapping.iconFile,
        definition.fileName,
        `"${name}" resolved to ${resolved.mapping.iconFile} instead of ${definition.fileName}`,
      );
      assert.equal(resolved.mapping.category, definition.category);
    }
  }
});

test('Microsoft 365 symbols never take over a real service name', () => {
  // The symbols carry workload aliases so palette search finds them, but a
  // symbol is a shape, not a service. If one of them entered the service map it
  // would replace the real product icon everywhere the AI names that product.
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    if (definition.kind !== 'symbol') continue;
    assert.equal(
      SERVICE_ICON_MAP[definition.serviceName],
      undefined,
      `${definition.serviceName} must not be in SERVICE_ICON_MAP`,
    );
    for (const alias of definition.aliases) {
      const resolved = resolveServiceIconMapping(alias);
      if (!resolved) continue;
      assert.notEqual(
        resolved.mapping.iconFile,
        definition.fileName,
        `"${alias}" resolves to the Microsoft 365 symbol ${definition.fileName}`,
      );
    }
  }

  // The anchors those aliases must keep pointing at.
  assert.equal(resolveServiceIconMapping('Microsoft Entra ID')?.mapping.category, 'identity');
  assert.equal(
    resolveServiceIconMapping('Microsoft Copilot Studio')?.mapping.iconFile,
    'copilot-studio',
  );
});

test('the Microsoft 365 workloads a generated diagram can name all resolve', () => {
  // Before these existed, a diagram that said "Exchange Online" or "Microsoft
  // Teams" rendered an empty box: the icon library had no Microsoft 365 set at
  // all, so nothing matched.
  for (const workload of [
    'Microsoft 365',
    'Microsoft 365 Tenant',
    'Microsoft 365 Admin Center',
    'Microsoft 365 Copilot',
    'Microsoft Teams',
    'Microsoft Teams Channel',
    'Microsoft Teams Chat',
    'Microsoft Teams Meeting',
    'Microsoft Teams Phone',
    'Microsoft Teams Rooms',
    'Microsoft Teams Shifts',
    'SharePoint Online',
    'OneDrive for Business',
    'Exchange Online',
    'Microsoft Outlook',
    'Microsoft Purview Information Protection',
    'Microsoft Purview Data Loss Prevention',
    'Microsoft Purview Compliance Manager',
    'Microsoft Defender for Office 365',
    'Microsoft Defender XDR',
    'Microsoft Intune',
    'Windows 365',
    'Microsoft Viva Connections',
    'Microsoft Viva Engage',
    'Microsoft Viva Learning',
    'Microsoft Viva Insights',
    'Microsoft Planner',
    'Microsoft To Do',
    'Microsoft Project',
    'Microsoft OneNote',
    'Microsoft Word',
    'Microsoft Excel',
    'Microsoft PowerPoint',
    'Microsoft Loop',
    'Microsoft Lists',
    'Microsoft Forms',
    'Microsoft Bookings',
    'Microsoft Stream',
    'Microsoft Clipchamp',
    'Microsoft Whiteboard',
    'Microsoft Sway',
    'Microsoft Search',
    'Microsoft Graph',
    'Microsoft Graph PowerShell',
    // Short forms and Japanese names an author is at least as likely to type.
    'Teams',
    'SharePoint',
    'OneDrive',
    'Exchange',
    'Outlook',
    'Intune',
    'Yammer',
    'Planner',
    'Loop',
    'Office 365',
    'M365',
    'チームズ',
    'シェアポイント',
  ]) {
    const resolved = resolveServiceIconMapping(workload);
    assert.ok(resolved, `"${workload}" does not resolve to any icon`);
    assert.equal(
      resolved.mapping.category,
      'microsoft 365',
      `"${workload}" resolved to ${resolved.mapping.category}/${resolved.mapping.iconFile}`,
    );
    assert.match(resolved.mapping.iconFile, /^m365-app-/);
  }
});

test('the new services join the shared service map without shadowing Azure or Fabric entries', () => {
  const seen = new Map<string, string>();
  for (const [serviceName, mapping] of Object.entries(SERVICE_ICON_MAP)) {
    const key = serviceName.toLocaleLowerCase();
    assert.ok(!seen.has(key), `duplicate service key in SERVICE_ICON_MAP: ${serviceName}`);
    seen.set(key, mapping.iconFile);
  }

  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    if (!definition.includeInServiceMap) continue;
    const mapping = SERVICE_ICON_MAP[definition.serviceName];
    assert.ok(mapping, `${definition.serviceName} is not in SERVICE_ICON_MAP`);
    assert.equal(mapping.iconFile, definition.fileName);
    // These products are licensed per user, never through an Azure meter.
    assert.equal(mapping.hasPricingData, false);
  }

  // Existing Azure and Fabric anchors must keep their icons.
  assert.equal(resolveServiceIconMapping('Microsoft Fabric')?.mapping.category, 'fabric');
  assert.equal(
    resolveServiceIconMapping('Microsoft Entra ID')?.mapping.category,
    'identity',
  );
});

test('the palette exposes the four new Microsoft product categories', () => {
  const ids = new Set(iconPaletteCategories.map(category => category.id));
  assert.ok(ids.has('power-platform'));
  assert.ok(ids.has('dynamics-365'));
  assert.ok(ids.has('microsoft-365'));
  assert.ok(ids.has('microsoft-copilot'));

  for (const id of ['power-platform', 'dynamics-365', 'microsoft-365', 'microsoft-copilot'] as const) {
    const category = iconPaletteCategories.find(entry => entry.id === id)!;
    assert.ok(category.label.en.length > 0 && category.label.ja.length > 0);
    assert.ok(category.description.en.length > 0 && category.description.ja.length > 0);
    assert.ok(category.keywords.length > 0);
  }

  // The user asked for Power Platform to keep its own name rather than share it.
  assert.equal(
    iconPaletteCategories.find(entry => entry.id === 'power-platform')!.label.en,
    'Power Platform',
  );
  assert.equal(
    iconPaletteCategories.find(entry => entry.id === 'microsoft-365')!.label.en,
    'Microsoft 365',
  );

  assert.equal(
    classifyIconPaletteCategory('power platform', 'Power Apps', 'power-apps'),
    'power-platform',
  );
  assert.equal(
    classifyIconPaletteCategory('dynamics 365', 'Dynamics 365 Sales', 'dynamics-365-sales'),
    'dynamics-365',
  );
  assert.equal(
    classifyIconPaletteCategory('microsoft 365', 'People Team', 'm365-people-team'),
    'microsoft-365',
  );
});

test('every Copilot and agent asset lands in the Copilot category', () => {
  const copilotIcons: Array<[string, string, string]> = [
    ['power platform', 'Copilot Studio', 'copilot-studio'],
    ['power platform', 'Agent 365', 'agent-365'],
    ['fabric', 'Copilot in Fabric', 'fabric-workload-copilot'],
    ['fabric', 'Fabric Data Agent', 'fabric-data-agent'],
    ['fabric', 'Operations Agent', 'fabric-item-operations-agent'],
    ['microsoft 365', 'Microsoft Copilot', 'm365-app-copilot'],
    ['microsoft 365', 'Bot', 'm365-bot'],
    ['ai + machine learning', 'Foundry Agent Service', '038470523-icon-service-Foundry-Agent-Service'],
    ['ai + machine learning', 'Bot Services', '10165-icon-service-Bot-Services'],
    ['new icons', 'Agentic Web Apps', '034296882-icon-service-Agentic-Web-Apps'],
  ];

  for (const [sourceCategory, displayName, fileName] of copilotIcons) {
    assert.equal(
      classifyIconPaletteCategory(sourceCategory, displayName, fileName),
      'microsoft-copilot',
      `${sourceCategory}/${fileName} should be a Copilot icon`,
    );
    assert.ok(
      existsSync(path.join(iconRoot, sourceCategory, `${fileName}.svg`)),
      `${sourceCategory}/${fileName}.svg is missing from the icon library`,
    );
  }
});

test('the AI prompts can emit the new icon categories', () => {
  const prompts = [
    'src/services/azureOpenAI.ts',
    'src/services/referenceArchitectureAI.ts',
  ];
  for (const relativePath of prompts) {
    const source = readFileSync(path.join(repoRoot, relativePath), 'utf8');
    assert.ok(
      source.includes('"power platform"'),
      `${relativePath} does not offer the "power platform" category`,
    );
    assert.ok(
      source.includes('"dynamics 365"'),
      `${relativePath} does not offer the "dynamics 365" category`,
    );
    assert.ok(
      source.includes('"microsoft 365"'),
      `${relativePath} does not offer the "microsoft 365" category`,
    );
  }
});

test('the manifest matches the catalog', () => {
  const manifest = readManifest();
  assert.equal(manifest.packageVersion, MICROSOFT_PRODUCT_ICON_PACKAGE_VERSION);
  assert.equal(manifest.iconCount, MICROSOFT_PRODUCT_ICON_CATALOG.length);
  assert.equal(manifest.files.length, MICROSOFT_PRODUCT_ICON_CATALOG.length);

  const manifestByPath = new Map(manifest.files.map(entry => [entry.path, entry]));
  for (const definition of MICROSOFT_PRODUCT_ICON_CATALOG) {
    const entry = manifestByPath.get(`${definition.category}/${definition.fileName}.svg`);
    assert.ok(entry, `${definition.fileName}.svg is missing from the manifest`);
    assert.equal(entry.serviceName, definition.serviceName);
    assert.equal(entry.family, definition.family);
    assert.match(entry.sha256, /^[a-f0-9]{64}$/);
  }
});
