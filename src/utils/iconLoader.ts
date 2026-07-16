// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { getFabricIconByFileName } from '../data/fabricIconCatalog';

export interface AzureIcon {
  id: string;
  name: string;
  category: string;
  path: string;
  searchTerms: string[];
}

export const iconCategories = [
  'ai + machine learning',
  'analytics',
  'app services',
  'azure ecosystem',
  'azure stack',
  'blockchain',
  'compute',
  'containers',
  'databases',
  'devops',
  'fabric',
  'general',
  'hybrid + multicloud',
  'identity',
  'integration',
  'intune',
  'iot',
  'management + governance',
  'menu',
  'migrate',
  'migration',
  'mixed reality',
  'mobile',
  'monitor',
  'networking',
  'new icons',
  'other',
  'security',
  'storage',
  'web',
];

// Keep legacy paths loadable for saved diagrams while exposing each service only
// in its canonical category in the left palette.
const HIDDEN_LEGACY_PALETTE_PATHS = new Set([
  '/Azure_Public_Service_Icons/Icons/general/10840-icon-service-Storage-Queue.svg',
  '/Azure_Public_Service_Icons/Icons/general/10841-icon-service-Table.svg',
  '/Azure_Public_Service_Icons/Icons/other/02989-icon-service-Container-Apps-Environments.svg',
  '/Azure_Public_Service_Icons/Icons/web/00049-icon-service-App-Service-Certificates.svg',
]);

const ICON_ROOT = '/Azure_Public_Service_Icons/Icons/';
const iconModules = import.meta.glob('/Azure_Public_Service_Icons/Icons/**/*.svg', {
  eager: false,
  query: '?url',
  import: 'default',
}) as Record<string, () => Promise<string>>;

let iconMetadataCache: Map<string, AzureIcon[]> | undefined;
const iconUrlCache = new Map<string, Promise<string>>();

function formatIconName(fileNameWithoutExtension: string): string {
  return fileNameWithoutExtension
    .replace(/^\d+-icon-service-/, '')
    .replace(/-/g, ' ')
    .split(' ')
    .map(word => {
      const upper = word.toUpperCase();
      if (['AI', 'ML', 'BI', 'CDN', 'SQL', 'IOT', 'API', 'VM', 'VMS', 'AKS', 'ACR', 'ACI', 'DB', 'KQL', 'RDL', 'RTI', 'FHIR'].includes(upper)) {
        return upper;
      }
      if (word.toLowerCase() === 'openai') return 'OpenAI';
      if (word.toLowerCase() === 'postgresql') return 'PostgreSQL';
      if (word.toLowerCase() === 'mysql') return 'MySQL';
      if (word.toLowerCase() === 'redis') return 'Redis';
      if (word.toLowerCase() === 'cosmos') return 'Cosmos';
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

function getIconMetadataCache(): Map<string, AzureIcon[]> {
  if (iconMetadataCache) return iconMetadataCache;

  const cache = new Map<string, AzureIcon[]>(
    iconCategories.map(category => [category, []]),
  );

  for (const path of Object.keys(iconModules)) {
    if (HIDDEN_LEGACY_PALETTE_PATHS.has(path) || !path.startsWith(ICON_ROOT)) continue;

    const relativePath = path.slice(ICON_ROOT.length);
    const separatorIndex = relativePath.indexOf('/');
    if (separatorIndex < 1) continue;

    const category = relativePath.slice(0, separatorIndex);
    const icons = cache.get(category);
    if (!icons) continue;

    const fileName = relativePath.slice(separatorIndex + 1);
    const fileNameWithoutExtension = fileName.replace(/\.svg$/i, '');
    const fabricDefinition = category === 'fabric'
      ? getFabricIconByFileName(fileNameWithoutExtension)
      : undefined;

    icons.push({
      id: fileNameWithoutExtension,
      name: fabricDefinition?.displayName ?? formatIconName(fileNameWithoutExtension),
      category,
      path,
      searchTerms: fabricDefinition
        ? [
            fabricDefinition.serviceName,
            fabricDefinition.group,
            fabricDefinition.kind,
            ...fabricDefinition.aliases,
          ]
        : [],
    });
  }

  for (const icons of cache.values()) {
    icons.sort((a, b) => a.name.localeCompare(b.name));
  }

  iconMetadataCache = cache;
  return cache;
}

export async function loadIconsFromCategory(category: string): Promise<AzureIcon[]> {
  return getIconMetadataCache().get(category) ?? [];
}

export async function loadIcon(path: string): Promise<string> {
  const iconModule = iconModules[path];
  if (!iconModule) return '';

  const cached = iconUrlCache.get(path);
  if (cached) return cached;

  const pending = iconModule().catch(error => {
    iconUrlCache.delete(path);
    console.error('Error loading icon:', error);
    return '';
  });
  iconUrlCache.set(path, pending);
  return pending;
}
