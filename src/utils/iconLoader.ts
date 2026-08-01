// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import azureIconManifest from '../../Azure_Public_Service_Icons/manifest.json';
import {
  classifyIconPaletteCategory,
  getIconPaletteCategory,
  iconPaletteCategories,
  type IconPaletteCategoryId,
} from '../data/iconCatalog';
import { getFabricIconByFileName } from '../data/fabricIconCatalog';
import { SERVICE_ICON_MAP } from '../data/serviceIconMapping';

export interface AzureIcon {
  id: string;
  name: string;
  /** Stable service key used for pricing and semantic classification. */
  serviceName: string;
  /** Physical icon folder retained for saved-diagram and pricing compatibility. */
  category: string;
  /** Curated, meaning-based category shown in the service palette. */
  paletteCategory: IconPaletteCategoryId;
  path: string;
  searchTerms: string[];
  source: 'official-azure' | 'fabric' | 'supplemental';
}

export interface IconLibraryStats {
  azureVersion: string;
  officialAzureIcons: number;
  fabricIcons: number;
  searchableIcons: number;
}

const ICON_ROOT = '/Azure_Public_Service_Icons/Icons/';
const iconModules = import.meta.glob('/Azure_Public_Service_Icons/Icons/**/*.svg', {
  eager: false,
  query: '?url',
  import: 'default',
}) as Record<string, () => Promise<string>>;

const officialRelativePaths = new Set(azureIconManifest.files.map(file => file.path));
const officialFileNames = new Set(
  azureIconManifest.files.map(file => file.path.slice(file.path.lastIndexOf('/') + 1)),
);

// These friendly filenames predate the official numbered package names. Keep
// them loadable for saved diagrams and service mappings, but show only the
// current official icon in the palette.
const LEGACY_FRIENDLY_PATHS = new Set([
  'ai + machine learning/azure-cognitive-search.svg',
  'ai + machine learning/azure-machine-learning.svg',
  'ai + machine learning/azure-openai.svg',
  'ai + machine learning/azure-speech.svg',
  'ai + machine learning/cognitive-services.svg',
  'ai + machine learning/computer-vision.svg',
  'ai + machine learning/custom-vision.svg',
  'ai + machine learning/document-intelligence.svg',
  'ai + machine learning/language.svg',
  'ai + machine learning/translator.svg',
  'analytics/azure-synapse-analytics.svg',
  'analytics/data-factory.svg',
  'analytics/event-hubs.svg',
  'analytics/stream-analytics.svg',
  'app services/app-service.svg',
  'app services/cdn-profiles.svg',
  'compute/azure-functions.svg',
  'compute/virtual-machines.svg',
  'containers/azure-kubernetes-service.svg',
  'containers/container-instances.svg',
  'containers/container-registry.svg',
  'databases/azure-cosmos-db.svg',
  'databases/azure-database-mysql.svg',
  'databases/redis-cache.svg',
  'databases/sql-database.svg',
  'integration/api-management.svg',
  'integration/logic-apps.svg',
  'integration/service-bus.svg',
  'monitor/application-insights.svg',
  'monitor/log-analytics.svg',
  'networking/application-gateway.svg',
  'networking/azure-front-door.svg',
  'security/key-vault.svg',
  'storage/storage-account.svg',
]);

const HIDDEN_LEGACY_PALETTE_PATHS = new Set([
  'general/10840-icon-service-Storage-Queue.svg',
  'general/10841-icon-service-Table.svg',
  'other/02989-icon-service-Container-Apps-Environments.svg',
  'web/00049-icon-service-App-Service-Certificates.svg',
]);

const aliasesByIcon = new Map<string, Set<string>>();
for (const [serviceName, mapping] of Object.entries(SERVICE_ICON_MAP)) {
  const key = `${mapping.category}/${mapping.iconFile}`;
  const aliases = aliasesByIcon.get(key) ?? new Set<string>();
  aliases.add(serviceName);
  aliases.add(mapping.displayName);
  mapping.aliases.forEach(alias => aliases.add(alias));
  aliasesByIcon.set(key, aliases);
}

export const iconCategories = [...new Set(
  Object.keys(iconModules).flatMap(path => {
    if (!path.startsWith(ICON_ROOT)) return [];
    const relativePath = path.slice(ICON_ROOT.length);
    const separatorIndex = relativePath.indexOf('/');
    return separatorIndex > 0 ? [relativePath.slice(0, separatorIndex)] : [];
  }),
)].sort((left, right) => left.localeCompare(right));

export const paletteCategories = iconPaletteCategories;

interface IconMetadataCache {
  all: AzureIcon[];
  bySourceCategory: Map<string, AzureIcon[]>;
  byPaletteCategory: Map<IconPaletteCategoryId, AzureIcon[]>;
}

let iconMetadataCache: IconMetadataCache | undefined;
const iconUrlCache = new Map<string, Promise<string>>();

function formatIconName(fileNameWithoutExtension: string): string {
  return fileNameWithoutExtension
    .replace(/^\d+\s*-icon-service-/i, '')
    .replace(/[-_]+/g, ' ')
    .trim()
    .split(/\s+/)
    .map(word => {
      const upper = word.toUpperCase();
      if ([
        'AI', 'ML', 'BI', 'CDN', 'SQL', 'IOT', 'API', 'VM', 'VMS', 'AKS',
        'ACR', 'ACI', 'DB', 'KQL', 'RDL', 'RTI', 'FHIR', 'VPN', 'DNS', 'SAP',
        'HPC', 'HSM', 'SSH', 'WAF', 'OS',
      ].includes(upper)) {
        return upper === 'IOT' ? 'IoT' : upper;
      }
      if (word.toLowerCase() === 'openai') return 'OpenAI';
      if (word.toLowerCase() === 'postgresql') return 'PostgreSQL';
      if (word.toLowerCase() === 'mysql') return 'MySQL';
      if (word.toLowerCase() === 'redis') return 'Redis';
      if (word.toLowerCase() === 'cosmos') return 'Cosmos';
      if (word.toLowerCase() === 'promethus') return 'Prometheus';
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join(' ');
}

function normalizeSearchValue(value: string): string {
  return value
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/[+&/_.(),:;()[\]{}-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isHiddenFromPalette(relativePath: string, fileName: string): boolean {
  if (
    LEGACY_FRIENDLY_PATHS.has(relativePath)
    || HIDDEN_LEGACY_PALETTE_PATHS.has(relativePath)
  ) {
    return true;
  }

  // V24 moved several icons out of "new icons", "other", and other legacy
  // folders. If the same official filename exists elsewhere, expose only its
  // current package location.
  return !officialRelativePaths.has(relativePath) && officialFileNames.has(fileName);
}

function getIconMetadataCache(): IconMetadataCache {
  if (iconMetadataCache) return iconMetadataCache;

  const bySourceCategory = new Map<string, AzureIcon[]>(
    iconCategories.map(category => [category, []]),
  );
  const byPaletteCategory = new Map<IconPaletteCategoryId, AzureIcon[]>(
    iconPaletteCategories.map(category => [category.id, []]),
  );
  const all: AzureIcon[] = [];

  for (const path of Object.keys(iconModules)) {
    if (!path.startsWith(ICON_ROOT)) continue;

    const relativePath = path.slice(ICON_ROOT.length);
    const separatorIndex = relativePath.indexOf('/');
    if (separatorIndex < 1) continue;

    const sourceCategory = relativePath.slice(0, separatorIndex);
    const fileName = relativePath.slice(separatorIndex + 1);
    if (isHiddenFromPalette(relativePath, fileName)) continue;

    const fileNameWithoutExtension = fileName.replace(/\.svg$/i, '');
    const fabricDefinition = sourceCategory === 'fabric'
      ? getFabricIconByFileName(fileNameWithoutExtension)
      : undefined;
    const name = fabricDefinition?.displayName ?? formatIconName(fileNameWithoutExtension);
    const paletteCategory = classifyIconPaletteCategory(
      sourceCategory,
      name,
      fileNameWithoutExtension,
    );
    const categoryDefinition = getIconPaletteCategory(paletteCategory);
    const mappedAliases = aliasesByIcon.get(`${sourceCategory}/${fileNameWithoutExtension}`);
    const searchTerms = [
      sourceCategory,
      fileNameWithoutExtension.replace(/[-_]+/g, ' '),
      categoryDefinition.label.en,
      categoryDefinition.label.ja,
      categoryDefinition.description.en,
      categoryDefinition.description.ja,
      ...categoryDefinition.keywords,
      ...(mappedAliases ? [...mappedAliases] : []),
      ...(fabricDefinition
        ? [
            fabricDefinition.serviceName,
            fabricDefinition.group,
            fabricDefinition.kind,
            ...fabricDefinition.aliases,
          ]
        : []),
    ];

    const icon: AzureIcon = {
      id: `${sourceCategory}/${fileNameWithoutExtension}`,
      name,
      serviceName: fabricDefinition?.serviceName ?? name,
      category: sourceCategory,
      paletteCategory,
      path,
      searchTerms,
      source: officialRelativePaths.has(relativePath)
        ? 'official-azure'
        : sourceCategory === 'fabric'
          ? 'fabric'
          : 'supplemental',
    };

    all.push(icon);
    bySourceCategory.get(sourceCategory)?.push(icon);
    byPaletteCategory.get(paletteCategory)?.push(icon);
  }

  const sortIcons = (icons: AzureIcon[]) => {
    icons.sort((left, right) => left.name.localeCompare(right.name));
  };
  sortIcons(all);
  bySourceCategory.forEach(sortIcons);
  byPaletteCategory.forEach(sortIcons);

  iconMetadataCache = { all, bySourceCategory, byPaletteCategory };
  return iconMetadataCache;
}

export function iconMatchesSearch(icon: AzureIcon, query: string): boolean {
  const tokens = normalizeSearchValue(query).split(' ').filter(Boolean);
  if (tokens.length === 0) return true;

  const searchable = normalizeSearchValue([
    icon.name,
    icon.category,
    ...icon.searchTerms,
  ].join(' '));
  return tokens.every(token => searchable.includes(token));
}

export async function loadIconsFromCategory(category: string): Promise<AzureIcon[]> {
  return getIconMetadataCache().bySourceCategory.get(category) ?? [];
}

export async function loadIconsFromPaletteCategory(
  category: IconPaletteCategoryId,
): Promise<AzureIcon[]> {
  return getIconMetadataCache().byPaletteCategory.get(category) ?? [];
}

export function getIconLibraryStats(): IconLibraryStats {
  const all = getIconMetadataCache().all;
  return {
    azureVersion: azureIconManifest.packageVersion,
    officialAzureIcons: azureIconManifest.officialIconCount,
    fabricIcons: all.filter(icon => icon.source === 'fabric').length,
    searchableIcons: all.length,
  };
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
