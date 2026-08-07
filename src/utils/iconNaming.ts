// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface SearchableIcon {
  name: string;
  searchTerms?: string[];
}

const ICON_DISPLAY_NAME_OVERRIDES: Record<string, string> = {
  '02658-icon-service-FHIR-Service': 'Azure Health Data Services FHIR service',
  '10044-icon-service-Cognitive-Search': 'Azure AI Search',
  '10162-icon-service-Cognitive-Services': 'Foundry Tools',
  '00819-icon-service-Form-Recognizers': 'Azure AI Document Intelligence',
  'azure-cognitive-search': 'Azure AI Search',
  'cognitive-services': 'Foundry Tools',
  'document-intelligence': 'Azure AI Document Intelligence',
};

const SUPERSEDED_ICON_FILES = new Set([
  '10212-icon-service-Azure-API-for-FHIR',
]);

export function getCurrentIconDisplayName(iconFile: string, derivedName: string): string {
  return ICON_DISPLAY_NAME_OVERRIDES[iconFile] || derivedName;
}

export function isSupersededIconFile(iconFile: string): boolean {
  return SUPERSEDED_ICON_FILES.has(iconFile);
}

export function matchesIconSearch(icon: SearchableIcon, term: string): boolean {
  const normalized = term.trim().toLowerCase();
  if (!normalized) return true;
  return [icon.name, ...(icon.searchTerms || [])]
    .some(value => value.toLowerCase().includes(normalized));
}
