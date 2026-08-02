// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';

const STORAGE_KEY = 'azurediagarm.icon-workspace.v1';
const MAX_FAVORITES = 300;
const MAX_RECENT = 24;
const MAX_COLLECTIONS = 20;
const MAX_COLLECTION_ICONS = 300;

export interface IconCollection {
  id: string;
  name: string;
  iconIds: string[];
  createdAt: string;
  updatedAt: string;
}

export interface IconWorkspaceState {
  favoriteIds: string[];
  recentIds: string[];
  collections: IconCollection[];
}

export const EMPTY_ICON_WORKSPACE: IconWorkspaceState = {
  favoriteIds: [],
  recentIds: [],
  collections: [],
};

function uniqueIds(value: unknown, limit: number): string[] {
  if (!Array.isArray(value)) return [];
  const seen = new Set<string>();
  const result: string[] = [];
  for (const entry of value) {
    if (typeof entry !== 'string') continue;
    const id = entry.trim().slice(0, 240);
    if (!id || seen.has(id)) continue;
    seen.add(id);
    result.push(id);
    if (result.length >= limit) break;
  }
  return result;
}

function normalizeCollection(value: unknown, index: number): IconCollection | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const raw = value as Partial<IconCollection>;
  const name = typeof raw.name === 'string' ? raw.name.trim().slice(0, 60) : '';
  if (!name) return null;
  const now = new Date().toISOString();
  return {
    id: typeof raw.id === 'string' && raw.id.trim()
      ? raw.id.trim().slice(0, 100)
      : `collection-${index + 1}`,
    name,
    iconIds: uniqueIds(raw.iconIds, MAX_COLLECTION_ICONS),
    createdAt: typeof raw.createdAt === 'string' ? raw.createdAt : now,
    updatedAt: typeof raw.updatedAt === 'string' ? raw.updatedAt : now,
  };
}

export function normalizeIconWorkspace(value: unknown): IconWorkspaceState {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ...EMPTY_ICON_WORKSPACE };
  }
  const raw = value as Partial<IconWorkspaceState>;
  const collections = Array.isArray(raw.collections)
    ? raw.collections
        .slice(0, MAX_COLLECTIONS)
        .map(normalizeCollection)
        .filter((collection): collection is IconCollection => Boolean(collection))
    : [];
  return {
    favoriteIds: uniqueIds(raw.favoriteIds, MAX_FAVORITES),
    recentIds: uniqueIds(raw.recentIds, MAX_RECENT),
    collections,
  };
}

export function loadIconWorkspace(): IconWorkspaceState {
  const raw = readLocalStorage(STORAGE_KEY);
  if (!raw) return { ...EMPTY_ICON_WORKSPACE };
  try {
    return normalizeIconWorkspace(JSON.parse(raw));
  } catch {
    return { ...EMPTY_ICON_WORKSPACE };
  }
}

export function saveIconWorkspace(state: IconWorkspaceState): void {
  writeLocalStorage(STORAGE_KEY, JSON.stringify(normalizeIconWorkspace(state)));
}

export function toggleFavoriteIcon(
  state: IconWorkspaceState,
  iconId: string,
): IconWorkspaceState {
  const favoriteIds = state.favoriteIds.includes(iconId)
    ? state.favoriteIds.filter((id) => id !== iconId)
    : [iconId, ...state.favoriteIds].slice(0, MAX_FAVORITES);
  return { ...state, favoriteIds };
}

export function recordRecentIcon(
  state: IconWorkspaceState,
  iconId: string,
): IconWorkspaceState {
  return {
    ...state,
    recentIds: [iconId, ...state.recentIds.filter((id) => id !== iconId)]
      .slice(0, MAX_RECENT),
  };
}

function newCollectionId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `collection-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

export function createIconCollection(
  state: IconWorkspaceState,
  name: string,
): { state: IconWorkspaceState; collection: IconCollection | null } {
  const normalizedName = name.trim().slice(0, 60);
  const duplicateName = state.collections.some(
    (collection) => collection.name.localeCompare(normalizedName, undefined, {
      sensitivity: 'accent',
    }) === 0,
  );
  if (!normalizedName || duplicateName || state.collections.length >= MAX_COLLECTIONS) {
    return { state, collection: null };
  }
  const now = new Date().toISOString();
  const collection: IconCollection = {
    id: newCollectionId(),
    name: normalizedName,
    iconIds: [],
    createdAt: now,
    updatedAt: now,
  };
  return {
    state: { ...state, collections: [...state.collections, collection] },
    collection,
  };
}

export function deleteIconCollection(
  state: IconWorkspaceState,
  collectionId: string,
): IconWorkspaceState {
  return {
    ...state,
    collections: state.collections.filter((collection) => collection.id !== collectionId),
  };
}

export function toggleIconInCollection(
  state: IconWorkspaceState,
  collectionId: string,
  iconId: string,
): IconWorkspaceState {
  const now = new Date().toISOString();
  return {
    ...state,
    collections: state.collections.map((collection) => {
      if (collection.id !== collectionId) return collection;
      const iconIds = collection.iconIds.includes(iconId)
        ? collection.iconIds.filter((id) => id !== iconId)
        : [iconId, ...collection.iconIds].slice(0, MAX_COLLECTION_ICONS);
      return { ...collection, iconIds, updatedAt: now };
    }),
  };
}
