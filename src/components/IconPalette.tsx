// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  Folder,
  FolderPlus,
  Grid3X3,
  History,
  List,
  PanelLeftClose,
  PanelLeftOpen,
  Plus,
  Search,
  Sparkles,
  Star,
  Trash2,
  X,
} from 'lucide-react';
import {
  iconMatchesSearch,
  loadIcon,
  loadIconsFromPaletteCategory,
  paletteCategories,
  type AzureIcon,
} from '../utils/iconLoader';
import {
  createIconCollection,
  deleteIconCollection,
  loadIconWorkspace,
  recordRecentIcon,
  saveIconWorkspace,
  toggleFavoriteIcon,
  toggleIconInCollection,
} from '../services/iconWorkspaceService';
import {
  deduplicatePaletteIcons,
  normalizeIconDiscoveryText,
  splitIconSearchHighlight,
} from '../utils/iconDiscovery';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import VirtualizedIconGrid from './VirtualizedIconGrid';
import ResponsiveDrawer from './ResponsiveDrawer';
import './IconPalette.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import type { IconPaletteCategoryId } from '../data/iconCatalog';

interface IconPaletteProps {
  forceCollapsed?: number;
  openSignal?: number;
  onAddIcon?: (icon: AzureIcon) => void;
}

type PaletteView = 'catalog' | 'favorites' | 'recent' | 'recommended' | 'collections';
type PaletteLayout = 'grid' | 'list';
const PALETTE_VIEW_ORDER: PaletteView[] = [
  'catalog',
  'favorites',
  'recent',
  'recommended',
  'collections',
];
const PALETTE_LAYOUT_STORAGE_KEY = 'azure-diagram-builder.paletteLayout.v1';
const RECOMMENDED_SERVICE_NAMES = [
  ['App Services'],
  ['Function Apps'],
  ['Virtual Machine'],
  ['Kubernetes Services'],
  ['SQL Database'],
  ['Azure Cosmos DB'],
  ['Storage Accounts'],
  ['Key Vaults'],
  ['API Management Services'],
  ['Azure Service Bus'],
  ['Front Door And CDN Profiles'],
  ['Application Insights'],
  ['Log Analytics Workspaces'],
  ['Virtual Networks'],
  ['Load Balancers'],
  ['Azure OpenAI'],
  ['Cognitive Search'],
] as const;

const IconPalette: React.FC<IconPaletteProps> = ({ forceCollapsed, openSignal, onAddIcon }) => {
  const { t, language } = useLanguage();
  const isCompactPalette = useMediaQuery(MEDIA_QUERIES.compactOrShortWorkspace);
  const [isCollapsed, setIsCollapsed] = useState(
    () => (
      typeof window !== 'undefined'
      && window.matchMedia(MEDIA_QUERIES.compactOrShortWorkspace).matches
    ),
  );
  const [expandedCategories, setExpandedCategories] = useState<Set<IconPaletteCategoryId>>(
    new Set(['ai']),
  );
  const [searchTerm, setSearchTerm] = useState('');
  const [activeView, setActiveView] = useState<PaletteView>('recommended');
  const [layout, setLayout] = useState<PaletteLayout>(() => (
    readLocalStorage(PALETTE_LAYOUT_STORAGE_KEY) === 'list' ? 'list' : 'grid'
  ));
  const [categoryIcons, setCategoryIcons] = useState<Map<IconPaletteCategoryId, AzureIcon[]>>(
    new Map(),
  );
  const [canonicalIconIds, setCanonicalIconIds] = useState<Map<string, string>>(new Map());
  const [iconUrls, setIconUrls] = useState<Map<string, string>>(new Map());
  const [workspace, setWorkspace] = useState(loadIconWorkspace);
  const [selectedCollectionId, setSelectedCollectionId] = useState(
    () => workspace.collections[0]?.id || '',
  );
  const [collectionTargetId, setCollectionTargetId] = useState<string | null>(null);
  const viewTabRefs = useRef<Partial<Record<PaletteView, HTMLButtonElement>>>({});

  const handleViewTabKeyDown = useCallback((
    event: React.KeyboardEvent<HTMLButtonElement>,
    currentView: PaletteView,
  ) => {
    const currentIndex = PALETTE_VIEW_ORDER.indexOf(currentView);
    let nextIndex: number | null = null;
    if (event.key === 'ArrowRight') {
      nextIndex = (currentIndex + 1) % PALETTE_VIEW_ORDER.length;
    } else if (event.key === 'ArrowLeft') {
      nextIndex = (currentIndex - 1 + PALETTE_VIEW_ORDER.length) % PALETTE_VIEW_ORDER.length;
    } else if (event.key === 'Home') {
      nextIndex = 0;
    } else if (event.key === 'End') {
      nextIndex = PALETTE_VIEW_ORDER.length - 1;
    }
    if (nextIndex === null) return;

    event.preventDefault();
    const nextView = PALETTE_VIEW_ORDER[nextIndex];
    setActiveView(nextView);
    window.requestAnimationFrame(() => viewTabRefs.current[nextView]?.focus());
  }, []);

  useEffect(() => {
    saveIconWorkspace(workspace);
  }, [workspace]);

  useEffect(() => {
    writeLocalStorage(PALETTE_LAYOUT_STORAGE_KEY, layout);
  }, [layout]);

  useEffect(() => {
    if (forceCollapsed) setIsCollapsed(true);
  }, [forceCollapsed]);

  useEffect(() => {
    if (openSignal) setIsCollapsed(false);
  }, [openSignal]);

  useEffect(() => {
    const compactViewport = window.matchMedia(MEDIA_QUERIES.compactOrShortWorkspace);
    const collapseForCompactViewport = (event: MediaQueryListEvent) => {
      if (event.matches) setIsCollapsed(true);
    };
    compactViewport.addEventListener('change', collapseForCompactViewport);
    return () => compactViewport.removeEventListener('change', collapseForCompactViewport);
  }, []);

  useEffect(() => {
    if (
      selectedCollectionId
      && workspace.collections.some((collection) => collection.id === selectedCollectionId)
    ) {
      return;
    }
    setSelectedCollectionId(workspace.collections[0]?.id || '');
  }, [selectedCollectionId, workspace.collections]);

  const loadIconUrls = useCallback(async (icons: AzureIcon[]) => {
    const unloaded = icons.filter((icon) => !iconUrls.has(icon.path));
    if (unloaded.length === 0) return;

    const loadedEntries = await Promise.all(
      unloaded.map(async (icon) => [icon.path, await loadIcon(icon.path)] as const),
    );

    setIconUrls((previous) => {
      const next = new Map(previous);
      let changed = false;
      for (const [path, url] of loadedEntries) {
        if (url && next.get(path) !== url) {
          next.set(path, url);
          changed = true;
        }
      }
      return changed ? next : previous;
    });
  }, [iconUrls]);

  useEffect(() => {
    let cancelled = false;

    const loadAllIconMetadata = async () => {
      const entries = await Promise.all(
        paletteCategories.map(async (category) => (
          [category.id, await loadIconsFromPaletteCategory(category.id)] as const
        )),
      );
      if (cancelled) return;

      const { icons, canonicalIdById } = deduplicatePaletteIcons(
        entries.flatMap(([, categoryItems]) => categoryItems),
      );
      const deduplicatedCategories = new Map<IconPaletteCategoryId, AzureIcon[]>(
        paletteCategories.map(category => [category.id, []]),
      );
      for (const icon of icons) {
        deduplicatedCategories.get(icon.paletteCategory)?.push(icon);
      }
      setCategoryIcons(deduplicatedCategories);
      setCanonicalIconIds(canonicalIdById);
    };

    void loadAllIconMetadata();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (canonicalIconIds.size === 0) return;
    setWorkspace((previous) => {
      const canonicalizeIds = (ids: string[]) => [...new Set(
        ids.map(id => canonicalIconIds.get(id) || id),
      )];
      const next = {
        ...previous,
        favoriteIds: canonicalizeIds(previous.favoriteIds),
        recentIds: canonicalizeIds(previous.recentIds),
        collections: previous.collections.map(collection => ({
          ...collection,
          iconIds: canonicalizeIds(collection.iconIds),
        })),
      };
      return JSON.stringify(next) === JSON.stringify(previous) ? previous : next;
    });
  }, [canonicalIconIds]);

  const allIcons = useMemo(
    () => paletteCategories.flatMap((category) => categoryIcons.get(category.id) || []),
    [categoryIcons],
  );
  const iconsById = useMemo(
    () => new Map(allIcons.map((icon) => [icon.id, icon])),
    [allIcons],
  );
  const iconsForIds = useCallback((ids: string[]) => (
    ids.map((id) => iconsById.get(id)).filter((icon): icon is AzureIcon => Boolean(icon))
  ), [iconsById]);
  const favoriteIcons = useMemo(
    () => iconsForIds(workspace.favoriteIds),
    [iconsForIds, workspace.favoriteIds],
  );
  const recentIcons = useMemo(
    () => iconsForIds(workspace.recentIds),
    [iconsForIds, workspace.recentIds],
  );
  const recommendedIcons = useMemo(() => {
    const iconsByName = new Map(
      allIcons.map(icon => [normalizeIconDiscoveryText(icon.name), icon]),
    );
    const selected = RECOMMENDED_SERVICE_NAMES.flatMap((candidateNames) => {
      for (const candidate of candidateNames) {
        const match = iconsByName.get(normalizeIconDiscoveryText(candidate));
        if (match) return [match];
      }
      return [];
    });
    return [...new Map(selected.map(icon => [icon.id, icon])).values()];
  }, [allIcons]);
  const selectedCollection = workspace.collections.find(
    (collection) => collection.id === selectedCollectionId,
  );
  const collectionIcons = useMemo(
    () => iconsForIds(selectedCollection?.iconIds || []),
    [iconsForIds, selectedCollection?.iconIds],
  );

  const searchIcons = useCallback((icons: AzureIcon[]) => {
    const query = searchTerm.trim();
    return query ? icons.filter((icon) => iconMatchesSearch(icon, query)) : icons;
  }, [searchTerm]);

  const catalogSearchResults = useMemo(
    () => searchIcons(allIcons),
    [allIcons, searchIcons],
  );
  const activeViewIcons = useMemo(() => {
    if (activeView === 'favorites') return searchIcons(favoriteIcons);
    if (activeView === 'recent') return searchIcons(recentIcons);
    if (activeView === 'recommended') return searchIcons(recommendedIcons);
    if (activeView === 'collections') return searchIcons(collectionIcons);
    return catalogSearchResults;
  }, [
    activeView,
    catalogSearchResults,
    collectionIcons,
    favoriteIcons,
    recentIcons,
    recommendedIcons,
    searchIcons,
  ]);

  const toggleCategory = (category: IconPaletteCategoryId) => {
    setExpandedCategories((previous) => {
      const next = new Set(previous);
      if (next.has(category)) next.delete(category);
      else next.add(category);
      return next;
    });
  };

  const markRecent = useCallback((iconId: string) => {
    setWorkspace((previous) => recordRecentIcon(previous, iconId));
  }, []);

  const onDragStart = useCallback((event: React.DragEvent, icon: AzureIcon) => {
    markRecent(icon.id);
    event.dataTransfer.setData('application/reactflow', 'azureNode');
    event.dataTransfer.setData('iconPath', icon.path);
    event.dataTransfer.setData('iconName', icon.name);
    event.dataTransfer.setData('iconServiceName', icon.serviceName);
    event.dataTransfer.setData('iconCategory', icon.category);
    event.dataTransfer.effectAllowed = 'move';
  }, [markRecent]);

  const addIconToCanvas = useCallback((icon: AzureIcon) => {
    markRecent(icon.id);
    onAddIcon?.(icon);
    if (window.matchMedia(MEDIA_QUERIES.compactOrShortWorkspace).matches) {
      setIsCollapsed(true);
    }
  }, [markRecent, onAddIcon]);

  const requestNewCollection = useCallback((iconId?: string) => {
    const name = window.prompt(localize(language, {
      en: 'Name the new icon collection:',
      ja: '新しいアイコン コレクション名:',
    }));
    if (!name) return;
    const result = createIconCollection(workspace, name);
    if (!result.collection) {
      window.alert(localize(language, {
        en: 'The collection could not be created. Use a unique name and keep no more than 20 collections.',
        ja: 'コレクションを作成できませんでした。名前を指定し、コレクション数を20件以下にしてください。',
      }));
      return;
    }
    const next = iconId
      ? toggleIconInCollection(result.state, result.collection.id, iconId)
      : result.state;
    setWorkspace(next);
    setSelectedCollectionId(result.collection.id);
    setActiveView('collections');
  }, [language, workspace]);

  const removeSelectedCollection = () => {
    if (!selectedCollection) return;
    if (!window.confirm(localize(language, {
      en: `Delete the "${selectedCollection.name}" collection?`,
      ja: `「${selectedCollection.name}」コレクションを削除しますか？`,
    }))) return;
    setWorkspace((previous) => deleteIconCollection(previous, selectedCollection.id));
  };

  const renderIcon = useCallback((icon: AzureIcon) => {
    const iconUrl = iconUrls.get(icon.path);
    const isFavorite = workspace.favoriteIds.includes(icon.id);
    const isCollected = workspace.collections.some(
      (collection) => collection.iconIds.includes(icon.id),
    );
    return (
      <div className="icon-item" key={icon.id}>
        <button
          type="button"
          className="icon-item-main"
          draggable
          onDragStart={(event) => onDragStart(event, icon)}
          onClick={() => addIconToCanvas(icon)}
          aria-label={t('palette.addService', { service: icon.name })}
          title={t('palette.addServiceHint', { service: icon.name })}
        >
          <span className="icon-add-indicator" aria-hidden="true">
            <Plus size={12} />
          </span>
          {iconUrl ? (
            <img
              src={iconUrl}
              alt=""
              className="icon-image"
              loading="lazy"
              decoding="async"
            />
          ) : (
            <div className="icon-placeholder" aria-hidden="true">
              {t('Loading...')}
            </div>
          )}
          <span className="icon-label">
            {splitIconSearchHighlight(icon.name, searchTerm).map((segment, index) => (
              segment.matched
                ? <mark key={`${segment.text}-${index}`}>{segment.text}</mark>
                : <React.Fragment key={`${segment.text}-${index}`}>{segment.text}</React.Fragment>
            ))}
          </span>
        </button>
        <div className="icon-item-actions">
          <button
            type="button"
            className={`icon-workspace-action${isFavorite ? ' active' : ''}`}
            onClick={() => setWorkspace((previous) => toggleFavoriteIcon(previous, icon.id))}
            title={localize(language, {
              en: isFavorite ? 'Remove from favorites' : 'Add to favorites',
              ja: isFavorite ? 'お気に入りから削除' : 'お気に入りに追加',
            })}
            aria-pressed={isFavorite}
          >
            <Star size={13} fill={isFavorite ? 'currentColor' : 'none'} />
          </button>
          <button
            type="button"
            className={`icon-workspace-action${isCollected ? ' active' : ''}`}
            onClick={() => setCollectionTargetId(icon.id)}
            title={localize(language, {
              en: 'Manage collection membership',
              ja: 'コレクションへの登録を管理',
            })}
          >
            <FolderPlus size={13} />
          </button>
        </div>
      </div>
    );
  }, [
    addIconToCanvas,
    iconUrls,
    language,
    onDragStart,
    searchTerm,
    t,
    workspace.collections,
    workspace.favoriteIds,
  ]);

  const viewCount = activeView === 'catalog'
    ? allIcons.length
    : activeView === 'favorites'
      ? favoriteIcons.length
      : activeView === 'recent'
        ? recentIcons.length
        : activeView === 'recommended'
          ? recommendedIcons.length
          : collectionIcons.length;
  const collectionTarget = collectionTargetId ? iconsById.get(collectionTargetId) : undefined;

  const palettePanel = (
      <div
        className={`icon-palette palette-layout-${layout} ${isCollapsed ? 'collapsed' : ''}`}
        role="region"
        aria-label={t('Microsoft Services')}
      >
      {isCollapsed ? (
        <button
          type="button"
          className="palette-open-toggle"
          onClick={() => setIsCollapsed(false)}
          title={t('Open services panel')}
          aria-label={t('Open services panel')}
          aria-controls="palette-view-panel"
          aria-expanded="false"
        >
          <PanelLeftOpen size={18} aria-hidden="true" />
          <span>{t('Open services panel')}</span>
        </button>
      ) : (
        <>
          <div className="palette-header">
            <div className="palette-title-row">
              <div>
                <h2>{t('Microsoft Services')}</h2>
              </div>
              <button
                type="button"
                className="palette-close-toggle"
                onClick={() => setIsCollapsed(true)}
                title={t('Close services panel')}
                aria-label={t('Close services panel')}
                aria-controls="palette-view-panel"
                aria-expanded="true"
              >
                <PanelLeftClose size={16} aria-hidden="true" />
                <span>{t('Close services panel')}</span>
              </button>
            </div>
            <div className="search-box">
              <Search size={16} aria-hidden="true" />
              <input
                type="search"
                placeholder={localize(language, {
                  en: 'Search Azure, Fabric, Microsoft 365, Copilot...',
                  ja: 'Azure、Fabric、Microsoft 365、Copilotを検索...',
                })}
                aria-label={t('palette.searchLabel')}
                value={searchTerm}
                onChange={(event) => setSearchTerm(event.target.value)}
              />
              {searchTerm && (
                <button
                  type="button"
                  className="search-clear"
                  onClick={() => setSearchTerm('')}
                  aria-label={localize(language, { en: 'Clear icon search', ja: 'アイコン検索をクリア' })}
                  title={localize(language, { en: 'Clear search', ja: '検索をクリア' })}
                >
                  <X size={14} />
                </button>
              )}
            </div>
            <div className="palette-view-tabs" role="tablist" aria-label={localize(language, {
              en: 'Icon library views',
              ja: 'アイコン ライブラリ表示',
            })}>
              {([
                ['catalog', Grid3X3, localize(language, { en: 'All', ja: 'すべて' }), allIcons.length],
                ['favorites', Star, localize(language, { en: 'Favorites', ja: 'お気に入り' }), favoriteIcons.length],
                ['recent', History, localize(language, { en: 'Recent', ja: '最近' }), recentIcons.length],
                ['recommended', Sparkles, localize(language, { en: 'Recommended', ja: '推奨' }), recommendedIcons.length],
                ['collections', Folder, localize(language, { en: 'Collections', ja: 'コレクション' }), workspace.collections.length],
              ] as const).map(([id, Icon, label, count]) => (
                <button
                  type="button"
                  key={id}
                  id={`palette-view-tab-${id}`}
                  role="tab"
                  className={activeView === id ? 'active' : ''}
                  aria-selected={activeView === id}
                  aria-controls="palette-view-panel"
                  tabIndex={activeView === id ? 0 : -1}
                  ref={(element) => {
                    viewTabRefs.current[id] = element || undefined;
                  }}
                  onClick={() => setActiveView(id)}
                  onKeyDown={(event) => handleViewTabKeyDown(event, id)}
                  title={`${label} (${count})`}
                >
                  <Icon
                    size={14}
                    fill={id === 'favorites' && activeView === id ? 'currentColor' : 'none'}
                    aria-hidden="true"
                  />
                  <span>{label}</span>
                  <b>{count}</b>
                </button>
              ))}
            </div>
            <div className="palette-view-meta">
              <div className="palette-search-summary" role="status" aria-live="polite">
                {searchTerm.trim()
                  ? localize(language, {
                      en: `${activeViewIcons.length} matching icons`,
                      ja: `${activeViewIcons.length} 件のアイコン`,
                    })
                  : localize(language, {
                      en: `${viewCount} icons in this view`,
                      ja: `この表示に ${viewCount} 件`,
                    })}
              </div>
              <div
                className="palette-layout-switch"
                role="group"
                aria-label={localize(language, {
                  en: 'Service display layout',
                  ja: 'サービス表示レイアウト',
                })}
              >
                <button
                  type="button"
                  className={layout === 'grid' ? 'active' : ''}
                  aria-pressed={layout === 'grid'}
                  onClick={() => setLayout('grid')}
                  title={localize(language, { en: 'Grid view', ja: 'グリッド表示' })}
                  aria-label={localize(language, { en: 'Grid view', ja: 'グリッド表示' })}
                >
                  <Grid3X3 size={14} />
                </button>
                <button
                  type="button"
                  className={layout === 'list' ? 'active' : ''}
                  aria-pressed={layout === 'list'}
                  onClick={() => setLayout('list')}
                  title={localize(language, { en: 'List view', ja: 'リスト表示' })}
                  aria-label={localize(language, { en: 'List view', ja: 'リスト表示' })}
                >
                  <List size={15} />
                </button>
              </div>
            </div>
            <p className="palette-help">{t('palette.interactionHint')}</p>
          </div>

          <div
            className="palette-content"
            id="palette-view-panel"
            role="tabpanel"
            aria-labelledby={`palette-view-tab-${activeView}`}
          >
            {activeView === 'recommended' && searchTerm.trim() === '' && (
              <div className="palette-recommended-intro">
                <Sparkles size={16} aria-hidden="true" />
                <span>{localize(language, {
                  en: 'Common building blocks for starting an Azure architecture.',
                  ja: 'Azure アーキテクチャを始めるための代表的なノードです。',
                })}</span>
              </div>
            )}

            {activeView === 'catalog' && searchTerm.trim() === '' && paletteCategories.map((category) => {
              const isExpanded = expandedCategories.has(category.id);
              const icons = categoryIcons.get(category.id) || [];
              return (
                <section key={category.id} className="category-section">
                  <button
                    type="button"
                    className="category-header"
                    onClick={() => toggleCategory(category.id)}
                    aria-expanded={isExpanded}
                    title={localize(language, category.description)}
                  >
                    {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    <span className="category-title">{localize(language, category.label)}</span>
                    <span className="icon-count">{icons.length}</span>
                  </button>
                  {isExpanded && (
                    <>
                      <p className="category-description">
                        {localize(language, category.description)}
                      </p>
                      <VirtualizedIconGrid
                        icons={icons}
                        renderIcon={renderIcon}
                        onVisibleIconsChange={loadIconUrls}
                        ariaLabel={localize(language, category.label)}
                        layout={layout}
                      />
                    </>
                  )}
                </section>
              );
            })}

            {activeView === 'catalog' && searchTerm.trim() !== '' && activeViewIcons.length > 0 && (
              <VirtualizedIconGrid
                icons={activeViewIcons}
                renderIcon={renderIcon}
                onVisibleIconsChange={loadIconUrls}
                ariaLabel={localize(language, { en: 'Icon search results', ja: 'アイコン検索結果' })}
                maxHeight={640}
                layout={layout}
              />
            )}

            {(activeView === 'favorites'
              || activeView === 'recent'
              || activeView === 'recommended')
              && activeViewIcons.length > 0 && (
              <VirtualizedIconGrid
                icons={activeViewIcons}
                renderIcon={renderIcon}
                onVisibleIconsChange={loadIconUrls}
                ariaLabel={activeView === 'favorites'
                  ? localize(language, { en: 'Favorite icons', ja: 'お気に入りアイコン' })
                  : activeView === 'recent'
                    ? localize(language, { en: 'Recently used icons', ja: '最近使用したアイコン' })
                    : localize(language, { en: 'Recommended services', ja: '推奨サービス' })}
                maxHeight={640}
                layout={layout}
              />
            )}

            {activeView === 'collections' && (
              <>
                <div className="collection-toolbar">
                  <select
                    value={selectedCollectionId}
                    onChange={(event) => setSelectedCollectionId(event.target.value)}
                    aria-label={localize(language, { en: 'Selected collection', ja: '選択中のコレクション' })}
                    disabled={workspace.collections.length === 0}
                  >
                    {workspace.collections.length === 0 && (
                      <option value="">{localize(language, { en: 'No collections', ja: 'コレクションなし' })}</option>
                    )}
                    {workspace.collections.map((collection) => (
                      <option key={collection.id} value={collection.id}>
                        {collection.name} ({collection.iconIds.length})
                      </option>
                    ))}
                  </select>
                  <button type="button" onClick={() => requestNewCollection()}>
                    <FolderPlus size={14} />
                    <span>{localize(language, { en: 'New', ja: '新規' })}</span>
                  </button>
                  <button
                    type="button"
                    className="danger"
                    onClick={removeSelectedCollection}
                    disabled={!selectedCollection}
                    title={localize(language, { en: 'Delete collection', ja: 'コレクションを削除' })}
                  >
                    <Trash2 size={14} />
                  </button>
                </div>
                {activeViewIcons.length > 0 && (
                  <VirtualizedIconGrid
                    icons={activeViewIcons}
                    renderIcon={renderIcon}
                    onVisibleIconsChange={loadIconUrls}
                    ariaLabel={selectedCollection?.name || localize(language, {
                      en: 'Icon collection',
                      ja: 'アイコン コレクション',
                    })}
                    maxHeight={580}
                    layout={layout}
                  />
                )}
              </>
            )}

            {activeViewIcons.length === 0 && (
              <div className="palette-empty-search">
                {searchTerm.trim()
                  ? localize(language, {
                      en: 'No icons match every search keyword in this view.',
                      ja: 'この表示では検索キーワードに一致するアイコンがありません。',
                    })
                  : activeView === 'favorites'
                    ? localize(language, {
                        en: 'Star frequently used services to keep them here.',
                        ja: 'よく使うサービスへ星を付けると、ここに表示されます。',
                      })
                    : activeView === 'recent'
                      ? localize(language, {
                          en: 'Icons appear here after you add or drag them.',
                          ja: 'アイコンを追加またはドラッグすると、ここに表示されます。',
                        })
                      : activeView === 'recommended'
                        ? localize(language, {
                            en: 'Recommended services are still loading.',
                            ja: '推奨サービスを読み込んでいます。',
                          })
                      : activeView === 'collections'
                        ? localize(language, {
                            en: 'Create a collection, then use the folder button on any icon to organize it.',
                            ja: 'コレクションを作成し、各アイコンのフォルダー ボタンで整理できます。',
                          })
                        : localize(language, {
                            en: 'No icons match every search keyword.',
                            ja: 'すべての検索キーワードに一致するアイコンはありません。',
                          })}
              </div>
            )}
          </div>

          {collectionTarget && (
            <aside
              className="collection-assignment-panel"
              role="dialog"
              aria-label={localize(language, {
                en: `Add ${collectionTarget.name} to collections`,
                ja: `${collectionTarget.name} のコレクション登録`,
              })}
            >
              <div className="collection-assignment-header">
                <strong>{collectionTarget.name}</strong>
                <button
                  type="button"
                  onClick={() => setCollectionTargetId(null)}
                  aria-label={t('Close')}
                  title={t('Close')}
                >
                  <X size={15} />
                </button>
              </div>
              <div className="collection-assignment-list">
                {workspace.collections.map((collection) => (
                  <label key={collection.id}>
                    <input
                      type="checkbox"
                      checked={collection.iconIds.includes(collectionTarget.id)}
                      onChange={() => setWorkspace((previous) => (
                        toggleIconInCollection(previous, collection.id, collectionTarget.id)
                      ))}
                    />
                    <span>{collection.name}</span>
                    <small>{collection.iconIds.length}</small>
                  </label>
                ))}
                {workspace.collections.length === 0 && (
                  <p>{localize(language, {
                    en: 'No collections yet.',
                    ja: 'コレクションはまだありません。',
                  })}</p>
                )}
              </div>
              <button
                type="button"
                className="collection-assignment-new"
                onClick={() => requestNewCollection(collectionTarget.id)}
              >
                <FolderPlus size={14} />
                {localize(language, { en: 'Create collection', ja: 'コレクションを作成' })}
              </button>
            </aside>
          )}
        </>
      )}
      </div>
  );

  if (!isCollapsed && isCompactPalette) {
    return (
      <ResponsiveDrawer
        isOpen
        modal
        placement="left"
        className="palette-drawer-shell"
        backdropClassName="palette-backdrop"
        ariaLabel={t('Microsoft Services')}
        onClose={() => setIsCollapsed(true)}
        backgroundSelectors={[
          '.app > .app-header',
          '.workspace > .canvas-container',
          '.workspace > .workflow-panel',
        ]}
      >
        {palettePanel}
      </ResponsiveDrawer>
    );
  }

  return palettePanel;
};

export default React.memo(IconPalette);
