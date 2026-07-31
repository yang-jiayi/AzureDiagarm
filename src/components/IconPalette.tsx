// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import {
  ChevronDown,
  ChevronRight,
  PanelLeftClose,
  PanelLeftOpen,
  Plus,
  Search,
  X,
} from 'lucide-react';
import {
  getIconLibraryStats,
  iconMatchesSearch,
  loadIcon,
  loadIconsFromPaletteCategory,
  paletteCategories,
  type AzureIcon,
} from '../utils/iconLoader';
import './IconPalette.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import type { IconPaletteCategoryId } from '../data/iconCatalog';

interface IconPaletteProps {
  forceCollapsed?: number;
  onAddIcon?: (icon: AzureIcon) => void;
}

const libraryStats = getIconLibraryStats();

const IconPalette: React.FC<IconPaletteProps> = ({ forceCollapsed, onAddIcon }) => {
  const { t, language } = useLanguage();
  const [isCollapsed, setIsCollapsed] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(max-width: 640px)').matches,
  );
  const [expandedCategories, setExpandedCategories] = useState<Set<IconPaletteCategoryId>>(
    new Set(['ai']),
  );
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryIcons, setCategoryIcons] = useState<Map<IconPaletteCategoryId, AzureIcon[]>>(
    new Map(),
  );
  const [iconUrls, setIconUrls] = useState<Map<string, string>>(new Map());

  useEffect(() => {
    if (forceCollapsed) setIsCollapsed(true);
  }, [forceCollapsed]);

  useEffect(() => {
    const mobileViewport = window.matchMedia('(max-width: 640px)');
    const collapseForMobile = (event: MediaQueryListEvent) => {
      if (event.matches) setIsCollapsed(true);
    };
    mobileViewport.addEventListener('change', collapseForMobile);
    return () => mobileViewport.removeEventListener('change', collapseForMobile);
  }, []);

  const loadIconUrls = useCallback(async (icons: AzureIcon[]) => {
    if (icons.length === 0) return;

    const loadedEntries = await Promise.all(
      icons.map(async icon => [icon.path, await loadIcon(icon.path)] as const),
    );

    setIconUrls(previous => {
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
  }, []);

  const toggleCategory = (category: IconPaletteCategoryId) => {
    const isExpanding = !expandedCategories.has(category);
    setExpandedCategories(previous => {
      const next = new Set(previous);
      if (next.has(category)) next.delete(category);
      else next.add(category);
      return next;
    });

    if (isExpanding) {
      void loadIconUrls(categoryIcons.get(category) || []);
    }
  };

  useEffect(() => {
    let cancelled = false;

    const loadAllIconMetadata = async () => {
      const entries = await Promise.all(
        paletteCategories.map(async category => (
          [category.id, await loadIconsFromPaletteCategory(category.id)] as const
        )),
      );
      if (cancelled) return;

      const nextCategoryIcons = new Map<IconPaletteCategoryId, AzureIcon[]>(entries);
      setCategoryIcons(nextCategoryIcons);
      void loadIconUrls(nextCategoryIcons.get('ai') || []);
    };

    void loadAllIconMetadata();
    return () => {
      cancelled = true;
    };
  }, [loadIconUrls]);

  const onDragStart = (event: React.DragEvent, icon: AzureIcon) => {
    event.dataTransfer.setData('application/reactflow', 'azureNode');
    event.dataTransfer.setData('iconPath', icon.path);
    event.dataTransfer.setData('iconName', icon.name);
    event.dataTransfer.setData('iconCategory', icon.category);
    event.dataTransfer.effectAllowed = 'move';
  };

  const addIconToCanvas = (icon: AzureIcon) => {
    onAddIcon?.(icon);
    if (window.matchMedia('(max-width: 640px)').matches) {
      setIsCollapsed(true);
    }
  };

  const matchingIconsByCategory = useMemo(() => {
    const next = new Map<IconPaletteCategoryId, AzureIcon[]>();
    for (const category of paletteCategories) {
      const icons = categoryIcons.get(category.id) || [];
      next.set(
        category.id,
        searchTerm.trim() === ''
          ? icons
          : icons.filter(icon => iconMatchesSearch(icon, searchTerm)),
      );
    }
    return next;
  }, [categoryIcons, searchTerm]);

  const filteredCategories = useMemo(() => paletteCategories.filter(category => (
    (matchingIconsByCategory.get(category.id) || []).length > 0
  )), [matchingIconsByCategory]);

  const resultCount = useMemo(() => (
    filteredCategories.reduce(
      (total, category) => total + (matchingIconsByCategory.get(category.id)?.length ?? 0),
      0,
    )
  ), [filteredCategories, matchingIconsByCategory]);

  useEffect(() => {
    if (searchTerm.trim() === '') return;

    const matchingCategories = filteredCategories.map(category => category.id);
    setExpandedCategories(previous => {
      const next = new Set(previous);
      matchingCategories.forEach(category => next.add(category));
      return next;
    });

    void loadIconUrls(
      matchingCategories.flatMap(category => matchingIconsByCategory.get(category) || []),
    );
  }, [filteredCategories, loadIconUrls, matchingIconsByCategory, searchTerm]);

  return (
    <div className={`icon-palette ${isCollapsed ? 'collapsed' : ''}`} aria-label={t("Azure Services")}>
      {isCollapsed ? (
        <button
          type="button"
          className="palette-open-toggle"
          onClick={() => setIsCollapsed(false)}
          title={t("Open services panel")}
          aria-label={t("Open services panel")}
          aria-controls="azure-services-palette-content"
          aria-expanded="false"
        >
          <PanelLeftOpen size={18} aria-hidden="true" />
          <span>{t("Open services panel")}</span>
        </button>
      ) : (
        <>
          <div className="palette-header">
            <div className="palette-title-row">
              <div>
                <h2>{t("Azure Services")}</h2>
                <div className="palette-library-meta">
                  {libraryStats.azureVersion} · {libraryStats.officialAzureIcons} Azure ·{' '}
                  {libraryStats.fabricIcons} Fabric
                </div>
              </div>
              <button
                type="button"
                className="palette-close-toggle"
                onClick={() => setIsCollapsed(true)}
                title={t("Close services panel")}
                aria-label={t("Close services panel")}
                aria-controls="azure-services-palette-content"
                aria-expanded="true"
              >
                <PanelLeftClose size={16} aria-hidden="true" />
                <span>{t("Close services panel")}</span>
              </button>
            </div>
            <div className="search-box">
              <Search size={16} aria-hidden="true" />
              <input
                type="search"
                placeholder={localize(language, {
                  en: 'Search name, acronym, or purpose...',
                  ja: '名前・略称・用途で検索...',
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
            <div className="palette-search-summary" role="status" aria-live="polite">
              {searchTerm.trim()
                ? localize(language, {
                    en: `${resultCount} matching icons`,
                    ja: `${resultCount} 件のアイコン`,
                  })
                : localize(language, {
                    en: `${libraryStats.searchableIcons} organized icons`,
                    ja: `${libraryStats.searchableIcons} 件を用途別に分類`,
                  })}
            </div>
            <p className="palette-help">{t('palette.interactionHint')}</p>
          </div>

          <div className="palette-content" id="azure-services-palette-content">
            {filteredCategories.length === 0 && (
              <div className="palette-empty-search">
                {localize(language, {
                  en: 'No icons match every search keyword.',
                  ja: 'すべての検索キーワードに一致するアイコンはありません。',
                })}
              </div>
            )}
            {filteredCategories.map(category => {
              const isExpanded = expandedCategories.has(category.id);
              const icons = matchingIconsByCategory.get(category.id) || [];

              return (
                <section key={category.id} className="category-section">
                  <button
                    type="button"
                    className="category-header"
                    onClick={() => toggleCategory(category.id)}
                    onPointerEnter={() => {
                      if (!isExpanded) void loadIconUrls(icons);
                    }}
                    onFocus={() => {
                      if (!isExpanded) void loadIconUrls(icons);
                    }}
                    aria-expanded={isExpanded}
                    title={localize(language, category.description)}
                  >
                    {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    <span className="category-title">
                      {localize(language, category.label)}
                    </span>
                    <span className="icon-count">{icons.length}</span>
                  </button>
                  {isExpanded && (
                    <>
                      <p className="category-description">
                        {localize(language, category.description)}
                      </p>
                      <div className="icons-grid">
                        {icons.map(icon => {
                          const iconUrl = iconUrls.get(icon.path);
                          return (
                            <button
                              type="button"
                              key={icon.id}
                              className="icon-item"
                              draggable
                              onDragStart={event => onDragStart(event, icon)}
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
                                <div className="icon-placeholder" aria-label={t("Loading...")}>
                                  {t("Loading...")}
                                </div>
                              )}
                              <span className="icon-label">{icon.name}</span>
                            </button>
                          );
                        })}
                      </div>
                    </>
                  )}
                </section>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
};

export default React.memo(IconPalette);
