// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useState } from 'react';
import { Search, ChevronDown, ChevronRight, PanelLeftClose, PanelLeftOpen } from 'lucide-react';
import { iconCategories, loadIconsFromCategory, AzureIcon, loadIcon } from '../utils/iconLoader';
import './IconPalette.css';
import { useLanguage } from '../i18n/LanguageContext';

interface IconPaletteProps {
  forceCollapsed?: number;
}

function iconMatchesSearch(icon: AzureIcon, term: string): boolean {
  return [icon.name, icon.category, ...icon.searchTerms]
    .some(value => value.toLowerCase().includes(term));
}

const IconPalette: React.FC<IconPaletteProps> = ({ forceCollapsed }) => {
  const { t, translate } = useLanguage();
  const [isCollapsed, setIsCollapsed] = useState(
    () => typeof window !== 'undefined' && window.matchMedia('(max-width: 640px)').matches
  );

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
  const [expandedCategories, setExpandedCategories] = useState<Set<string>>(new Set(['ai + machine learning']));
  const [searchTerm, setSearchTerm] = useState('');
  const [categoryIcons, setCategoryIcons] = useState<Map<string, AzureIcon[]>>(new Map());
  const [iconUrls, setIconUrls] = useState<Map<string, string>>(new Map());

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

  const toggleCategory = (category: string) => {
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

  // Load all icon metadata on mount so search works across all categories
  useEffect(() => {
    let cancelled = false;

    const loadAllIconMetadata = async () => {
      const entries = await Promise.all(
        iconCategories.map(async category => (
          [category, await loadIconsFromCategory(category)] as const
        )),
      );
      if (cancelled) return;

      const newCategoryIcons = new Map<string, AzureIcon[]>(entries);
      setCategoryIcons(newCategoryIcons);

      const initialIcons = newCategoryIcons.get('ai + machine learning') || [];
      void loadIconUrls(initialIcons);
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

  const filteredCategories = useMemo(() => iconCategories.filter(cat => {
    if (searchTerm === '') return true;
    const term = searchTerm.toLowerCase();
    if (cat.toLowerCase().includes(term)) return true;
    const icons = categoryIcons.get(cat) || [];
    return icons.some(icon => iconMatchesSearch(icon, term));
  }), [categoryIcons, searchTerm]);

  useEffect(() => {
    if (searchTerm === '') return;

    const term = searchTerm.toLowerCase();
    const categoriesToExpand: string[] = [];
    const matchingIcons: AzureIcon[] = [];

    iconCategories.forEach(cat => {
      const icons = categoryIcons.get(cat) || [];
      const matches = icons.filter(icon => iconMatchesSearch(icon, term));
      if (matches.length > 0) {
        if (!expandedCategories.has(cat)) {
          categoriesToExpand.push(cat);
        }
        matchingIcons.push(...matches);
      }
    });

    if (categoriesToExpand.length > 0) {
      setExpandedCategories(prev => {
        const next = new Set(prev);
        categoriesToExpand.forEach(c => next.add(c));
        return next;
      });
    }

    void loadIconUrls(matchingIcons);
  }, [categoryIcons, expandedCategories, loadIconUrls, searchTerm]);

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
              <h2>{t("Azure Services")}</h2>
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
              <Search size={16} />
              <input
                type="text"
                placeholder={t("Search services...")}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
          </div>

          <div className="palette-content" id="azure-services-palette-content">
            {filteredCategories.map((category) => {
              const isExpanded = expandedCategories.has(category);
              const icons = categoryIcons.get(category) || [];
              const filteredIcons = icons.filter(icon =>
                searchTerm === '' || iconMatchesSearch(icon, searchTerm.toLowerCase())
              );

              return (
                <div key={category} className="category-section">
                  <button
                    type="button"
                    className="category-header"
                    onClick={() => toggleCategory(category)}
                    onPointerEnter={() => {
                      if (!isExpanded) void loadIconUrls(icons);
                    }}
                    onFocus={() => {
                      if (!isExpanded) void loadIconUrls(icons);
                    }}
                    aria-expanded={isExpanded}
                  >
                    {isExpanded ? <ChevronDown size={16} /> : <ChevronRight size={16} />}
                    <span className="category-title">{translate(category)}</span>
                    {isExpanded && <span className="icon-count">{t("(")}{filteredIcons.length}{t(")")}</span>}
                  </button>
                  {isExpanded && (
                    <div className="icons-grid">
                      {filteredIcons.length > 0 ? (
                        filteredIcons.map((icon) => {
                          const iconUrl = iconUrls.get(icon.path);
                          return (
                            <div
                              key={icon.id}
                              className="icon-item"
                              draggable
                              onDragStart={(e) => onDragStart(e, icon)}
                              title={icon.name}
                            >
                              {iconUrl ? (
                                <img
                                  src={iconUrl}
                                  alt={icon.name}
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
                            </div>
                          );
                        })
                      ) : (
                        <div className="no-icons">{t("Loading icons...")}</div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </>
      )}
    </div>
  );
};

export default React.memo(IconPalette);
