// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Command, Plus, Search, X } from 'lucide-react';
import {
  iconMatchesSearch,
  loadIconsFromPaletteCategory,
  paletteCategories,
  type AzureIcon,
} from '../utils/iconLoader';
import {
  deduplicatePaletteIcons,
  normalizeIconDiscoveryText,
  splitIconSearchHighlight,
} from '../utils/iconDiscovery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import ResponsiveDrawer from './ResponsiveDrawer';
import './CommandPalette.css';

export interface CommandPaletteAction {
  id: string;
  label: string;
  description?: string;
  keywords?: string[];
  group: string;
  shortcut?: string;
  icon?: React.ReactNode;
  disabled?: boolean;
  run: () => void | Promise<void>;
}

interface CommandPaletteProps {
  isOpen: boolean;
  onClose: () => void;
  commands: CommandPaletteAction[];
  onAddService: (icon: AzureIcon) => void;
}

type PaletteResult =
  | { kind: 'command'; key: string; command: CommandPaletteAction }
  | { kind: 'service'; key: string; icon: AzureIcon };

let serviceCatalogPromise: Promise<AzureIcon[]> | null = null;

function loadServiceCatalog(): Promise<AzureIcon[]> {
  if (!serviceCatalogPromise) {
    serviceCatalogPromise = Promise.all(
      paletteCategories.map(category => loadIconsFromPaletteCategory(category.id)),
    ).then(categoryIcons => deduplicatePaletteIcons(categoryIcons.flat()).icons);
  }
  return serviceCatalogPromise;
}

export default function CommandPalette({
  isOpen,
  onClose,
  commands,
  onAddService,
}: CommandPaletteProps) {
  const { language } = useLanguage();
  const [query, setQuery] = useState('');
  const [services, setServices] = useState<AzureIcon[]>([]);
  const [activeIndex, setActiveIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!isOpen) return;
    setQuery('');
    setActiveIndex(0);
    let cancelled = false;
    void loadServiceCatalog().then((catalog) => {
      if (!cancelled) setServices(catalog);
    });
    const focusTimer = window.setTimeout(() => inputRef.current?.focus(), 0);
    return () => {
      cancelled = true;
      window.clearTimeout(focusTimer);
    };
  }, [isOpen]);

  const results = useMemo<PaletteResult[]>(() => {
    const normalizedQuery = normalizeIconDiscoveryText(query);
    const tokens = normalizedQuery.split(' ').filter(Boolean);
    const matchingCommands = commands.filter((command) => {
      if (tokens.length === 0) return true;
      const searchable = normalizeIconDiscoveryText([
        command.label,
        command.description || '',
        ...(command.keywords || []),
      ].join(' '));
      return tokens.every(token => searchable.includes(token));
    }).map(command => ({
      kind: 'command' as const,
      key: `command:${command.id}`,
      command,
    }));
    const matchingServices = normalizedQuery.length < 2
      ? []
      : services
          .filter(icon => iconMatchesSearch(icon, normalizedQuery))
          .slice(0, 18)
          .map(icon => ({
            kind: 'service' as const,
            key: `service:${icon.id}`,
            icon,
          }));
    return [...matchingCommands, ...matchingServices];
  }, [commands, query, services]);

  useEffect(() => {
    setActiveIndex(current => Math.min(current, Math.max(0, results.length - 1)));
  }, [results.length]);

  const executeResult = useCallback((result: PaletteResult | undefined) => {
    if (!result) return;
    if (result.kind === 'command' && result.command.disabled) return;
    onClose();
    window.setTimeout(() => {
      if (result.kind === 'service') onAddService(result.icon);
      else void Promise.resolve(result.command.run());
    }, 0);
  }, [onAddService, onClose]);

  const handleInputKeyDown = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      setActiveIndex(current => results.length === 0 ? 0 : (current + 1) % results.length);
    } else if (event.key === 'ArrowUp') {
      event.preventDefault();
      setActiveIndex(current => results.length === 0
        ? 0
        : (current - 1 + results.length) % results.length);
    } else if (event.key === 'Enter') {
      event.preventDefault();
      executeResult(results[activeIndex]);
    }
  };

  return (
    <ResponsiveDrawer
      isOpen={isOpen}
      modal
      placement="center"
      className="command-palette"
      backdropClassName="command-palette-backdrop"
      ariaLabel={localize(language, {
        en: 'Command palette',
        ja: 'コマンド パレット',
      })}
      onClose={onClose}
      backgroundSelectors={[
        '.app > .app-header',
        '.app > .workspace',
        '.app > .arch-chat-panel',
      ]}
      data-testid="command-palette"
    >
      <div className="command-palette-header">
        <span className="command-palette-title">
          <Command size={19} aria-hidden="true" />
          {localize(language, { en: 'Commands and services', ja: 'コマンドとサービス' })}
        </span>
        <button
          type="button"
          onClick={onClose}
          aria-label={localize(language, { en: 'Close command palette', ja: 'コマンド パレットを閉じる' })}
        >
          <X size={18} aria-hidden="true" />
        </button>
      </div>
      <div className="command-palette-search">
        <Search size={18} aria-hidden="true" />
        <input
          ref={inputRef}
          type="search"
          value={query}
          onChange={event => {
            setQuery(event.target.value);
            setActiveIndex(0);
          }}
          onKeyDown={handleInputKeyDown}
          placeholder={localize(language, {
            en: 'Search commands or Microsoft services...',
            ja: 'コマンドまたは Microsoft サービスを検索...',
          })}
          aria-label={localize(language, {
            en: 'Search commands and services',
            ja: 'コマンドとサービスを検索',
          })}
          role="combobox"
          aria-expanded="true"
          aria-controls="command-palette-results"
          aria-activedescendant={results[activeIndex]?.key}
          autoComplete="off"
        />
        <kbd>Ctrl K</kbd>
      </div>
      <div
        id="command-palette-results"
        className="command-palette-results"
        role="listbox"
        aria-label={localize(language, { en: 'Command results', ja: 'コマンド結果' })}
      >
        {results.length === 0 ? (
          <div className="command-palette-empty">
            {localize(language, {
              en: 'No matching command or service.',
              ja: '一致するコマンドまたはサービスがありません。',
            })}
          </div>
        ) : results.map((result, index) => {
          const selected = index === activeIndex;
          if (result.kind === 'service') {
            return (
              <button
                type="button"
                id={result.key}
                key={result.key}
                role="option"
                aria-selected={selected}
                className={selected ? 'active' : ''}
                onMouseEnter={() => setActiveIndex(index)}
                onClick={() => executeResult(result)}
              >
                <span className="command-palette-icon command-palette-service-icon">
                  <Plus size={17} aria-hidden="true" />
                </span>
                <span className="command-palette-copy">
                  <strong>
                    {splitIconSearchHighlight(result.icon.name, query).map((segment, segmentIndex) => (
                      segment.matched
                        ? <mark key={`${segment.text}-${segmentIndex}`}>{segment.text}</mark>
                        : <React.Fragment key={`${segment.text}-${segmentIndex}`}>{segment.text}</React.Fragment>
                    ))}
                  </strong>
                  <small>
                    {localize(language, {
                      en: `Add Microsoft service · ${result.icon.category}`,
                      ja: `Microsoft サービスを追加 · ${result.icon.category}`,
                    })}
                  </small>
                </span>
                <span className="command-palette-group">
                  {localize(language, { en: 'Service', ja: 'サービス' })}
                </span>
              </button>
            );
          }

          return (
            <button
              type="button"
              id={result.key}
              key={result.key}
              role="option"
              aria-selected={selected}
              aria-disabled={result.command.disabled || undefined}
              disabled={result.command.disabled}
              className={selected ? 'active' : ''}
              onMouseEnter={() => setActiveIndex(index)}
              onClick={() => executeResult(result)}
            >
              <span className="command-palette-icon">
                {result.command.icon || <Command size={17} aria-hidden="true" />}
              </span>
              <span className="command-palette-copy">
                <strong>{result.command.label}</strong>
                {result.command.description && <small>{result.command.description}</small>}
              </span>
              <span className="command-palette-group">{result.command.group}</span>
              {result.command.shortcut && <kbd>{result.command.shortcut}</kbd>}
            </button>
          );
        })}
      </div>
      <div className="command-palette-footer">
        <span>↑↓ {localize(language, { en: 'Navigate', ja: '移動' })}</span>
        <span>Enter {localize(language, { en: 'Run', ja: '実行' })}</span>
        <span>Esc {localize(language, { en: 'Close', ja: '閉じる' })}</span>
      </div>
    </ResponsiveDrawer>
  );
}
