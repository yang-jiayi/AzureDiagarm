// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useMemo, useState } from 'react';
import { createPortal } from 'react-dom';
import { Check, LayoutTemplate, Search, X } from 'lucide-react';
import { ARCHITECTURE_TEMPLATES, type ArchitectureTemplate } from '../data/architectureTemplates';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './TemplateGallery.css';

interface TemplateGalleryProps {
  isOpen: boolean;
  onClose: () => void;
  onApply: (template: ArchitectureTemplate) => void;
}

function TemplatePreview({ template }: { template: ArchitectureTemplate }) {
  const nodes = template.diagram.nodes;
  const positions = new Map(nodes.map(node => [node.id, node.position]));
  const minX = Math.min(...nodes.map(node => node.position.x));
  const maxX = Math.max(...nodes.map(node => node.position.x));
  const minY = Math.min(...nodes.map(node => node.position.y));
  const maxY = Math.max(...nodes.map(node => node.position.y));
  const scaleX = (x: number) => 24 + ((x - minX) / Math.max(1, maxX - minX)) * 272;
  const scaleY = (y: number) => 28 + ((y - minY) / Math.max(1, maxY - minY)) * 104;

  return (
    <svg className="template-preview" viewBox="0 0 320 160" role="img" aria-hidden="true">
      {template.diagram.edges.map(item => {
        const source = positions.get(item.source);
        const target = positions.get(item.target);
        if (!source || !target) return null;
        return (
          <line
            key={item.id}
            x1={scaleX(source.x)}
            y1={scaleY(source.y)}
            x2={scaleX(target.x)}
            y2={scaleY(target.y)}
          />
        );
      })}
      {nodes.map(item => (
        <g key={item.id} transform={`translate(${scaleX(item.position.x)}, ${scaleY(item.position.y)})`}>
          <rect x="-16" y="-12" width="32" height="24" rx="6" style={{ fill: template.accent }} />
          <circle cx="10" cy="-8" r="4" />
        </g>
      ))}
    </svg>
  );
}

export default function TemplateGallery({ isOpen, onClose, onApply }: TemplateGalleryProps) {
  const { language } = useLanguage();
  const [query, setQuery] = useState('');
  const [selectedId, setSelectedId] = useState(ARCHITECTURE_TEMPLATES[0].id);
  const dialogRef = useModalFocus<HTMLElement>(isOpen);
  useEscapeKey(isOpen, onClose);

  const filteredTemplates = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase();
    if (!normalized) return ARCHITECTURE_TEMPLATES;
    return ARCHITECTURE_TEMPLATES.filter(template => (
      [
        localize(language, template.name),
        localize(language, template.description),
        ...template.tags.map(tag => localize(language, tag)),
      ].join(' ').toLocaleLowerCase().includes(normalized)
    ));
  }, [language, query]);
  const selected = filteredTemplates.find(template => template.id === selectedId)
    ?? filteredTemplates[0]
    ?? null;

  if (!isOpen) return null;

  return createPortal(
    <div className="template-gallery-overlay" onClick={onClose}>
      <section
        ref={dialogRef}
        className="template-gallery-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="template-gallery-title"
        tabIndex={-1}
        onClick={event => event.stopPropagation()}
      >
        <header className="template-gallery-header">
          <div>
            <span>{localize(language, { en: 'Starter architectures', ja: 'スターター アーキテクチャ' })}</span>
            <h2 id="template-gallery-title">
              {localize(language, { en: 'Choose a template', ja: 'テンプレートを選択' })}
            </h2>
          </div>
          <button type="button" onClick={onClose} aria-label={localize(language, { en: 'Close template gallery', ja: 'テンプレート ギャラリーを閉じる' })}>
            <X size={20} />
          </button>
        </header>

        <label className="template-gallery-search">
          <Search size={17} aria-hidden="true" />
          <span className="sr-only">{localize(language, { en: 'Search templates', ja: 'テンプレートを検索' })}</span>
          <input
            value={query}
            onChange={event => setQuery(event.target.value)}
            placeholder={localize(language, { en: 'Search by workload or pattern', ja: 'ワークロードまたはパターンで検索' })}
          />
        </label>

        <div className="template-gallery-layout">
          <div className="template-gallery-list" role="listbox" aria-label={localize(language, { en: 'Architecture templates', ja: 'アーキテクチャ テンプレート' })}>
            {filteredTemplates.map(template => {
              const isSelected = template.id === selected?.id;
              return (
                <button
                  key={template.id}
                  type="button"
                  className={`template-card${isSelected ? ' selected' : ''}`}
                  role="option"
                  aria-selected={isSelected}
                  onClick={() => setSelectedId(template.id)}
                >
                  <TemplatePreview template={template} />
                  <span className="template-card-title">
                    <LayoutTemplate size={17} aria-hidden="true" />
                    {localize(language, template.name)}
                    {isSelected && <Check size={16} aria-hidden="true" />}
                  </span>
                  <span>{localize(language, template.description)}</span>
                </button>
              );
            })}
            {filteredTemplates.length === 0 && (
              <p className="template-gallery-empty">
                {localize(language, { en: 'No templates match that search.', ja: '一致するテンプレートがありません。' })}
              </p>
            )}
          </div>

          <aside className="template-gallery-detail">
            {selected ? (
              <>
                <TemplatePreview template={selected} />
                <h3>{localize(language, selected.name)}</h3>
                <p>{localize(language, selected.description)}</p>
                <div className="template-gallery-tags">
                  {selected.tags.map(tag => (
                    <span key={localize(language, tag)}>{localize(language, tag)}</span>
                  ))}
                </div>
                <dl>
                  <div>
                    <dt>{localize(language, { en: 'Services', ja: 'サービス' })}</dt>
                    <dd>{selected.diagram.nodes.length}</dd>
                  </div>
                  <div>
                    <dt>{localize(language, { en: 'Connections', ja: '接続' })}</dt>
                    <dd>{selected.diagram.edges.length}</dd>
                  </div>
                </dl>
              </>
            ) : (
              <p className="template-gallery-empty">
                {localize(language, {
                  en: 'Clear the search to preview a starter architecture.',
                  ja: '検索をクリアするとスターター アーキテクチャをプレビューできます。',
                })}
              </p>
            )}
          </aside>
        </div>

        <footer className="template-gallery-footer">
          <button type="button" className="btn btn-secondary" onClick={onClose}>
            {localize(language, { en: 'Cancel', ja: 'キャンセル' })}
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => {
              if (selected) onApply(selected);
            }}
            disabled={!selected}
          >
            <LayoutTemplate size={17} />
            {localize(language, { en: 'Use this template', ja: 'このテンプレートを使用' })}
          </button>
        </footer>
      </section>
    </div>,
    document.body,
  );
}
