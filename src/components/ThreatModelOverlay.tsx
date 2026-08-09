// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useRef } from 'react';
import { Eye, ShieldAlert, X } from 'lucide-react';
import type { ThreatKind, ThreatLevel, ThreatMarker } from '../utils/threatModel';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './ThreatModelOverlay.css';

interface ThreatModelOverlayProps {
  enabled: boolean;
  markers: ThreatMarker[];
  onClose: () => void;
}

const COLORS: Record<ThreatLevel, string> = {
  high: '#dc2626',
  medium: '#d97706',
  low: '#2563eb',
  control: '#047857',
};

const GLOWS: Record<ThreatLevel, string> = {
  high: '0 0 0 7px rgba(220, 38, 38, 0.18)',
  medium: '0 0 0 7px rgba(217, 119, 6, 0.18)',
  low: '0 0 0 7px rgba(37, 99, 235, 0.18)',
  control: '0 0 0 7px rgba(4, 120, 87, 0.18)',
};

export default function ThreatModelOverlay({ enabled, markers, onClose }: ThreatModelOverlayProps) {
  const { language } = useLanguage();
  const overlayRef = useRef<HTMLElement>(null);

  useEffect(() => {
    if (!enabled) return;
    const canvas = overlayRef.current?.parentElement;
    if (!canvas) return;

    const markerById = new Map(markers.map(marker => [marker.nodeId, marker]));
    const restorers: Array<() => void> = [];
    canvas.querySelectorAll<HTMLElement>('.react-flow__node[data-id]').forEach((node) => {
      const marker = markerById.get(node.dataset.id || '');
      if (!marker) return;
      const previous = ['outline', 'outline-offset', 'box-shadow'].map(property => ({
        property,
        value: node.style.getPropertyValue(property),
        priority: node.style.getPropertyPriority(property),
      }));
      node.style.setProperty('outline', `3px solid ${COLORS[marker.level]}`, 'important');
      node.style.setProperty('outline-offset', '4px');
      node.style.setProperty('box-shadow', GLOWS[marker.level], 'important');
      restorers.push(() => {
        previous.forEach(({ property, value, priority }) => {
          if (value) node.style.setProperty(property, value, priority);
          else node.style.removeProperty(property);
        });
      });
    });

    return () => restorers.forEach(restore => restore());
  }, [enabled, markers]);

  if (!enabled) return null;

  const kindLabel = (kind: ThreatKind) => ({
    internet: localize(language, { en: 'Internet exposure', ja: 'インターネット公開' }),
    data: localize(language, { en: 'Sensitive data', ja: '機密データ' }),
    identity: localize(language, { en: 'Identity control', ja: 'ID 制御' }),
    secrets: localize(language, { en: 'Secrets control', ja: 'シークレット制御' }),
    observability: localize(language, { en: 'Detection control', ja: '検出制御' }),
  })[kind];

  return (
    <aside
      ref={overlayRef}
      className="threat-model-overlay"
      aria-label={localize(language, { en: 'Threat model overlay', ja: '脅威モデル オーバーレイ' })}
    >
        <header>
          <div>
            <span><Eye size={14} /> {localize(language, { en: 'Overlay active', ja: 'オーバーレイ有効' })}</span>
            <strong><ShieldAlert size={17} /> {localize(language, { en: 'Threat model', ja: '脅威モデル' })}</strong>
          </div>
          <button type="button" onClick={onClose} aria-label={localize(language, { en: 'Hide threat model overlay', ja: '脅威モデル オーバーレイを非表示' })}>
            <X size={16} />
          </button>
        </header>
        <p>
          {localize(language, {
            en: `${markers.length} architecture controls or review points highlighted.`,
            ja: `${markers.length} 件のアーキテクチャ制御または確認ポイントを強調表示しています。`,
          })}
        </p>
        <ul>
          {markers.slice(0, 6).map(marker => (
            <li key={`${marker.nodeId}-${marker.kind}`}>
              <span style={{ background: COLORS[marker.level] }} aria-hidden="true" />
              <div>
                <strong>{marker.serviceName}</strong>
                <small>{kindLabel(marker.kind)}</small>
              </div>
            </li>
          ))}
        </ul>
    </aside>
  );
}
