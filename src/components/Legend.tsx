// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { Zap, DollarSign, ChevronDown, ChevronUp } from 'lucide-react';
import './Legend.css';
import { useLanguage } from '../i18n/LanguageContext';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';
import { MEDIA_QUERIES } from '../styles/breakpoints';

const LEGEND_COLLAPSED_STORAGE_KEY = 'azure-diagram-builder.legendCollapsed.v1';

interface LegendProps {
  forceCollapsed?: number;
}

const Legend: React.FC<LegendProps> = ({ forceCollapsed }) => {
  const { t } = useLanguage();
  const [isCollapsed, setIsCollapsed] = useState(() => {
    const stored = readLocalStorage(LEGEND_COLLAPSED_STORAGE_KEY);
    return stored === null ? true : stored === '1';
  });

  useEffect(() => {
    if (forceCollapsed) setIsCollapsed(true);
  }, [forceCollapsed]);

  useEffect(() => {
    const mobileViewport = window.matchMedia(MEDIA_QUERIES.compactOrLowHeight);
    const collapseForMobile = (event: MediaQueryListEvent) => {
      if (event.matches) setIsCollapsed(true);
    };
    mobileViewport.addEventListener('change', collapseForMobile);
    return () => mobileViewport.removeEventListener('change', collapseForMobile);
  }, []);

  const toggleLegend = () => {
    setIsCollapsed((current) => {
      const next = !current;
      writeLocalStorage(LEGEND_COLLAPSED_STORAGE_KEY, next ? '1' : '0');
      return next;
    });
  };

  return (
    <section className={`legend ${isCollapsed ? 'collapsed' : ''}`} aria-label={t("LEGEND")}>
      <button
        type="button"
        className="legend-header"
        onClick={toggleLegend}
        aria-expanded={!isCollapsed}
        aria-controls="diagram-legend-content"
        title={isCollapsed ? t('legend.showDetails') : t('legend.hideDetails')}
      >
        <span className="legend-title">{t("LEGEND")}</span>
        <span className="legend-toggle" aria-hidden="true">
          {isCollapsed ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
        </span>
      </button>
      
      {!isCollapsed && (
        <div className="legend-content" id="diagram-legend-content">
          <div className="legend-section">
            <div className="legend-section-title">{t("Connection Types")}</div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="currentColor" strokeWidth="2" />
              </svg>
              <div className="legend-description">
                <strong>{t("Synchronous")}</strong>
                <span>{t("Real-time, request-response (HTTP, SQL)")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="currentColor" strokeWidth="2" strokeDasharray="5, 5" />
              </svg>
              <div className="legend-description">
                <strong>{t("Asynchronous")}</strong>
                <span>{t("Message-based, event-driven (queues, events)")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="currentColor" strokeWidth="2" strokeDasharray="2, 4" opacity="0.65" />
              </svg>
              <div className="legend-description">
                <strong>{t("Optional")}</strong>
                <span>{t("Conditional, fallback paths")}</span>
              </div>
            </div>

            <div className="legend-item">
              <svg width="40" height="16" className="legend-line legend-line-security">
                <line x1="4" y1="8" x2="36" y2="8" stroke="currentColor" strokeWidth="2.5" strokeDasharray="2, 3" />
              </svg>
              <div className="legend-description">
                <strong>{t("Security")}</strong>
                <span>{t("Identity, trust, and policy enforcement")}</span>
              </div>
            </div>

            <div className="legend-item">
              <svg width="40" height="16" className="legend-line legend-line-telemetry">
                <line x1="4" y1="8" x2="36" y2="8" stroke="currentColor" strokeWidth="2" strokeDasharray="7, 3, 2, 3" />
              </svg>
              <div className="legend-description">
                <strong>{t("Telemetry")}</strong>
                <span>{t("Metrics, logs, traces, and diagnostics")}</span>
              </div>
            </div>
          </div>

          <div className="legend-section">
            <div className="legend-section-title">{t("Service Categories")}</div>
            
            <div className="legend-item">
              <div className="legend-color-box category-web"></div>
              <div className="legend-description">
                <strong>{t("Web & Frontend")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-compute"></div>
              <div className="legend-description">
                <strong>{t("Compute & API")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-data"></div>
              <div className="legend-description">
                <strong>{t("Data & Storage")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-ai"></div>
              <div className="legend-description">
                <strong>{t("AI & Analytics")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-iot"></div>
              <div className="legend-description">
                <strong>{t("IoT & Devices")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-security"></div>
              <div className="legend-description">
                <strong>{t("Security & Identity")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-operations"></div>
              <div className="legend-description">
                <strong>{t("Monitoring & Ops")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box category-networking"></div>
              <div className="legend-description">
                <strong>{t("Networking & Integration")}</strong>
              </div>
            </div>
          </div>

          <div className="legend-section">
            <div className="legend-section-title">{t("Pricing Types")}</div>
            
            <div className="legend-item">
              <div className="legend-badge fixed-pricing">
                <DollarSign size={12} />
                {' '}{t("$XX")}{' '}</div>
              <div className="legend-description">
                <strong>{t("Fixed Pricing")}</strong>
                <span>{t("Predictable monthly cost")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-badge usage-pricing">
                <Zap size={12} />
                {' '}{t("~$XX")}{' '}</div>
              <div className="legend-description">
                <strong>{t("Usage-Based")}</strong>
                <span>{t("Varies with consumption")}</span>
              </div>
            </div>
          </div>

          <div className="legend-section">
            <div className="legend-section-title">{t("Cost Levels")}</div>
            
            <div className="legend-item">
              <div className="legend-badge cost-low">
                <DollarSign size={12} />
              </div>
              <div className="legend-description">
                <strong>{t("Free / Low")}</strong>
                <span>{t("Under $100/month")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-badge cost-medium">
                <DollarSign size={12} />
              </div>
              <div className="legend-description">
                <strong>{t("Medium")}</strong>
                <span>{t("$100 - $500/month")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-badge cost-high">
                <DollarSign size={12} />
              </div>
              <div className="legend-description">
                <strong>{t("High")}</strong>
                <span>{t("$500 - $1,000/month")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-badge cost-very-high">
                <DollarSign size={12} />
              </div>
              <div className="legend-description">
                <strong>{t("Very High")}</strong>
                <span>{t("Over $1,000/month")}</span>
              </div>
            </div>
          </div>
        </div>
      )}
    </section>
  );
};

export default Legend;
