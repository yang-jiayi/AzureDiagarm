// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { Zap, DollarSign } from 'lucide-react';
import './Legend.css';
import { useLanguage } from '../i18n/LanguageContext';

interface LegendProps {
  forceCollapsed?: number;
}

const Legend: React.FC<LegendProps> = ({ forceCollapsed }) => {
  const { t } = useLanguage();
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

  return (
    <div className={`legend ${isCollapsed ? 'collapsed' : ''}`}>
      <div className="legend-header" onClick={() => setIsCollapsed(!isCollapsed)}>
        <span className="legend-title">{t("LEGEND")}</span>
        <span className="legend-toggle">{isCollapsed ? '▲' : '▼'}</span>
      </div>
      
      {!isCollapsed && (
        <div className="legend-content">
          <div className="legend-section">
            <div className="legend-section-title">{t("Connection Types")}</div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="#0078d4" strokeWidth="2" />
              </svg>
              <div className="legend-description">
                <strong>{t("Synchronous")}</strong>
                <span>{t("Real-time, request-response (HTTP, SQL)")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="#0078d4" strokeWidth="2" strokeDasharray="5, 5" />
              </svg>
              <div className="legend-description">
                <strong>{t("Asynchronous")}</strong>
                <span>{t("Message-based, event-driven (queues, events)")}</span>
              </div>
            </div>
            
            <div className="legend-item">
              <svg width="40" height="16" className="legend-line">
                <line x1="4" y1="8" x2="36" y2="8" stroke="#0078d4" strokeWidth="2" strokeDasharray="2, 4" opacity="0.6" />
              </svg>
              <div className="legend-description">
                <strong>{t("Optional")}</strong>
                <span>{t("Conditional, fallback paths")}</span>
              </div>
            </div>
          </div>

          <div className="legend-section">
            <div className="legend-section-title">{t("Service Categories")}</div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#6b7280' }}></div>
              <div className="legend-description">
                <strong>{t("Web & Frontend")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#0078d4' }}></div>
              <div className="legend-description">
                <strong>{t("Compute & API")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#10b981' }}></div>
              <div className="legend-description">
                <strong>{t("Data & Storage")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#f59e0b' }}></div>
              <div className="legend-description">
                <strong>{t("AI & Analytics")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#f97316' }}></div>
              <div className="legend-description">
                <strong>{t("IoT & Devices")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#ef4444' }}></div>
              <div className="legend-description">
                <strong>{t("Security & Identity")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#8b5cf6' }}></div>
              <div className="legend-description">
                <strong>{t("Monitoring & Ops")}</strong>
              </div>
            </div>
            
            <div className="legend-item">
              <div className="legend-color-box" style={{ backgroundColor: '#06b6d4' }}></div>
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
    </div>
  );
};

export default Legend;
