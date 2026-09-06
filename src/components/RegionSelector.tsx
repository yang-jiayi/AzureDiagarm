// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import { setActiveRegion, getActiveRegion, AVAILABLE_REGIONS, AzureRegion, RegionInfo } from '../services/regionalPricingService';
import { trackRegionChange } from '../services/telemetryService';
import './RegionSelector.css';
import { useLanguage } from '../i18n/LanguageContext';

interface RegionSelectorProps {
  isActive?: boolean;
  onRegionChange?: (region: AzureRegion) => void;
  region?: AzureRegion;
  isUpdating?: boolean;
}

const RegionSelector: React.FC<RegionSelectorProps> = ({
  isActive = true, onRegionChange, region, isUpdating = false,
}) => {
  const { t, translate } = useLanguage();
  const [selectedRegion, setSelectedRegion] = useState<AzureRegion>(getActiveRegion());
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    if (!isActive || isUpdating) setIsOpen(false);
  }, [isActive, isUpdating]);

  const handleRegionSelect = (nextRegion: AzureRegion) => {
    if (!isActive || isUpdating) return;
    if (region === undefined) {
      setSelectedRegion(nextRegion);
      setActiveRegion(nextRegion);
    }
    setIsOpen(false);
    trackRegionChange(nextRegion);
    
    if (onRegionChange) {
      onRegionChange(nextRegion);
    }
  };

  const currentRegion = region ?? selectedRegion;
  const currentRegionInfo = AVAILABLE_REGIONS.find(r => r.id === currentRegion);

  return (
    <div className="region-selector">
      <button 
        type="button"
        className="region-selector-button"
        disabled={!isActive || isUpdating}
        aria-busy={isUpdating}
        aria-expanded={isOpen}
        aria-controls="pricing-region-options"
        onClick={() => setIsOpen(!isOpen)}
        title={t('pricing.regionDescription')}
        aria-label={t('pricing.regionAriaLabel', {
          region: currentRegionInfo?.displayName ?? currentRegion,
        })}
        aria-haspopup="listbox"
      >
        <span className="region-flag">{currentRegionInfo?.flag}</span>
        <span className="region-name">{currentRegionInfo?.displayName}</span>
        <span className="region-arrow">{isOpen ? '▲' : '▼'}</span>
      </button>
      
      {isOpen && (
        <div className="region-dropdown" id="pricing-region-options" role="listbox" aria-label={t('pricing.regionLabel')}>
          <div className="region-dropdown-header">
            <strong>{t('pricing.regionLabel')}</strong>
            <span>{t('pricing.regionDescription')}</span>
          </div>
          {AVAILABLE_REGIONS.map((region: RegionInfo) => (
            <button
              type="button"
              key={region.id}
              className={`region-option ${currentRegion === region.id ? 'selected' : ''}`}
              onClick={() => handleRegionSelect(region.id)}
              role="option"
              aria-selected={currentRegion === region.id}
            >
              <span className="region-flag">{region.flag}</span>
              <div className="region-info">
                <div className="region-display-name">
                  {region.displayName}
                  <span className={`region-type-badge region-type-${region.regionType.toLowerCase()}`}>
                    {translate(region.regionType)}
                  </span>
                </div>
                <div className="region-location">{translate(region.location)}{t(",")}{' '}{translate(region.geography)}</div>
              </div>
              {currentRegion === region.id && <span className="checkmark">{t("✓")}</span>}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

export default RegionSelector;
