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
}

const RegionSelector: React.FC<RegionSelectorProps> = ({ isActive = true, onRegionChange }) => {
  const { t, translate } = useLanguage();
  const [selectedRegion, setSelectedRegion] = useState<AzureRegion>(getActiveRegion());
  const [isOpen, setIsOpen] = useState(false);

  useEffect(() => {
    if (!isActive) setIsOpen(false);
  }, [isActive]);

  const handleRegionSelect = (region: AzureRegion) => {
    setSelectedRegion(region);
    setActiveRegion(region);
    setIsOpen(false);
    trackRegionChange(region);
    
    if (onRegionChange) {
      onRegionChange(region);
    }
  };

  const currentRegionInfo = AVAILABLE_REGIONS.find(r => r.id === selectedRegion);

  return (
    <div className="region-selector">
      <button 
        className="region-selector-button"
        onClick={() => setIsOpen(!isOpen)}
        title={t("Select pricing region")}
      >
        <span className="region-flag">{currentRegionInfo?.flag}</span>
        <span className="region-name">{currentRegionInfo?.displayName}</span>
        <span className="region-arrow">{isOpen ? '▲' : '▼'}</span>
      </button>
      
      {isOpen && (
        <div className="region-dropdown">
          {AVAILABLE_REGIONS.map((region: RegionInfo) => (
            <button
              key={region.id}
              className={`region-option ${selectedRegion === region.id ? 'selected' : ''}`}
              onClick={() => handleRegionSelect(region.id)}
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
              {selectedRegion === region.id && <span className="checkmark">{t("✓")}</span>}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

export default RegionSelector;
