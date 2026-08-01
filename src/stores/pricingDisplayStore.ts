// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Pricing display preferences
 *
 * UI-only preference controlling whether cost estimates are shown on the
 * canvas. Deliberately separate from the `stylePreset` ('detailed' |
 * 'presentation') switch: presentation mode restyles the whole diagram
 * (edge weights, label opacity) and hiding cost is only a side effect of it.
 * Users asked to drop the cost figure while keeping the detailed styling:
 *
 *   "cost as a fixed value is not acceptable, i would rather hide it or make
 *    it configurable"
 *
 * Estimates are indicative catalog values even when the user adjusts the
 * tier, quantity, or price, so hiding them is a legitimate choice rather than
 * a cosmetic one.
 *
 * Defaults to ON so existing behaviour is unchanged. Persists to localStorage.
 */

import { useState, useEffect } from 'react';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';

export interface PricingDisplayPrefs {
  /** Show cost badges on nodes and the cost summary in the toolbar. */
  showCostBadges: boolean;
}

const STORAGE_KEY = 'azure-diagrams-pricing-display-prefs';

const DEFAULT_PREFS: PricingDisplayPrefs = {
  showCostBadges: true,
};

function loadPrefs(): PricingDisplayPrefs {
  try {
    const stored = readLocalStorage(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      return {
        showCostBadges: typeof parsed.showCostBadges === 'boolean'
          ? parsed.showCostBadges
          : DEFAULT_PREFS.showCostBadges,
      };
    }
  } catch {
    // Invalid preference data falls back to the behavior-safe default.
  }
  return { ...DEFAULT_PREFS };
}

function savePrefs(prefs: PricingDisplayPrefs): void {
  writeLocalStorage(STORAGE_KEY, JSON.stringify(prefs));
}

let currentPrefs: PricingDisplayPrefs = loadPrefs();
const listeners: Set<(prefs: PricingDisplayPrefs) => void> = new Set();

function notifyListeners() {
  listeners.forEach(listener => listener(currentPrefs));
}

/** Non-hook accessor for services / non-React callers. */
export function getPricingDisplayPrefs(): PricingDisplayPrefs {
  return { ...currentPrefs };
}

/** Update one or more display preferences. */
export function updatePricingDisplayPrefs(updates: Partial<PricingDisplayPrefs>): void {
  currentPrefs = { ...currentPrefs, ...updates };
  savePrefs(currentPrefs);
  notifyListeners();
}

/**
 * React hook for pricing display preferences.
 * Provides reactive updates when preferences change in any component.
 */
export function usePricingDisplayPrefs(): [PricingDisplayPrefs, (updates: Partial<PricingDisplayPrefs>) => void] {
  const [prefs, setPrefs] = useState<PricingDisplayPrefs>(currentPrefs);

  useEffect(() => {
    const listener = (newPrefs: PricingDisplayPrefs) => setPrefs(newPrefs);
    listeners.add(listener);
    return () => {
      listeners.delete(listener);
    };
  }, []);

  return [prefs, updatePricingDisplayPrefs];
}
