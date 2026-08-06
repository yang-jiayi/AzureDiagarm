// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect } from 'react';
import { X } from 'lucide-react';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import ResponsiveDrawer from './ResponsiveDrawer';
import './ResponsiveRibbonSurface.css';

interface ResponsiveRibbonSurfaceProps {
  isOpen: boolean;
  onClose: () => void;
  children: React.ReactNode;
}

export default function ResponsiveRibbonSurface({
  isOpen,
  onClose,
  children,
}: ResponsiveRibbonSurfaceProps) {
  const { language } = useLanguage();
  const isMobile = useMediaQuery(MEDIA_QUERIES.compact);

  useEffect(() => {
    if (!isMobile && isOpen) onClose();
  }, [isMobile, isOpen, onClose]);

  if (!isMobile) return <>{children}</>;

  const label = localize(language, {
    en: 'Ribbon commands',
    ja: 'リボン コマンド',
  });

  const handleClickCapture = (event: React.MouseEvent<HTMLDivElement>) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const control = target.closest<HTMLElement>('button, label');
    if (!control || control.closest('.mobile-ribbon-sheet-header')) return;
    if (
      control.matches(
        '.ribbon-tab, .toolbar-group-label, [aria-haspopup="menu"], '
        + '[aria-haspopup="listbox"], .model-popover-trigger, .region-selector-button',
      )
    ) {
      return;
    }
    window.setTimeout(onClose, 0);
  };

  return (
    <ResponsiveDrawer
      isOpen={isOpen}
      modal
      placement="bottom"
      className="mobile-ribbon-drawer"
      backdropClassName="mobile-ribbon-backdrop"
      ariaLabel={label}
      onClose={onClose}
      backgroundSelectors={[
        '.app-header .header-brand',
        '.app-header .header-identity-actions',
        '.app-header .mobile-command-bar',
        '.app-header .header-collapse-toggle',
        '.app > .workspace',
      ]}
    >
      <div onClickCapture={handleClickCapture}>
        <div className="mobile-ribbon-sheet-header">
          <strong>{label}</strong>
          <button type="button" onClick={onClose} aria-label={localize(language, {
            en: 'Close ribbon commands',
            ja: 'リボン コマンドを閉じる',
          })}>
            <X size={18} aria-hidden="true" />
          </button>
        </div>
        {children}
      </div>
    </ResponsiveDrawer>
  );
}
