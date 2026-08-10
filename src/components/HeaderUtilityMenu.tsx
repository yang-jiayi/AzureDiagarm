// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useRef, useState } from 'react';
import { History, Info, Languages, MoreHorizontal, ScanSearch } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import LanguageSwitch from './LanguageSwitch';

interface HeaderUtilityMenuProps {
  onOpenAbout: () => void;
  onOpenRecentWork: () => void;
  onOpenQualityDoctor: () => void;
}

export default function HeaderUtilityMenu({
  onOpenAbout,
  onOpenRecentWork,
  onOpenQualityDoctor,
}: HeaderUtilityMenuProps) {
  const { language } = useLanguage();
  const [isOpen, setIsOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const text = (en: string, ja: string) => localize(language, { en, ja });

  useEffect(() => {
    if (!isOpen) return;
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setIsOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setIsOpen(false);
    };
    document.addEventListener('pointerdown', closeOnOutsideClick);
    document.addEventListener('keydown', closeOnEscape);
    return () => {
      document.removeEventListener('pointerdown', closeOnOutsideClick);
      document.removeEventListener('keydown', closeOnEscape);
    };
  }, [isOpen]);

  return (
    <div className="header-utility-menu" ref={rootRef}>
      <button
        type="button"
        className="header-utility-button"
        onClick={() => setIsOpen(current => !current)}
        aria-haspopup="dialog"
        aria-expanded={isOpen}
        title={text('More application options', 'その他のアプリケーション オプション')}
      >
        <MoreHorizontal size={18} />
        <span>{text('More', 'その他')}</span>
      </button>
      {isOpen && (
        <div
          className="header-utility-menu-popover"
          role="dialog"
          aria-label={text('More application options', 'その他のアプリケーション オプション')}
        >
          <button
            type="button"
            className="header-utility-menu-item"
            onClick={() => {
              setIsOpen(false);
              onOpenRecentWork();
            }}
          >
            <History size={17} aria-hidden="true" />
            <span>{text('Resume recent work', '最近の作業を再開')}</span>
          </button>
          <button
            type="button"
            className="header-utility-menu-item"
            onClick={() => {
              setIsOpen(false);
              onOpenQualityDoctor();
            }}
          >
            <ScanSearch size={17} aria-hidden="true" />
            <span>{text('Diagram Quality Doctor', 'ダイアグラム品質診断')}</span>
          </button>
          <button
            type="button"
            className="header-utility-menu-item"
            onClick={() => {
              setIsOpen(false);
              onOpenAbout();
            }}
          >
            <Info size={17} aria-hidden="true" />
            <span>{text('About this application', 'このアプリについて')}</span>
          </button>
          <div className="header-utility-menu-language">
            <span>
              <Languages size={17} aria-hidden="true" />
              {text('Language', '言語')}
            </span>
            <LanguageSwitch />
          </div>
        </div>
      )}
    </div>
  );
}
