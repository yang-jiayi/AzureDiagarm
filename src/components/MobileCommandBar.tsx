// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Menu, MessagesSquare, PanelLeftOpen, Search, Maximize2, Minimize2 } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './MobileCommandBar.css';

interface MobileCommandBarProps {
  activeSection: string;
  commandSheetOpen: boolean;
  focusMode: boolean;
  chatOpen: boolean;
  onOpenCommands: () => void;
  onOpenCommandPalette: () => void;
  onOpenServices: () => void;
  onToggleChat: () => void;
  onToggleFocus: () => void;
}

export default function MobileCommandBar({
  activeSection,
  commandSheetOpen,
  focusMode,
  chatOpen,
  onOpenCommands,
  onOpenCommandPalette,
  onOpenServices,
  onToggleChat,
  onToggleFocus,
}: MobileCommandBarProps) {
  const { language } = useLanguage();

  return (
    <nav
      className="mobile-command-bar"
      aria-label={localize(language, {
        en: 'Mobile command bar',
        ja: 'モバイル コマンドバー',
      })}
    >
      <button
        type="button"
        onClick={onOpenCommands}
        aria-expanded={commandSheetOpen}
        aria-controls="application-toolbar"
      >
        <Menu size={17} aria-hidden="true" />
        <span>{activeSection}</span>
      </button>
      <button
        type="button"
        onClick={onOpenCommandPalette}
        aria-keyshortcuts="Control+K Meta+K"
      >
        <Search size={17} aria-hidden="true" />
        <span>{localize(language, { en: 'Search', ja: '検索' })}</span>
      </button>
      <button type="button" onClick={onOpenServices}>
        <PanelLeftOpen size={17} aria-hidden="true" />
        <span>{localize(language, { en: 'Services', ja: 'サービス' })}</span>
      </button>
      <button
        type="button"
        onClick={onToggleChat}
        aria-pressed={chatOpen}
      >
        <MessagesSquare size={17} aria-hidden="true" />
        <span>{localize(language, { en: 'Chat', ja: 'Chat' })}</span>
      </button>
      <button
        type="button"
        onClick={onToggleFocus}
        aria-pressed={focusMode}
      >
        {focusMode
          ? <Maximize2 size={17} aria-hidden="true" />
          : <Minimize2 size={17} aria-hidden="true" />}
        <span>{localize(language, {
          en: focusMode ? 'Exit' : 'Focus',
          ja: focusMode ? '解除' : '集中',
        })}</span>
      </button>
    </nav>
  );
}
