// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { ShieldCheck, X } from 'lucide-react';
import './ValidationHandoffToast.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';

interface ValidationHandoffToastProps {
  isOpen: boolean;
  isModification: boolean;
  isChatOpen: boolean;
  onValidate: () => void;
  onDismiss: () => void;
}

const ValidationHandoffToast: React.FC<ValidationHandoffToastProps> = ({
  isOpen,
  isModification,
  isChatOpen,
  onValidate,
  onDismiss,
}) => {
  const { language } = useLanguage();
  if (!isOpen) return null;

  return (
    <aside
      className={`validation-handoff${isChatOpen ? ' validation-handoff-with-chat' : ''}`}
      aria-label={localize(language, {
        en: 'Validate generated architecture',
        ja: '生成されたアーキテクチャを検証',
      })}
    >
      <button
        type="button"
        className="validation-handoff-close"
        onClick={onDismiss}
        title={localize(language, { en: 'Not now', ja: '後で' })}
        aria-label={localize(language, {
          en: 'Dismiss validation suggestion',
          ja: '検証の提案を閉じる',
        })}
      >
        <X size={16} />
      </button>

      <div className="validation-handoff-icon" aria-hidden="true">
        <ShieldCheck size={22} />
      </div>
      <div className="validation-handoff-content">
        <strong>
          {isModification
            ? localize(language, { en: 'Architecture updated', ja: 'アーキテクチャを更新しました' })
            : localize(language, { en: 'Architecture generated', ja: 'アーキテクチャを生成しました' })}
        </strong>
        <span>
          {localize(language, {
            en: 'Check Well-Architected readiness before you export or share it.',
            ja: 'エクスポートまたは共有する前に、Well-Architected の準備状況を確認します。',
          })}
        </span>
        <div className="validation-handoff-actions">
          <button type="button" className="validation-handoff-primary" onClick={onValidate}>
            <ShieldCheck size={16} />
            {localize(language, { en: 'Validate now', ja: '今すぐ検証' })}
          </button>
          <button type="button" className="validation-handoff-secondary" onClick={onDismiss}>
            {localize(language, { en: 'Not now', ja: '後で' })}
          </button>
        </div>
      </div>
    </aside>
  );
};

export default ValidationHandoffToast;