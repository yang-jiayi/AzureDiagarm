// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { MessageSquare } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import './FeedbackFab.css';

interface FeedbackFabProps {
  onClick: () => void;
  pulse?: boolean;
}

export default function FeedbackFab({ onClick, pulse = false }: FeedbackFabProps) {
  const { t } = useLanguage();

  return (
    <button
      type="button"
      className={`feedback-fab${pulse ? ' pulse-once' : ''}`}
      onClick={onClick}
      title={t('Share feedback')}
    >
      <MessageSquare size={18} aria-hidden="true" />
      {t('Feedback')}
    </button>
  );
}
