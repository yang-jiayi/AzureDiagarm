// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState } from 'react';
import { X, MessageSquarePlus, CheckCircle2 } from 'lucide-react';
import { submitFeedback, FeedbackContext } from '../services/feedbackService';
import './FeedbackToast.css';
import { useLanguage } from '../i18n/LanguageContext';

interface FeedbackToastProps {
  isOpen: boolean;
  onClose: () => void;
  /** Open the full modal (for a comment), carrying the rating the user just gave. */
  onAddComment: (rating: number) => void;
  context?: FeedbackContext;
}

const RATINGS: { value: number; emoji: string; label: string }[] = [
  { value: 1, emoji: '😞', label: 'Very unhappy' },
  { value: 2, emoji: '🙁', label: 'Unhappy' },
  { value: 3, emoji: '😐', label: 'Neutral' },
  { value: 4, emoji: '🙂', label: 'Happy' },
  { value: 5, emoji: '🤩', label: 'Love it' },
];

const FeedbackToast: React.FC<FeedbackToastProps> = ({ isOpen, onClose, onAddComment, context }) => {
  const { t, translate } = useLanguage();
  const [rated, setRated] = useState<number | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleRate = async (rating: number) => {
    setError(null);
    setIsSubmitting(true);
    try {
      await submitFeedback({ rating, category: 'Quick rating', comment: '', context });
      setRated(rating);
      window.setTimeout(() => onClose(), 4500);
    } catch (submitError) {
      console.error('[feedback] quick rating failed:', submitError);
      setError(translate('Feedback could not be sent. Please try again.'));
    } finally {
      setIsSubmitting(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className={`feedback-toast ${rated ? 'rated' : ''}`} role="dialog" aria-label={t("Quick feedback")}>
      <button className="feedback-toast-close" onClick={onClose} title={t("Dismiss")} aria-label={t("Dismiss")}>
        <X size={16} />
      </button>

      {rated ? (
        <div className="feedback-toast-thanks">
          <CheckCircle2 size={18} className="feedback-toast-thanks-icon" />
          <span>{t("Thanks! Want to add a quick note?")}</span>
          <button className="feedback-toast-link" onClick={() => onAddComment(rated)}>
            <MessageSquarePlus size={15} />
            {' '}{t("Add a comment")}{' '}</button>
        </div>
      ) : (
        <>
          <div className="feedback-toast-prompt">{t("How did that diagram turn out?")}</div>
          <div className="feedback-toast-ratings" role="radiogroup" aria-label={t("Rating")}>
            {RATINGS.map((r) => (
              <button
                key={r.value}
                type="button"
                role="radio"
                aria-checked="false"
                aria-label={translate(r.label)}
                title={translate(r.label)}
                className="feedback-toast-rating"
                onClick={() => void handleRate(r.value)}
                disabled={isSubmitting}
              >
                {r.emoji}
              </button>
            ))}
          </div>
          {error && <div className="feedback-toast-error">{error}</div>}
        </>
      )}
    </div>
  );
};

export default FeedbackToast;
