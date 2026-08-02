// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { X, MessageSquare, Send, CheckCircle2 } from 'lucide-react';
import { submitFeedback, FeedbackContext } from '../services/feedbackService';
import './FeedbackModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';

interface FeedbackModalProps {
  isOpen: boolean;
  onClose: () => void;
  context?: FeedbackContext;
  /** Pre-selected rating carried over from the quick toast, if any. */
  preselectedRating?: number;
}

const RATINGS: { value: number; emoji: string; label: string }[] = [
  { value: 1, emoji: '😞', label: 'Very unhappy' },
  { value: 2, emoji: '🙁', label: 'Unhappy' },
  { value: 3, emoji: '😐', label: 'Neutral' },
  { value: 4, emoji: '🙂', label: 'Happy' },
  { value: 5, emoji: '🤩', label: 'Love it' },
];

const CATEGORIES = [
  'General',
  'Bug / something broke',
  'Feature request',
  'Diagram quality',
  'Performance',
  'Other',
];

const FeedbackModal: React.FC<FeedbackModalProps> = ({ isOpen, onClose, context, preselectedRating }) => {
  const { t, translate } = useLanguage();
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);
  const [rating, setRating] = useState<number | null>(null);
  const [category, setCategory] = useState(CATEGORIES[0]);
  const [comment, setComment] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // When the modal is opened from the quick toast, seed the rating the user
  // already gave so they don't have to pick it twice.
  useEffect(() => {
    if (isOpen && preselectedRating != null) {
      setRating(preselectedRating);
    }
  }, [isOpen, preselectedRating]);

  const reset = () => {
    setRating(null);
    setCategory(CATEGORIES[0]);
    setComment('');
    setIsSubmitting(false);
    setSubmitted(false);
    setError(null);
  };

  const handleClose = () => {
    reset();
    onClose();
  };
  useEscapeKey(isOpen, handleClose);

  const handleSubmit = async () => {
    if (rating === null) {
      setError(translate('Please pick a rating so we know how you feel.'));
      return;
    }
    setError(null);
    setIsSubmitting(true);

    try {
      await submitFeedback({ rating, category, comment, context });
      setSubmitted(true);
    } catch (submitError) {
      console.error('[feedback] submit failed:', submitError);
      setError(translate('Feedback could not be sent. Please try again.'));
    } finally {
      setIsSubmitting(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={handleClose}>
      <div
        ref={dialogRef}
        className="modal-content feedback-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Share Feedback")}
        tabIndex={-1}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <h2>
            <MessageSquare size={24} />
            {' '}{t("Share Feedback")}{' '}</h2>
          <button className="modal-close" onClick={handleClose} title={t("Close")} aria-label={t("Close")}>
            <X size={24} />
          </button>
        </div>

        {submitted ? (
          <div className="modal-body feedback-thanks">
            <CheckCircle2 size={48} className="feedback-thanks-icon" />
            <h3>{t("Thank you!")}</h3>
            <p>{t("Your feedback helps us improve the Azure Architecture Diagram Builder.")}</p>
            <button className="btn-primary" onClick={handleClose}>{t("Done")}</button>
          </div>
        ) : (
          <>
            <div className="modal-body">
              <p className="feedback-intro">
                {' '}{t("How is your experience so far? Your input shapes what we build next.")}{' '}</p>

              <div className="form-group">
                <label>{t("How do you feel about the app?")}</label>
                <div className="feedback-ratings" role="radiogroup" aria-label={t("Rating")}>
                  {RATINGS.map((r) => (
                    <button
                      key={r.value}
                      type="button"
                      role="radio"
                      aria-checked={rating === r.value}
                      aria-label={translate(r.label)}
                      title={translate(r.label)}
                      className={`feedback-rating ${rating === r.value ? 'selected' : ''}`}
                      onClick={() => { setRating(r.value); setError(null); }}
                      disabled={isSubmitting}
                    >
                      <span className="feedback-rating-emoji">{r.emoji}</span>
                    </button>
                  ))}
                </div>
              </div>

              <div className="form-group">
                <label htmlFor="feedback-category">{t("Category")}</label>
                <select
                  id="feedback-category"
                  className="feedback-category"
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  disabled={isSubmitting}
                >
                  {CATEGORIES.map((c) => (
                    <option key={c} value={c}>{translate(c)}</option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label htmlFor="feedback-comment">
                  {' '}{t("Tell us more (optional)")}{' '}<span className="label-hint">{t("What worked well, what was confusing, what you'd love to see")}</span>
                </label>
                <textarea
                  id="feedback-comment"
                  className="feedback-comment"
                  placeholder={t("e.g., The diagram generation is great, but I'd love to export to Visio...")}
                  value={comment}
                  onChange={(e) => setComment(e.target.value)}
                  rows={4}
                  maxLength={1000}
                  disabled={isSubmitting}
                />
                <div className="character-count">{comment.length}{t("/1000")}</div>
              </div>

              {error && <div className="feedback-error">{error}</div>}

              <div className="feedback-hint">
                {' '}{t("🔒 We collect your rating and comment to improve the app. Don't include sensitive information.")}{' '}</div>
            </div>

            <div className="modal-actions">
              <button className="btn-secondary" onClick={handleClose} disabled={isSubmitting}>
                {' '}{t("Cancel")}{' '}</button>
              <button className="btn-primary" onClick={handleSubmit} disabled={isSubmitting}>
                {isSubmitting ? (
                  <>
                    <div className="spinner-small"></div>
                    {' '}{t("Sending...")}{' '}</>
                ) : (
                  <>
                    <Send size={18} />
                    {' '}{t("Send Feedback")}{' '}</>
                )}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

export default FeedbackModal;
