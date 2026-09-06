// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import { X, MessageSquare, Send, CheckCircle2 } from 'lucide-react';
import { submitFeedback, buildFeedbackPayload, getFeedbackPolicy, deleteFeedback, FeedbackContext, FeedbackPolicy, FeedbackReceipt } from '../services/feedbackService';
import './FeedbackModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import ModalScaffold from './ModalScaffold';

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

const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// The follow-up contact option is disabled by default. It only renders,
// validates, and submits when explicitly enabled — the backend gates it too.
const FEEDBACK_CONTACT_ENABLED = import.meta.env.VITE_FEEDBACK_CONTACT_ENABLED === 'true';

const FeedbackModal: React.FC<FeedbackModalProps> = ({ isOpen, onClose, context, preselectedRating }) => {
  const { t, translate, language } = useLanguage();
  const [rating, setRating] = useState<number | null>(null);
  const [category, setCategory] = useState(CATEGORIES[0]);
  const [comment, setComment] = useState('');
  const [contactConsent, setContactConsent] = useState(false);
  const [contactEmail, setContactEmail] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [includeMetadata, setIncludeMetadata] = useState(false);
  const [policy, setPolicy] = useState<FeedbackPolicy | null>(null);
  const [receipt, setReceipt] = useState<FeedbackReceipt | null>(null);
  const [deleted, setDeleted] = useState(false);
  const contactAvailable = FEEDBACK_CONTACT_ENABLED && policy?.contactEnabled === true;
  const feedbackInput = {
    rating: rating ?? 0,
    category,
    comment,
    context,
    includeMetadata,
    contact: contactAvailable ? { consent: contactConsent, email: contactEmail } : undefined,
  };

  useEffect(() => {
    if (!isOpen) return;
    setIncludeMetadata(false);
    setPolicy(null);
    const controller = new AbortController();
    getFeedbackPolicy(controller.signal).then(setPolicy).catch(() => { /* Display unavailable policy, not a guessed retention period. */ });
    return () => controller.abort();
  }, [isOpen]);

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
    setContactConsent(false);
    setContactEmail('');
    setIsSubmitting(false);
    setSubmitted(false);
    setError(null);
    setIncludeMetadata(false);
    setReceipt(null);
    setDeleted(false);
  };

  const handleClose = () => {
    if (isSubmitting) return;
    reset();
    onClose();
  };

  const handleSubmit = async () => {
    if (isSubmitting) return;
    if (rating === null) {
      setError(localize(language, {
        en: 'Please pick a rating so we know how you feel.',
        ja: '評価を選択してください（ご感想の把握に使用します）。',
      }));
      return;
    }
    const normalizedEmail = contactEmail.trim();
    if (contactAvailable && contactConsent && (!EMAIL_PATTERN.test(normalizedEmail) || normalizedEmail.length > 254)) {
      setError(localize(language, {
        en: 'Enter a valid email address so we can follow up.',
        ja: 'フォローアップできるよう、有効なメールアドレスを入力してください。',
      }));
      return;
    }
    setError(null);
    setIsSubmitting(true);

    try {
      setReceipt(await submitFeedback(feedbackInput));
      setSubmitted(true);
    } catch (submitError) {
      console.error('[feedback] submit failed:', submitError);
      setError(translate('Feedback could not be sent. Please try again.'));
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleDelete = async () => {
    if (!receipt || isSubmitting) return;
    setError(null);
    setIsSubmitting(true);
    try { await deleteFeedback(receipt.id); setDeleted(true); }
    catch { setError(localize(language, { en: 'Deletion failed. Please retry.', ja: '削除できませんでした。もう一度お試しください。' })); }
    finally { setIsSubmitting(false); }
  };

  if (!isOpen) return null;

  return (
    <ModalScaffold
      isOpen={isOpen}
      onClose={handleClose}
      className="feedback-modal"
      ariaLabel={t("Share Feedback")}
      closeOnBackdrop={!isSubmitting}
      closeOnEscape={!isSubmitting}
    >
      <div className="modal-header">
          <h2>
            <MessageSquare size={24} />
            {' '}{t("Share Feedback")}{' '}</h2>
          <button
            className="modal-close"
            onClick={handleClose}
            title={t("Close")}
            aria-label={t("Close")}
            disabled={isSubmitting}
          >
            <X size={24} />
          </button>
      </div>

        {submitted ? (
          <div className="modal-body feedback-thanks">
            <CheckCircle2 size={48} className="feedback-thanks-icon" />
            <h3>{t("Thank you!")}</h3>
            <p>
              {contactAvailable && contactConsent && receipt?.emailDelivered
                ? translate('Your feedback was saved. The maintainer may contact you about this submission.')
                : t("Your feedback helps us improve the Microsoft Product Architecture Diagram Builder.")}
            </p>
            {receipt && <p className="feedback-receipt">
              {localize(language, { en: 'Feedback ID:', ja: 'フィードバック ID:' })} <code>{receipt.id}</code>
            </p>}
            {receipt?.canDelete && !deleted && <button className="azd-button azd-button--secondary" onClick={handleDelete} disabled={isSubmitting}>
              {localize(language, { en: 'Delete my archived feedback', ja: '保存されたフィードバックを削除' })}
            </button>}
            {deleted && <p role="status">{localize(language, { en: 'Deleted from the application archive. Email copies are not deleted.', ja: 'アプリの保存先から削除しました。メールのコピーは削除されません。' })}</p>}
            {receipt?.emailDelivered && <p>{localize(language, { en: 'A copy was sent by email. Mailbox retention is managed separately.', ja: 'コピーをメールで送信しました。メールの保持期間は別途管理されます。' })}</p>}
            {error && <div className="feedback-error azd-callout azd-callout--danger" role="alert">{error}</div>}
            <button
              className="azd-button azd-button--primary"
              onClick={handleClose}
              disabled={isSubmitting}
            >
              {t("Done")}
            </button>
          </div>
        ) : (
          <>
            <div className="modal-body">
              <p className="feedback-intro">
                {' '}{t("How is your experience so far? Your input shapes what we build next.")}{' '}</p>

              <div className="form-group azd-field">
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

              <div className="form-group azd-field">
                <label htmlFor="feedback-category">{t("Category")}</label>
                <select
                  id="feedback-category"
                  className="feedback-category azd-control"
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  disabled={isSubmitting}
                >
                  {CATEGORIES.map((c) => (
                    <option key={c} value={c}>{translate(c)}</option>
                  ))}
                </select>
              </div>

              <div className="form-group azd-field">
                <label htmlFor="feedback-comment">
                  {' '}{t("Tell us more (optional)")}{' '}<span className="label-hint">{t("What worked well, what was confusing, what you'd love to see")}</span>
                </label>
                <textarea
                  id="feedback-comment"
                  className="feedback-comment azd-control"
                  placeholder={t("e.g., The diagram generation is great, but I'd love to export to Visio...")}
                  value={comment}
                  onChange={(e) => setComment(e.target.value)}
                  rows={4}
                  maxLength={1000}
                  disabled={isSubmitting}
                />
                <div className="character-count azd-character-count">{comment.length}{t("/1000")}</div>
              </div>

              {contactAvailable && (
                <div className="feedback-contact">
                  <div className="feedback-contact-heading">{translate('Open to a follow-up?')}</div>
                  <label className="feedback-contact-consent" htmlFor="feedback-contact-consent">
                    <input
                      id="feedback-contact-consent"
                      type="checkbox"
                      checked={contactConsent}
                      onChange={(event) => {
                        setContactConsent(event.target.checked);
                        setError(null);
                      }}
                      disabled={isSubmitting}
                    />
                    <span>{translate('You may contact me about this feedback')}</span>
                  </label>
                  {contactConsent && (
                    <label className="feedback-contact-email" htmlFor="feedback-contact-email">
                      {translate('Email address')}
                      <input
                        id="feedback-contact-email"
                        className="azd-control"
                        type="email"
                        value={contactEmail}
                        onChange={(event) => {
                          setContactEmail(event.target.value);
                          setError(null);
                        }}
                        autoComplete="email"
                        maxLength={254}
                        required
                        aria-describedby="feedback-contact-email-hint"
                        placeholder="name@company.com"
                        disabled={isSubmitting}
                      />
                      <span id="feedback-contact-email-hint">{translate('Required when follow-up is enabled. Used only to contact you about this feedback.')}</span>
                    </label>
                  )}
                </div>
              )}
              <label className="feedback-metadata-choice">
                <input type="checkbox" checked={includeMetadata} disabled={isSubmitting} onChange={event => setIncludeMetadata(event.target.checked)} />
                {localize(language, {
                  en: 'Include optional diagnostics: service count, model, and site origin only.',
                  ja: '任意の診断情報を含める: サービス数、モデル、サイトのオリジンのみ。',
                })}
              </label>
              <details className="feedback-payload">
                <summary>{localize(language, { en: 'Preview exactly what will be submitted', ja: '送信する内容を確認' })}</summary>
                <pre>{JSON.stringify(buildFeedbackPayload(feedbackInput), null, 2)}</pre>
              </details>
              <p className="feedback-privacy">
                {policy
                  ? localize(language, {
                    en: policy.archiveEnabled
                      ? `New archived feedback expires after ${policy.retentionDays} days (Table Storage cleanup runs every 15 minutes while the server is running). Signed-in owners and administrators can delete archived feedback.`
                      : 'No application archive is configured. Email-only feedback cannot be deleted through this app.',
                    ja: policy.archiveEnabled
                      ? `新しく保存されるフィードバックは ${policy.retentionDays} 日後に期限切れになります（Table Storage はサーバー稼働中に15分間隔で削除）。サインインした所有者と管理者は保存データを削除できます。`
                      : 'アプリの保存先が未構成です。メールのみのフィードバックはこのアプリでは削除できません。',
                  })
                  : localize(language, { en: 'Retention policy is currently unavailable.', ja: '現在、保持ポリシーを取得できません。' })}
                {' '}{localize(language, {
                  en: 'Email and backup retention follow their own policies; app deletion does not delete those copies. Basic rating/category analytics are separate and contain no email, comment or prompt text.',
                  ja: 'メールとバックアップには個別の保持ポリシーが適用され、アプリでの削除ではそのコピーは削除されません。評価・カテゴリの基本分析は別途記録され、メールアドレス、コメントやプロンプト本文は含まれません。',
                })}
              </p>

              {error && (
                <div className="feedback-error azd-callout azd-callout--danger" role="alert">
                  {error}
                </div>
              )}

              <div className="feedback-hint azd-callout">
                {' '}{contactAvailable
                  ? localize(language, {
                      en: `🔒 Follow-up is optional. With your consent, your address is delivered only by email to the maintainer for follow-up within ${policy?.contactRetentionDays ?? 180} days; the archive stores consent metadata, not your address. It is not sent to analytics. Don't include other sensitive information.`,
                      ja: `🔒 フォローアップは任意です。同意した場合、${policy?.contactRetentionDays ?? 180} 日以内の連絡のためアドレスを管理者へメールでのみ送信します。アプリにはアドレスではなく同意情報を保存し、分析には送信しません。その他の機微な情報は入力しないでください。`,
                    })
                  : localize(language, {
                      en: "🔒 We collect your rating and comment to improve the app. Don't include sensitive information.",
                      ja: '🔒 アプリ改善のため評価とコメントを収集します。機微な情報は入力しないでください。',
                    })}{' '}</div>
            </div>

            <div className="modal-actions">
              <button className="azd-button azd-button--secondary" onClick={handleClose} disabled={isSubmitting}>
                {' '}{t("Cancel")}{' '}</button>
              <button className="azd-button azd-button--primary" onClick={handleSubmit} disabled={isSubmitting}>
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
    </ModalScaffold>
  );
};

export default FeedbackModal;
