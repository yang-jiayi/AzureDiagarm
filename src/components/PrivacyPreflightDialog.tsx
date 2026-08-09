// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AlertTriangle, CheckCircle2, Eye, ShieldCheck, Sparkles, X } from 'lucide-react';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import type { SensitiveFinding, SensitiveFindingKind } from '../utils/privacyPreflight';
import './PrivacyPreflightDialog.css';

interface PrivacyPreflightDialogProps {
  isOpen: boolean;
  purpose: 'export' | 'share' | 'review';
  findings: SensitiveFinding[];
  canAnonymize: boolean;
  threatOverlayEnabled: boolean;
  onThreatOverlayChange: (enabled: boolean) => void;
  onCancel: () => void;
  onProceed: () => void;
  onAnonymize: () => void;
}

export default function PrivacyPreflightDialog({
  isOpen,
  purpose,
  findings,
  canAnonymize,
  threatOverlayEnabled,
  onThreatOverlayChange,
  onCancel,
  onProceed,
  onAnonymize,
}: PrivacyPreflightDialogProps) {
  const { language } = useLanguage();
  const dialogRef = useModalFocus<HTMLElement>(isOpen);
  useEscapeKey(isOpen, onCancel);
  if (!isOpen) return null;

  const kindLabel = (kind: SensitiveFindingKind) => ({
    credential: localize(language, { en: 'Credential', ja: '資格情報' }),
    'connection-string': localize(language, { en: 'Connection string', ja: '接続文字列' }),
    email: localize(language, { en: 'Email address', ja: 'メール アドレス' }),
    'private-address': localize(language, { en: 'Private address', ja: 'プライベート アドレス' }),
    'resource-id': localize(language, { en: 'Azure resource ID', ja: 'Azure リソース ID' }),
    'internal-host': localize(language, { en: 'Internal host', ja: '内部ホスト' }),
  })[kind];

  return (
    <div className="privacy-preflight-overlay" onClick={onCancel}>
      <section
        ref={dialogRef}
        className="privacy-preflight-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="privacy-preflight-title"
        tabIndex={-1}
        onClick={event => event.stopPropagation()}
      >
        <header>
          <div>
            <span><ShieldCheck size={15} /> {localize(language, { en: 'Privacy preflight', ja: 'プライバシー事前確認' })}</span>
            <h2 id="privacy-preflight-title">
              {findings.length > 0
                ? localize(language, { en: 'Review sensitive information', ja: '機密情報を確認' })
                : localize(language, { en: 'Ready to share', ja: '共有準備完了' })}
            </h2>
          </div>
          <button type="button" onClick={onCancel} aria-label={localize(language, { en: 'Close privacy preflight', ja: 'プライバシー事前確認を閉じる' })}>
            <X size={20} />
          </button>
        </header>

        <div className="privacy-preflight-body">
          <div className={`privacy-preflight-summary${findings.length > 0 ? ' warning' : ' clear'}`}>
            {findings.length > 0 ? <AlertTriangle size={21} /> : <CheckCircle2 size={21} />}
            <div>
              <strong>{findings.length > 0
                ? localize(language, {
                  en: `${findings.length} potential sensitive value${findings.length === 1 ? '' : 's'} found`,
                  ja: `${findings.length} 件の機密情報候補を検出`,
                })
                : localize(language, { en: 'No known sensitive patterns found', ja: '既知の機密パターンは検出されませんでした' })}</strong>
              <span>{localize(language, {
                en: 'Detection is a safety aid, not a guarantee. Review the diagram before external distribution.',
                ja: '検出は安全支援であり保証ではありません。外部配布前に図を確認してください。',
              })}</span>
            </div>
          </div>

          {findings.length > 0 && (
            <ul className="privacy-finding-list">
              {findings.slice(0, 20).map(finding => (
                <li key={finding.id}>
                  <span className={`privacy-finding-severity ${finding.severity}`}>{finding.severity}</span>
                  <div>
                    <strong>{kindLabel(finding.kind)}</strong>
                    <small>{finding.location}</small>
                    <code>{finding.preview}</code>
                  </div>
                </li>
              ))}
            </ul>
          )}

          {findings.length > 0 && !canAnonymize && (
            <p className="privacy-anonymize-note">
              {localize(language, {
                en: 'Open this cloud diagram before creating an anonymized working copy.',
                ja: '匿名化した作業コピーを作成するには、先にこのクラウド図面を開いてください。',
              })}
            </p>
          )}

          <label className="privacy-threat-toggle">
            <input
              type="checkbox"
              checked={threatOverlayEnabled}
              onChange={event => onThreatOverlayChange(event.target.checked)}
            />
            <Eye size={18} />
            <span>
              <strong>{localize(language, { en: 'Show threat-model overlay', ja: '脅威モデル オーバーレイを表示' })}</strong>
              <small>{localize(language, {
                en: 'Highlight review points and controls. This aid is not a formal threat assessment.',
                ja: '確認ポイントと制御を強調します。正式な脅威評価ではありません。',
              })}</small>
            </span>
          </label>
        </div>

        <footer>
          <button type="button" className="btn btn-secondary" onClick={onCancel}>
            {localize(language, { en: 'Cancel', ja: 'キャンセル' })}
          </button>
          {findings.length > 0 && (
            <button
              type="button"
              className="btn btn-secondary privacy-anonymize"
              onClick={onAnonymize}
              disabled={!canAnonymize}
              title={!canAnonymize
                ? localize(language, {
                  en: 'Open this cloud diagram before creating an anonymized working copy.',
                  ja: '匿名化した作業コピーを作成するには、先にこのクラウド図面を開いてください。',
                })
                : undefined}
            >
              <Sparkles size={17} />
              {localize(language, { en: 'Anonymize diagram', ja: '図を匿名化' })}
            </button>
          )}
          <button type="button" className="btn btn-primary" onClick={onProceed}>
            {purpose === 'review'
              ? localize(language, { en: 'Done', ja: '完了' })
              : localize(language, { en: 'Proceed as shown', ja: '表示内容のまま続行' })}
          </button>
        </footer>
      </section>
    </div>
  );
}
