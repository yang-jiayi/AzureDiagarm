import { useState } from 'react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useModalFocus } from '../hooks/useModalFocus';
import type { DiagramDraft } from '../services/draftStorage';
import type { DraftSaveStatus } from '../hooks/useDraftAutosave';
import './DraftStatus.css';

interface StatusProps {
  status: DraftSaveStatus;
  savedAt?: number;
  error: Error | null;
  onRetry: () => void;
}

export function DraftStatus({ status, savedAt, error, onRetry }: StatusProps) {
  const { language } = useLanguage();
  const label = localize(language, {
    en: status === 'loading' ? 'Opening local draft'
      : status === 'saving' ? 'Saving on this device...'
        : status === 'saved' ? 'Saved on this device'
          : status === 'error' ? 'Draft not saved' : 'Local autosave ready',
    ja: status === 'loading' ? '下書きを読み込み中'
      : status === 'saving' ? 'この端末に保存中...'
        : status === 'saved' ? 'この端末に保存済み'
          : status === 'error' ? '下書きを保存できません' : '端末内の自動保存が有効',
  });
  return (
    <span className={`draft-status draft-status--${status}`}>
      <span role="status" title={error?.message ?? (savedAt ? new Date(savedAt).toLocaleString() : undefined)}>{label}</span>
      {error && (
        <button type="button" onClick={onRetry} title={error.message}>
          {localize(language, { en: 'Retry', ja: '再試行' })}
        </button>
      )}
    </span>
  );
}

interface RecoveryProps {
  draft: DiagramDraft;
  error: Error | null;
  onRestore: () => void;
  onDiscard: () => Promise<void>;
}

export function DraftRecoveryDialog({ draft, error, onRestore, onDiscard }: RecoveryProps) {
  const { language } = useLanguage();
  const [confirmDiscard, setConfirmDiscard] = useState(false);
  const [busy, setBusy] = useState(false);
  const dialogRef = useModalFocus<HTMLElement>(true, () => setConfirmDiscard(false), { closeOnEscape: !busy });
  const download = () => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(draft.document, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = 'recovered-diagram.json';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  return (
    <div className="draft-recovery-overlay">
      <section ref={dialogRef} role="dialog" aria-modal="true" aria-labelledby="draft-recovery-title" tabIndex={-1} className="draft-recovery-dialog">
        <h2 id="draft-recovery-title">{localize(language, { en: 'Resume your saved draft', ja: '保存された下書きから再開' })}</h2>
        <p><strong>{draft.document.titleBlockData.architectureName}</strong></p>
        <p>{new Date(draft.updatedAt).toLocaleString()} · {draft.document.nodes.length} {localize(language, { en: 'shapes', ja: '図形' })}</p>
        <p>{localize(language, {
          en: 'This draft is stored in this browser, not in the cloud. Clearing browser data removes it. Download a file when you need a portable backup.',
          ja: 'この下書きはクラウドではなく、このブラウザーに保存されています。ブラウザーのデータ削除で失われるため、バックアップにはファイルをダウンロードしてください。',
        })}</p>
        {error && <p role="alert">{error.message}</p>}
        <div className="draft-recovery-actions">
          <button type="button" onClick={onRestore} disabled={busy} autoFocus>{localize(language, { en: 'Restore draft', ja: '下書きを復元' })}</button>
          <button type="button" onClick={download}>{localize(language, { en: 'Download a copy', ja: 'コピーをダウンロード' })}</button>
          <button type="button" disabled={busy} onClick={() => setConfirmDiscard(true)}>{localize(language, { en: 'Start without this draft', ja: 'この下書きを使わず新規作成' })}</button>
        </div>
        {confirmDiscard && <div className="draft-discard-confirmation">
          <p>{localize(language, { en: 'Delete this local draft? Download a copy first if you need to keep it.', ja: 'この下書きを削除しますか？残す場合は先にコピーをダウンロードしてください。' })}</p>
          <button type="button" disabled={busy} onClick={async () => {
            setBusy(true);
            try { await onDiscard(); } finally { setBusy(false); }
          }}>{localize(language, { en: 'Delete draft and start new', ja: '下書きを削除して新規作成' })}</button>
          <button type="button" onClick={() => setConfirmDiscard(false)}>{localize(language, { en: 'Keep draft', ja: '下書きを残す' })}</button>
        </div>}
      </section>
    </div>
  );
}
