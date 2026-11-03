// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AlertTriangle, CheckCircle2, Eye, ShieldCheck, Sparkles, X } from 'lucide-react';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import type { RecentWorkSyncState } from '../services/recentWorkService';
import type { SensitiveFinding, SensitiveFindingKind } from '../utils/privacyPreflight';
import './PrivacyPreflightDialog.css';

interface AIConnectionDisclosure {
  source: 'managed' | 'bring-your-own';
  model: string;
  provider?: 'azure-openai' | 'openai';
  endpoint?: string;
}

interface PrivacyPreflightDialogProps {
  isOpen: boolean;
  purpose: 'export' | 'share' | 'review';
  findings: SensitiveFinding[];
  canAnonymize: boolean;
  cloudStorageState: RecentWorkSyncState;
  aiConnection: AIConnectionDisclosure;
  shareRole?: 'viewer' | 'editor';
  shareDocumentName?: string;
  shareVersionCount?: number;
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
  cloudStorageState,
  aiConnection,
  shareRole,
  shareDocumentName,
  shareVersionCount,
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
  const title = purpose === 'review'
    ? localize(language, { en: 'Privacy review', ja: 'プライバシー確認' })
    : findings.length > 0
      ? localize(language, { en: 'Review sensitive information', ja: '機密情報を確認' })
      : localize(language, { en: 'Review data handling', ja: 'データの取り扱いを確認' });
  const storageLabel = purpose === 'share' && shareDocumentName
    ? localize(language, {
        en: `Cloud share target: ${shareDocumentName}. The editor’s current local/cloud save state is separate from this share target.`,
        ja: `クラウド共有対象: ${shareDocumentName}。エディターで現在開いている図面のローカル/クラウド保存状態は、この共有対象とは別です。`,
      })
    : {
        local: localize(language, { en: 'Local browser draft (IndexedDB); no cloud document is linked.', ja: 'ブラウザー内のローカル下書き（IndexedDB）。クラウド図面は接続されていません。' }),
        saving: localize(language, { en: 'Cloud sync is saving this document.', ja: 'クラウド同期で図面を保存中です。' }),
        synced: localize(language, { en: 'Cloud document is synced through this application’s owner-scoped diagram API. Production uses private Azure Blob storage; local development may use in-memory storage.', ja: 'このアプリの所有者単位の図面 API にクラウド図面を同期済みです。本番環境はプライベート Azure Blob Storage を使用し、ローカル開発ではメモリ内ストレージの場合があります。' }),
        readonly: localize(language, { en: 'Cloud document is open read-only.', ja: 'クラウド図面は読み取り専用で開かれています。' }),
        offline: localize(language, { en: 'Cloud document is offline; recent local edits may not be synced.', ja: 'クラウド図面はオフラインです。最近のローカル編集は未同期の可能性があります。' }),
        unavailable: localize(language, { en: 'Cloud storage is unavailable; the current local state may not be synced.', ja: 'クラウドストレージを利用できません。現在のローカル状態は未同期の可能性があります。' }),
        conflict: localize(language, { en: 'Cloud sync has a conflict; local changes are not confirmed as saved.', ja: 'クラウド同期で競合が発生しています。ローカル変更の保存は確認されていません。' }),
        error: localize(language, { en: 'Cloud sync failed; local changes are not confirmed as saved.', ja: 'クラウド同期に失敗しました。ローカル変更の保存は確認されていません。' }),
    }[cloudStorageState];
  const actionDisclosure = purpose === 'review'
    ? localize(language, {
      en: 'This sensitive-data scan runs in this browser. Opening this review does not send the diagram to an AI provider; any cloud autosave shown below is a separate operation.',
      ja: '機密情報の検出はこのブラウザー内で実行されます。この確認を開いても図面は AI プロバイダーへ送信されません。下記のクラウド自動保存は別の処理です。',
    })
    : purpose === 'export'
      ? localize(language, {
        en: 'The selected export creates a file in the browser’s download location on this device. The format determines which diagram details are included; this action does not upload the file.',
        ja: '選択した形式のファイルを、このデバイス上のブラウザーのダウンロード先に作成します。含まれる図面情報は形式によって異なります。この操作でファイルはアップロードされません。',
      })
      : localize(language, {
        en: `The link grants ${shareRole === 'editor' ? 'editor' : 'viewer'} access to cloud diagram “${shareDocumentName || 'the selected diagram'}” and its comments${shareVersionCount === undefined ? '' : `, plus ${shareVersionCount} history version${shareVersionCount === 1 ? '' : 's'}`}. Recipients still need Microsoft Entra ID assignment to AzureDiagarm and the link. The application does not configure link expiry; the owner can revoke it. Creating the link does not email it to recipients.`,
        ja: `共有リンクではクラウド図面「${shareDocumentName || '選択した図面'}」とそのコメント${shareVersionCount === undefined ? '' : `、履歴 ${shareVersionCount} 件`}に${shareRole === 'editor' ? '編集者' : '閲覧者'}としてアクセスできます。受信者には Microsoft Entra ID で AzureDiagarm への割り当てとリンクの両方が必要です。アプリケーションではリンクの有効期限を設定していません。所有者はリンクを取り消せます。作成時にリンクをメール送信することはありません。`,
      });
  const aiDestination = aiConnection.source === 'managed'
    ? localize(language, {
      en: `Managed ${aiConnection.model} through this application’s /api/openai proxy to its configured Azure OpenAI deployment; the exact endpoint hostname is not exposed to this browser.`,
      ja: `このアプリの /api/openai プロキシを経由して設定済み Azure OpenAI デプロイへ送信するマネージド ${aiConnection.model}。正確なエンドポイントのホスト名はこのブラウザーには公開されていません。`,
    })
    : localize(language, {
      en: `BYO ${aiConnection.provider === 'openai' ? 'OpenAI' : 'Azure OpenAI'}${aiConnection.endpoint ? ` at ${aiConnection.endpoint}` : ''} through this application’s /api/openai proxy (model ${aiConnection.model || 'not configured'}).`,
      ja: `このアプリの /api/openai プロキシ経由で送信する BYO ${aiConnection.provider === 'openai' ? 'OpenAI' : 'Azure OpenAI'}${aiConnection.endpoint ? `（${aiConnection.endpoint}）` : ''}（モデル: ${aiConnection.model || '未設定'}）。`,
    });

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
              {title}
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

          <section className="privacy-data-policy" aria-labelledby="privacy-data-policy-title">
            <h3 id="privacy-data-policy-title">
              {localize(language, { en: 'Data handling', ja: 'データの取り扱い' })}
            </h3>
            <dl>
              <div>
                <dt>{localize(language, { en: 'This action', ja: '今回の操作' })}</dt>
                <dd>{actionDisclosure}</dd>
              </div>
              <div>
                <dt>{localize(language, {
                  en: purpose === 'share' ? 'Share target' : 'Current storage',
                  ja: purpose === 'share' ? '共有対象' : '現在の保存先',
                })}</dt>
                <dd>{storageLabel}</dd>
              </div>
              <div>
                <dt>{localize(language, { en: 'AI destination when you submit a request', ja: 'AI リクエスト送信時の送信先' })}</dt>
                <dd>{aiDestination}</dd>
                <dd>{localize(language, {
                  en: 'Only an explicit AI submission sends its prompt, relevant diagram context, and supported images. The app sets store:false on the API request; provider retention and logging policies are separate and are not guaranteed by this app. A BYO credential is sent with the request to this app’s server, but is not stored as a job input.',
                  ja: '明示的に AI リクエストを送信した場合に限り、プロンプト、関連する図面コンテキスト、対応画像が送信されます。アプリは API リクエストに store:false を指定しますが、プロバイダー側の保持・ログ方針は別であり、このアプリは保証しません。BYO 資格情報はリクエストとともにこのアプリのサーバーへ送信されますが、ジョブ入力として保存されません。',
                })}</dd>
              </div>
              <div>
                <dt>{localize(language, { en: 'Retention', ja: '保持期間' })}</dt>
                <dd>{localize(language, {
                  en: 'The client defines no fixed expiry for local drafts. The diagram API defines no document retention period; cloud retention depends on the deployment’s storage policy. AI job results are API-retrievable for 1 hour, idempotency records for 24 hours, and production Blob soft-deletion may retain deleted job data for 30 days. These AI-job periods do not describe diagram retention.',
                  ja: 'ローカル下書きにクライアント側の固定期限はありません。図面 API に文書保持期間の定義はなく、クラウド保持はデプロイ環境のストレージ ポリシーによります。AI ジョブ結果は API から 1 時間、冪等性記録は 24 時間取得可能で、本番 Blob のソフト削除では削除済みジョブデータが 30 日保持される場合があります。これらは AI ジョブの期間であり、図面の保持期間ではありません。',
                })}</dd>
              </div>
            </dl>
          </section>

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
