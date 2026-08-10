// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useState } from 'react';
import { X, DownloadCloud, RefreshCw, AlertTriangle, LogIn } from 'lucide-react';
import { AzureImportDisabledError } from '../services/azureImport';
import {
  isDelegatedMode,
  ensureSignedIn,
  getSubscriptions,
  getResourceGroups,
  type AzureSubscription,
  type AzureResourceGroup,
} from '../services/azureImportProvider';
import { getSignedInName } from '../services/msalAuth';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import ModalScaffold from './ModalScaffold';
import './AzureImportModal.css';

interface AzureImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  /** Runs the query + deterministic mapping + apply; resolves when done. */
  onImport: (subscriptionId: string, resourceGroup: string) => Promise<void>;
}

function errorMessage(value: unknown, fallback: string): string {
  return value instanceof Error && value.message ? value.message : fallback;
}

const AzureImportModal: React.FC<AzureImportModalProps> = ({ isOpen, onClose, onImport }) => {
  const { language } = useLanguage();
  const text = useCallback(
    (en: string, ja: string) => localize(language, { en, ja }),
    [language],
  );
  const delegated = isDelegatedMode();
  const [account, setAccount] = useState<string | undefined>(undefined);
  const [needsSignIn, setNeedsSignIn] = useState(delegated);
  const [signingIn, setSigningIn] = useState(false);
  const [subs, setSubs] = useState<AzureSubscription[]>([]);
  const [groups, setGroups] = useState<AzureResourceGroup[]>([]);
  const [subId, setSubId] = useState('');
  const [rg, setRg] = useState('');
  const [loadingSubs, setLoadingSubs] = useState(false);
  const [loadingGroups, setLoadingGroups] = useState(false);
  const [importing, setImporting] = useState(false);
  const [disabled, setDisabled] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadSubs = useCallback(() => {
    setLoadingSubs(true);
    setError(null);
    getSubscriptions()
      .then((s) => {
        setSubs(s);
        if (s.length === 1) setSubId(s[0].subscriptionId);
      })
      .catch((e) => {
        if (e instanceof AzureImportDisabledError) setDisabled(true);
        else setError(errorMessage(e, text('Failed to list subscriptions', 'Subscriptionの一覧取得に失敗しました')));
      })
      .finally(() => setLoadingSubs(false));
  }, [text]);

  // On open: server mode loads subs immediately; delegated mode loads subs only
  // once the user is signed in (otherwise show the sign-in gate).
  useEffect(() => {
    if (!isOpen) return;
    setError(null);
    setDisabled(false);
    if (!delegated) { setNeedsSignIn(false); loadSubs(); return; }
    getSignedInName()
      .then((name) => {
        if (name) { setAccount(name); setNeedsSignIn(false); loadSubs(); }
        else { setNeedsSignIn(true); }
      })
      .catch((e) => setError(errorMessage(e, text('Failed to initialize Azure sign-in', 'Azureサインインの初期化に失敗しました'))));
  }, [delegated, isOpen, loadSubs, text]);

  // Load resource groups when a subscription is chosen.
  useEffect(() => {
    if (!subId) { setGroups([]); setRg(''); return; }
    setRg('');
    setLoadingGroups(true);
    setError(null);
    getResourceGroups(subId)
      .then(setGroups)
      .catch((e) => setError(errorMessage(e, text('Failed to list resource groups', 'Resource Groupの一覧取得に失敗しました'))))
      .finally(() => setLoadingGroups(false));
  }, [subId, text]);

  if (!isOpen) return null;

  const handleSignIn = async () => {
    setSigningIn(true);
    setError(null);
    try {
      const name = await ensureSignedIn();
      setAccount(name);
      setNeedsSignIn(false);
      loadSubs();
    } catch (e: unknown) {
      setError(errorMessage(e, text('Sign-in failed', 'サインインに失敗しました')));
    } finally {
      setSigningIn(false);
    }
  };

  const handleImport = async () => {
    if (!subId || !rg) return;
    setImporting(true);
    setError(null);
    try {
      await onImport(subId, rg);
      onClose();
    } catch (e: unknown) {
      setError(errorMessage(e, text('Import failed', 'インポートに失敗しました')));
    } finally {
      setImporting(false);
    }
  };

  return (
    <ModalScaffold
      isOpen={isOpen}
      onClose={onClose}
      className="azure-import-modal"
      ariaLabelledBy="azure-import-title"
      closeOnBackdrop={!importing}
      closeOnEscape={!importing}
    >
      <div className="modal-header">
          <h2 id="azure-import-title">
            <DownloadCloud size={24} />
            {text('Import from Azure', 'Azureからインポート')}
          </h2>
          <button
            className="modal-close"
            onClick={onClose}
            title={text('Close', '閉じる')}
            aria-label={text('Close Import from Azure', 'Azureからインポートを閉じる')}
            disabled={importing}
          >
            <X size={24} />
          </button>
      </div>

      <div className="modal-body">
          {disabled ? (
            <div className="azimp-disabled azd-callout azd-callout--warning">
              <AlertTriangle size={20} />
              <div>
                <p>
                  <strong>
                    {text(
                      'Azure import is disabled on the server.',
                      'Azureインポートはサーバーで無効になっています。',
                    )}
                  </strong>
                </p>
                <p className="azimp-muted">
                  {text(
                    'Reverse-engineering a live resource group uses the server identity to enumerate resources, so it is off by default. To enable it for local or self-hosted use, set ',
                    '稼働中のResource Groupをリバースエンジニアリングすると、サーバーIDでリソースを列挙するため、既定では無効です。ローカルまたはセルフホスト環境で有効にするには、',
                  )}
                  <code>AZURE_IMPORT_ENABLED=true</code>
                  {text(
                    ' on the token server (with az login or a Reader-scoped managed identity), then reopen this dialog.',
                    'をToken Serverに設定し（az loginまたはReaderスコープのManaged Identityを使用）、このダイアログを開き直してください。',
                  )}
                </p>
              </div>
            </div>
          ) : needsSignIn ? (
            <div className="azimp-signin">
              <p className="azimp-intro">
                {text(
                  'Sign in with your Azure account to reverse-engineer a resource group you can access. We request read-only ',
                  'Azureアカウントでサインインし、アクセス可能なResource Groupをリバースエンジニアリングします。読み取り専用の',
                )}
                <strong>Azure Service Management</strong>
                {text(
                  ' access and query only what your permissions allow. Nothing is stored.',
                  'アクセスのみを要求し、ユーザーの権限で許可された情報だけを照会します。情報は保存されません。',
                )}
              </p>
              <button
                className="azd-button azd-button--primary azimp-signin-btn"
                onClick={handleSignIn}
                disabled={signingIn}
              >
                <LogIn size={16} />
                {signingIn
                  ? text('Signing in…', 'サインイン中…')
                  : text('Sign in to Azure', 'Azureにサインイン')}
              </button>
            </div>
          ) : (
            <>
              <p className="azimp-intro">
                {text(
                  'Reverse-engineer a deployed resource group into a diagram using deterministic Azure Resource Graph mapping.',
                  'デプロイ済みResource GroupをAzure Resource Graphで決定論的にマッピングし、実環境を反映した図に変換します。',
                )}
                {account && <> {text('Signed in as', 'サインイン中:')} <strong>{account}</strong>.</>}
              </p>

              <div className="form-group azd-field">
                <label htmlFor="azimp-sub">{text('Subscription', 'Subscription')}</label>
                <select
                  id="azimp-sub"
                  className="azd-control"
                  value={subId}
                  onChange={(e) => setSubId(e.target.value)}
                  disabled={loadingSubs || importing}
                >
                  <option value="">
                    {loadingSubs
                      ? text('Loading subscriptions…', 'Subscriptionを読み込み中…')
                      : text('Select a subscription…', 'Subscriptionを選択…')}
                  </option>
                  {subs.map((s) => (
                    <option key={s.subscriptionId} value={s.subscriptionId}>
                      {s.displayName} ({s.subscriptionId.slice(0, 8)}…)
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group azd-field">
                <label htmlFor="azimp-rg">{text('Resource group', 'Resource Group')}</label>
                <select
                  id="azimp-rg"
                  className="azd-control"
                  value={rg}
                  onChange={(e) => setRg(e.target.value)}
                  disabled={!subId || loadingGroups || importing}
                >
                  <option value="">
                    {!subId
                      ? text('Choose a subscription first', '最初にSubscriptionを選択してください')
                      : loadingGroups
                        ? text('Loading resource groups…', 'Resource Groupを読み込み中…')
                        : text('Select a resource group…', 'Resource Groupを選択…')}
                  </option>
                  {groups.map((g) => (
                    <option key={g.name} value={g.name}>{g.name} · {g.location}</option>
                  ))}
                </select>
              </div>

              {importing && (
                <div
                  className="azimp-progress azd-callout azd-callout--info"
                  role="status"
                  aria-live="polite"
                >
                  <RefreshCw size={16} className="spin-icon" />
                  {text('Scanning', 'スキャン中:')} <strong>{rg}</strong>
                  {text(' and building the diagram…', '。図を作成しています…')}
                </div>
              )}
            </>
          )}

          {error && (
            <div className="azimp-error azd-callout azd-callout--danger" role="alert">
              <AlertTriangle size={16} /> {error}
            </div>
          )}
      </div>

      <div className="modal-actions">
          <button className="btn-secondary" onClick={onClose} disabled={importing}>
            {text('Cancel', 'キャンセル')}
          </button>
          {!disabled && !needsSignIn && (
            <button className="btn-primary" onClick={handleImport} disabled={!subId || !rg || importing}>
              {importing
                ? text('Importing…', 'インポート中…')
                : text('Import resource group', 'Resource Groupをインポート')}
            </button>
          )}
      </div>
    </ModalScaffold>
  );
};

export default AzureImportModal;
