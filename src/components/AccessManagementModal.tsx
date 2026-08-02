import React, { FormEvent, useCallback, useEffect, useRef, useState } from 'react';
import { AlertCircle, LogOut, RefreshCw, ShieldCheck, Trash2, UserPlus, X } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import {
  addAllowedUser,
  listAllowedUsers,
  removeAllowedUser,
  type AccessIdentity,
  type AllowedUser,
} from '../services/accessControlService';
import { OperationGeneration } from '../utils/operationGeneration';
import './AccessManagementModal.css';

interface AccessManagementModalProps {
  isOpen: boolean;
  identity: AccessIdentity;
  onClose: () => void;
}

const AccessManagementModal: React.FC<AccessManagementModalProps> = ({
  isOpen,
  identity,
  onClose,
}) => {
  const { language } = useLanguage();
  const text = useCallback(
    (en: string, ja: string) => localize(language, { en, ja }),
    [language],
  );
  const [users, setUsers] = useState<AllowedUser[]>([]);
  const [email, setEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const isOpenRef = useRef(isOpen);
  const loadGenerationRef = useRef(new OperationGeneration());
  const mutationGenerationRef = useRef(new OperationGeneration());

  isOpenRef.current = isOpen;

  const loadUsers = useCallback(async () => {
    const generation = loadGenerationRef.current.advance();
    setLoading(true);
    setError('');
    try {
      const nextUsers = await listAllowedUsers();
      if (!isOpenRef.current || !loadGenerationRef.current.isCurrent(generation)) return;
      setUsers(nextUsers);
    } catch (loadError) {
      if (!isOpenRef.current || !loadGenerationRef.current.isCurrent(generation)) return;
      console.error('[access] failed to load users:', loadError);
      setError(text('The access list could not be loaded.', 'アクセス許可リストを読み込めませんでした。'));
    } finally {
      if (isOpenRef.current && loadGenerationRef.current.isCurrent(generation)) {
        setLoading(false);
      }
    }
  }, [text]);

  useEffect(() => {
    if (!isOpen) {
      loadGenerationRef.current.advance();
      mutationGenerationRef.current.advance();
      setLoading(false);
      setSaving(false);
      return;
    }
    setEmail('');
    setNotice('');
    void loadUsers();
  }, [isOpen, loadUsers]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !saving) onClose();
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose, saving]);

  const handleAdd = async (event: FormEvent) => {
    event.preventDefault();
    const normalized = email.trim().toLowerCase();
    if (!normalized) {
      setError(text('Enter an email address.', 'メールアドレスを入力してください。'));
      return;
    }

    setSaving(true);
    setError('');
    setNotice('');
    const generation = mutationGenerationRef.current.advance();
    try {
      await addAllowedUser(normalized);
      if (!isOpenRef.current || !mutationGenerationRef.current.isCurrent(generation)) return;
      setEmail('');
      setNotice(text('Access was granted.', 'アクセスを許可しました。'));
      await loadUsers();
    } catch (saveError) {
      if (!isOpenRef.current || !mutationGenerationRef.current.isCurrent(generation)) return;
      console.error('[access] failed to add user:', saveError);
      setError(text(
        'Access could not be granted. The address may already be listed.',
        'アクセスを許可できませんでした。このメールアドレスは既に登録されている可能性があります。',
      ));
    } finally {
      if (isOpenRef.current && mutationGenerationRef.current.isCurrent(generation)) {
        setSaving(false);
      }
    }
  };

  const handleRemove = async (user: AllowedUser) => {
    const confirmed = window.confirm(text(
      `Remove ${user.email} from the access list?`,
      `${user.email} をアクセス許可リストから削除しますか？`,
    ));
    if (!confirmed) return;

    setSaving(true);
    setError('');
    setNotice('');
    const generation = mutationGenerationRef.current.advance();
    try {
      await removeAllowedUser(user.email);
      if (!isOpenRef.current || !mutationGenerationRef.current.isCurrent(generation)) return;
      setNotice(text('Access was removed.', 'アクセス許可を削除しました。'));
      await loadUsers();
    } catch (removeError) {
      if (!isOpenRef.current || !mutationGenerationRef.current.isCurrent(generation)) return;
      console.error('[access] failed to remove user:', removeError);
      setError(text('Access could not be removed.', 'アクセス許可を削除できませんでした。'));
    } finally {
      if (isOpenRef.current && mutationGenerationRef.current.isCurrent(generation)) {
        setSaving(false);
      }
    }
  };

  if (!isOpen) return null;

  return (
    <div className="access-modal-overlay" onClick={() => !saving && onClose()}>
      <section
        className="access-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="access-modal-title"
        onClick={(event) => event.stopPropagation()}
      >
        <header className="access-modal-header">
          <div>
            <div className="access-modal-eyebrow">
              <ShieldCheck size={16} />
              Microsoft Entra ID
            </div>
            <h2 id="access-modal-title">
              {text('Application access', 'アプリのアクセス管理')}
            </h2>
            <p>
              {text(
                'Only the listed email addresses can open the application.',
                '許可リストに登録されたメールアドレスだけがアプリを開けます。',
              )}
            </p>
          </div>
          <button
            type="button"
            className="access-icon-button"
            onClick={onClose}
            disabled={saving}
            aria-label={text('Close', '閉じる')}
            title={text('Close', '閉じる')}
          >
            <X size={22} />
          </button>
        </header>

        <div className="access-admin-summary">
          <div>
            <span>{text('Signed in administrator', 'サインイン中の管理者')}</span>
            <strong>{identity.email}</strong>
          </div>
          <a href="/.auth/logout?post_logout_redirect_uri=%2F">
            <LogOut size={16} />
            {text('Sign out', 'サインアウト')}
          </a>
        </div>

        <form className="access-add-form" onSubmit={handleAdd}>
          <label htmlFor="access-email">
            {text('Grant access by email address', 'メールアドレスでアクセスを許可')}
          </label>
          <div>
            <input
              id="access-email"
              type="email"
              value={email}
              onChange={(event) => setEmail(event.target.value)}
              placeholder="user@example.com"
              maxLength={254}
              autoComplete="off"
              disabled={saving}
            />
            <button type="submit" disabled={saving || email.trim().length === 0}>
              <UserPlus size={17} />
              {text('Add', '追加')}
            </button>
          </div>
        </form>

        {(error || notice) && (
          <div
            className={`access-message ${error ? 'error' : 'success'}`}
            role={error ? 'alert' : 'status'}
          >
            {error && <AlertCircle size={17} />}
            <span>{error || notice}</span>
          </div>
        )}

        <div className="access-list-heading">
          <div>
            <h3>{text('Allowed users', 'アクセス許可ユーザー')}</h3>
            <span>{users.length}</span>
          </div>
          <button
            type="button"
            className="access-refresh-button"
            onClick={() => void loadUsers()}
            disabled={loading || saving}
          >
            <RefreshCw size={16} className={loading ? 'spinning' : ''} />
            {text('Refresh', '更新')}
          </button>
        </div>

        <div className="access-user-list" aria-busy={loading}>
          {loading && users.length === 0 ? (
            <div className="access-empty">{text('Loading access list...', 'アクセス許可リストを読み込んでいます...')}</div>
          ) : (
            users.map((user) => (
              <div className="access-user-row" key={user.email}>
                <div className="access-user-avatar" aria-hidden="true">
                  {user.email.charAt(0).toUpperCase()}
                </div>
                <div className="access-user-details">
                  <strong>{user.email}</strong>
                  <span>
                    {user.isAdmin
                      ? text('Permanent administrator', '常設管理者')
                      : text(
                          `Added ${user.addedAt ? new Date(user.addedAt).toLocaleString(language === 'ja' ? 'ja-JP' : 'en-US') : ''}`,
                          `追加日時: ${user.addedAt ? new Date(user.addedAt).toLocaleString('ja-JP') : ''}`,
                        )}
                  </span>
                </div>
                {user.isAdmin ? (
                  <span className="access-admin-badge">{text('Admin', '管理者')}</span>
                ) : (
                  <button
                    type="button"
                    className="access-remove-button"
                    onClick={() => void handleRemove(user)}
                    disabled={saving}
                    aria-label={text(`Remove ${user.email}`, `${user.email} を削除`)}
                    title={text('Remove access', 'アクセス許可を削除')}
                  >
                    <Trash2 size={17} />
                  </button>
                )}
              </div>
            ))
          )}
        </div>
      </section>
    </div>
  );
};

export default AccessManagementModal;
