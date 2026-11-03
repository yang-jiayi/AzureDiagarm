// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useEffect, useState } from 'react';
import { ChevronDown, RefreshCw } from 'lucide-react';
import { getAIBudget, type AIBudget } from '../services/aiBudgetService';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './AIBudgetStatus.css';

export default function AIBudgetStatus() {
  const { language } = useLanguage();
  const [budget, setBudget] = useState<AIBudget | null>(null);
  const [loading, setLoading] = useState(true);
  const [refresh, setRefresh] = useState(0);

  useEffect(() => {
    let controller: AbortController | undefined;
    const update = async () => {
      if (document.hidden) return;
      controller?.abort();
      const requestController = new AbortController();
      controller = requestController;
      setLoading(true);
      try {
        setBudget(await getAIBudget(requestController.signal));
      } catch (error) {
        if (!(error instanceof DOMException && error.name === 'AbortError')) setBudget(null);
      } finally {
        if (controller === requestController) setLoading(false);
      }
    };
    void update();
    const interval = window.setInterval(() => void update(), 30_000);
    window.addEventListener('focus', update);
    return () => {
      controller?.abort();
      window.clearInterval(interval);
      window.removeEventListener('focus', update);
    };
  }, [refresh]);

  const summary = loading
    ? localize(language, { en: 'Checking AI status', ja: 'AI 状態を確認中' })
    : budget
      ? localize(language, { en: 'AI usage', ja: 'AI 利用状況' })
      : localize(language, { en: 'AI status unavailable', ja: 'AI 状態を取得できません' });
  const usage = budget
    ? localize(language, {
        en: `${budget.remainingTokens.toLocaleString(language)} tokens remaining today · ${budget.concurrentRequests}/${budget.concurrentLimit} requests running${budget.mode === 'local' ? ' · local estimate' : ''}`,
        ja: `本日の残り ${budget.remainingTokens.toLocaleString(language)} トークン · 実行中 ${budget.concurrentRequests}/${budget.concurrentLimit}${budget.mode === 'local' ? ' · ローカル見積' : ''}`,
      })
    : localize(language, {
        en: 'Sign in or retry shortly to check AI usage.',
        ja: 'サインインするか、しばらくしてから AI 利用状況を再確認してください。',
      });
  const reset = budget
    ? localize(language, {
        en: `Resets ${new Date(budget.resetAt).toLocaleString(language)} (midnight UTC). This is not a monetary balance.`,
        ja: `${new Date(budget.resetAt).toLocaleString(language)} にリセット（UTC 午前0時）。金額の残高ではありません。`,
      })
    : null;

  return (
    <details className="ai-budget-status">
      <summary aria-live="polite">
        <span>{summary}</span>
        <ChevronDown size={14} aria-hidden="true" />
      </summary>
      <div className="ai-budget-status-panel">
        <p>{usage}</p>
        {reset && <small>{reset}</small>}
        <button
          type="button"
          onClick={() => setRefresh(value => value + 1)}
          aria-label={localize(language, { en: 'Refresh AI usage details', ja: 'AI 利用状況を更新' })}
        >
          <RefreshCw size={14} aria-hidden="true" />
          {localize(language, { en: 'Refresh', ja: '更新' })}
        </button>
      </div>
    </details>
  );
}
