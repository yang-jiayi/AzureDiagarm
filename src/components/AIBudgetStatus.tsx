// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
import { useEffect, useState } from 'react';
import { getAIBudget, type AIBudget } from '../services/aiBudgetService';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './AIBudgetStatus.css';

export default function AIBudgetStatus() {
  const { language } = useLanguage();
  const [budget, setBudget] = useState<AIBudget | null>(null);
  const [refresh, setRefresh] = useState(0);
  useEffect(() => {
    let controller: AbortController | undefined;
    const update = async () => {
      if (document.hidden) return;
      controller?.abort();
      controller = new AbortController();
      try { setBudget(await getAIBudget(controller.signal)); }
      catch (error) { if (!(error instanceof DOMException && error.name === 'AbortError')) setBudget(null); }
    };
    void update();
    const interval = window.setInterval(() => void update(), 30_000);
    window.addEventListener('focus', update);
    return () => { controller?.abort(); window.clearInterval(interval); window.removeEventListener('focus', update); };
  }, [refresh]);
  const label = budget
    ? localize(language, {
      en: `AI: ${budget.remainingTokens.toLocaleString(language)} tokens left today · ${budget.concurrentRequests}/${budget.concurrentLimit} running${budget.mode === 'local' ? ' (local)' : ''}`,
      ja: `AI: 本日の残り ${budget.remainingTokens.toLocaleString(language)} トークン · 実行中 ${budget.concurrentRequests}/${budget.concurrentLimit}${budget.mode === 'local' ? '（ローカル）' : ''}`,
    })
    : localize(language, { en: 'AI budget unavailable', ja: 'AI の利用枠を取得できません' });
  const title = budget
    ? localize(language, {
      en: `Refresh budget. Input/output reservations are included. Resets ${new Date(budget.resetAt).toLocaleString(language)} (midnight UTC). Not a monetary balance.`,
      ja: `利用枠を更新。入力・出力の予約分を含みます。リセット: ${new Date(budget.resetAt).toLocaleString(language)}（UTC 午前0時）。金額の残高ではありません。`,
    })
    : localize(language, { en: 'Refresh AI budget. Sign in or retry shortly.', ja: 'AI 利用枠を更新。サインインするか、しばらくして再試行してください。' });
  return <button type="button" className="ai-budget-status" onClick={() => setRefresh(value => value + 1)} title={title} aria-label={`${label}. ${title}`}>
    {label}
  </button>;
}
