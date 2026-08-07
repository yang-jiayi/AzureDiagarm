// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useRef, useEffect, useCallback } from 'react';
import { Sparkles, X, Send, Loader2, AlertCircle, MessageSquare, ChevronDown, ChevronUp, Shield, Activity, DollarSign, Wrench, Zap, Lightbulb, type LucideIcon } from 'lucide-react';
import { generateArchitectureWithAI, generateFollowUpSuggestions, isAzureOpenAIConfigured } from '../services/azureOpenAI';
import { useModelSettings, MODEL_CONFIG } from '../stores/modelSettingsStore';
import {
  getBYOAIProviderLabel,
  isBYOAIReady,
  useBYOAISettings,
} from '../stores/byoAISettingsStore';
import {
  buildModificationPrompt,
  summarizeArchitectureChange,
  CurrentArchitecture,
} from '../services/modificationPrompt';
import './ArchitectureChatPanel.css';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';
import { OperationGeneration } from '../utils/operationGeneration';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';
import { MEDIA_QUERIES } from '../styles/breakpoints';
import ResponsiveDrawer from './ResponsiveDrawer';

interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'error';
  text: string;
  ts: number;
}

interface ArchitectureChatPanelProps {
  isOpen: boolean;
  onClose: () => void;
  currentArchitecture: CurrentArchitecture;
  diagramKey: string;
  /** Applies a generated architecture to the canvas (App's handleAIGenerate). */
  onApply: (architecture: any, prompt: string, autoSnapshot?: boolean) => void | Promise<void>;
}

const CHAT_PANEL_WIDTH_KEY = 'azure-diagram-builder.chatPanelWidth.v1';
const DEFAULT_CHAT_PANEL_WIDTH = 460;
const MIN_CHAT_PANEL_WIDTH = 360;
const MAX_CHAT_PANEL_WIDTH = 720;
function clampChatPanelWidth(width: number): number {
  return Math.min(MAX_CHAT_PANEL_WIDTH, Math.max(MIN_CHAT_PANEL_WIDTH, width));
}

// Cold start: when the canvas is empty, offer complete starter architectures
// so Chat works as a first-class entry point (not just a refinement tool).
const STARTER_SUGGESTIONS: LocalizedText[] = [
  {
    en: 'Customer-facing web app on App Service with Azure SQL and Azure Cache for Redis, fronted by Azure Front Door with WAF and Application Insights',
    ja: 'Azure Front Door（WAF）とApplication Insightsを備えた、App Service、Azure SQL、Azure Cache for Redisによる顧客向けWebアプリ',
  },
  {
    en: 'Order-processing pipeline: Service Bus queue to Azure Functions to Cosmos DB, with dead-lettering and secrets in Key Vault',
    ja: 'Service Bus QueueからAzure Functions、Cosmos DBへ連携し、Dead-letter QueueとKey Vaultを備えた注文処理パイプライン',
  },
  {
    en: 'Internal REST API on Container Apps backed by Azure SQL, secured with Microsoft Entra ID and API Management',
    ja: 'Azure SQLをバックエンドに使用し、Microsoft Entra IDとAPI Managementで保護するContainer Apps上の社内REST API',
  },
  {
    en: 'Document-processing workflow: Blob Storage triggers Azure Functions to run Azure AI Document Intelligence, with results in Cosmos DB',
    ja: 'Blob StorageをトリガーにAzure FunctionsでAzure AI Document Intelligenceを実行し、結果をCosmos DBに保存するドキュメント処理',
  },
  {
    en: 'Secure AI assistant: Azure OpenAI behind private endpoints, exposed through API Management with managed identity',
    ja: 'Private Endpoint配下のAzure OpenAIをAPI ManagementとManaged Identityで公開する安全なAIアシスタント',
  },
];

// Cold start (advanced): richer, enterprise-grade patterns revealed behind a
// "More ideas" toggle so first-timers aren't overwhelmed but power users can
// see the tool's ceiling.
const ADVANCED_STARTER_SUGGESTIONS: LocalizedText[] = [
  {
    en: 'Multi-region active-active e-commerce platform: Azure Front Door, AKS in paired regions, geo-replicated Cosmos DB, and Service Bus for order events',
    ja: 'Azure Front Door、ペアリージョンのAKS、geo-replicationされたCosmos DB、注文イベント用Service Busによるマルチリージョンactive-active EC基盤',
  },
  {
    en: 'HIPAA-compliant healthcare data platform: private-endpoint ingestion, AKS clinical workloads, Azure Health Data Services FHIR service, and Microsoft Purview governance',
    ja: 'Private Endpoint経由の取り込み、AKS医療ワークロード、Azure Health Data Services FHIR service、Microsoft PurviewガバナンスによるHIPAA準拠医療データ基盤',
  },
  {
    en: 'Enterprise landing zone: hub-and-spoke with Azure Firewall, Bastion, Private DNS zones, and centralized Log Analytics',
    ja: 'Azure Firewall、Bastion、Private DNS Zone、集中管理されたLog AnalyticsによるHub-and-Spoke構成のEnterprise Landing Zone',
  },
  {
    en: 'Real-time fraud detection: Event Hubs to Stream Analytics to Azure Machine Learning scoring to Cosmos DB, with Event Grid alerting',
    ja: 'Event Hubs、Stream Analytics、Azure Machine Learningスコアリング、Cosmos DB、Event Grid通知によるリアルタイム不正検知',
  },
  {
    en: 'RAG knowledge platform: Azure OpenAI, Azure AI Search, and Cosmos DB, with Azure Functions ingestion and private endpoints',
    ja: 'Azure OpenAI、Azure AI Search、Cosmos DB、Azure Functions取り込みをすべてPrivate Endpointで保護するRAGナレッジ基盤',
  },
  {
    en: 'Event-driven microservices on AKS with KEDA autoscaling from Service Bus, Key Vault CSI Driver, and private-link Azure Container Registry',
    ja: 'Service Bus連動KEDA自動スケーリング、Key Vault CSI Driver、Private Link対応Azure Container Registryを備えたAKS上のイベント駆動型マイクロサービス',
  },
];

// Warm start: once a diagram exists, offer incremental refinements. Used as a
// fallback when no context-aware "what's missing" suggestions apply.
const REFINE_SUGGESTIONS: LocalizedText[] = [
  {
    en: 'Add Azure Front Door with WAF in front of the web tier',
    ja: 'Web層の前段にWAF付きAzure Front Doorを追加する',
  },
  {
    en: 'Make it zone-redundant for high availability',
    ja: '高可用性のためにゾーン冗長構成にする',
  },
  {
    en: 'Add a Redis cache between the API and the database',
    ja: 'APIとデータベースの間にRedis Cacheを追加する',
  },
  {
    en: 'Add monitoring with Application Insights and Log Analytics',
    ja: 'Application InsightsとLog Analyticsで監視を追加する',
  },
  {
    en: 'Put private endpoints on the data services',
    ja: 'データ サービスにPrivate Endpointを追加する',
  },
];

// Context-aware refinement suggestions: inspect the services already on the
// canvas and propose the most valuable *missing* Well-Architected additions
// (security, reliability, observability). Falls back to the static list when
// nothing obvious is missing so the panel is never empty.
function computeRefineSuggestions(nodes: any[], language: 'en' | 'ja'): string[] {
  const labels = nodes
    .filter((n) => n?.type === 'azureNode')
    .map((n) => String(n?.data?.label || '').toLowerCase());
  const has = (...needles: string[]) =>
    labels.some((l) => needles.some((needle) => l.includes(needle)));

  const suggestions: string[] = [];

  // Security / identity
  if (!has('key vault')) {
    suggestions.push(localize(language, {
      en: 'Add Key Vault and use managed identities for secrets',
      ja: 'Key Vaultを追加し、シークレットへのアクセスにはManaged Identityを使用する',
    }));
  }
  if (!has('private endpoint', 'private link')) {
    suggestions.push(localize(language, REFINE_SUGGESTIONS[4]));
  }
  if (!has('front door', 'application gateway', 'firewall', 'waf')) {
    suggestions.push(localize(language, {
      en: 'Add Azure Front Door with a WAF in front of the web tier',
      ja: 'Web層の前段にWAF付きAzure Front Doorを追加する',
    }));
  }
  // Observability
  if (!has('application insights', 'monitor', 'log analytics')) {
    suggestions.push(localize(language, REFINE_SUGGESTIONS[3]));
  }
  // Reliability
  suggestions.push(localize(language, REFINE_SUGGESTIONS[1]));
  if (!has('redis', 'cache')) {
    suggestions.push(localize(language, REFINE_SUGGESTIONS[2]));
  }

  const deduped = Array.from(new Set(suggestions));
  if (deduped.length >= 3) return deduped.slice(0, 5);
  // Pad with static defaults if the diagram already covers most pillars.
  return Array.from(new Set([
    ...deduped,
    ...REFINE_SUGGESTIONS.map(suggestion => localize(language, suggestion)),
  ])).slice(0, 5);
}

const uid = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

// Well-Architected pillar tagging for suggestion chips (Tier 4). Heuristic
// keyword match maps each suggestion to a pillar so chips carry a small icon.
type Pillar = 'security' | 'reliability' | 'cost' | 'operations' | 'performance';
const PILLAR_META: Record<Pillar, { label: LocalizedText; Icon: LucideIcon; className: string }> = {
  security: { label: { en: 'Security', ja: 'セキュリティ' }, Icon: Shield, className: 'pillar-security' },
  reliability: { label: { en: 'Reliability', ja: '信頼性' }, Icon: Activity, className: 'pillar-reliability' },
  cost: { label: { en: 'Cost', ja: 'コスト' }, Icon: DollarSign, className: 'pillar-cost' },
  operations: { label: { en: 'Operations', ja: '運用' }, Icon: Wrench, className: 'pillar-operations' },
  performance: { label: { en: 'Performance', ja: 'パフォーマンス' }, Icon: Zap, className: 'pillar-performance' },
};
function pillarFor(text: string): Pillar {
  const t = text.toLowerCase();
  if (/(private|key vault|waf|firewall|defender|encrypt|rbac|identity|secret|auth|ddos|network isolation)/.test(t)) return 'security';
  if (/(zone|redundan|availability|failover|backup|geo|replica|resilien|disaster|multi-region|sla)/.test(t)) return 'reliability';
  if (/(cost|budget|reserved|spot|right-?siz|cheaper|save money|lower tier)/.test(t)) return 'cost';
  if (/(cache|redis|cdn|front door|latency|throughput|scale out|accelerat|performance)/.test(t)) return 'performance';
  return 'operations';
}

const ArchitectureChatPanel: React.FC<ArchitectureChatPanelProps> = ({
  isOpen,
  onClose,
  currentArchitecture,
  diagramKey,
  onApply,
}) => {
  const { t, translate, language } = useLanguage();
  const isCompactChat = useMediaQuery(MEDIA_QUERIES.workspace);
  const [panelWidth, setPanelWidth] = useState(() => {
    const stored = readLocalStorage(CHAT_PANEL_WIDTH_KEY);
    const parsed = stored === null ? Number.NaN : Number(stored);
    return Number.isFinite(parsed)
      ? clampChatPanelWidth(parsed)
      : DEFAULT_CHAT_PANEL_WIDTH;
  });
  const resizeStateRef = useRef<{ startX: number; startWidth: number } | null>(null);

  useEffect(() => {
    document.documentElement.style.setProperty('--arch-chat-width', `${panelWidth}px`);
  }, [panelWidth]);

  const updatePanelWidth = useCallback((width: number) => {
    const next = clampChatPanelWidth(width);
    setPanelWidth(next);
    writeLocalStorage(CHAT_PANEL_WIDTH_KEY, String(next));
  }, []);

  const handleResizePointerDown = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    resizeStateRef.current = { startX: event.clientX, startWidth: panelWidth };
  }, [panelWidth]);

  const handleResizePointerMove = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    const state = resizeStateRef.current;
    if (!state) return;
    updatePanelWidth(state.startWidth + (state.startX - event.clientX));
  }, [updatePanelWidth]);

  const handleResizePointerUp = useCallback((event: React.PointerEvent<HTMLDivElement>) => {
    if (resizeStateRef.current) {
      event.currentTarget.releasePointerCapture(event.pointerId);
      resizeStateRef.current = null;
    }
  }, []);

  const handleResizeKeyDown = useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
    if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
    event.preventDefault();
    updatePanelWidth(panelWidth + (event.key === 'ArrowLeft' ? 24 : -24));
  }, [panelWidth, updatePanelWidth]);

  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState('');
  const [isSending, setIsSending] = useState(false);
  const [showAdvanced, setShowAdvanced] = useState(false);
  // Suggestions the user has already picked this session, so follow-up chips
  // keep advancing instead of re-offering the same ideas.
  const [usedSuggestions, setUsedSuggestions] = useState<Set<string>>(new Set());
  // Tier 3: change-specific follow-ups from a fast model, keyed to the assistant
  // turn they were generated for. Null until (and unless) they arrive.
  const [modelFollowUps, setModelFollowUps] = useState<{ forMsgId: string; items: string[] } | null>(null);
  // Tier 4: loading flags for the background follow-up fetch and the
  // "What would you add?" single-best-recommendation button.
  const [followUpsLoading, setFollowUpsLoading] = useState(false);
  const [askingBest, setAskingBest] = useState(false);
  const [modelSettings] = useModelSettings();
  const byoSnapshot = useBYOAISettings();

  const threadRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const latestFollowUpRequestRef = useRef<string | null>(null);
  const diagramKeyRef = useRef(diagramKey);
  const sendGenerationRef = useRef(new OperationGeneration());
  const bestSuggestionGenerationRef = useRef(new OperationGeneration());

  diagramKeyRef.current = diagramKey;

  const byoReady = byoSnapshot.settings.enabled && isBYOAIReady();
  const configured = isAzureOpenAIConfigured() || byoReady;
  const hasDiagram = currentArchitecture.nodes.some((n) => n.type === 'azureNode');
  const modelName = byoReady
    ? `${getBYOAIProviderLabel(byoSnapshot.settings.provider)} · ${byoSnapshot.settings.model}`
    : (MODEL_CONFIG[modelSettings.model]?.displayName || modelSettings.model);

  const markUsed = (s: string) =>
    setUsedSuggestions((prev) => (prev.has(s) ? prev : new Set(prev).add(s)));

  // Live, context-aware follow-ups shown under the latest reply during an active
  // chat. Recomputed from the current (post-change) canvas, so they evolve as the
  // diagram grows; already-picked ideas are filtered out.
  const staticFollowUps = hasDiagram
    ? computeRefineSuggestions(currentArchitecture.nodes, language)
        .filter((s) => !usedSuggestions.has(s))
        .slice(0, 3)
    : [];
  // Tier 3: prefer the model's change-specific follow-ups when available; fall
  // back to the static rule-based chips otherwise.
  const dynamicFollowUps = (modelFollowUps?.items || [])
    .filter((s) => !usedSuggestions.has(s))
    .slice(0, 3);
  const followUps = dynamicFollowUps.length ? dynamicFollowUps : staticFollowUps;

  // Auto-scroll to the newest message.
  useEffect(() => {
    if (threadRef.current) {
      threadRef.current.scrollTop = threadRef.current.scrollHeight;
    }
  }, [messages, isSending]);

  useEffect(() => {
    sendGenerationRef.current.advance();
    bestSuggestionGenerationRef.current.advance();
    latestFollowUpRequestRef.current = null;
    setMessages([]);
    setInput('');
    setIsSending(false);
    setShowAdvanced(false);
    setUsedSuggestions(new Set());
    setModelFollowUps(null);
    setFollowUpsLoading(false);
    setAskingBest(false);
  }, [diagramKey]);

  // Focus the composer when the panel opens.
  useEffect(() => {
    if (isOpen) {
      const t = setTimeout(() => inputRef.current?.focus(), 120);
      return () => clearTimeout(t);
    }
  }, [isOpen]);

  const send = useCallback(
    async (raw: string) => {
      const text = raw.trim();
      if (!text || isSending) return;

      const requestGeneration = sendGenerationRef.current.advance();
      const requestDiagramKey = diagramKey;
      setInput('');
      latestFollowUpRequestRef.current = null;
      setModelFollowUps(null);
      setFollowUpsLoading(false);
      const userMsg: ChatMessage = { id: uid(), role: 'user', text, ts: Date.now() };
      setMessages((prev) => [...prev, userMsg]);
      setIsSending(true);

      // Snapshot the canvas state BEFORE applying so we can diff for a summary.
      const before: CurrentArchitecture = {
        nodes: currentArchitecture.nodes,
        edges: currentArchitecture.edges,
        architectureName: currentArchitecture.architectureName,
      };

      // Recent user instructions help the model resolve references.
      const recentRequests = [...messages, userMsg]
        .filter((m) => m.role === 'user')
        .slice(-5)
        .map((m) => m.text);

      try {
        const prompt = buildModificationPrompt(before, text, recentRequests.slice(0, -1), language);
        const result = await generateArchitectureWithAI(prompt, undefined, undefined, language);
        if (
          diagramKeyRef.current !== requestDiagramKey
          || !sendGenerationRef.current.isCurrent(requestGeneration)
        ) return;

        await onApply(result, text, true);
        if (
          diagramKeyRef.current !== requestDiagramKey
          || !sendGenerationRef.current.isCurrent(requestGeneration)
        ) return;

        const summary = summarizeArchitectureChange(before, result, language);
        const asstId = uid();
        setMessages((prev) => [
          ...prev,
          { id: asstId, role: 'assistant', text: summary, ts: Date.now() },
        ]);

        // Tier 3: fetch change-specific follow-ups in the background (non-blocking).
        // The static rule-based chips render immediately; these replace them when
        // they arrive. Uses result.services (post-change) to avoid stale state.
        const nextServices = Array.isArray((result as any)?.services)
          ? (result as any).services
              .map((s: any) => String(s?.label ?? s?.name ?? s?.service ?? '').trim())
              .filter(Boolean)
          : [];
        setModelFollowUps(null);
        setFollowUpsLoading(true);
        latestFollowUpRequestRef.current = asstId;
        void generateFollowUpSuggestions({ services: nextServices, lastChange: summary, recentRequests, language })
          .then((items) => {
            if (
              diagramKeyRef.current === requestDiagramKey
              && sendGenerationRef.current.isCurrent(requestGeneration)
              && latestFollowUpRequestRef.current === asstId
              && items.length
            ) {
              setModelFollowUps({ forMsgId: asstId, items });
            }
          })
          .catch(() => { /* fall back to static chips */ })
          .finally(() => {
            if (
              diagramKeyRef.current === requestDiagramKey
              && sendGenerationRef.current.isCurrent(requestGeneration)
              && latestFollowUpRequestRef.current === asstId
            ) {
              setFollowUpsLoading(false);
            }
          });
      } catch (err: any) {
        if (
          diagramKeyRef.current !== requestDiagramKey
          || !sendGenerationRef.current.isCurrent(requestGeneration)
          || err?.name === 'CloudDiagramOperationCancelledError'
        ) return;
        setMessages((prev) => [
          ...prev,
          {
            id: uid(),
            role: 'error',
            text: err?.message
              ? translate(err.message)
              : localize(language, {
                  en: 'Something went wrong updating the diagram. Please try again.',
                  ja: '図の更新中に問題が発生しました。もう一度お試しください。',
                }),
            ts: Date.now(),
          },
        ]);
      } finally {
        if (
          diagramKeyRef.current === requestDiagramKey
          && sendGenerationRef.current.isCurrent(requestGeneration)
        ) {
          setIsSending(false);
        }
      }
    },
    [diagramKey, isSending, messages, currentArchitecture, onApply, language, translate],
  );

  // Tier 4: "What would you add?" — ask the model for the single highest-impact
  // next step (from the current canvas) and apply it like a chip click.
  const handleAskBest = async () => {
    if (isSending || askingBest || !configured) return;
    const requestGeneration = bestSuggestionGenerationRef.current.advance();
    const requestDiagramKey = diagramKey;
    setAskingBest(true);
    try {
      const services = currentArchitecture.nodes
        .filter((n) => n.type === 'azureNode')
        .map((n) => String(n.data?.label || '').trim())
        .filter(Boolean);
      const recent = messages.filter((m) => m.role === 'user').slice(-4).map((m) => m.text);
      const best = await generateFollowUpSuggestions({
        services,
        lastChange: '',
        recentRequests: recent,
        count: 1,
        language,
      });
      if (
        diagramKeyRef.current !== requestDiagramKey
        || !bestSuggestionGenerationRef.current.isCurrent(requestGeneration)
      ) return;
      if (best[0]) {
        markUsed(best[0]);
        await send(best[0]);
      }
    } finally {
      if (
        diagramKeyRef.current === requestDiagramKey
        && bestSuggestionGenerationRef.current.isCurrent(requestGeneration)
      ) {
        setAskingBest(false);
      }
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
      e.preventDefault();
      send(input);
    }
  };

  return (
    <ResponsiveDrawer
        isOpen={isOpen}
        modal={isCompactChat}
        placement="right"
        className="arch-chat-panel"
        role="complementary"
        backdropClassName="arch-chat-backdrop"
        ariaLabel={t("Architecture chat")}
        onClose={onClose}
        backgroundSelectors={[
          '.app > .app-header',
          '.app > .workspace',
        ]}
        style={{ '--arch-chat-width': `${panelWidth}px` } as React.CSSProperties}
      >
        <div
          className="arch-chat-resizer"
          role="separator"
          aria-orientation="vertical"
          aria-label={localize(language, {
            en: 'Resize Architecture Chat',
            ja: 'Architecture Chat の幅を変更',
          })}
          aria-valuemin={MIN_CHAT_PANEL_WIDTH}
          aria-valuemax={MAX_CHAT_PANEL_WIDTH}
          aria-valuenow={panelWidth}
          tabIndex={0}
          onPointerDown={handleResizePointerDown}
          onPointerMove={handleResizePointerMove}
          onPointerUp={handleResizePointerUp}
          onPointerCancel={handleResizePointerUp}
          onKeyDown={handleResizeKeyDown}
        />
      <div className="arch-chat-header">
        <div className="arch-chat-title">
          <MessageSquare size={18} />
          <span>{t("Architecture Chat")}</span>
        </div>
        <button className="arch-chat-close" onClick={onClose} title={t("Close chat")} aria-label={t("Close chat")}>
          <X size={18} />
        </button>
      </div>

      <div className="arch-chat-subhead">
        <Sparkles size={13} />
        <span>
          {hasDiagram
            ? <>{localize(language, { en: 'Refine your diagram in natural language ·', ja: '自然言語で図を調整 ·' })}{' '}<strong>{modelName}</strong></>
            : <>{localize(language, { en: 'Describe it, I’ll draw it — then refine in natural language ·', ja: '要件を説明すると図を作成し、そのまま自然言語で調整できます ·' })}{' '}<strong>{modelName}</strong></>}
        </span>
      </div>

      <div className="arch-chat-thread" ref={threadRef}>
        {messages.length === 0 && (
          <div className="arch-chat-empty">
            <p className="arch-chat-empty-title">
              {hasDiagram
                ? localize(language, { en: 'Describe a change and I’ll update the diagram.', ja: '変更内容を入力すると図を更新します。' })
                : localize(language, { en: 'Start by describing what you want to build — I’ll draw the first version, then we refine it together.', ja: '作成したい内容を入力してください。最初の図を作成し、その後一緒に調整できます。' })}
            </p>
            <p className="arch-chat-empty-sub">
              {hasDiagram
                ? localize(language, { en: 'Every change is saved to version history, so you can experiment freely.', ja: '各変更はバージョン履歴に保存されるため、自由に試せます。' })
                : localize(language, { en: 'Pick a starter below or type your own. Every step is saved to version history.', ja: '下の例を選ぶか、要件を入力してください。各手順はバージョン履歴に保存されます。' })}
            </p>
            <div className="arch-chat-suggestions">
              {(hasDiagram
                ? computeRefineSuggestions(currentArchitecture.nodes, language)
                : STARTER_SUGGESTIONS.map(suggestion => localize(language, suggestion))
              ).map((s) => (
                <button
                  key={s}
                  className="arch-chat-chip"
                  disabled={isSending || !configured}
                  onClick={() => { markUsed(s); send(s); }}
                >
                  {s}
                </button>
              ))}

              {!hasDiagram && showAdvanced && ADVANCED_STARTER_SUGGESTIONS.map(item => {
                const s = localize(language, item);
                return (
                <button
                  key={s}
                  className="arch-chat-chip arch-chat-chip-advanced"
                  disabled={isSending || !configured}
                  onClick={() => { markUsed(s); send(s); }}
                >
                  {s}
                </button>
                );
              })}

              {!hasDiagram && (
                <button
                  type="button"
                  className="arch-chat-more-toggle"
                  onClick={() => setShowAdvanced((v) => !v)}
                  aria-expanded={showAdvanced}
                >
                  {showAdvanced
                    ? <><ChevronUp size={15} /> {' '}{t("Fewer ideas")}</>
                    : <><ChevronDown size={15} /> {' '}{t("More ideas — enterprise patterns")}</>}
                </button>
              )}
            </div>
          </div>
        )}

        {messages.map((m) => (
          <div key={m.id} className={`arch-chat-msg arch-chat-msg-${m.role}`}>
            {m.role === 'error' && <AlertCircle size={15} className="arch-chat-msg-icon" />}
            {m.role === 'assistant' && <Sparkles size={15} className="arch-chat-msg-icon" />}
            <div className="arch-chat-bubble">{m.text}</div>
          </div>
        ))}

        {isSending && (
          <div className="arch-chat-msg arch-chat-msg-assistant">
            <Loader2 size={15} className="arch-chat-msg-icon spin" />
            <div className="arch-chat-bubble arch-chat-bubble-pending">{t("Updating the diagram…")}</div>
          </div>
        )}

        {messages.length > 0 && configured && !isSending && (hasDiagram ? followUps.length > 0 : true) && (
          <div className="arch-chat-followups">
            <div className="arch-chat-followups-label">
              <Sparkles size={12} />
              {hasDiagram
                ? (followUpsLoading && dynamicFollowUps.length === 0
                    ? <>{t("Finding tailored suggestions…")}{' '}<Loader2 size={11} className="spin" /></>
                    : <>{t("Suggested next steps")}</>)
                : <>{t("Start a new architecture")}</>}
            </div>
            <div
              className="arch-chat-suggestions arch-chat-suggestions-inline"
              role="group"
              aria-label={hasDiagram ? t("Suggested follow-ups") : t("Starter architectures")}
            >
              {(hasDiagram
                ? followUps
                : STARTER_SUGGESTIONS.map(suggestion => localize(language, suggestion))
              ).map((s) => {
                const meta = hasDiagram ? PILLAR_META[pillarFor(s)] : null;
                const Icon = meta?.Icon;
                return (
                  <button
                    key={s}
                    className={`arch-chat-chip arch-chat-chip-followup${meta ? ` ${meta.className}` : ''}`}
                    disabled={isSending || !configured}
                    title={meta
                      ? localize(language, {
                          en: `${localize(language, meta.label)} improvement`,
                          ja: `${localize(language, meta.label)}の改善`,
                        })
                      : undefined}
                    onClick={() => { markUsed(s); send(s); }}
                  >
                    {Icon && <Icon size={12} className="arch-chat-chip-icon" />}
                    {s}
                  </button>
                );
              })}

              {hasDiagram && (
                <button
                  type="button"
                  className="arch-chat-chip arch-chat-chip-ask"
                  disabled={isSending || askingBest || !configured}
                  title={t("Ask the model for the single highest-impact improvement")}
                  onClick={handleAskBest}
                >
                  {askingBest
                    ? <Loader2 size={12} className="spin arch-chat-chip-icon" />
                    : <Lightbulb size={12} className="arch-chat-chip-icon" />}
                  {' '}{t("What would you add?")}{' '}</button>
              )}
            </div>
          </div>
        )}
      </div>

      <div className="arch-chat-composer">
        {!configured && (
          <div className="arch-chat-warning">
            <AlertCircle size={14} /> {' '}
            {localize(language, {
              en: 'No AI model is configured.',
              ja: 'AI モデルが設定されていません。',
            })}{' '}
          </div>
        )}
        <div className="arch-chat-input-row">
          <textarea
            ref={inputRef}
            className="arch-chat-input"
            placeholder={hasDiagram ? t("e.g. add a load balancer in front of the VMs") : t("Describe your architecture…")}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            rows={2}
            disabled={isSending || !configured}
          />
          <button
            className="arch-chat-send"
            onClick={() => send(input)}
            disabled={isSending || !configured || !input.trim()}
            title={t("Send (Enter)")}
            aria-label={t("Send")}
          >
            {isSending ? <Loader2 size={18} className="spin" /> : <Send size={18} />}
          </button>
        </div>
        <div className="arch-chat-hint">{t("Enter to send · Shift+Enter for a new line · each change is auto-saved to version history")}</div>
      </div>
    </ResponsiveDrawer>
  );
};

export default ArchitectureChatPanel;
