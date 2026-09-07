// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useRef, useEffect, useCallback } from 'react';
import { Sparkles, X, Send, Loader2, AlertCircle, MessageSquare, ChevronDown, ChevronUp, Shield, Activity, DollarSign, Wrench, Zap, Lightbulb, type LucideIcon } from 'lucide-react';
import { generateArchitectureWithAI, generateFollowUpSuggestions, isAzureOpenAIConfigured, throwIfGenerationAborted } from '../services/azureOpenAI';
import { useModelSettings } from '../stores/modelSettingsStore';
import { captureRuntimeModelOverride, getEffectiveAIModelInfo, type RuntimeModelOverride } from '../services/aiModelRuntime';
import { useBYOAISettings } from '../stores/byoAISettingsStore';
import { useRuntimeConfig } from '../services/runtimeConfig';
import AIConnectionSelector from './AIConnectionSelector';
import {
  buildModificationPrompt,
  architectureFingerprint,
  summarizeArchitectureChange,
  CurrentArchitecture,
} from '../services/modificationPrompt';
import './ArchitectureChatPanel.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';
import { OperationGeneration } from '../utils/operationGeneration';
import { readLocalStorage, writeLocalStorage } from '../utils/safeStorage';
import ResponsiveDrawer from './ResponsiveDrawer';
import { useMediaQuery } from '../hooks/useMediaQuery';
import { MEDIA_QUERIES } from '../styles/breakpoints';

interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'error';
  text: string;
  ts: number;
  submittedModel?: string;
}

interface ArchitectureChatPanelProps {
  isOpen: boolean;
  onClose: () => void;
  onConfigureConnections?: () => void;
  currentArchitecture: CurrentArchitecture;
  diagramKey: string;
  /** Applies a generated architecture to the canvas (App's handleAIGenerate). */
  onApply: (architecture: any, prompt: string, autoSnapshot?: boolean, baseRevision?: number, signal?: AbortSignal) => boolean | void | Promise<boolean | void>;
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
  onConfigureConnections,
  currentArchitecture,
  diagramKey,
  onApply,
}) => {
  const { t, translate, language } = useLanguage();
  const isCompactWorkspace = useMediaQuery(MEDIA_QUERIES.workspace);
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
  // Tier 3: AI-generated, change-specific follow-ups, keyed to the assistant
  // turn they were generated for. Null until (and unless) they arrive.
  const [modelFollowUps, setModelFollowUps] = useState<{ forMsgId: string; items: string[] } | null>(null);
  // Tier 4: loading flags for the background follow-up fetch and the
  // "What would you add?" single-best-recommendation button.
  const [followUpsLoading, setFollowUpsLoading] = useState(false);
  const [askingBest, setAskingBest] = useState(false);
  const [canRetry, setCanRetry] = useState(false);
  useModelSettings();
  const { storageError } = useBYOAISettings();
  useRuntimeConfig();
  const [submittedModel, setSubmittedModel] = useState('');

  const threadRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const latestFollowUpRequestRef = useRef<string | null>(null);
  const diagramKeyRef = useRef(diagramKey);
  const sendGenerationRef = useRef(new OperationGeneration());
  const bestSuggestionGenerationRef = useRef(new OperationGeneration());
  const requestRef = useRef<AbortController | null>(null);
  const followUpControllerRef = useRef<AbortController | null>(null);
  const lastPromptRef = useRef('');
  const architectureRef = useRef(currentArchitecture);
  architectureRef.current = currentArchitecture;
  const focusOpenerRef = useRef<Element | null>(null);
  const openRef = useRef(false);
  if (isOpen && !openRef.current && typeof document !== 'undefined') {
    focusOpenerRef.current = document.activeElement;
  }
  openRef.current = isOpen;
  const cancelRequest = useCallback(() => {
    sendGenerationRef.current.advance();
    bestSuggestionGenerationRef.current.advance();
    const active = requestRef.current;
    requestRef.current = null;
    active?.abort();
    followUpControllerRef.current?.abort();
    followUpControllerRef.current = null;
    latestFollowUpRequestRef.current = null;
    setIsSending(false);
    setAskingBest(false);
    setFollowUpsLoading(false);
    if (active && lastPromptRef.current) setCanRetry(true);
  }, []);
  useEffect(() => {
    if (!isOpen) cancelRequest();
  }, [isOpen, cancelRequest]);
  useEffect(() => () => {
    requestRef.current?.abort();
    requestRef.current = null;
    followUpControllerRef.current?.abort();
    latestFollowUpRequestRef.current = null;
  }, []);

  diagramKeyRef.current = diagramKey;

  const configured = isAzureOpenAIConfigured();
  const hasDiagram = currentArchitecture.nodes.some((n) => n.type === 'azureNode');
  const connectionInfo = getEffectiveAIModelInfo('architectureGeneration');
  const modelName = connectionInfo.displayName;

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
    cancelRequest();
    latestFollowUpRequestRef.current = null;
    setMessages([]);
    setInput('');
    setIsSending(false);
    setShowAdvanced(false);
    setUsedSuggestions(new Set());
    setModelFollowUps(null);
    setFollowUpsLoading(false);
    setAskingBest(false);
    setCanRetry(false);
    lastPromptRef.current = '';
  }, [diagramKey, cancelRequest]);

  // Focus the composer when the panel opens.
  useEffect(() => {
    if (isOpen) {
      const opener = focusOpenerRef.current;
      const t = setTimeout(() => {
        if (document.activeElement === opener && !document.querySelector('[aria-modal="true"]')) {
          inputRef.current?.focus({ preventScroll: true });
        }
      }, 120);
      return () => clearTimeout(t);
    }
  }, [isOpen]);

  const send = useCallback(
    async (raw: string, context = currentArchitecture, continuation?: AbortController, capturedModel?: RuntimeModelOverride) => {
      const text = raw.trim();
      if (!openRef.current || !text || (requestRef.current && requestRef.current !== continuation)) return;
      const controller = continuation ?? new AbortController();
      if (controller.signal.aborted) return;
      requestRef.current = controller;
      const requestGeneration = sendGenerationRef.current.advance();
      const requestDiagramKey = diagramKey;
      const active = () => openRef.current && requestRef.current === controller && !controller.signal.aborted
        && diagramKeyRef.current === requestDiagramKey
        && sendGenerationRef.current.isCurrent(requestGeneration);

      setInput(text);
      lastPromptRef.current = text;
      setCanRetry(false);
      followUpControllerRef.current?.abort();
      latestFollowUpRequestRef.current = null;
      setModelFollowUps(null);
      setFollowUpsLoading(false);
      const userMsg: ChatMessage = { id: uid(), role: 'user', text, ts: Date.now() };
      setMessages((prev) => [...prev, userMsg]);
      setIsSending(true);

      // Snapshot the canvas state BEFORE applying so we can diff for a summary.
      const before: CurrentArchitecture = {
        nodes: context.nodes,
        edges: context.edges,
        architectureName: context.architectureName,
        revision: context.revision,
      };
      const baselineFingerprint = architectureFingerprint(before);

      // Recent user instructions help the model resolve references.
      const recentRequests = [...messages, userMsg]
        .filter((m) => m.role === 'user')
        .slice(-5)
        .map((m) => m.text);

      try {
        const modelOverride = capturedModel ?? captureRuntimeModelOverride('architectureGeneration');
        const submittedName = modelOverride.connection?.displayName
          ?? getEffectiveAIModelInfo('architectureGeneration').displayName;
        setSubmittedModel(submittedName);
        const prompt = buildModificationPrompt(before, text, recentRequests.slice(0, -1), language);
        const result = await generateArchitectureWithAI(prompt, modelOverride, undefined, language, { signal: controller.signal });
        if (!active()) return;
        if (architectureFingerprint(architectureRef.current) !== baselineFingerprint) {
          throw new Error(localize(language, {
            en: 'The diagram changed while this request was running. Your edits were preserved. Review the request and try again.',
            ja: 'リクエストの実行中に図が変更されました。編集は保持されています。内容を確認して再試行してください。',
          }));
        }

        const applied = await onApply(result, text, true, before.revision, controller.signal);
        if (!active()) return;
        if (applied === false) {
          setCanRetry(true);
          return;
        }
        // Let the accepted subset reach props; the proposal may contain
        // changes the user explicitly rejected in the review dialog.
        await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
        if (!active()) return;
        setInput('');

        const acceptedServices = architectureRef.current.nodes
          .filter(node => node.type === 'azureNode')
          .map(node => ({ name: String(node.data?.label || node.data?.serviceName || '').trim() }))
          .filter(service => service.name);
        const summary = summarizeArchitectureChange(before, { services: acceptedServices }, language);
        const asstId = uid();
        setMessages((prev) => [
          ...prev,
          { id: asstId, role: 'assistant', text: summary, ts: Date.now(), submittedModel: submittedName },
        ]);

        // Tier 3: fetch change-specific follow-ups in the background (non-blocking).
        // The static rule-based chips render immediately; these replace them when
        // they arrive. Use the applied subset, never the unreviewed proposal.
        const nextServices = acceptedServices.map(service => service.name);
        setModelFollowUps(null);
        setFollowUpsLoading(true);
        latestFollowUpRequestRef.current = asstId;
        const followUpController = new AbortController();
        followUpControllerRef.current = followUpController;
        void generateFollowUpSuggestions({ services: nextServices, lastChange: summary, recentRequests, language, signal: followUpController.signal, modelOverride })
          .then((items) => {
            if (
              openRef.current && !followUpController.signal.aborted
              && diagramKeyRef.current === requestDiagramKey
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
              openRef.current && !followUpController.signal.aborted
              && diagramKeyRef.current === requestDiagramKey
              && sendGenerationRef.current.isCurrent(requestGeneration)
              && latestFollowUpRequestRef.current === asstId
            ) {
              setFollowUpsLoading(false);
            }
          });
      } catch (err: any) {
        if (!active()) return;
        setCanRetry(true);
        if (err?.name === 'AbortError' || err?.name === 'CloudDiagramOperationCancelledError') return;
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
        if (requestRef.current === controller) {
          requestRef.current = null;
          setIsSending(false);
        }
      }
    },
    [diagramKey, messages, currentArchitecture, onApply, language, translate],
  );

  // Tier 4: "What would you add?" — ask the model for the single highest-impact
  // next step (from the current canvas) and apply it like a chip click.
  const handleAskBest = async () => {
    if (!openRef.current || requestRef.current || !configured) return;
    const controller = new AbortController();
    requestRef.current = controller;
    const before = { ...currentArchitecture };
    const baselineFingerprint = architectureFingerprint(before);
    const requestGeneration = bestSuggestionGenerationRef.current.advance();
    const requestDiagramKey = diagramKey;
    setAskingBest(true);
    try {
      const modelOverride = captureRuntimeModelOverride('architectureGeneration');
      setSubmittedModel(modelOverride.connection?.displayName
        ?? getEffectiveAIModelInfo('architectureGeneration').displayName);
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
        signal: controller.signal,
        modelOverride,
      });
      throwIfGenerationAborted(controller.signal);
      if (requestRef.current !== controller) return;
      if (
        diagramKeyRef.current !== requestDiagramKey
        || !bestSuggestionGenerationRef.current.isCurrent(requestGeneration)
      ) return;
      if (architectureFingerprint(architectureRef.current) !== baselineFingerprint) {
        throw new Error('The diagram changed while finding a suggestion.');
      }
      if (best[0]) {
        markUsed(best[0]);
        await send(best[0], before, controller, modelOverride);
      }
    } catch {
      if (!controller.signal.aborted && requestRef.current === controller) {
        setMessages(previous => [...previous, {
          id: uid(), role: 'error', ts: Date.now(),
          text: localize(language, { en: 'Could not suggest a change. Please try again.', ja: '変更を提案できませんでした。もう一度お試しください。' }),
        }]);
      }
    } finally {
      if (requestRef.current === controller) requestRef.current = null;
      if (
        !controller.signal.aborted
        && openRef.current
        &&
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
        modal={isCompactWorkspace}
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
        <button className="arch-chat-close" onClick={() => { cancelRequest(); onClose(); }} title={t("Close chat")} aria-label={t("Close chat")}>
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

      <AIConnectionSelector compact disabled={isSending || askingBest}
        onConfigureConnections={onConfigureConnections} />

      <div className="arch-chat-thread" ref={threadRef}>
        {messages.length === 0 && (
          <div className="arch-chat-empty">
            <p className="arch-chat-empty-title">
              {hasDiagram
                ? localize(language, { en: 'Describe a change and review the proposal before applying it.', ja: '変更内容を入力し、提案を確認してから図に適用できます。' })
                : localize(language, { en: 'Start by describing what you want to build — I’ll draw the first version, then we refine it together.', ja: '作成したい内容を入力してください。最初の図を作成し、その後一緒に調整できます。' })}
            </p>
            <p className="arch-chat-empty-sub">
              {hasDiagram
                ? localize(language, { en: 'Choose which changes to accept. Accepted edits can be undone.', ja: '適用する変更を選択できます。適用した編集は元に戻せます。' })
                : localize(language, { en: 'Pick a starter below or type your own. Review the generated diagram before applying it.', ja: '下の例を選ぶか、要件を入力してください。生成された図を確認してから適用できます。' })}
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
            <div className="arch-chat-bubble">{m.text}
              {m.submittedModel && <small className="arch-chat-provenance">
                {localize(language, { en: 'Submitted with', ja: '送信時のモデル' })}: {m.submittedModel}
              </small>}
            </div>
          </div>
        ))}

        {isSending && (
          <div className="arch-chat-msg arch-chat-msg-assistant">
            <Loader2 size={15} className="arch-chat-msg-icon spin" />
            <div className="arch-chat-bubble arch-chat-bubble-pending">{t("Updating the diagram…")}
              {submittedModel && <small className="arch-chat-provenance">{submittedModel}</small>}
            </div>
          </div>
        )}

        {messages.length > 0 && configured && !isSending && !canRetry && (hasDiagram ? followUps.length > 0 : true) && (
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
        {(isSending || askingBest) && <button type="button" className="btn btn-secondary" onClick={cancelRequest}>
          {localize(language, { en: 'Cancel request', ja: 'リクエストをキャンセル' })}
        </button>}
        {canRetry && !isSending && !askingBest && <button type="button" className="btn btn-secondary"
          disabled={!configured} onClick={() => void send(input.trim() || lastPromptRef.current)}>
          {localize(language, { en: 'Retry request', ja: 'リクエストを再試行' })}
        </button>}
        {!configured && !storageError && (
          <div className="arch-chat-warning">
            <AlertCircle size={14} /> {' '}
            {connectionInfo.source === 'bring-your-own' ? localize(language, {
              en: 'The selected BYO connection is unavailable. Enter its key, verify it in AI connections, or explicitly choose managed Astra. Requests will not switch providers automatically.',
              ja: '選択中の BYO 接続は利用できません。AI 接続の設定でキーを入力して確認するか、管理対象の Astra を明示的に選択してください。プロバイダーは自動で切り替わりません。',
            }) : localize(language, {
              en: 'GPT-6 Astra is not configured. Contact the application administrator to configure the managed Astra deployment.',
              ja: 'GPT-6 Astraが設定されていません。管理対象のAstraデプロイを設定するようアプリケーション管理者に連絡してください。',
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
        <div className="arch-chat-hint">{localize(language, {
          en: 'Enter to send · Shift+Enter for a new line · Review changes before applying',
          ja: 'Enterで送信 · Shift+Enterで改行 · 変更は確認してから適用',
        })}</div>
      </div>
    </ResponsiveDrawer>
  );
};

export default ArchitectureChatPanel;
