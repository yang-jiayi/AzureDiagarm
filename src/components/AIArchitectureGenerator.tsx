// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Sparkles, X, Loader2, Clock, Zap, Brain, Network, PenTool, Layers } from 'lucide-react';
import { generateArchitectureWithAI, isAzureOpenAIConfigured, AIMetrics, analyzeArchitectureDiagramImage, ModelOverride } from '../services/azureOpenAI';
import { generateReferenceArchitectureWithAI } from '../services/referenceArchitectureAI';
import { generateBlueprintArchitectureWithAI } from '../services/blueprintArchitectureAI';
import { generateComponentManifest, ComponentManifest } from '../services/componentManifestAI';
import { exportReferenceArchitectureAsPng } from '../utils/exportReferencePng';
import { exportBlueprintArchitectureAsPng } from '../utils/exportBlueprintPng';
import ImageUploader from './ImageUploader';
import {
  useModelSettings,
  MODEL_CONFIG,
  FeatureType,
  getAvailableModels,
  getModelSettingsForFeature,
  ModelType,
  ReasoningEffort,
  FEATURE_CONFIG,
  getReasoningEffortLabel,
  getSupportedReasoningEfforts,
  isModelAvailable,
  updateFeatureOverride,
} from '../stores/modelSettingsStore';
import { trackImageImport } from '../services/telemetryService';
import { buildModificationPrompt } from '../services/modificationPrompt';
import './AIArchitectureGenerator.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';
import { readBooleanPreference, readLocalStorage, writeLocalStorage } from '../utils/safeStorage';

type GenerationMode = 'topology' | 'reference' | 'blueprint' | 'both';

// Blueprint diagrams require general-purpose OpenAI models.
// - Non-OpenAI partner deployments (DeepSeek, Grok, Mistral, Kimi, etc. —
//   identified by apiFormat: 'chat-completions') run under stricter Azure AI
//   Content Safety configurations that block the blueprint system prompt as
//   adversarial.
// - Codex-tuned variants (e.g. gpt-5.2-codex, gpt-5.3-codex) are optimized for
//   coding tasks and tend to refuse non-code architecture-diagram prompts with
//   "I'm sorry, ..." responses.
const isBlueprintCapableModel = (m: ModelType): boolean =>
  MODEL_CONFIG[m].apiFormat !== 'chat-completions' && !m.includes('codex');
const modeRequiresOpenAI = (m: GenerationMode): boolean =>
  m === 'blueprint' || m === 'both';

interface PromptCategory {
  category: LocalizedText;
  color: string;
  prompts: LocalizedText[];
}

const CATEGORIZED_PROMPTS: PromptCategory[] = [
  {
    category: { en: 'Web & Microservices', ja: 'Webとマイクロサービス' },
    color: '#3b82f6',
    prompts: [
      {
        en: 'A web application with a React frontend, Node.js backend API, PostgreSQL database, and blob storage for images',
        ja: 'Reactフロントエンド、Node.jsバックエンドAPI、PostgreSQLデータベース、画像用Blob Storageを使用するWebアプリ',
      },
      {
        en: 'A microservices architecture with container apps, API gateway, message queue, and Redis cache',
        ja: 'Container Apps、API Gateway、メッセージ キュー、Redis Cacheを使用するマイクロサービス アーキテクチャ',
      },
    ],
  },
  {
    category: { en: 'Security & Networking', ja: 'セキュリティとネットワーク' },
    color: '#ef4444',
    prompts: [
      {
        en: 'A zero trust enterprise network with Azure Firewall, Application Gateway with WAF, Private Link for PaaS services, Bastion for VM access, Microsoft Entra ID with Conditional Access, and Microsoft Defender for Cloud — segmented into DMZ, application, and data tiers',
        ja: 'Azure Firewall、WAF付きApplication Gateway、PaaSサービス向けPrivate Link、VMアクセス向けBastion、Conditional Access付きMicrosoft Entra ID、Microsoft Defender for Cloudを使用し、DMZ、アプリケーション、データの各層に分離したZero Trustエンタープライズ ネットワーク',
      },
      {
        en: 'A security operations center architecture with Microsoft Sentinel for SIEM, Log Analytics, Microsoft Defender for Cloud, Azure Key Vault, Azure Monitor, automation playbooks with Logic Apps, and integration with Microsoft Entra ID for identity threat detection',
        ja: 'SIEMとしてMicrosoft Sentinelを使用し、Log Analytics、Microsoft Defender for Cloud、Azure Key Vault、Azure Monitor、Logic Appsによる自動化プレイブック、ID脅威検出向けMicrosoft Entra ID連携を備えたセキュリティ運用センター（SOC）アーキテクチャ',
      },
    ],
  },
  {
    category: { en: 'AI & Cognitive', ja: 'AIとCognitive Services' },
    color: '#8b5cf6',
    prompts: [
      {
        en: 'A machine learning pipeline with data ingestion, training, and inference endpoints',
        ja: 'データ取り込み、学習、推論エンドポイントを含む機械学習パイプライン',
      },
      {
        en: 'An intelligent customer service chatbot using Azure OpenAI for conversations, Language for sentiment analysis, Speech Services for voice input/output, and Translator for multi-language support, with chat history in Cosmos DB and API Management for external access',
        ja: '会話にAzure OpenAI、感情分析にLanguage、音声入出力にSpeech Services、多言語対応にTranslatorを使用するインテリジェントなカスタマーサービス チャットボット。チャット履歴をCosmos DBに保存し、API Managementで外部公開する',
      },
      {
        en: 'A smart document processing platform that uses Computer Vision to analyze uploaded images, Document Intelligence to extract form data, Language to classify and summarize content, all coordinated through Azure Functions with results stored in Cosmos DB and searchable via Cognitive Search',
        ja: 'Computer Visionで画像を分析し、Document Intelligenceでフォーム データを抽出し、Languageで分類と要約を行うドキュメント処理プラットフォーム。Azure Functionsで処理を連携し、結果をCosmos DBに保存してCognitive Searchで検索可能にする',
      },
      {
        en: 'A content moderation system for social media using Computer Vision to scan images, Language for text analysis and content safety checks, Azure OpenAI for context understanding, with real-time processing via Event Hubs and results stored in SQL Database with API Management exposing moderation APIs',
        ja: 'Computer Visionによる画像スキャン、Languageによるテキスト分析とContent Safetyチェック、Azure OpenAIによる文脈理解を行うソーシャル メディア向けコンテンツ モデレーション システム。Event Hubsでリアルタイム処理し、結果をSQL Databaseへ保存してAPI ManagementでモデレーションAPIを公開する',
      },
    ],
  },
  {
    category: { en: 'E-commerce', ja: 'Eコマース' },
    color: '#f59e0b',
    prompts: [
      {
        en: 'A Black Friday-ready e-commerce platform handling 50,000 orders/hour peak with real-time inventory sync across 12 regional warehouses, ML-powered fraud detection scoring each transaction in under 200ms, personalized recommendations engine, multi-currency payment processing with PCI-DSS compliance, abandoned cart recovery workflows, using Azure Kubernetes Service for microservices, Cosmos DB for product catalog with global distribution, Redis Cache for session and cart state, Service Bus for order orchestration, Azure Functions for inventory webhooks, Cognitive Search for faceted product search, and CDN with dynamic site acceleration',
        ja: 'Black Fridayのピーク時に毎時50,000件の注文を処理し、12地域の倉庫間でリアルタイム在庫同期、200ms未満のML不正検知、パーソナライズ推薦、PCI-DSS準拠の多通貨決済、カート放棄回復を実行するEコマース プラットフォーム。マイクロサービスにAzure Kubernetes Service、グローバル商品カタログにCosmos DB、セッションとカート状態にRedis Cache、注文オーケストレーションにService Bus、在庫WebhookにAzure Functions、ファセット商品検索にCognitive Search、動的サイト高速化にCDNを使用する',
      },
    ],
  },
  {
    category: { en: 'Healthcare', ja: '医療' },
    color: '#22c55e',
    prompts: [
      {
        en: 'A HIPAA-compliant healthcare data platform integrating EHR systems via HL7 FHIR R4 APIs, medical imaging PACS with DICOM support storing 500TB of radiology images, real-time clinical decision support, patient portal with secure messaging, audit logging for all PHI access, disaster recovery with 15-minute RPO, using Azure API for FHIR, Azure Health Data Services for DICOM, Blob Storage with immutable retention for images, Cosmos DB for patient timelines, Azure Functions for HL7v2 to FHIR transformation, Logic Apps for clinical workflows, Key Vault for encryption key management, and Microsoft Defender for Cloud for continuous compliance monitoring',
        ja: 'HL7 FHIR R4 APIによるEHR連携、DICOM対応PACSへの500TBの放射線画像保存、リアルタイム臨床意思決定支援、安全なメッセージ機能付き患者ポータル、全PHIアクセスの監査ログ、RPO 15分の災害復旧を備えたHIPAA準拠の医療データ プラットフォーム。Azure API for FHIR、Azure Health Data Services for DICOM、イミュータブル保持付きBlob Storage、患者タイムライン用Cosmos DB、HL7v2からFHIRへの変換用Azure Functions、臨床ワークフロー用Logic Apps、暗号鍵管理用Key Vault、継続的なコンプライアンス監視用Microsoft Defender for Cloudを使用する',
      },
      {
        en: 'An eventing architecture for healthcare imaging with high throughput (50,000-75,000 events/sec), large payloads up to 10MB, strict message ordering, cloud-to-on-prem bridging via VPN Gateway, managed services only (no self-managed Kafka), 99.99% availability SLO, supporting 250M studies, 2.5M daily volume, 5M daily notifications, with Event Hubs for ingestion, Service Bus for routing, Azure Functions for processing, Cosmos DB for metadata, Blob Storage for images, and Log Analytics for monitoring',
        ja: '毎秒50,000～75,000件のイベント、最大10MBのペイロード、厳密なメッセージ順序、VPN Gatewayによるクラウドとオンプレミスの接続、マネージド サービスのみの構成、99.99%の可用性SLOに対応し、2.5億件の検査、日次250万件の処理、日次500万件の通知を扱う医療画像イベント アーキテクチャ。取り込みにEvent Hubs、ルーティングにService Bus、処理にAzure Functions、メタデータにCosmos DB、画像にBlob Storage、監視にLog Analyticsを使用する',
      },
    ],
  },
  {
    category: { en: 'Data & Analytics', ja: 'データと分析' },
    color: '#06b6d4',
    prompts: [
      {
        en: 'A data lakehouse with Azure Data Lake Storage, Synapse Analytics for SQL and Spark queries, Data Factory for ETL pipelines, and Power BI for dashboards',
        ja: 'Azure Data Lake Storage、SQLとSparkクエリ用Synapse Analytics、ETLパイプライン用Data Factory、ダッシュボード用Power BIを使用するデータ レイクハウス',
      },
      {
        en: 'A real-time analytics pipeline using Event Hubs for ingestion, Stream Analytics for windowed aggregations, Cosmos DB for serving layer, and Azure Monitor for pipeline health',
        ja: '取り込みにEvent Hubs、ウィンドウ集計にStream Analytics、配信層にCosmos DB、パイプラインの正常性監視にAzure Monitorを使用するリアルタイム分析パイプライン',
      },
      {
        en: 'A data warehouse with Azure SQL Database, Data Factory for scheduled imports from multiple sources, Purview for data governance and cataloging, and Power BI embedded reports',
        ja: 'Azure SQL Database、複数のデータ ソースからの定期取り込み用Data Factory、データ ガバナンスとカタログ用Purview、Power BI Embeddedレポートを使用するデータ ウェアハウス',
      },
    ],
  },
  {
    category: { en: 'Microsoft Fabric', ja: 'Microsoft Fabric' },
    color: '#0d9488',
    prompts: [
      {
        en: 'A Microsoft Fabric medallion lakehouse: Data Factory ingestion into OneLake, Bronze/Silver/Gold Lakehouses processed with Fabric Notebooks and Dataflow Gen2, a Warehouse for curated marts, and a Power BI Report via a Direct Lake Semantic Model, running on a Fabric Capacity F2',
        ja: 'Microsoft FabricのメダリオンLakehouse。Data FactoryでOneLakeへ取り込み、Fabric NotebookとDataflow Gen2でBronze、Silver、Gold Lakehouseを処理し、整備済みデータ マートにWarehouseを使用する。Direct Lake Semantic Model経由でPower BI Reportを提供し、Fabric Capacity F2で実行する',
      },
      {
        en: 'An end-to-end Microsoft Fabric analytics platform: ingest on-prem SQL via a Fabric Data Pipeline and Mirrored Database, stream IoT telemetry through an Eventstream into an Eventhouse with a KQL Database, land data in OneLake, build Bronze/Silver/Gold Lakehouses, expose a Semantic Model to a Power BI Report and a Real-Time Dashboard, and add a Fabric Data Agent for natural-language Q&A — on a Fabric Capacity F64',
        ja: 'エンドツーエンドのMicrosoft Fabric分析プラットフォーム。Fabric Data PipelineとMirrored DatabaseでオンプレミスSQLを取り込み、IoTテレメトリをEventstreamからKQL Databaseを持つEventhouseへストリーミングし、OneLakeへ保存する。Bronze、Silver、Gold Lakehouseを構築し、Semantic ModelをPower BI ReportとReal-Time Dashboardへ公開し、自然言語Q&A用Fabric Data Agentを追加してFabric Capacity F64で実行する',
      },
      {
        en: 'A real-time intelligence solution in Microsoft Fabric: Eventstream ingestion into an Eventhouse and KQL Database, a Real-Time Dashboard for live KPIs, and a Lakehouse plus Power BI Report for historical analysis, on a Fabric Capacity',
        ja: 'Microsoft FabricのReal-Time Intelligenceソリューション。EventstreamからEventhouseとKQL Databaseへ取り込み、ライブKPI用Real-Time Dashboard、履歴分析用LakehouseとPower BI ReportをFabric Capacity上で実行する',
      },
    ],
  },
  {
    category: { en: 'IoT', ja: 'IoT' },
    color: '#14b8a6',
    prompts: [
      {
        en: 'An industrial IoT predictive maintenance platform for a manufacturing facility with 5,000+ sensors generating telemetry every 5 seconds, requiring real-time anomaly detection with sub-second latency, batch analytics for trend analysis, secure device provisioning and management, OT/IT network segregation with Private Link, 99.9% uptime SLA, 6-month hot storage and 7-year cold retention, using IoT Hub for ingestion, Stream Analytics for real-time processing, Azure ML for predictive models, Data Lake for raw storage, Synapse Analytics for reporting, Time Series Insights for dashboards, and Digital Twins for facility modeling',
        ja: '5,000台以上のセンサーが5秒ごとにテレメトリを生成する製造施設向け産業IoT予知保全プラットフォーム。1秒未満のリアルタイム異常検知、傾向分析用のバッチ分析、安全なデバイス プロビジョニングと管理、Private LinkによるOT/ITネットワーク分離、99.9%の稼働率SLA、6か月のホット ストレージと7年のコールド保持を実現する。取り込みにIoT Hub、リアルタイム処理にStream Analytics、予測モデルにAzure ML、生データ保存にData Lake、レポートにSynapse Analytics、ダッシュボードにTime Series Insights、施設モデルにDigital Twinsを使用する',
      },
    ],
  },
];

interface AIArchitectureGeneratorProps {
  onGenerate: (architecture: any, prompt: string, autoSnapshot: boolean, referenceImageUrl?: string) => void | Promise<void>;
  /**
   * Called when a Reference Architecture has been generated. Reference mode
   * intentionally does NOT push a topology onto the canvas (the transformed
   * topology is low-fidelity and confuses users); the PNG is the deliverable.
   * App uses this to stash the ref so the toolbar can re-export the PNG.
   */
  onReferenceArchitecture?: (ref: any) => void;
  /**
   * Called when a Blueprint Architecture has been generated. Like reference
   * mode, blueprint mode does NOT push a topology onto the canvas; the PNG is
   * the deliverable. App stashes the blueprint so the toolbar can re-export.
   */
  onBlueprintArchitecture?: (bp: any) => void;
  currentArchitecture?: {
    nodes: any[];
    edges: any[];
    architectureName: string;
  };
}

const AIArchitectureGenerator: React.FC<AIArchitectureGeneratorProps> = ({ onGenerate, onReferenceArchitecture, onBlueprintArchitecture, currentArchitecture }) => {
  const { t, translate, language } = useLanguage();
  const [isOpen, setIsOpen] = useState(false);
  const [description, setDescription] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [error, setError] = useState('');
  const [aiMetrics, setAiMetrics] = useState<AIMetrics | null>(null);
  const [isAnalyzingImage, setIsAnalyzingImage] = useState(false);
  const [imageAnalyzed, setImageAnalyzed] = useState(false);
  const [uploadedImageUrl, setUploadedImageUrl] = useState<string | null>(null);
  const [mode, setMode] = useState<GenerationMode>(() => {
    const saved = readLocalStorage('aiGenerator.mode');
    if (saved === 'blueprint' || saved === 'both') return saved;
    // Reference mode is hidden in the UI; migrate any stale persisted value.
    if (saved === 'reference') return 'blueprint';
    return 'topology';
  });

  // When mode === 'both', run topology + blueprint generations in parallel
  // (default) or sequentially. Persisted.
  const [bothInParallel, setBothInParallel] = useState<boolean>(() => {
    return readBooleanPreference('aiGenerator.bothInParallel', true);
  });
  const handleBothInParallelChange = (checked: boolean) => {
    setBothInParallel(checked);
    writeLocalStorage('aiGenerator.bothInParallel', JSON.stringify(checked));
  };

  const handleModeChange = (m: GenerationMode) => {
    setMode(m);
    writeLocalStorage('aiGenerator.mode', m);
  };


  // Opt-in: also download an editorial PNG when generating in reference mode.
  // Model settings from reactive hook (stays in sync with dropdown)
  const [modelSettings] = useModelSettings();

  // Blueprint generation requires a general-purpose OpenAI deployment.
  // Keep the architecture model untouched and correct only the blueprint override.
  useEffect(() => {
    if (!modeRequiresOpenAI(mode)) return;
    const blueprintSettings = getModelSettingsForFeature('blueprint');
    if (isBlueprintCapableModel(blueprintSettings.model)) return;
    const fallback = getAvailableModels().find(isBlueprintCapableModel);
    if (!fallback) return;
    const cfg = MODEL_CONFIG[fallback];
    updateFeatureOverride('blueprint', {
      model: fallback,
      reasoningEffort: cfg.isReasoning
        ? (cfg.defaultReasoningEffort ?? blueprintSettings.reasoningEffort)
        : undefined,
    });
  }, [mode, modelSettings.featureOverrides, modelSettings.model, modelSettings.reasoningEffort]);
  
  // Auto-snapshot preference (stored in localStorage)
  const [autoSnapshot, setAutoSnapshot] = useState<boolean>(() => {
    return readBooleanPreference('aiGenerator.autoSnapshot', true);
  });

  // Blueprint legend position preference (stored in localStorage). 'auto'
  // picks bottom vs right based on aspect ratio.
  const [legendPosition, setLegendPosition] = useState<'auto' | 'bottom' | 'right'>(() => {
    const saved = readLocalStorage('aiGenerator.blueprintLegendPosition');
    if (saved === 'bottom' || saved === 'right' || saved === 'auto') return saved;
    return 'auto';
  });
  const handleLegendPositionChange = (v: 'auto' | 'bottom' | 'right') => {
    setLegendPosition(v);
    writeLocalStorage('aiGenerator.blueprintLegendPosition', v);
  };

  // Save preference to localStorage when it changes
  const handleAutoSnapshotChange = (checked: boolean) => {
    setAutoSnapshot(checked);
    writeLocalStorage('aiGenerator.autoSnapshot', JSON.stringify(checked));
  };

  const closeTimerRef = useRef<number | null>(null);
  const cancelScheduledClose = useCallback(() => {
    if (closeTimerRef.current !== null) {
      window.clearTimeout(closeTimerRef.current);
      closeTimerRef.current = null;
    }
  }, []);
  const closeModal = useCallback(() => {
    cancelScheduledClose();
    setIsOpen(false);
  }, [cancelScheduledClose]);
  const scheduleClose = useCallback(() => {
    cancelScheduledClose();
    closeTimerRef.current = window.setTimeout(() => {
      closeTimerRef.current = null;
      setIsOpen(false);
      setAiMetrics(null);
      setUploadedImageUrl(null);
    }, 45_000);
  }, [cancelScheduledClose]);

  useEffect(() => cancelScheduledClose, [cancelScheduledClose]);

  // Handle image analysis result
  const handleImageAnalyzed = (analyzedDescription: string) => {
    // Prepend or replace the description with the analyzed content
    const prefix = localize(language, {
      en: '🖼️ [Analyzed from uploaded diagram]\n\n',
      ja: '🖼️ [アップロードした図から分析]\n\n',
    });
    setDescription(prefix + analyzedDescription);
    setImageAnalyzed(true);
    trackImageImport();
  };

  // Wrapper to pass to ImageUploader
  const handleAnalyzeImage = async (base64: string, mimeType: string) => {
    const result = await analyzeArchitectureDiagramImage(base64, mimeType, language);
    return { description: result.description };
  };

  const categorizedPrompts = CATEGORIZED_PROMPTS.map(group => ({
    category: localize(language, group.category),
    color: group.color,
    prompts: group.prompts.map(prompt => localize(language, prompt)),
  }));

  const handleGenerate = async () => {
    if (!description.trim()) {
      setError(translate('Please describe your architecture'));
      return;
    }

    if (!isAzureOpenAIConfigured()) {
      setError(translate('Azure OpenAI is not configured. Please check your environment variables.'));
      return;
    }

    setIsGenerating(true);
    setError('');
    setAiMetrics(null); // Clear previous metrics
    
    const currentModelSettings: ModelOverride = getModelSettingsForFeature('architectureGeneration');
    console.log(`🎯 Generate clicked: default model=${modelSettings.model}, effective model=${currentModelSettings.model}, reasoning=${currentModelSettings.reasoningEffort}, overrides=${JSON.stringify(modelSettings.featureOverrides)}`);

    // Use the same effective feature setting shown in the modal. If a stale
    // setting points to an incompatible model, fall back to the deployed
    // recommendation so blueprint generation still succeeds.
    const blueprintModelSettings: ModelOverride = (() => {
      const configured = getModelSettingsForFeature('blueprint');
      if (isBlueprintCapableModel(configured.model)) {
        return {
          model: configured.model,
          reasoningEffort: configured.reasoningEffort,
        };
      }
      const rec = FEATURE_CONFIG.blueprint.recommendedModel;
      if (isModelAvailable(rec) && isBlueprintCapableModel(rec)) {
        const cfg = MODEL_CONFIG[rec];
        return {
          model: rec,
          reasoningEffort: cfg.isReasoning
            ? (FEATURE_CONFIG.blueprint.recommendedReasoning || modelSettings.reasoningEffort)
            : modelSettings.reasoningEffort,
        };
      }
      return currentModelSettings;
    })();
    console.log(`📐 Blueprint model: ${blueprintModelSettings.model} (reasoning=${blueprintModelSettings.reasoningEffort})`);

    try {
      // ── Reference (Editorial) mode — PNG is the sole deliverable.
      // We deliberately do NOT push a topology onto the canvas: the
      // transformed network-flow view is low fidelity for editorial inputs
      // and confuses users. Instead we notify App so it can enable the
      // toolbar “Export Editorial PNG” action, then render + download.
      if (mode === 'reference') {
        const ref = await generateReferenceArchitectureWithAI(description, currentModelSettings, language);
        if (ref.metrics) setAiMetrics(ref.metrics);

        // Stash the ref for the toolbar re-export button (if App provided it).
        onReferenceArchitecture?.(ref);

        // Always export the PNG — it is the only artifact produced in this mode.
        try {
          await exportReferenceArchitectureAsPng(ref);
        } catch (err) {
          console.warn('Reference architecture PNG export failed:', err);
          setError(translate('PNG export failed. See console for details.'));
        }

        setDescription('');
        scheduleClose();
        return;
      }

      // ── Blueprint (Whiteboard) mode — PNG is the sole deliverable.
      // Hand-drawn / sketchnote-style nested zones with numbered, labeled
      // arrows. Like reference mode, we do not touch the ReactFlow canvas.
      if (mode === 'blueprint') {
        const bp = await generateBlueprintArchitectureWithAI(
          description,
          blueprintModelSettings,
          undefined,
          language,
        );
        if (bp.metrics) setAiMetrics(bp.metrics);

        onBlueprintArchitecture?.(bp);

        try {
          await exportBlueprintArchitectureAsPng(bp, { legendPosition });
        } catch (err) {
          console.warn('Blueprint architecture PNG export failed:', err);
          setError(translate('PNG export failed. See console for details.'));
        }

        setDescription('');
        scheduleClose();
        return;
      }

      // ── Both mode — generate topology AND blueprint from the same prompt.
      // Topology renders on the canvas (via onGenerate); blueprint is stashed
      // and (when autoSnapshot is on) auto-downloaded as PNG. The toolbar's
      // "Export Blueprint PNG" remains available either way.
      if (mode === 'both') {
        // Build the same enriched context the topology branch uses below.
        let bothContextPrompt = description;
        if (currentArchitecture && currentArchitecture.nodes.length > 0) {
          const groups = currentArchitecture.nodes
            .filter((n) => n.type === 'groupNode')
            .map((n) => ({ name: n.data.label, id: n.id }));
          const groupNameMap = new Map(groups.map((g) => [g.id, g.name]));
          const services = currentArchitecture.nodes
            .filter((n) => n.type === 'azureNode')
            .map((n) => {
              const groupName = n.parentNode ? groupNameMap.get(n.parentNode) : null;
              return { name: n.data.label, group: groupName || null };
            });
          const connections = currentArchitecture.edges.map((e) => {
            const fromNode = currentArchitecture.nodes.find((n) => n.id === e.source);
            const toNode = currentArchitecture.nodes.find((n) => n.id === e.target);
            return `${fromNode?.data.label || e.source} → ${toNode?.data.label || e.target}${e.label ? ` (${e.label})` : ''}`;
          });
          const servicesList = services.map((s) => `${s.name}${s.group ? ` [${s.group}]` : ''}`).join(', ');
          bothContextPrompt = `MODIFY EXISTING ARCHITECTURE: "${currentArchitecture.architectureName}"\nServices: ${servicesList}\n${groups.length > 0 ? `Groups: ${groups.map((g) => g.name).join(', ')}` : ''}\n${connections.length > 0 ? `Connections: ${connections.join('; ')}` : ''}\n\nCHANGE REQUESTED: ${description}\n\nIMPORTANT: Return the COMPLETE architecture JSON (all services, groups, connections, workflow). Keep everything unchanged EXCEPT what the user requested. Only add, modify, or remove what was asked.`;
        }

        const topoCall = (m?: ComponentManifest) =>
          generateArchitectureWithAI(bothContextPrompt, currentModelSettings, m, language);
        const bpCall = (m?: ComponentManifest) =>
          generateBlueprintArchitectureWithAI(bothContextPrompt, blueprintModelSettings, m, language);

        const t0 = performance.now();
        // Pre-pass: extract a canonical component manifest so topology and
        // blueprint agree on the set of services, zones, and on-prem actors.
        let manifest: ComponentManifest | undefined;
        try {
          manifest = await generateComponentManifest(bothContextPrompt, currentModelSettings, language);
          console.log(
            `📋 Manifest: ${manifest.components.length} components across ${manifest.zones.length} zones (${manifest.metrics?.totalTokens ?? '?'} tokens, ${Math.round((manifest.metrics?.elapsedTimeMs ?? 0) / 100) / 10}s)`,
          );
        } catch (err) {
          console.warn('Component manifest pre-pass failed; falling back to independent generation:', err);
          manifest = undefined;
        }

        let topoResult: any = null;
        let bpResult: any = null;
        let topoFailure: unknown = null;
        let bpFailure: unknown = null;
        if (bothInParallel) {
          const [topologyOutcome, blueprintOutcome] = await Promise.allSettled([
            topoCall(manifest),
            bpCall(manifest),
          ]);
          if (topologyOutcome.status === 'fulfilled') topoResult = topologyOutcome.value;
          else topoFailure = topologyOutcome.reason;
          if (blueprintOutcome.status === 'fulfilled') bpResult = blueprintOutcome.value;
          else bpFailure = blueprintOutcome.reason;
        } else {
          try {
            topoResult = await topoCall(manifest);
          } catch (error) {
            topoFailure = error;
          }
          try {
            bpResult = await bpCall(manifest);
          } catch (error) {
            bpFailure = error;
          }
        }
        if (!topoResult && !bpResult) {
          throw topoFailure || bpFailure || new Error('Topology and Blueprint generation failed.');
        }
        const wallElapsed = performance.now() - t0;

        // Combined metrics: sum tokens (including manifest); wall-clock
        // elapsed reflects actual perceived time (manifest + max(topo, bp)
        // for parallel; manifest + topo + bp for sequential).
        const tm = topoResult?.metrics;
        const bm = bpResult?.metrics;
        const mm = manifest?.metrics;
        if (tm || bm || mm) {
          setAiMetrics({
            elapsedTimeMs: Math.round(wallElapsed),
            promptTokens: (tm?.promptTokens || 0) + (bm?.promptTokens || 0) + (mm?.promptTokens || 0),
            completionTokens: (tm?.completionTokens || 0) + (bm?.completionTokens || 0) + (mm?.completionTokens || 0),
            totalTokens: (tm?.totalTokens || 0) + (bm?.totalTokens || 0) + (mm?.totalTokens || 0),
          } as AIMetrics);
        }

        // Push topology to canvas first.
        // Inject the manifest title (when available) so the canvas banner /
        // title block stop reading "Untitled Architecture" after generation.
        if (manifest?.title && topoResult && typeof topoResult === 'object') {
          if (!topoResult.architectureName || /untitled/i.test(String(topoResult.architectureName))) {
            topoResult.architectureName = manifest.title;
          }
        }
        if (topoResult) {
          await onGenerate(topoResult, description, autoSnapshot, uploadedImageUrl || undefined);
        }
        if (bpResult) {
          // Stash blueprint for the toolbar re-export button.
          onBlueprintArchitecture?.(bpResult);
        }

        // Auto-download the blueprint PNG when the user has autoSnapshot on
        // (matches the existing "auto" behavior they're already used to).
        if (autoSnapshot && bpResult) {
          try {
            await exportBlueprintArchitectureAsPng(bpResult, { legendPosition });
          } catch (err) {
            console.warn('Blueprint architecture PNG export failed:', err);
            setError(translate('Blueprint PNG export failed. See console for details.'));
          }
        }

        if (topoFailure || bpFailure) {
          const failedOutput = topoFailure ? 'Topology' : 'Blueprint';
          const cause = topoFailure || bpFailure;
          const detail = cause instanceof Error ? cause.message : String(cause);
          throw new Error(`${failedOutput} generation failed, but the other output was created successfully: ${detail}`);
        }

        setDescription('');
        scheduleClose();
        return;
      }

      // Build context about existing architecture if present
      let contextPrompt = description;
      
      if (currentArchitecture && currentArchitecture.nodes.length > 0) {
        contextPrompt = buildModificationPrompt(currentArchitecture, description, [], language);
      }
      
      // Call Azure OpenAI to generate architecture
      const result = await generateArchitectureWithAI(
        contextPrompt,
        currentModelSettings,
        undefined,
        language,
      );
      
      // Store AI metrics if available
      if (result.metrics) {
        setAiMetrics(result.metrics);
      }
      
      await onGenerate(result, description, autoSnapshot, uploadedImageUrl || undefined);
      setDescription('');
      
      // Close modal shortly after successful generation
      scheduleClose(); // Give user 45 seconds to review results or type a modification
    } catch (err: any) {
      setError(err.message ? translate(err.message) : translate('Failed to generate architecture. Please try again.'));
    } finally {
      setIsGenerating(false);
    }
  };

  const useExample = (example: string) => {
    setDescription(example);
  };

  const renderFeatureModelControls = (feature: FeatureType, openAiOnly: boolean) => {
    const featureSettings = getModelSettingsForFeature(feature);
    const config = MODEL_CONFIG[featureSettings.model];
    const label = mode === 'both'
      ? translate(FEATURE_CONFIG[feature].displayName)
      : t("Model:");

    return (
      <div className="ai-modal-model-group" key={feature}>
        <span className="ai-modal-model-label">{label}</span>
        <select
          className="ai-modal-model-select"
          value={featureSettings.model}
          onChange={(e) => {
            const next = e.target.value as ModelType;
            const nextConfig = MODEL_CONFIG[next];
            updateFeatureOverride(feature, {
              model: next,
              reasoningEffort: nextConfig.isReasoning
                ? (nextConfig.defaultReasoningEffort ?? featureSettings.reasoningEffort)
                : undefined,
            });
          }}
          disabled={isGenerating}
          aria-label={`${translate(FEATURE_CONFIG[feature].displayName)} - ${t("Select AI model")}`}
        >
          {getAvailableModels()
            .filter((model) => !openAiOnly || isBlueprintCapableModel(model))
            .map((model) => (
              <option key={model} value={model}>
                {MODEL_CONFIG[model].displayName}
              </option>
            ))}
        </select>
        {config.isReasoning && (
          <>
            <span className="ai-modal-model-label">{t("Reasoning:")}</span>
            <select
              className="ai-modal-model-select"
              value={featureSettings.reasoningEffort}
              onChange={(e) =>
                updateFeatureOverride(feature, {
                  model: featureSettings.model,
                  reasoningEffort: e.target.value as ReasoningEffort,
                })
              }
              disabled={isGenerating}
              aria-label={`${translate(FEATURE_CONFIG[feature].displayName)} - ${t("Select reasoning effort")}`}
            >
              {getSupportedReasoningEfforts(featureSettings.model).map(level => (
                <option key={level} value={level}>
                  {t(getReasoningEffortLabel(level))}
                </option>
              ))}
            </select>
          </>
        )}
      </div>
    );
  };

  return (
    <>
      <button
        className="btn btn-ai btn-generate-ai"
        onClick={() => {
          cancelScheduledClose();
          setIsOpen(true);
          // Reset state when opening modal
          setError('');
          setImageAnalyzed(false);
        }}
        title={t("Generate architecture with AI")}
      >
        <Sparkles size={18} />
        {' '}{t("Generate with AI")}{' '}</button>

      {isOpen && createPortal(
        <div className="modal-overlay" onClick={closeModal}>
          <div className="modal-content ai-architecture-modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <div className="modal-title">
                <Sparkles size={20} />
                <h2>{t("AI Architecture Generator")}</h2>
              </div>
              <button
                className="modal-close"
                onClick={closeModal}
                title={t("Close")}
                aria-label={t("Close")}
              >
                <X size={20} />
              </button>
            </div>

            <div className="modal-body">
             <div className="modal-body-grid">
              <div className="modal-col modal-col-left">
              <p className="modal-description">
                {mode === 'reference' ? (
                  localize(language, {
                    en: 'Describe the workload in natural language, and AI will generate a publication-style Reference Architecture with Ingest, Process, and Serve stages, a foundation strip, and cross-cutting governance rails in the style of the Azure Architecture Center.',
                    ja: 'ワークロードを自然な言葉で説明すると、AIが取り込み、処理、提供の各ステージ、基盤領域、横断的なガバナンス領域を備えたAzure Architecture Center形式のReference Architectureを生成します。',
                  })
                ) : mode === 'blueprint' ? (
                  localize(language, {
                    en: 'Describe the workload, and AI will sketch a whiteboard-style Blueprint with nested Azure, VNet, and on-premises zones plus numbered, labeled arrows that show the end-to-end flow.',
                    ja: 'ワークロードを説明すると、AIがAzure、VNet、オンプレミスのネストされた領域と、エンドツーエンドのフローを示す番号・ラベル付き矢印を備えたホワイトボード形式のBlueprintを作成します。',
                  })
                ) : mode === 'both' ? (
                  localize(language, {
                    en: 'Generate both an editable, deployable topology on the canvas and a polished whiteboard-style Blueprint PNG from the same prompt.',
                    ja: '同じプロンプトから、キャンバス上で編集できるデプロイ可能なトポロジと、共有用に整えたホワイトボード形式のBlueprint PNGを両方生成します。',
                  })
                ) : (
                  localize(language, {
                    en: 'Describe your Azure architecture in natural language, and AI will generate a diagram with the appropriate services and connections. You can also upload a screenshot, whiteboard photo, or diagram exported from another tool and rebuild it as an editable architecture.',
                    ja: 'Azureアーキテクチャを自然な言葉で説明すると、AIが適切なサービスと接続を含む図を生成します。スクリーンショット、ホワイトボード写真、または他のツールからエクスポートした図をアップロードし、編集可能なアーキテクチャとして再構築することもできます。',
                  })
                )}
              </p>

              <div className="form-group">
                <label htmlFor="architecture-description">{t("Architecture Description or Modification")}</label>
                <textarea
                  id="architecture-description"
                  className="form-textarea"
                  placeholder={imageAnalyzed 
                    ? t("AI has analyzed your diagram. Review the description above, make any adjustments, then click Generate.")
                    : t("Describe a new architecture or request changes to the current diagram. Example: I need a web app with a frontend, API backend, SQL database, and blob storage...")}
                  value={description}
                  onChange={(e) => {
                    setDescription(e.target.value);
                    // Clear imageAnalyzed flag if user clears the text
                    if (!e.target.value.includes('[Analyzed from uploaded diagram]')) {
                      setImageAnalyzed(false);
                    }
                  }}
                  rows={imageAnalyzed ? 10 : 6}
                  disabled={isGenerating || isAnalyzingImage}
                />
              </div>

              <ImageUploader
                onImageAnalyzed={handleImageAnalyzed}
                onImageDataUrl={setUploadedImageUrl}
                onAnalyzing={setIsAnalyzingImage}
                onError={setError}
                disabled={isGenerating}
                analyzeImage={handleAnalyzeImage}
              />

              {error && (
                <div className="error-message">
                  {error}
                </div>
              )}

              {aiMetrics && (
                <div className="similar-architectures">
                  <h3>{t("✓ Architecture generated successfully!")}</h3>
                  <div className="ai-metrics">
                    <span className="metric">
                      <Clock size={14} />
                      {(aiMetrics.elapsedTimeMs / 1000).toFixed(1)}{t("s")}{' '}</span>
                    <span className="metric">
                      <Zap size={14} />
                      {aiMetrics.promptTokens.toLocaleString()} {' '}{t("in →")}{' '}{aiMetrics.completionTokens.toLocaleString()} {' '}{t("out (")}{aiMetrics.totalTokens.toLocaleString()} {' '}{t("total)")}{' '}</span>
                  </div>
                </div>
              )}
              </div>
              <div className="modal-col modal-col-right">
              <div className="mode-toggle" role="tablist" aria-label={t("Generation mode")}>
                <button
                  role="tab"
                  aria-selected={mode === 'topology'}
                  className={`mode-toggle-btn ${mode === 'topology' ? 'active' : ''}`}
                  onClick={() => handleModeChange('topology')}
                  disabled={isGenerating}
                  type="button"
                >
                  <Network size={16} />
                  <span className="mode-label">{t("Topology")}</span>
                  <span className="mode-sub">{localize(language, { en: 'Deployable network diagram', ja: 'デプロイ可能なネットワーク図' })}</span>
                </button>
                {/* Reference (swim-lane) mode hidden — Blueprint replaces it. Code path kept for now in case we want to restore. */}
                <button
                  role="tab"
                  aria-selected={mode === 'blueprint'}
                  className={`mode-toggle-btn ${mode === 'blueprint' ? 'active' : ''}`}
                  onClick={() => handleModeChange('blueprint')}
                  disabled={isGenerating}
                  type="button"
                >
                  <PenTool size={16} />
                  <span className="mode-label">{t("Blueprint")}{' '}<span className="mode-badge-beta">{t("BETA")}</span></span>
                  <span className="mode-sub">{localize(language, { en: 'Hand-drawn whiteboard diagram', ja: '手描きホワイトボード図' })}</span>
                </button>
                <button
                  role="tab"
                  aria-selected={mode === 'both'}
                  className={`mode-toggle-btn ${mode === 'both' ? 'active' : ''}`}
                  onClick={() => handleModeChange('both')}
                  disabled={isGenerating}
                  type="button"
                >
                  <Layers size={16} />
                  <span className="mode-label">{t("Both")}{' '}<span className="mode-badge-beta">{t("BETA")}</span></span>
                  <span className="mode-sub">{t("Topology + Blueprint")}</span>
                </button>
              </div>
              <div className="example-prompts">
                <h3>{t("Example Prompts")}</h3>
                <div className="example-list">
                  {categorizedPrompts.map((group) => (
                    <div key={group.category} className="example-category">
                      <div
                        className="example-category-label"
                        style={{
                          color: group.color,
                          textTransform: language === 'ja' ? 'none' : undefined,
                          letterSpacing: language === 'ja' ? 'normal' : undefined,
                        }}
                      >
                        <span className="example-category-dot" style={{ backgroundColor: group.color }} />
                        {group.category}
                      </div>
                      {group.prompts.map((prompt, idx) => (
                        <button
                          key={idx}
                          className="example-button"
                          style={{ borderLeftColor: group.color }}
                          onClick={() => useExample(prompt)}
                          disabled={isGenerating}
                        >
                          {prompt}
                        </button>
                      ))}
                    </div>
                  ))}
                </div>
              </div>
              </div>
             </div>
            </div>

            <div className="modal-footer">
              <div className="ai-modal-active-model">
                <Brain size={20} />
                <div className="ai-modal-model-controls">
                  {mode === 'both' ? (
                    <>
                      {renderFeatureModelControls('architectureGeneration', false)}
                      {renderFeatureModelControls('blueprint', true)}
                    </>
                  ) : renderFeatureModelControls(
                    mode === 'blueprint' ? 'blueprint' : 'architectureGeneration',
                    mode === 'blueprint',
                  )}
                </div>
                <span className="model-change-hint">
                  {mode === 'both'
                    ? t("Each output uses its feature-specific model.")
                    : mode === 'blueprint'
                      ? t("Blueprint mode supports general-purpose OpenAI models only (partner and Codex models are filtered out).")
                      : t("Also configurable in toolbar → AI Model")}
                </span>
              </div>
              {currentArchitecture && currentArchitecture.nodes.length > 0 && (
                <div className="auto-snapshot-option">
                  <label className="checkbox-label">
                    <input
                      type="checkbox"
                      checked={autoSnapshot}
                      onChange={(e) => handleAutoSnapshotChange(e.target.checked)}
                      disabled={isGenerating}
                    />
                    <span>{t("Auto-save snapshot before regenerating")}</span>
                  </label>
                  <p className="checkbox-hint">
                    {' '}{t("Automatically saves your current diagram to version history before generating a new one")}{' '}</p>
                </div>
              )}
              {mode === 'both' && (
                <div className="auto-snapshot-option">
                  <label className="checkbox-label">
                    <input
                      type="checkbox"
                      checked={bothInParallel}
                      onChange={(e) => handleBothInParallelChange(e.target.checked)}
                      disabled={isGenerating}
                    />
                    <span>{t("Run topology and blueprint in parallel")}</span>
                  </label>
                  <p className="checkbox-hint">
                    {' '}{t("Parallel ≈ half the wall-time (recommended on high-quota deployments). Uncheck to run sequentially if your model deployment has tight rate limits.")}{' '}</p>
                </div>
              )}
              {(mode === 'blueprint' || mode === 'both') && (
                <div className="auto-snapshot-option">
                  <label className="checkbox-label" style={{ alignItems: 'center', gap: 8 }}>
                    <span>{t("Blueprint legend position:")}</span>
                    <select
                      className="ai-modal-model-select"
                      value={legendPosition}
                      onChange={(e) => handleLegendPositionChange(e.target.value as 'auto' | 'bottom' | 'right')}
                      disabled={isGenerating}
                      aria-label={t("Blueprint legend position")}
                    >
                      <option value="auto">{t("Auto (by aspect ratio)")}</option>
                      <option value="bottom">{t("Bottom (full-width canvas)")}</option>
                      <option value="right">{t("Right (taller canvas)")}</option>
                    </select>
                  </label>
                  <p className="checkbox-hint">
                    {' '}{t("Auto picks \"bottom\" for wide diagrams and \"right\" for square / tall ones.")}{' '}</p>
                </div>
              )}
              <div className="modal-footer-actions">
                <button
                  className="btn btn-secondary"
                  onClick={closeModal}
                  disabled={isGenerating}
                >
                  {aiMetrics ? t("Close") : t("Cancel")}
                </button>
                <button
                  className="btn btn-primary btn-generate-ai"
                  onClick={handleGenerate}
                  disabled={isGenerating || isAnalyzingImage || !description.trim()}
                  style={{ display: aiMetrics ? 'none' : 'flex' }}
                >
                  {isGenerating ? (
                    <>
                      <Loader2 size={18} className="spinner" />
                      {' '}{t("Generating...")}{' '}</>
                  ) : (
                    <>
                      <Sparkles size={18} />
                      {' '}{t("Generate Architecture")}{' '}</>
                  )}
                </button>
              </div>
            </div>
          </div>
        </div>,
        document.body
      )}
    </>
  );
};

export default AIArchitectureGenerator;
