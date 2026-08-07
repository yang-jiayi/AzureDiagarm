// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import {
  BookOpen, Check, CheckCircle2, CircleDollarSign, ClipboardCheck, Download,
  Copy, FileCode2, GitCompare, History, Image as ImageIcon, LayoutDashboard,
  Lightbulb, Map, MessageSquare, MousePointer2, Presentation,
  Rocket, Route, ShieldCheck, Sparkles, UploadCloud, Wrench, X,
} from 'lucide-react';
import { trackHelpOpened } from '../services/telemetryService';
import './GuidedHelpPanel.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';

interface GuidedHelpPanelProps {
  isOpen: boolean;
  onClose: () => void;
}

type SectionId = 'quick-start' | 'paths' | 'create' | 'assess' | 'deliver' | 'prompts' | 'faq';

const CHECKLIST_STORAGE_KEY = 'help.onboarding.v3';

const SECTIONS: { id: SectionId; label: string; icon: React.ReactNode }[] = [
  { id: 'quick-start', label: 'Start Here', icon: <Rocket size={16} /> },
  { id: 'paths', label: 'Choose a Path', icon: <Route size={16} /> },
  { id: 'create', label: 'Create & Refine', icon: <Sparkles size={16} /> },
  { id: 'assess', label: 'Assess', icon: <ShieldCheck size={16} /> },
  { id: 'deliver', label: 'Deliver', icon: <Presentation size={16} /> },
  { id: 'prompts', label: 'Prompt Lab', icon: <Lightbulb size={16} /> },
  { id: 'faq', label: 'FAQ & Safety', icon: <BookOpen size={16} /> },
];

const FIRST_TOUR = [
  { id: 'create', title: '1. Create', detail: 'Choose Guided Chat, Generate Diagram, or Import Existing.' },
  { id: 'refine', title: '2. Refine', detail: 'Continue in Guided Chat or edit nodes, groups, and connections directly.' },
  { id: 'validate', title: '3. Validate & Improve', detail: 'Review findings, apply selected changes, and revalidate when the design changes.' },
  { id: 'deliver', title: '4. Share or Build', detail: 'Export a review artifact or create deployment guidance when the design is ready.' },
];

const PATHS = [
  {
    icon: <MessageSquare size={20} />,
    title: 'Start from an idea',
    label: 'Recommended for first-time users',
    steps: ['Open Guided Chat', 'Describe the outcome and constraints', 'Create and refine in one conversation'],
  },
  {
    icon: <ImageIcon size={20} />,
    title: 'Start from a brief or sketch',
    label: 'Detailed requirements or image',
    steps: ['Open Generate Diagram', 'Choose Topology, Blueprint, or both', 'Review, then continue in Guided Chat or on canvas'],
  },
  {
    icon: <FileCode2 size={20} />,
    title: 'Start from current estate',
    label: 'IaC or live Azure',
    steps: ['Import Bicep, Terraform, or ARM', 'Or use Import from Azure', 'Correct inferred relationships on canvas'],
  },
  {
    icon: <Presentation size={20} />,
    title: 'Prepare a review or workshop',
    label: 'Customer-ready flow',
    steps: ['Create and correct the concept', 'Validate and improve iteratively', 'Share a review artifact or build deployment guidance'],
  },
];

const CREATE_FEATURES = [
  { icon: <MessageSquare size={18} />, title: 'Guided Chat', body: 'Best for conversational creation and ongoing refinement. Build from empty or change the current canvas in one thread; existing manual positions are retained during modifications.' },
  { icon: <Sparkles size={18} />, title: 'Generate Diagram', body: 'Best when you have detailed requirements, an existing sketch, or need explicit output controls. Choose Topology, Blueprint, or Both, then hand off to Guided Chat or canvas review.' },
  { icon: <UploadCloud size={18} />, title: 'Import', body: 'Reconstruct a diagram image, parse Bicep/Terraform/ARM, or sign in to reverse-engineer a live Azure resource group.' },
  { icon: <GitCompare size={18} />, title: 'Compare Models', body: 'Run one prompt across several models, inspect latency/tokens/topology differences, and apply the result you prefer.' },
  { icon: <MousePointer2 size={18} />, title: 'Edit on canvas', body: 'Drag services, resize groups, edit labels, reconnect edges, align selections, and choose a layout preset or edge style.' },
  { icon: <History size={18} />, title: 'Version History', body: 'A snapshot is saved before AI regeneration. Save named checkpoints and restore prior versions when an experiment does not work.' },
];

const ASSESS_FEATURES = [
  { icon: <ShieldCheck size={18} />, title: 'Well-Architected validation', body: 'Review Cost Optimization, Operational Excellence, Performance Efficiency, Reliability, and Security. Apply selected recommendations, review the resulting iteration, and revalidate after material changes.' },
  { icon: <GitCompare size={18} />, title: 'Compare Validation', body: 'Ask multiple models to review the same architecture, compare findings, and use consensus to separate recurring gaps from model-specific opinions.' },
  { icon: <CircleDollarSign size={18} />, title: 'Cost and region', body: 'Inspect per-service monthly estimates across eight regions and switch between PAYG and 1-year savings. Usage-based values remain indicative.' },
  { icon: <ClipboardCheck size={18} />, title: 'Validation timing', body: 'Review and refine the generated concept first, then use the Validate & Improve journey stage before sharing or building. Revalidate after material changes.' },
];

const DELIVER_FEATURES = [
  { icon: <Download size={18} />, title: 'Editable formats', body: 'Use Visio (VSDX), Draw.io, JSON, or interactive HTML when another person needs to continue editing.' },
  { icon: <Presentation size={18} />, title: 'Presentation formats', body: 'Export PNG, SVG, a PowerPoint slide, or a customer deck. Use Export background to choose Plain (recommended), Dots, or Grid without changing the editing canvas.' },
  { icon: <Map size={18} />, title: 'Workflow outputs', body: 'Export a Markdown narrative or animated workflow, and use Narrate when the Speech presenter is available.' },
  { icon: <FileCode2 size={18} />, title: 'Deployment Guide', body: 'Generate a Microsoft Learn-grounded runbook and Bicep starters. Review all commands, sizing, identities, and safeguards before deployment.' },
  { icon: <CircleDollarSign size={18} />, title: 'Cost package', body: 'Download CSV or the all-formats ZIP with summaries, analysis, JSON, and multi-region comparison.' },
  { icon: <LayoutDashboard size={18} />, title: 'Demo mode', body: 'Use Focus, Hide Toolbar, collapse groups, mini-map navigation, and Fit to view to present large diagrams clearly.' },
];

const STRUCTURED_PROMPT = `Design an Azure architecture for the following workload.

Outcome: <business outcome in one sentence>
Users and channels: <who uses it, how, and approximate scale>
Data sources and destinations: <systems, documents, events, databases>
AI capabilities: <retrieval, conversation, prediction, vision, automation>
Existing investments to reuse: <Azure, Microsoft 365, Fabric, Dynamics>
Constraints:
  - Region(s): <primary and DR>
  - Identity: <Entra ID, external identities, hybrid>
  - Compliance and data sensitivity: <requirements>
  - Availability and scale: <targets>
  - Budget or delivery stage: <demo, pilot, production>
Out of scope: <explicit exclusions>

Group the design by responsibility, show primary data flows, and include identity, observability, and secrets management where required.`;

const EXAMPLE_PROMPTS = [
  'Internal RAG assistant grounded on SharePoint and policy documents, available in Teams, secured with Entra ID, with citations and feedback telemetry.',
  'Event-driven order processing at 50K orders/hour using API Management, Service Bus, Functions, Cosmos DB, Key Vault, and Application Insights.',
  'Import and modernize a three-tier application into Container Apps with private connectivity, managed identity, Azure SQL, Redis, and Front Door with WAF.',
  'AI Discovery Cards workshop concept: reduce claims triage time using Document Intelligence, anomaly detection, human review, Fabric analytics, and D365 integration.',
];

const STRUCTURED_PROMPT_JA = `次のワークロード向けに Azure アーキテクチャを設計してください。

成果: <ビジネス成果を1文で記載>
ユーザーとチャネル: <利用者、利用方法、おおよその規模>
データの入力元と出力先: <システム、文書、イベント、データベース>
AI機能: <検索、会話、予測、画像、オートメーション>
再利用する既存資産: <Azure、Microsoft 365、Fabric、Dynamics>
制約:
  - リージョン: <プライマリとDR>
  - ID: <Entra ID、外部ID、ハイブリッド>
  - コンプライアンスとデータ機密性: <要件>
  - 可用性と規模: <目標>
  - 予算または提供段階: <デモ、パイロット、本番>
対象外: <明示的な除外事項>

責任範囲ごとに設計をグループ化し、主要なデータ フローを示し、必要に応じてID、監視、シークレット管理を含めてください。`;

const GUIDED_HELP_JAPANESE: Record<string, string> = {
  'Help and Learn': 'ヘルプと学習',
  'Help & Learn': 'ヘルプと学習',
  'From first prompt to architecture handoff': '最初のプロンプトからアーキテクチャの引き渡しまで',
  'Close': '閉じる',
  'Close help': 'ヘルプを閉じる',
  'Start Here': 'はじめに',
  'Choose a Path': '開始方法',
  'Create & Refine': '作成と改善',
  'Assess': '評価',
  'Deliver': '提供',
  'Prompt Lab': 'プロンプト ラボ',
  'FAQ & Safety': 'FAQと安全な利用',
  'First tour': '最初のツアー',
  'Your first 10 minutes': '最初の10分',
  'Build confidence by completing one full loop': '一連の流れを完了して操作を身につける',
  'Create, correct, validate, inspect cost, and export. Mark each task as you try it—the checklist is saved in this browser.': '作成、修正、検証、コスト確認、エクスポートまで試してください。完了した項目はブラウザーに保存されます。',
  'Create or import a diagram': '図を作成またはインポート',
  'Use Chat, Generate with AI, an image, an IaC template, or Import from Azure.': 'Chat、AIで生成、画像、IaCテンプレート、またはAzureからのインポートを使用します。',
  'Make one targeted change': '対象を絞って1件変更',
  'Ask Chat for a change or edit nodes, groups, and connections directly.': 'Chatに変更を依頼するか、ノード、グループ、接続を直接編集します。',
  'Run Well-Architected validation': 'Well-Architected検証を実行',
  'Review the five pillars and identify the most important gaps.': '5つの柱を確認し、重要な不足を特定します。',
  'Inspect cost and region': 'コストとリージョンを確認',
  'Choose a region, compare PAYG and 1-year savings, and inspect service estimates.': 'リージョンを選択し、従量課金と1年予約の節約額、サービス見積もりを確認します。',
  'Export a useful artifact': '用途に合う成果物をエクスポート',
  'Try PowerPoint, Visio, Draw.io, PNG, HTML, workflow Markdown, or JSON.': 'PowerPoint、Visio、Draw.io、PNG、HTML、ワークフローMarkdown、JSONを試します。',
  'Important:': '重要:',
  'AI output is a starting hypothesis. Correct it with domain experts, validate assumptions, and review generated deployment content before use.': 'AI出力は検討の出発点です。利用前に専門家と修正し、前提を検証し、生成されたデプロイ内容をレビューしてください。',
  'Choose how you want to start': '開始方法を選択',
  'Choose a starting point': '開始点を選択',
  'You do not need to begin with a perfect prompt': '最初から完璧なプロンプトは必要ありません',
  'Start from an idea': 'アイデアから開始',
  'Fastest path': '最短ルート',
  'Open Chat': 'Chatを開く',
  'Describe the outcome and constraints': '成果と制約を説明',
  'Refine in plain English': '自然な言葉で改善',
  'Start from a sketch': 'スケッチから開始',
  'Whiteboard or screenshot': 'ホワイトボードまたはスクリーンショット',
  'Open Generate with AI': 'AIで生成を開く',
  'Upload the image': '画像をアップロード',
  'Review the reconstructed services and flows': '再構成されたサービスとフローを確認',
  'Start from current estate': '現在の環境から開始',
  'IaC or live Azure': 'IaCまたは稼働中のAzure',
  'Import Bicep, Terraform, or ARM': 'Bicep、Terraform、ARMをインポート',
  'Or use Import from Azure': 'またはAzureからインポート',
  'Correct inferred relationships on canvas': '推定された関係をキャンバス上で修正',
  'Prepare a review or workshop': 'レビューまたはワークショップを準備',
  'Customer-ready flow': '顧客向けの流れ',
  'Generate and correct the concept': 'コンセプトを生成して修正',
  'Validate and inspect cost': '検証してコストを確認',
  'Export a customer deck or editable artifact': '顧客向け資料または編集可能な成果物をエクスポート',
  'Create and refine': '作成と改善',
  'Move between AI and direct canvas editing': 'AIとキャンバスの直接編集を使い分ける',
  'Use AI for acceleration, then use the canvas for precision. Targeted follow-up requests preserve your existing manual layout.': 'AIで作業を加速し、キャンバスで正確に調整します。対象を絞った追加依頼では既存の手動レイアウトが維持されます。',
  'Architecture Chat': 'アーキテクチャ Chat',
  'Build from empty or refine the current canvas. Existing manual positions are retained during AI modifications, and each change is snapshotted.': '空の状態から作成するか、現在のキャンバスを改善します。AI変更時も手動位置を維持し、変更ごとにスナップショットを保存します。',
  'Generate with AI': 'AIで生成',
  'Choose Topology for an editable canvas, Blueprint for a whiteboard-style PNG, or Both for the two views together.': '編集可能なキャンバスはTopology、ホワイトボード形式のPNGはBlueprint、両方必要な場合はBothを選択します。',
  'Import': 'インポート',
  'Reconstruct a diagram image, parse Bicep/Terraform/ARM, or sign in to reverse-engineer a live Azure resource group.': '図の画像を再構成し、Bicep/Terraform/ARMを解析するか、サインインして稼働中のAzureリソース グループを可視化します。',
  'Compare Models': 'モデル比較',
  'Run one prompt across several models, inspect latency/tokens/topology differences, and apply the result you prefer.': '同じプロンプトを複数モデルで実行し、待ち時間、トークン、トポロジの違いを確認して、選択した結果を適用します。',
  'Edit on canvas': 'キャンバス上で編集',
  'Drag services, resize groups, edit labels, reconnect edges, align selections, and choose a layout preset or edge style.': 'サービスの移動、グループのサイズ変更、ラベル編集、接続変更、整列、レイアウトやエッジ スタイルの選択ができます。',
  'Version History': 'バージョン履歴',
  'A snapshot is saved before AI regeneration. Save named checkpoints and restore prior versions when an experiment does not work.': 'AI再生成前にスナップショットを保存します。名前付きチェックポイントを保存し、試行がうまくいかない場合は以前の版を復元できます。',
  'Turn a diagram into a review conversation': '図をレビューの対話につなげる',
  'Validation and cost are decision aids. They expose assumptions and tradeoffs; they do not replace sizing, security review, or architecture approval.': '検証とコストは意思決定の補助です。前提とトレードオフを明確にしますが、サイジング、セキュリティ レビュー、アーキテクチャ承認の代わりにはなりません。',
  'Well-Architected validation': 'Well-Architected検証',
  'Review Cost Optimization, Operational Excellence, Performance Efficiency, Reliability, and Security. Apply selected recommendations to create a new iteration.': 'コスト最適化、オペレーショナル エクセレンス、パフォーマンス効率、信頼性、セキュリティを確認し、選択した推奨事項を次の版に適用します。',
  'Compare Validation': '検証比較',
  'Ask multiple models to review the same architecture, compare findings, and use consensus to separate recurring gaps from model-specific opinions.': '複数モデルで同じアーキテクチャをレビューし、結果を比較して、共通の不足とモデル固有の意見を分けます。',
  'Cost and region': 'コストとリージョン',
  'Inspect per-service monthly estimates across eight regions and switch between PAYG and 1-year savings. Usage-based values remain indicative.': '8リージョンのサービス別月額見積もりを確認し、従量課金と1年予約を切り替えます。使用量ベースの値は概算です。',
  'Validation handoff': '検証への引き渡し',
  'After generation, use Validate now to check readiness before sharing. A concept diagram is still a hypothesis—not an approved production design.': '生成後に「今すぐ検証」で共有前の準備状況を確認します。コンセプト図は仮説であり、承認済みの本番設計ではありません。',
  'Choose an output for the next person': '次の利用者に合う出力を選択',
  'Export based on what the recipient needs to do next: present, edit, review, estimate, or continue implementation planning.': '受け手が次に行う作業に応じて、プレゼン、編集、レビュー、見積もり、実装計画向けの形式を選びます。',
  'Editable formats': '編集可能な形式',
  'Use Visio (VSDX), Draw.io, JSON, or interactive HTML when another person needs to continue editing.': '他の利用者が編集を続ける場合はVisio (VSDX)、Draw.io、JSON、対話型HTMLを使用します。',
  'Presentation formats': 'プレゼンテーション形式',
  'Export PNG, SVG, a PowerPoint slide, or a multi-slide customer architecture deck.': 'PNG、SVG、PowerPointスライド、複数スライドの顧客向けアーキテクチャ資料を出力します。',
  'Workflow outputs': 'ワークフロー出力',
  'Export a Markdown narrative or animated workflow, and use Narrate when the Speech presenter is available.': 'Markdownの説明またはアニメーション ワークフローを出力し、音声プレゼンターが利用可能な場合はナレーションを使用します。',
  'Deployment Guide': 'デプロイ ガイド',
  'Generate a Microsoft Learn-grounded runbook and Bicep starters. Review all commands, sizing, identities, and safeguards before deployment.': 'Microsoft Learnに基づく手順書とBicepのひな型を生成します。デプロイ前にコマンド、サイジング、ID、安全策をすべて確認してください。',
  'Cost package': 'コスト パッケージ',
  'Download CSV or the all-formats ZIP with summaries, analysis, JSON, and multi-region comparison.': '概要、分析、JSON、複数リージョン比較を含むCSVまたは全形式ZIPをダウンロードします。',
  'Demo mode': 'デモ モード',
  'Use Focus, Hide Toolbar, collapse groups, mini-map navigation, and Fit to view to present large diagrams clearly.': 'フォーカス、ツールバー非表示、グループ折りたたみ、ミニマップ、全体表示を使って大きな図を見やすく提示します。',
  'Describe intent and constraints—not a shopping list': 'サービス一覧ではなく目的と制約を説明する',
  'A useful prompt names the outcome, users, data, existing investments, and non-functional constraints. You can leave unknowns explicit.': '有用なプロンプトには、成果、ユーザー、データ、既存資産、非機能要件を記載します。不明点は不明のまま明示できます。',
  'Copied': 'コピーしました',
  'Copy template': 'テンプレートをコピー',
  'Quick examples': '簡単な例',
  'Best follow-ups are specific: “keep existing positions,” “use private endpoints for data services,” “show a pilot under $500/month,” or “replace App Service with Container Apps.”': '追加依頼は具体的にします。例:「既存位置を維持」「データ サービスにPrivate Endpointを使用」「月額500ドル未満のパイロットを提示」「App ServiceをContainer Appsに置換」。',
  'FAQ and responsible use': 'FAQと責任ある利用',
  'Know what the tool does—and what still needs review': 'ツールの役割と追加レビューが必要な範囲を理解する',
  'Which model should I use?': 'どのモデルを使用すべきですか?',
  'Use the selected default for most work. Compare models when the architecture is consequential or outputs vary. Higher reasoning can improve complex designs but usually takes longer.': '通常は選択済みの既定モデルを使用します。重要な設計や結果に差がある場合はモデルを比較します。推論を強くすると複雑な設計の品質が上がる場合がありますが、時間も長くなります。',
  'How do I correct an AI result?': 'AIの結果をどのように修正しますか?',
  'Use Chat for a targeted change, then edit directly on canvas. Existing positions are preserved during refinements. Version History lets you restore an earlier state.': 'Chatで対象を絞って変更し、その後キャンバス上で直接編集します。改善時は既存位置が維持され、バージョン履歴から以前の状態を復元できます。',
  'Can I import existing infrastructure?': '既存インフラをインポートできますか?',
  'Yes. Import Bicep, Terraform, ARM, an architecture image, or a live Azure resource group. Review inferred connections and unsupported resources.': 'はい。Bicep、Terraform、ARM、アーキテクチャ画像、稼働中のAzureリソース グループをインポートできます。推定された接続と未対応リソースを確認してください。',
  'Are the costs authoritative?': 'コストは確定値ですか?',
  'No. They are indicative catalog-based estimates. Confirm SKU, quantity, usage, discounts, networking, support, and regional availability in the Azure Pricing Calculator.': 'いいえ。カタログに基づく概算です。SKU、数量、使用量、割引、ネットワーク、サポート、リージョン提供状況をAzure料金計算ツールで確認してください。',
  'Does a WAF score approve the design?': 'WAFスコアで設計が承認されますか?',
  'No. It is a structured review aid based on visible topology and model context. Validate findings with architects, security, operations, and workload owners.': 'いいえ。表示されたトポロジとモデル コンテキストに基づくレビュー補助です。アーキテクト、セキュリティ、運用、ワークロード所有者と結果を検証してください。',
  'Can I deploy the generated Bicep directly?': '生成されたBicepをそのままデプロイできますか?',
  'Treat it as starter IaC. Review API versions, identities, network controls, naming, policy, sizing, dependencies, and destructive operations before deployment.': 'IaCのひな型として扱ってください。デプロイ前にAPIバージョン、ID、ネットワーク制御、命名、ポリシー、サイジング、依存関係、破壊的操作を確認します。',
  'What information should I avoid entering?': '入力を避けるべき情報は何ですか?',
  'Do not enter passwords, keys, tokens, regulated personal data, confidential customer content, or production data unless your organization has explicitly approved that use.': '組織が明示的に承認していない限り、パスワード、キー、トークン、規制対象の個人データ、顧客の機密情報、本番データを入力しないでください。',
  'Trusted references': '信頼できる参考資料',
  'From Prompt to Production—AADB overview': 'From Prompt to Production - AADB概要',
  'Azure Well-Architected Framework': 'Azure Well-Architected Framework',
  'Azure Pricing Calculator': 'Azure料金計算ツール',
  'Azure Architecture Center': 'Azure Architecture Center',
  'Still stuck or found something wrong? Close Help and use the Feedback button in the lower-right corner.': '解決しない場合や問題を見つけた場合は、ヘルプを閉じて右下の「フィードバック」ボタンをご利用ください。',
  'Internal RAG assistant grounded on SharePoint and policy documents, available in Teams, secured with Entra ID, with citations and feedback telemetry.': 'SharePointと社内規程文書を根拠にし、Teamsから利用でき、Entra IDで保護され、引用とフィードバック計測を備えた社内RAGアシスタント。',
  'Event-driven order processing at 50K orders/hour using API Management, Service Bus, Functions, Cosmos DB, Key Vault, and Application Insights.': 'API Management、Service Bus、Functions、Cosmos DB、Key Vault、Application Insightsを使用し、1時間あたり5万件を処理するイベント駆動型注文処理。',
  'Import and modernize a three-tier application into Container Apps with private connectivity, managed identity, Azure SQL, Redis, and Front Door with WAF.': '3層アプリケーションをインポートし、プライベート接続、マネージドID、Azure SQL、Redis、WAF付きFront Doorを備えたContainer Appsへモダナイズ。',
  'AI Discovery Cards workshop concept: reduce claims triage time using Document Intelligence, anomaly detection, human review, Fabric analytics, and D365 integration.': 'Document Intelligence、異常検知、人によるレビュー、Fabric分析、D365連携を使用して請求トリアージ時間を短縮するAI Discovery Cardsワークショップ案。',
};

function readCompletedTour(): Set<string> {
  try {
    const saved = JSON.parse(localStorage.getItem(CHECKLIST_STORAGE_KEY) || '[]');
    return new Set(Array.isArray(saved) ? saved.filter((value) => typeof value === 'string') : []);
  } catch {
    return new Set();
  }
}

const GuidedHelpPanel: React.FC<GuidedHelpPanelProps> = ({ isOpen, onClose }) => {
  const { language, translate: translateFallback } = useLanguage();
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);
  useEscapeKey(isOpen, onClose);
  const translate = (text: string): string => {
    if (language !== 'ja') return text;
    if (text === STRUCTURED_PROMPT) return STRUCTURED_PROMPT_JA;
    return GUIDED_HELP_JAPANESE[text] ?? translateFallback(text);
  };
  const [section, setSection] = useState<SectionId>('quick-start');
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const [completedTour, setCompletedTour] = useState<Set<string>>(readCompletedTour);

  useEffect(() => {
    if (isOpen) trackHelpOpened('quick-start');
  }, [isOpen]);

  const goToSection = (id: SectionId) => {
    setSection(id);
    trackHelpOpened(id);
  };

  const toggleTourItem = (id: string) => {
    setCompletedTour((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id); else next.add(id);
      try { localStorage.setItem(CHECKLIST_STORAGE_KEY, JSON.stringify([...next])); } catch { /* ignore */ }
      return next;
    });
  };

  const copyText = async (text: string, key: string) => {
    let copied = false;
    try {
      await navigator.clipboard.writeText(text);
      copied = true;
    } catch {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.setAttribute('readonly', '');
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      copied = document.execCommand('copy');
      textarea.remove();
    }

    if (copied) {
      setCopiedKey(key);
      window.setTimeout(() => setCopiedKey((current) => current === key ? null : current), 1600);
    }
  };

  if (!isOpen) return null;

  const completedCount = FIRST_TOUR.filter((item) => completedTour.has(item.id)).length;

  return (
    <div
      ref={dialogRef}
      className="guided-help-overlay"
      role="dialog"
      aria-modal="true"
      aria-label={translate('Help and Learn')}
      tabIndex={-1}
      onClick={onClose}
    >
      <div className="guided-help-modal" onClick={(event) => event.stopPropagation()}>
        <header className="guided-help-header">
          <div className="guided-help-title">
            <BookOpen size={21} />
            <div>
              <strong>{translate('Help & Learn')}</strong>
              <span>{translate('From first prompt to architecture handoff')}</span>
            </div>
          </div>
          <button className="guided-help-close" onClick={onClose} title={translate('Close')} aria-label={translate('Close help')}><X size={18} /></button>
        </header>

        <div className="guided-help-body">
          <nav className="guided-help-nav" aria-label={translate('Help sections')}>
            {SECTIONS.map((item) => (
              <button key={item.id} className={`guided-help-nav-item${section === item.id ? ' active' : ''}`} onClick={() => goToSection(item.id)}>
                {item.icon}<span>{translate(item.label)}</span>
              </button>
            ))}
            <div className="guided-help-nav-progress">
              <span>{translate('First tour')}</span><strong>{completedCount}/{FIRST_TOUR.length}</strong>
              <div><i style={{ width: `${completedCount * 100 / FIRST_TOUR.length}%` }} /></div>
            </div>
          </nav>

          <main className="guided-help-content">
            {section === 'quick-start' && (
              <section className="guided-help-section">
                <p className="guided-help-eyebrow">{translate('Your first 10 minutes')}</p>
                <h2>{translate('Build confidence by completing one full loop')}</h2>
                <p className="guided-help-lead">{translate('Create, correct, validate, inspect cost, and export. Mark each task as you try it—the checklist is saved in this browser.')}</p>
                <div className="guided-help-checklist">
                  {FIRST_TOUR.map((item, index) => {
                    const done = completedTour.has(item.id);
                    return <button key={item.id} className={done ? 'done' : ''} onClick={() => toggleTourItem(item.id)} aria-pressed={done}>
                      <span className="guided-help-check-state">{done ? <CheckCircle2 size={20} /> : index + 1}</span>
                      <span><strong>{translate(item.title)}</strong><small>{translate(item.detail)}</small></span>
                    </button>;
                  })}
                </div>
                <div className="guided-help-callout"><ShieldCheck size={18} /><span><strong>{translate('Important:')}</strong> {translate('AI output is a starting hypothesis. Correct it with domain experts, validate assumptions, and review generated deployment content before use.')}</span></div>
                <div className="guided-help-next"><button onClick={() => goToSection('paths')}>{translate('Choose how you want to start')} <Route size={16} /></button></div>
              </section>
            )}

            {section === 'paths' && (
              <section className="guided-help-section">
                <p className="guided-help-eyebrow">{translate('Choose a starting point')}</p>
                <h2>{translate('You do not need to begin with a perfect prompt')}</h2>
                <div className="guided-help-paths">{PATHS.map((path) => <article key={path.title}>
                  <div className="guided-help-path-icon">{path.icon}</div><span>{translate(path.label)}</span><h3>{translate(path.title)}</h3>
                  <ol>{path.steps.map((step) => <li key={step}>{translate(step)}</li>)}</ol>
                </article>)}</div>
              </section>
            )}

            {section === 'create' && <FeatureSection eyebrow="Create and refine" title="Move between AI and direct canvas editing" intro="Use AI for acceleration, then use the canvas for precision. Targeted follow-up requests preserve your existing manual layout." features={CREATE_FEATURES} translate={translate} />}
            {section === 'assess' && <FeatureSection eyebrow="Assess" title="Turn a diagram into a review conversation" intro="Validation and cost are decision aids. They expose assumptions and tradeoffs; they do not replace sizing, security review, or architecture approval." features={ASSESS_FEATURES} translate={translate} />}
            {section === 'deliver' && <FeatureSection eyebrow="Deliver" title="Choose an output for the next person" intro="Export based on what the recipient needs to do next: present, edit, review, estimate, or continue implementation planning." features={DELIVER_FEATURES} translate={translate} />}

            {section === 'prompts' && (
              <section className="guided-help-section">
                <p className="guided-help-eyebrow">{translate('Prompt Lab')}</p><h2>{translate('Describe intent and constraints—not a shopping list')}</h2>
                <p className="guided-help-lead">{translate('A useful prompt names the outcome, users, data, existing investments, and non-functional constraints. You can leave unknowns explicit.')}</p>
                <div className="guided-help-template"><pre>{translate(STRUCTURED_PROMPT)}</pre><button onClick={() => copyText(translate(STRUCTURED_PROMPT), 'template')}>{copiedKey === 'template' ? <Check size={16} /> : <Copy size={16} />}{copiedKey === 'template' ? translate('Copied') : translate('Copy template')}</button></div>
                <h3>{translate('Quick examples')}</h3><div className="guided-help-prompts">{EXAMPLE_PROMPTS.map((prompt, index) => <button key={prompt} onClick={() => copyText(translate(prompt), `example-${index}`)}><span>{translate(prompt)}</span>{copiedKey === `example-${index}` ? <Check size={16} /> : <Copy size={16} />}</button>)}</div>
                <div className="guided-help-callout"><Wrench size={18} /><span>{translate('Best follow-ups are specific: “keep existing positions,” “use private endpoints for data services,” “show a pilot under $500/month,” or “replace App Service with Container Apps.”')}</span></div>
              </section>
            )}

            {section === 'faq' && (
              <section className="guided-help-section">
                <p className="guided-help-eyebrow">{translate('FAQ and responsible use')}</p><h2>{translate('Know what the tool does—and what still needs review')}</h2>
                <div className="guided-help-faq">
                  <Faq q={translate('Which model should I use?')} a={translate('Use the selected default for most work. Compare models when the architecture is consequential or outputs vary. Higher reasoning can improve complex designs but usually takes longer.')} />
                  <Faq q={translate('What is the difference between Guided Chat and Generate Diagram?')} a={translate('Guided Chat is best for conversational creation and repeated refinement. Generate Diagram is best for detailed prompts, uploaded sketches, model selection, and choosing Topology or Blueprint output. Both create an editable result and can continue in Guided Chat.')} />
                  <Faq q={translate('How do I remove the dots from an export?')} a={translate('Open Export and set Export background to Plain (recommended). You can also choose Dots or Grid. This affects visual exports only; the editing canvas remains dotted.')} />
                  <Faq q={translate('How do I correct an AI result?')} a={translate('Use Chat for a targeted change, then edit directly on canvas. Existing positions are preserved during refinements. Version History lets you restore an earlier state.')} />
                  <Faq q={translate('Can I import existing infrastructure?')} a={translate('Yes. Import Bicep, Terraform, ARM, an architecture image, or a live Azure resource group. Review inferred connections and unsupported resources.')} />
                  <Faq q={translate('Are the costs authoritative?')} a={translate('No. They are indicative catalog-based estimates. Confirm SKU, quantity, usage, discounts, networking, support, and regional availability in the Azure Pricing Calculator.')} />
                  <Faq q={translate('Does a WAF score approve the design?')} a={translate('No. It is a structured review aid based on visible topology and model context. Validate findings with architects, security, operations, and workload owners.')} />
                  <Faq q={translate('Can I deploy the generated Bicep directly?')} a={translate('Treat it as starter IaC. Review API versions, identities, network controls, naming, policy, sizing, dependencies, and destructive operations before deployment.')} />
                  <Faq q={translate('What information should I avoid entering?')} a={translate('Do not enter passwords, keys, tokens, regulated personal data, confidential customer content, or production data unless your organization has explicitly approved that use.')} />
                </div>
                <h3>{translate('Trusted references')}</h3>
                <ul className="guided-help-resources">
                  <li><a href="https://techcommunity.microsoft.com/blog/azurearchitectureblog/from-prompt-to-production-building-azure-architecture-diagrams-with-ai/4520336" target="_blank" rel="noopener noreferrer">{translate('From Prompt to Production—AADB overview')}</a></li>
                  <li><a href="https://learn.microsoft.com/azure/well-architected/" target="_blank" rel="noopener noreferrer">{translate('Azure Well-Architected Framework')}</a></li>
                  <li><a href="https://azure.microsoft.com/pricing/calculator/" target="_blank" rel="noopener noreferrer">{translate('Azure Pricing Calculator')}</a></li>
                  <li><a href="https://learn.microsoft.com/azure/architecture/" target="_blank" rel="noopener noreferrer">{translate('Azure Architecture Center')}</a></li>
                </ul>
                <p className="guided-help-feedback">{translate('Still stuck or found something wrong? Close Help and use the Feedback button in the lower-right corner.')}</p>
              </section>
            )}
          </main>
        </div>
      </div>
    </div>
  );
};

function FeatureSection({ eyebrow, title, intro, features, translate }: { eyebrow: string; title: string; intro: string; features: typeof CREATE_FEATURES; translate: (text: string) => string }) {
  return <section className="guided-help-section"><p className="guided-help-eyebrow">{translate(eyebrow)}</p><h2>{translate(title)}</h2><p className="guided-help-lead">{translate(intro)}</p><div className="guided-help-features">{features.map((feature) => <article key={feature.title}><div>{feature.icon}</div><span><strong>{translate(feature.title)}</strong><small>{translate(feature.body)}</small></span></article>)}</div></section>;
}

function Faq({ q, a }: { q: string; a: string }) {
  return <article><h3>{q}</h3><p>{a}</p></article>;
}

export default GuidedHelpPanel;