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
import { helpJapanese } from '../i18n/helpJapanese';
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
    return helpJapanese[text] ?? translateFallback(text);
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