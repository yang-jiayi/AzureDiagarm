// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import {
  X, Rocket, Sparkles, MessageSquare, Image as ImageIcon, PenTool, ShieldCheck,
  GitCompare, DollarSign, FileText, Mic, History, Lightbulb, BookOpen, Copy, Check,
  Download, Map,
} from 'lucide-react';
import { trackHelpOpened } from '../services/telemetryService';
import './HelpLearnPanel.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize, type LocalizedText } from '../i18n/localization';

interface HelpLearnPanelProps {
  isOpen: boolean;
  onClose: () => void;
}

type SectionId = 'quick-start' | 'features' | 'prompts' | 'tips' | 'resources';

const SECTIONS: { id: SectionId; label: string; icon: React.ReactNode }[] = [
  { id: 'quick-start', label: 'Quick Start', icon: <Rocket size={16} /> },
  { id: 'features', label: 'Feature Tour', icon: <Sparkles size={16} /> },
  { id: 'prompts', label: 'Example Prompts', icon: <Lightbulb size={16} /> },
  { id: 'tips', label: 'Tips & FAQ', icon: <BookOpen size={16} /> },
  { id: 'resources', label: 'Resources', icon: <FileText size={16} /> },
];

const EXAMPLE_PROMPTS: LocalizedText[] = [
  {
    en: 'A web application with a React frontend, Node.js backend API, PostgreSQL database, and blob storage for images',
    ja: 'Reactフロントエンド、Node.jsバックエンドAPI、PostgreSQLデータベース、画像用Blob Storageを使用するWebアプリケーション',
  },
  {
    en: 'A microservices architecture with Container Apps, API gateway, message queue, and Redis cache',
    ja: 'Container Apps、API Gateway、メッセージ キュー、Redis Cacheを使用するマイクロサービス アーキテクチャ',
  },
  {
    en: 'A zero trust enterprise network with Azure Firewall, Application Gateway with WAF, Private Link for PaaS, Bastion for VM access, and Microsoft Entra ID with Conditional Access',
    ja: 'Azure Firewall、WAF付きApplication Gateway、PaaS向けPrivate Link、VMアクセス向けBastion、Conditional Access付きMicrosoft Entra IDを使用するZero Trustエンタープライズ ネットワーク',
  },
  {
    en: 'A HIPAA-compliant healthcare platform with FHIR APIs, de-identification pipeline, and audit logging',
    ja: 'FHIR APIs、匿名化パイプライン、監査ログを備えたHIPAA準拠の医療プラットフォーム',
  },
  {
    en: 'An event-driven order processing system handling 50K orders/hour with Event Hubs, Functions, and Cosmos DB',
    ja: 'Event Hubs、Azure Functions、Cosmos DBを使用し、1時間あたり5万件の注文を処理するイベント駆動型の注文処理システム',
  },
];

const FEATURES: { icon: React.ReactNode; title: string; body: string }[] = [
  { icon: <Sparkles size={18} />, title: 'AI Architecture Generation', body: 'Describe your architecture in plain English and pick a model — the AI lays out a complete, grouped diagram. Upload an existing diagram image to recreate it as editable nodes.' },
  { icon: <MessageSquare size={18} />, title: 'Architecture Chat', body: 'Your fastest starting point — the Chat panel opens automatically with ready‑made starter patterns. Build or refine in plain English (“add Front Door with WAF”, then “make it zone‑redundant”), and it even suggests the most valuable missing pieces based on what’s on the canvas. Every change is auto‑saved to version history.' },
  { icon: <ImageIcon size={18} />, title: 'Image Import', body: 'Drop a screenshot, whiteboard photo, or exported PNG into the AI generator and it reconstructs the architecture with proper Azure service mapping.' },
  { icon: <PenTool size={18} />, title: 'Blueprint Diagrams', body: 'Generate a hand‑drawn, whiteboard‑style blueprint PNG with numbered, labeled flows — great for presentations. Use Topology, Blueprint, or Both.' },
  { icon: <ShieldCheck size={18} />, title: 'Well‑Architected Validation', body: 'Score your design across the five WAF pillars, review findings, and regenerate an improved architecture from selected recommendations.' },
  { icon: <GitCompare size={18} />, title: 'Multi‑Model Comparison', body: 'Run the same prompt across models and compare service counts, tokens, latency, and WAF scores side‑by‑side, then apply the winner.' },
  { icon: <DollarSign size={18} />, title: 'Cost Estimation', body: 'See estimated monthly cost across 9 Azure regions with per‑service breakdowns, and export CSV / multi‑format cost reports.' },
  { icon: <FileText size={18} />, title: 'Deployment Guides', body: 'Generate step‑by‑step deployment docs with Bicep templates — now grounded in official Microsoft Learn docs with citations.' },
  { icon: <Download size={18} />, title: 'Export Anywhere', body: 'Export to PNG, SVG, Visio (VSDX), Draw.io, PowerPoint, interactive HTML, a Markdown workflow narrative, or a re‑importable JSON manifest — plus CSV/ZIP cost reports.' },
  { icon: <Map size={18} />, title: 'Navigate Large Diagrams', body: 'Drag or scroll the mini‑map (bottom‑right) to move around, click Fit‑to‑view to frame everything, or use Focus mode / Hide Toolbar to maximize canvas space for demos.' },
  { icon: <Mic size={18} />, title: 'Avatar Presenter', body: 'Have a photorealistic avatar narrate the model comparison or walk through each workflow step with live closed captions.' },
  { icon: <History size={18} />, title: 'Version History', body: 'A snapshot is auto‑saved before each AI regeneration. Browse and restore previous versions at any time.' },
];

const HelpLearnPanel: React.FC<HelpLearnPanelProps> = ({ isOpen, onClose }) => {
  const { t, translate, language } = useLanguage();
  const [section, setSection] = useState<SectionId>('quick-start');
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);

  // Telemetry: panel open.
  useEffect(() => {
    if (isOpen) trackHelpOpened('quick-start');
  }, [isOpen]);

  // Close on Escape.
  useEffect(() => {
    if (!isOpen) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [isOpen, onClose]);

  const goToSection = (id: SectionId) => {
    setSection(id);
    trackHelpOpened(id);
  };

  const copyPrompt = async (text: string, index: number) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedIndex(index);
      setTimeout(() => setCopiedIndex((c) => (c === index ? null : c)), 1500);
    } catch {
      /* clipboard may be blocked; ignore */
    }
  };

  if (!isOpen) return null;

  return (
    <div className="help-overlay" role="dialog" aria-modal="true" aria-label={t("Help and Learn")} onClick={onClose}>
      <div className="help-modal" onClick={(e) => e.stopPropagation()}>
        <div className="help-header">
          <div className="help-title">
            <BookOpen size={20} />
            <span>{t("Help & Learn")}</span>
          </div>
          <button className="help-close" onClick={onClose} title={t("Close")} aria-label={t("Close help")}>
            <X size={18} />
          </button>
        </div>

        <div className="help-body">
          <nav className="help-nav" aria-label={t("Help sections")}>
            {SECTIONS.map((s) => (
              <button
                key={s.id}
                className={`help-nav-item${section === s.id ? ' active' : ''}`}
                onClick={() => goToSection(s.id)}
              >
                {s.icon}
                <span>{translate(s.label)}</span>
              </button>
            ))}
          </nav>

          <div className="help-content">
            {section === 'quick-start' && (
              <div className="help-section">
                <h3>{t("Get your first diagram in 3 steps")}</h3>
                <ol className="help-steps">
                  <li>
                    {localize(language, {
                      en: 'Start it. Describe your architecture in the Chat panel (it opens automatically), or open Generate with AI to enter a prompt or upload a diagram image, then select a model.',
                      ja: '開始します。自動的に開くChatパネルにアーキテクチャを入力するか、AIで生成を開いてプロンプトまたは図の画像を指定し、モデルを選択します。',
                    })}
                  </li>
                  <li>
                    {localize(language, {
                      en: 'Refine it. Keep chatting to request changes, such as “add a load balancer in front of the VMs.” The canvas updates, and every change is saved to Version History.',
                      ja: '調整します。Chatで「VMの前にLoad Balancerを追加」などの変更を依頼します。キャンバスが更新され、各変更はVersion Historyに保存されます。',
                    })}
                  </li>
                  <li>
                    {localize(language, {
                      en: 'Use it. Validate the design against the Well-Architected Framework, estimate costs, generate a Deployment Guide, or export to PNG, Visio, Draw.io, or PowerPoint.',
                      ja: '活用します。Well-Architected Frameworkに照らして設計を検証し、コストを見積もり、Deployment Guideを生成するか、PNG、Visio、Draw.io、PowerPointへエクスポートします。',
                    })}
                  </li>
                </ol>
                <p className="help-callout">
                  {localize(language, { en: '💡 New here? Review the', ja: '💡 初めての方は、' })}{' '}
                  <button className="help-link" onClick={() => goToSection('features')}>{t("Feature Tour")}</button>
                  {' '}{localize(language, { en: 'and try an', ja: 'を確認して、' })}{' '}
                  <button className="help-link" onClick={() => goToSection('prompts')}>{t("Example Prompt")}</button>
                  {localize(language, { en: '.', ja: 'もお試しください。' })}
                </p>
              </div>
            )}

            {section === 'features' && (
              <div className="help-section">
                <h3>{t("What the tool can do")}</h3>
                <div className="help-feature-list">
                  {FEATURES.map((f) => (
                    <div key={f.title} className="help-feature">
                      <div className="help-feature-icon">{f.icon}</div>
                      <div>
                        <div className="help-feature-title">{translate(f.title)}</div>
                        <div className="help-feature-body">{translate(f.body)}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {section === 'prompts' && (
              <div className="help-section">
                <h3>{t("Example prompts")}</h3>
                <p className="help-muted">
                  {localize(language, {
                    en: 'Select a prompt to copy it, then paste it into Generate with AI.',
                    ja: 'プロンプトを選択してコピーし、AIで生成に貼り付けます。',
                  })}
                </p>
                <div className="help-prompts">
                  {EXAMPLE_PROMPTS.map((p, i) => {
                    const prompt = localize(language, p);
                    return (
                    <button key={i} className="help-prompt" onClick={() => copyPrompt(prompt, i)} title={t("Copy prompt")}>
                      <span>{prompt}</span>
                      {copiedIndex === i ? <Check size={16} className="help-prompt-icon copied" /> : <Copy size={16} className="help-prompt-icon" />}
                    </button>
                    );
                  })}
                </div>
              </div>
            )}

            {section === 'tips' && (
              <div className="help-section">
                <h3>{t("Tips & FAQ")}</h3>
                <div className="help-faq">
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("Which model should I use?")}</div>
                    <div className="help-faq-a">{t("Start with GPT-5.6 Sol at Low reasoning for balanced quality. Increase reasoning effort for complex enterprise architectures; lower levels are faster and cheaper for quick drafts.")}</div>
                  </div>
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("My diagram isn’t quite right — what now?")}</div>
                    <div className="help-faq-a">
                      {localize(language, {
                        en: 'Use Chat to make targeted edits instead of regenerating from scratch. You can also drag nodes, edit connections, and resize groups directly on the canvas.',
                        ja: 'Chatを使うと、最初から再生成せず対象を絞って編集できます。キャンバス上でノードの移動、接続の編集、グループのサイズ変更もできます。',
                      })}
                    </div>
                  </div>
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("I can’t see my whole diagram — how do I navigate?")}</div>
                    <div className="help-faq-a">
                      {localize(language, {
                        en: 'Scroll to zoom, right-click and drag to pan, or drag the mini-map in the lower-right. Use Fit-to-view to frame the whole diagram, and Focus or Hide Toolbar for more space.',
                        ja: 'スクロールで拡大・縮小し、右クリック ドラッグまたは右下のミニマップで移動します。「全体を表示」で図全体を収め、「フォーカス」または「ツールバーを非表示」で表示領域を広げられます。',
                      })}
                    </div>
                  </div>
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("Can I import existing infrastructure?")}</div>
                    <div className="help-faq-a">{t("Yes — import a Bicep, Terraform, or ARM template, or upload a diagram image, and the AI reconstructs an editable diagram.")}</div>
                  </div>
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("How do I undo an AI change?")}</div>
                    <div className="help-faq-a">
                      {localize(language, {
                        en: 'Open Version History. A snapshot is saved automatically before every AI regeneration, so you can restore any prior state.',
                        ja: 'Version Historyを開きます。AIで再生成する前にスナップショットが自動保存されるため、以前の状態を復元できます。',
                      })}
                    </div>
                  </div>
                  <div className="help-faq-item">
                    <div className="help-faq-q">{t("Are deployment guides accurate?")}</div>
                    <div className="help-faq-a">{t("They’re grounded in official Microsoft Learn docs with citations. Always review commands and Bicep before running them in your environment.")}</div>
                  </div>
                </div>
              </div>
            )}

            {section === 'resources' && (
              <div className="help-section">
                <h3>{t("Resources")}</h3>
                <ul className="help-resources">
                  <li>
                    <a href="https://techcommunity.microsoft.com/blog/azurearchitectureblog/from-prompt-to-production-building-azure-architecture-diagrams-with-ai/4520336" target="_blank" rel="noopener noreferrer">
                      {' '}{t("From Prompt to Production — blog post")}{' '}</a>
                  </li>
                  <li>
                    {localize(language, {
                      en: 'Send feedback at any time with the Feedback button in the lower-right.',
                      ja: '右下の「フィードバック」ボタンから、いつでもご意見を送信できます。',
                    })}
                  </li>
                </ul>
                <p className="help-muted">{t("More guided content is on the way — see the Help & Learn plan in the repo docs.")}</p>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default HelpLearnPanel;
