// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useMemo, useRef, useState } from 'react';
import { Check, Copy, Download, FileCode, GitCompare, Upload, X } from 'lucide-react';
import './IaCRoundTripModal.css';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import type {
  DriftAction,
  DriftPlanSummary,
  IaCBaseline,
  IaCBaselineResource,
  IaCComparisonReport,
  MatchConfidence,
  StarterTemplate,
  StarterTemplateFormat,
} from '../services/iacRoundTrip';

interface IaCRoundTripModalProps {
  isOpen: boolean;
  onClose: () => void;
  baseline: IaCBaseline | null;
  comparison: IaCComparisonReport | null;
  driftPlan: DriftPlanSummary | null;
  onImportDriftPlan: (event: React.ChangeEvent<HTMLInputElement>) => void;
  onClearDriftPlan: () => void;
  onDownloadStarter: (format: StarterTemplateFormat) => void;
  bicepStarter: StarterTemplate;
  terraformStarter: StarterTemplate;
  diagramServiceCount: number;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const IaCRoundTripModal: React.FC<IaCRoundTripModalProps> = ({
  isOpen,
  onClose,
  baseline,
  comparison,
  driftPlan,
  onImportDriftPlan,
  onClearDriftPlan,
  onDownloadStarter,
  bicepStarter,
  terraformStarter,
  diagramServiceCount,
}) => {
  const { language, t } = useLanguage();
  const [copiedKey, setCopiedKey] = useState<string | null>(null);
  const planInputRef = useRef<HTMLInputElement | null>(null);
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);

  useEscapeKey(isOpen, onClose);

  const l = (en: string, ja: string) => localize(language, { en, ja });

  const bicepWhatIfCommand = `az deployment group what-if --resource-group <resource-group-name> --template-file ${bicepStarter.fileName} --parameters @<parameters-file.json>`;
  const terraformPlanCommand = `terraform plan -out=tfplan\nterraform show -json tfplan > drift-plan.json`;

  const driftOrder = useMemo(
    () => (['create', 'update', 'delete', 'replace', 'no-op', 'other'] satisfies DriftAction[]),
    [],
  );

  if (!isOpen) return null;

  const confidenceLabel = (confidence: MatchConfidence) => {
    switch (confidence) {
      case 'exact':
        return l('Exact', '一致');
      case 'normalized':
        return l('Alias-normalized', '別名正規化');
      case 'approximate':
        return l('Approximate', '概算一致');
      case 'unmapped':
        return l('Unmapped hint', '未マップの推定');
    }
  };

  const actionLabel = (action: DriftAction) => {
    switch (action) {
      case 'create':
        return l('Create', '作成');
      case 'update':
        return l('Update', '更新');
      case 'delete':
        return l('Delete', '削除');
      case 'replace':
        return l('Replace', '再作成');
      case 'no-op':
        return l('No-op', '変更なし');
      default:
        return l('Review', '要確認');
    }
  };

  const renderResourceTitle = (resource: IaCBaselineResource) => {
    const name = resource.resourceName || resource.logicalName;
    if (resource.mappedService) {
      return `${resource.mappedService} • ${name}`;
    }
    return `${name} • ${resource.providerType}`;
  };

  const copyText = async (key: string, text: string) => {
    await navigator.clipboard.writeText(text);
    setCopiedKey(key);
    window.setTimeout(() => setCopiedKey(null), 2000);
  };

  const starterCards = [
    {
      key: 'bicep' as const,
      title: 'Bicep',
      template: bicepStarter,
      description: l('Safe design-time starter for az deployment what-if review.', 'az deployment what-if のレビュー向け安全な設計時スターターです。'),
    },
    {
      key: 'terraform' as const,
      title: 'Terraform',
      template: terraformStarter,
      description: l('Safe starter for terraform plan + show -json review.', 'terraform plan と show -json のレビュー向け安全なスターターです。'),
    },
  ];

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal-content iac-roundtrip-modal"
        role="dialog"
        aria-modal="true"
        aria-label={l('IaC round-trip and drift', 'IaC ラウンドトリップとドリフト')}
        tabIndex={-1}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-header">
          <div className="iac-roundtrip-title">
            <h2>{l('IaC Round-trip & Drift', 'IaC ラウンドトリップとドリフト')}</h2>
            <p>
              {l(
                'Deterministic baseline comparison, plan summary, and starter IaC export.',
                '決定論的なベースライン比較、プラン要約、スターター IaC エクスポートをまとめて確認できます。',
              )}
            </p>
          </div>
          <button className="modal-close" onClick={onClose} title={t('Close')} aria-label={t('Close')}>
            <X size={22} />
          </button>
        </div>

        <div className="modal-body">
          <div className="iac-roundtrip-summary">
            <div className="iac-summary-card">
              <div className="iac-summary-label">{l('Baseline', 'ベースライン')}</div>
              <div className="iac-summary-value">{baseline ? baseline.resourceCount : 0}</div>
              <div className="iac-summary-note">
                {baseline
                  ? `${baseline.formatLabel} • ${baseline.sourceFiles.length} ${l('file(s)', 'ファイル')}`
                  : l('Import a template to capture a baseline.', 'テンプレートをインポートするとベースラインを保持します。')}
              </div>
            </div>
            <div className="iac-summary-card">
              <div className="iac-summary-label">{l('Matched', '一致')}</div>
              <div className="iac-summary-value">{comparison?.matched.length || 0}</div>
              <div className="iac-summary-note">
                {comparison
                  ? `${comparison.approximateMatches} ${l('approximate', '概算一致')}`
                  : l('No baseline comparison yet.', 'まだベースライン比較はありません。')}
              </div>
            </div>
            <div className="iac-summary-card">
              <div className="iac-summary-label">{l('Missing from diagram', '図にないソース')}</div>
              <div className="iac-summary-value">{comparison?.sourceOnly.length || 0}</div>
              <div className="iac-summary-note">
                {l('Baseline resources without a current node match.', '現在の図に一致ノードがないベースライン リソースです。')}
              </div>
            </div>
            <div className="iac-summary-card">
              <div className="iac-summary-label">{l('New in design', '設計で追加')}</div>
              <div className="iac-summary-value">{comparison?.diagramOnly.length || 0}</div>
              <div className="iac-summary-note">
                {`${diagramServiceCount} ${l('diagram service node(s)', '図のサービス ノード')}`}
              </div>
            </div>
          </div>

          <section className="iac-section">
            <div className="iac-section-header">
              <h3>{l('Imported baseline', 'インポート済みベースライン')}</h3>
            </div>
            {baseline ? (
              <>
                <div className="iac-inline-list">
                  <span>{baseline.formatLabel}</span>
                  <span>•</span>
                  <span>{new Date(baseline.capturedAt).toLocaleString(language === 'ja' ? 'ja-JP' : 'en-US')}</span>
                </div>
                <ul className="iac-file-list">
                  {baseline.sourceFiles.map((file) => (
                    <li key={file.filename}>
                      <strong>{file.filename}</strong>
                      <span>{formatBytes(file.size)}</span>
                    </li>
                  ))}
                </ul>
                {baseline.warnings.length > 0 && (
                  <div className="iac-callout iac-callout-warning">
                    <strong>{l('Parsing notes', '解析メモ')}</strong>
                    <ul>
                      {baseline.warnings.map((warning) => (
                        <li key={warning}>{warning}</li>
                      ))}
                    </ul>
                  </div>
                )}
              </>
            ) : (
              <div className="iac-empty-state">
                {l(
                  'No IaC baseline is attached to this diagram yet. Import Bicep, Terraform, ARM JSON, or Terraform state to seed a round-trip baseline.',
                  'この図にはまだ IaC ベースラインがありません。Bicep、Terraform、ARM JSON、Terraform state をインポートしてラウンドトリップ ベースラインを作成してください。',
                )}
              </div>
            )}
          </section>

          <section className="iac-section">
            <div className="iac-section-header">
              <h3>{l('Round-trip comparison', 'ラウンドトリップ比較')}</h3>
            </div>
            {comparison ? (
              <div className="iac-comparison-grid">
                <div className="iac-column">
                  <h4>{l('Matched resources', '一致したリソース')}</h4>
                  {comparison.matched.length === 0 ? (
                    <div className="iac-empty-state">{l('No deterministic matches yet.', '決定論的な一致はまだありません。')}</div>
                  ) : (
                    <ul className="iac-resource-list">
                      {comparison.matched.map((match) => (
                        <li key={`${match.baseline.id}-${match.diagram.id}`} className="iac-resource-item">
                          <div className="iac-resource-top">
                            <strong>{renderResourceTitle(match.baseline)}</strong>
                            <span className={`iac-badge iac-badge-${match.confidence}`}>{confidenceLabel(match.confidence)}</span>
                          </div>
                          <div className="iac-resource-meta">
                            <span>{match.diagram.label}</span>
                            <span>•</span>
                            <span>{match.reason}</span>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
                <div className="iac-column">
                  <h4>{l('Source-only', 'ソースのみ')}</h4>
                  {comparison.sourceOnly.length === 0 ? (
                    <div className="iac-empty-state">{l('Every baseline resource has a current match.', 'すべてのベースライン リソースに現在の一致があります。')}</div>
                  ) : (
                    <ul className="iac-resource-list">
                      {comparison.sourceOnly.map((resource) => (
                        <li key={resource.id} className="iac-resource-item">
                          <strong>{renderResourceTitle(resource)}</strong>
                          <div className="iac-resource-meta">
                            <span>{resource.sourceFile}</span>
                            <span>•</span>
                            <span>{resource.notes || l('Missing from the current diagram.', '現在の図にはありません。')}</span>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
                <div className="iac-column">
                  <h4>{l('Diagram-only', '図のみ')}</h4>
                  {comparison.diagramOnly.length === 0 ? (
                    <div className="iac-empty-state">{l('No new service nodes beyond the baseline.', 'ベースラインを超える新しいサービス ノードはありません。')}</div>
                  ) : (
                    <ul className="iac-resource-list">
                      {comparison.diagramOnly.map((resource) => (
                        <li key={resource.id} className="iac-resource-item">
                          <strong>{resource.label}</strong>
                          <div className="iac-resource-meta">
                            <span>{resource.mappedService || resource.serviceName}</span>
                            <span>•</span>
                            <span>{l('New or unmatched in the current design.', '現在の設計で新規または未一致です。')}</span>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
            ) : (
              <div className="iac-empty-state">
                {l(
                  'Open this after a successful template import to compare the live design against the original IaC baseline.',
                  'テンプレートのインポート成功後に開くと、現在の設計を元の IaC ベースラインと比較できます。',
                )}
              </div>
            )}
          </section>

          <section className="iac-section">
            <div className="iac-section-header">
              <h3>{l('Deployment-plan drift import', 'デプロイ プランのドリフト取り込み')}</h3>
              <div className="iac-header-actions">
                <input
                  ref={planInputRef}
                  type="file"
                  accept=".json"
                  onChange={onImportDriftPlan}
                  style={{ display: 'none' }}
                />
                <button className="btn btn-secondary" onClick={() => planInputRef.current?.click()}>
                  <Upload size={16} />
                  {l('Import plan JSON', 'プラン JSON を取り込む')}
                </button>
                {driftPlan && (
                  <button className="btn btn-secondary" onClick={onClearDriftPlan}>
                    <X size={16} />
                    {l('Clear session plan', 'このセッションのプランを消去')}
                  </button>
                )}
              </div>
            </div>
            <div className="iac-callout">
              {l(
                'Accepted formats: Azure what-if JSON (changes / properties.changes) and Terraform terraform show -json output (resource_changes). Review every plan before apply; plans can be stale minutes later.',
                '対応形式: Azure what-if JSON（changes / properties.changes）と Terraform の terraform show -json 出力（resource_changes）。どのプランも apply 前に必ず見直してください。プランは数分で古くなる場合があります。',
              )}
            </div>
            {driftPlan ? (
              <>
                <div className="iac-inline-list">
                  <span>{driftPlan.kind === 'azure-what-if' ? 'Azure what-if' : 'Terraform plan'}</span>
                  <span>•</span>
                  <span>{driftPlan.sourceFile}</span>
                  <span>•</span>
                  <span>{new Date(driftPlan.importedAt).toLocaleString(language === 'ja' ? 'ja-JP' : 'en-US')}</span>
                </div>
                <div className="iac-drift-badges">
                  {driftOrder.map((action) => (
                    <span key={action} className={`iac-badge iac-badge-drift iac-badge-${action.replace(/[^a-z]/g, '')}`}>
                      {actionLabel(action)} {driftPlan.changeCounts[action]}
                    </span>
                  ))}
                </div>
                <ul className="iac-resource-list">
                  {driftPlan.changes.map((change) => (
                    <li key={change.id} className="iac-resource-item">
                      <div className="iac-resource-top">
                        <strong>{change.address}</strong>
                        <span className={`iac-badge iac-badge-drift iac-badge-${change.action.replace(/[^a-z]/g, '')}`}>{change.actionText}</span>
                      </div>
                      <div className="iac-resource-meta">
                        <span>{change.resourceType || l('Unknown type', '不明な型')}</span>
                        {change.resourceName && (
                          <>
                            <span>•</span>
                            <span>{change.resourceName}</span>
                          </>
                        )}
                        {change.rawActions.length > 0 && (
                          <>
                            <span>•</span>
                            <span>{change.rawActions.join(', ')}</span>
                          </>
                        )}
                      </div>
                    </li>
                  ))}
                </ul>
              </>
            ) : (
              <div className="iac-empty-state">
                {l(
                  'No deployment plan has been imported in this session.',
                  'このセッションではまだデプロイ プランを取り込んでいません。',
                )}
              </div>
            )}
          </section>

          <section className="iac-section">
            <div className="iac-section-header">
              <h3>{l('Safe review commands', '安全なレビュー コマンド')}</h3>
            </div>
            <div className="iac-command-grid">
              <div className="iac-command-card">
                <div className="iac-command-header">
                  <h4>{l('Bicep / ARM what-if', 'Bicep / ARM what-if')}</h4>
                  <button className="copy-button" onClick={() => copyText('whatif', bicepWhatIfCommand)} title={t('Copy to clipboard')}>
                    {copiedKey === 'whatif' ? <Check size={14} /> : <Copy size={14} />}
                  </button>
                </div>
                <pre>{bicepWhatIfCommand}</pre>
              </div>
              <div className="iac-command-card">
                <div className="iac-command-header">
                  <h4>{l('Terraform plan → JSON', 'Terraform plan → JSON')}</h4>
                  <button className="copy-button" onClick={() => copyText('tfplan', terraformPlanCommand)} title={t('Copy to clipboard')}>
                    {copiedKey === 'tfplan' ? <Check size={14} /> : <Copy size={14} />}
                  </button>
                </div>
                <pre>{terraformPlanCommand}</pre>
              </div>
            </div>
          </section>

          <section className="iac-section">
            <div className="iac-section-header">
              <h3>{l('Starter IaC export', 'スターター IaC エクスポート')}</h3>
            </div>
            <div className="iac-starter-grid">
              {starterCards.map((card) => (
                <div key={card.key} className="iac-starter-card">
                  <div className="iac-resource-top">
                    <strong>{card.title}</strong>
                    <span className="iac-inline-list">
                      <span>{card.template.supportedResourceCount} {l('supported', '対応')}</span>
                      <span>•</span>
                      <span>{card.template.todoCount} TODO</span>
                    </span>
                  </div>
                  <p>{card.description}</p>
                  <div className="iac-inline-list">
                    <FileCode size={15} />
                    <span>{card.template.fileName}</span>
                  </div>
                  <button className="btn btn-secondary" onClick={() => onDownloadStarter(card.key)}>
                    <Download size={16} />
                    {l('Download starter', 'スターターをダウンロード')}
                  </button>
                </div>
              ))}
            </div>
            <div className="iac-callout">
              {l(
                'Unsupported services are preserved as explicit TODO comments and are never silently dropped from the exported starter files.',
                '未対応サービスはエクスポートしたスターター ファイル内で明示的な TODO コメントとして保持され、黙って削除されることはありません。',
              )}
            </div>
          </section>
        </div>

        <div className="modal-actions">
          <button className="btn btn-secondary" onClick={onClose}>
            <GitCompare size={16} />
            {l('Close drift report', 'ドリフト レポートを閉じる')}
          </button>
        </div>
      </div>
    </div>
  );
};

export default IaCRoundTripModal;
