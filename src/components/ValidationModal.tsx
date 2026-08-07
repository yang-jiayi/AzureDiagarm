// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useState } from 'react';
import { X, AlertTriangle, CheckCircle, Info, Download, RefreshCw, Clock, Zap, Database, Cpu, Search, Wrench, ExternalLink, Play } from 'lucide-react';
import { ArchitectureValidation, ValidationFinding, formatValidationReport } from '../services/architectureValidator';
import { generateModelFilename } from '../utils/modelNaming';
import { scoreToBand, summarizeGaps, formatGapsSummary, formatPillarGaps } from '../services/wafMaturity';
import { useValidationDisplayPrefs } from '../stores/validationDisplayStore';
import './ValidationModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';

/**
 * Props for ValidationModal component
 */
interface ValidationModalProps {
  validation: ArchitectureValidation | null; // Validation results from GPT-5.2 agent
  isOpen: boolean; // Controls modal visibility
  onClose: () => void; // Handler for closing modal
  isLoading?: boolean; // Shows loading state during validation
  isStale?: boolean; // Previous result no longer matches the modified architecture
  onApplyRecommendations?: (selectedFindings: ValidationFinding[]) => void; // Handler for applying selected recommendations
  onRevalidate?: () => void; // Optional handler to rerun validation
}

/**
 * Modal displaying Azure Well-Architected Framework validation results.
 * Shows overall score, pillar-specific assessments, findings, and quick wins.
 * Includes download functionality for markdown report.
 */
const ValidationModal: React.FC<ValidationModalProps> = ({ validation, isOpen, onClose, isLoading, isStale, onApplyRecommendations, onRevalidate }) => {
  const { t, translate, language } = useLanguage();
  // Track selected findings for applying recommendations
  const [selectedFindings, setSelectedFindings] = useState<Set<string>>(new Set());
  // Display preference: show the raw 0-100 score alongside the maturity band
  const [displayPrefs, setDisplayPrefs] = useValidationDisplayPrefs();
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);
  useEscapeKey(isOpen, onClose);

  // Finding keys are positional (`pillar-0-0`), so a selection made against an
  // earlier result would silently carry over to a different finding in the next
  // run. Reset whenever a new validation is shown or the modal is reopened.
  useEffect(() => {
    setSelectedFindings(new Set());
  }, [validation?.timestamp, isOpen]);

  if (!isOpen) return null;

  /**
   * Toggle selection of a finding
   */
  const toggleFinding = (findingKey: string) => {
    setSelectedFindings(prev => {
      const next = new Set(prev);
      if (next.has(findingKey)) {
        next.delete(findingKey);
      } else {
        next.add(findingKey);
      }
      return next;
    });
  };

  /**
   * Get all findings as a flat array with unique keys
   */
  const getAllFindings = (): Array<ValidationFinding & { key: string }> => {
    if (!validation) return [];
    
    const findings: Array<ValidationFinding & { key: string }> = [];
    
    // Add pillar findings
    validation.pillars.forEach((pillar, pIndex) => {
      pillar.findings.forEach((finding, fIndex) => {
        findings.push({
          ...finding,
          key: `pillar-${pIndex}-${fIndex}`
        });
      });
    });
    
    // Add quick wins
    validation.quickWins.forEach((win, wIndex) => {
      findings.push({
        ...win,
        key: `quickwin-${wIndex}`
      });
    });
    
    return findings;
  };

  /**
   * Apply selected recommendations
   */
  const handleApplyRecommendations = () => {
    const allFindings = getAllFindings();
    const selected = allFindings.filter(f => selectedFindings.has(f.key));
    
    if (onApplyRecommendations && selected.length > 0) {
      setSelectedFindings(new Set());
      onApplyRecommendations(selected);
    }
  };

  /**
   * Returns appropriate icon component for finding severity level
   */
  /**
   * Returns appropriate icon component for finding severity level
   */
  const getSeverityIcon = (severity: ValidationFinding['severity']) => {
    switch (severity) {
      case 'critical': return <AlertTriangle className="severity-icon critical" />;
      case 'high': return <AlertTriangle className="severity-icon high" />;
      case 'medium': return <Info className="severity-icon medium" />;
      case 'low': return <CheckCircle className="severity-icon low" />;
    }
  };

  const renderFindingDetails = (finding: ValidationFinding) => (
    <div className="finding-content">
      <p className="finding-issue"><strong>{t("Issue:")}</strong> {finding.issue}</p>
      <p className="finding-recommendation">
        <strong>{t("Recommendation:")}</strong> {finding.recommendation}
      </p>
      {finding.evidence && finding.evidence.length > 0 && (
        <div className="finding-evidence">
          <strong>
            <Search size={14} />
            {localize(language, { en: 'Diagram evidence', ja: '図面上の根拠' })}
          </strong>
          <ul>
            {finding.evidence.map((item, index) => <li key={index}>{item}</li>)}
          </ul>
        </div>
      )}
      {finding.remediation && finding.remediation.length > 0 && (
        <div className="finding-remediation">
          <strong>
            <Wrench size={14} />
            {localize(language, { en: 'Remediation steps', ja: '修正手順' })}
          </strong>
          <ol>
            {finding.remediation.map((step, index) => <li key={index}>{step}</li>)}
          </ol>
        </div>
      )}
      {finding.resources && finding.resources.length > 0 && (
        <p className="finding-resources">
          <strong>{t("Affected:")}</strong> {finding.resources.join(', ')}
        </p>
      )}
      <div className="finding-action-row">
        {finding.referenceUrl && (
          <a href={finding.referenceUrl} target="_blank" rel="noopener noreferrer">
            <ExternalLink size={13} />
            {localize(language, { en: 'Microsoft Learn', ja: 'Microsoft Learn' })}
          </a>
        )}
        {onApplyRecommendations && finding.applyAction && (
          <button
            type="button"
            onClick={() => onApplyRecommendations([finding])}
            title={finding.applyAction.serviceType
              ? localize(language, {
                  en: `Apply by adding or configuring ${finding.applyAction.serviceType}`,
                  ja: `${finding.applyAction.serviceType}を追加または構成して適用`,
                })
              : finding.applyAction.label}
          >
            <Play size={13} />
            {finding.applyAction.label}
          </button>
        )}
      </div>
    </div>
  );

  /**
   * Downloads validation results as markdown file with timestamp
   */
  const handleDownload = () => {
    if (!validation) return;
    const ts = new Date(validation.timestamp).getTime();
    const report = formatValidationReport(validation);
    
    // Download markdown report
    const blob = new Blob([report], { type: 'text/markdown' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = generateModelFilename('architecture-validation', 'md', ts);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    // Download diagram PNG if captured
    if (validation.diagramImageDataUrl) {
      const imgLink = document.createElement('a');
      imgLink.href = validation.diagramImageDataUrl;
      imgLink.download = generateModelFilename('architecture-validation-diagram', 'png', ts);
      document.body.appendChild(imgLink);
      imgLink.click();
      document.body.removeChild(imgLink);
    }
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal-content validation-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("🔍 Architecture Validation")}
        tabIndex={-1}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <h2>{t("🔍 Architecture Validation")}</h2>
          <div className="modal-header-actions">
            {validation && (
              <label className="score-toggle" title={t("Show the underlying 0-100 numeric score alongside the maturity band")}>
                <input
                  type="checkbox"
                  checked={displayPrefs.showNumericScore}
                  onChange={(e) => setDisplayPrefs({ showNumericScore: e.target.checked })}
                />
                <span>{t("Show numeric score")}</span>
              </label>
            )}
            <button className="modal-close" onClick={onClose} title={t("Close")} aria-label={t("Close")}>
              <X size={24} />
            </button>
          </div>
        </div>

        {isLoading ? (
          <div className="modal-loading">
            <div className="spinner"></div>
            <div className="loading-content">
              <h3>{t("Analyzing architecture against Azure Well-Architected Framework...")}</h3>
              <p className="loading-description">
                {' '}{t("Running hybrid analysis: instant rule-based checks against")}{' '}{'>'}{t("65 curated WAF rules, followed by AI-powered contextual refinement for architecture-specific insights.")}{' '}</p>
              <div className="pillars-info">
                <h4>{t("Five Pillars of Azure Well-Architected Framework:")}</h4>
                <ul className="pillars-list">
                  <li>
                    <strong>{t("Cost Optimization")}</strong> {' '}{t("- Manage costs to maximize value delivered")}{' '}</li>
                  <li>
                    <strong>{t("Operational Excellence")}</strong> {' '}{t("- Operations processes that keep systems running in production")}{' '}</li>
                  <li>
                    <strong>{t("Performance Efficiency")}</strong> {' '}{t("- Ability to scale and adapt to changes in load")}{' '}</li>
                  <li>
                    <strong>{t("Reliability")}</strong> {' '}{t("- Ability to recover from failures and continue to function")}{' '}</li>
                  <li>
                    <strong>{t("Security")}</strong> {' '}{t("- Protect applications and data from threats")}{' '}</li>
                </ul>
              </div>
              <p className="validation-dismiss-hint">
                {' '}{t("You may close this panel at any time — once complete, reopen your results using the")}{' '}<strong>{t("Validation Score")}</strong> {' '}{t("button in the toolbar.")}{' '}</p>
            </div>
          </div>
        ) : validation ? (
          <>
            <div className="modal-body">
            {isStale && (
              <div className="validation-stale-notice" role="status">
                <AlertTriangle size={18} />
                <span>
                  <strong>{translate('Architecture changed.')}</strong> {translate('These results describe the version before recommendations were applied. Revalidate to assess the updated diagram.')}
                </span>
              </div>
            )}
            {/* Scope note - sets expectations for workshop facilitators and users */}
            <p className="validation-scope-note">
              <strong>{t("Scope:")}</strong> {' '}{t("Designed for")}{' '}<strong>{t("greenfield Azure")}</strong> {' '}{t("designs. This is a diagram-only, design-time signal to guide new architectures — not an audit of a deployed environment, and not for direct deployment into existing, complex environments without further review.")}{' '}</p>
            {/* Overall Assessment - maturity band (numeric score optional) */}
            {(() => {
              const overall = scoreToBand(validation.overallScore);
              const allFindings = validation.pillars.flatMap(p => p.findings);
              const gaps = summarizeGaps(allFindings);
              return (
            <div className="validation-score">
              <div 
                className="score-circle" 
                style={{ 
                  background: `conic-gradient(${overall.color} ${validation.overallScore * 3.6}deg, #e5e7eb 0deg)` 
                }}
                title={displayPrefs.showNumericScore ? undefined : t("Diagram-only, design-time signal — not a deployed-environment audit")}
              >
                <div className="score-inner">
                  {displayPrefs.showNumericScore ? (
                    <>
                      <span className="score-value">{validation.overallScore}</span>
                      <span className="score-label">{t("/100")}</span>
                    </>
                  ) : (
                    <span className="score-band-mark" style={{ color: overall.color }}>
                      {translate(overall.short)}
                    </span>
                  )}
                </div>
              </div>
              <div className="score-summary">
                <h3>{t("Overall Assessment")}</h3>
                <div className="maturity-headline">
                  <span className="maturity-band-pill" style={{ borderColor: overall.color, color: overall.color }}>
                    {translate(overall.label)}
                  </span>
                  <span className="gaps-summary">{formatGapsSummary(gaps, language)}</span>
                  {displayPrefs.showNumericScore && (
                    <span className="numeric-score-aside">{validation.overallScore}{t("/100")}</span>
                  )}
                </div>
                <p>{validation.summary}</p>
                {validation.metrics && (
                  <div className="ai-metrics-validation">
                    <span className="metric">
                      <Clock size={14} />
                      {(validation.metrics.elapsedTimeMs / 1000).toFixed(1)}{t("s")}{' '}</span>
                    <span className="metric">
                      <Zap size={14} />
                      {validation.metrics.promptTokens.toLocaleString()} {' '}{t("in →")}{' '}{validation.metrics.completionTokens.toLocaleString()} {' '}{t("out (")}{validation.metrics.totalTokens.toLocaleString()} {' '}{t("total)")}{' '}</span>
                    {(validation as any).hybridMetadata && (
                      <span className="metric hybrid-metric">
                        <Database size={14} />
                        {(validation as any).hybridMetadata.localFindings} {' '}{t("local rules (")}{(validation as any).hybridMetadata.localElapsedMs}{t("ms) + AI refinement")}{' '}</span>
                    )}
                  </div>
                )}
              </div>
            </div>
              );
            })()}

            {/* Five Pillars Section - Individual assessments for each WAF pillar */}
            <div className="pillars-section">
              <h3>{t("Five Pillars Assessment")}</h3>
              {validation.pillars.map((pillar, index) => {
                const pillarBand = scoreToBand(pillar.score);
                const pillarGaps = summarizeGaps(pillar.findings);
                return (
                <div key={index} className="pillar-card">
                  <div className="pillar-header">
                    <h4>{translate(pillar.pillar)}</h4>
                    <div className="pillar-assessment">
                      <span
                        className="maturity-band-pill small"
                        style={{ borderColor: pillarBand.color, color: pillarBand.color }}
                      >
                        {translate(pillarBand.label)}
                      </span>
                      <span className="pillar-gaps">{formatPillarGaps(pillarGaps, language)}</span>
                      {displayPrefs.showNumericScore && (
                        <span 
                          className="pillar-score"
                          style={{ color: pillarBand.color }}
                        >
                          {pillar.score}{t("/100")}{' '}</span>
                      )}
                    </div>
                  </div>
                  
                  {pillar.findings.length > 0 && (
                    <div className="findings-list">
                      {pillar.findings.map((finding, fIndex) => {
                        const findingKey = `pillar-${index}-${fIndex}`;
                        const isSelected = selectedFindings.has(findingKey);
                        
                        return (
                          <div key={fIndex} className={`finding-item severity-${finding.severity} ${isSelected ? 'selected' : ''}`}>
                            <div className="finding-header">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={() => toggleFinding(findingKey)}
                                className="finding-checkbox"
                              />
                              {getSeverityIcon(finding.severity)}
                              <span className="finding-category">{finding.category}</span>
                              <span className={`severity-badge ${finding.severity}`}>
                                {translate(finding.severity)}
                              </span>
                              {finding.source && (
                                <span className={`source-badge ${finding.source}`}>
                                  {finding.source === 'rule-based' ? <Database size={12} /> : <Cpu size={12} />}
                                  {finding.source === 'rule-based'
                                    ? localize(language, { en: 'Rule', ja: 'ルール' })
                                    : 'AI'}
                                </span>
                              )}
                            </div>
                            {renderFindingDetails(finding)}
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
                );
              })}
            </div>

            {/* Quick Wins Section - High-priority actionable items */}
            {validation.quickWins.length > 0 && (
              <div className="quickwins-section">
                <h3>{t("⚡ Quick Wins")}</h3>
                <div className="quickwins-list">
                  {validation.quickWins.map((win, index) => (
                    <div key={index} className="quickwin-item">
                      <div className="quickwin-header">
                        <CheckCircle className="quickwin-icon" />
                        <span className="quickwin-category">{win.category}</span>
                      </div>
                      {renderFindingDetails(win)}
                    </div>
                  ))}
                </div>
              </div>
            )}

          </div>

          {/* Action Buttons - Sticky footer with download and close */}
          <div className="modal-actions">
            <button className="btn-secondary" onClick={handleDownload}>
              <Download size={18} />
              {' '}{t("Download Report")}{' '}</button>
            {onRevalidate && (
              <button className="btn-secondary" onClick={onRevalidate} disabled={!!isLoading} title={t("Run validation again")}>
                <RefreshCw size={18} />
                {' '}{t("Revalidate")}{' '}</button>
            )}
            {selectedFindings.size > 0 && onApplyRecommendations && (
              <button className="btn-success" onClick={handleApplyRecommendations}>
                <RefreshCw size={18} />
                {localize(language, {
                  en: `Apply ${selectedFindings.size} recommendation${selectedFindings.size > 1 ? 's' : ''}`,
                  ja: `${selectedFindings.size}件の推奨事項を適用`,
                })}
              </button>
            )}
            <button className="btn-primary" onClick={onClose}>
              {' '}{t("Hide")}{' '}</button>
          </div>
        </>
        ) : (
          <div className="modal-empty">
            <p>{t("No validation results available.")}</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default ValidationModal;
