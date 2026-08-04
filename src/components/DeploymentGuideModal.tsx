// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState } from 'react';
import { X, Download, Copy, Check, ChevronDown, ChevronUp, FileCode, Package, Clock, Zap } from 'lucide-react';
import { DeploymentGuide, downloadDeploymentGuide, downloadBicepTemplate, downloadAllBicepTemplates, BicepModule } from '../services/deploymentGuideGenerator';
import './DeploymentGuideModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';

interface DeploymentGuideModalProps {
  guide: DeploymentGuide | null;
  isOpen: boolean;
  onClose: () => void;
  isLoading?: boolean;
}

const DeploymentGuideModal: React.FC<DeploymentGuideModalProps> = ({ guide, isOpen, onClose, isLoading }) => {
  const { t, language } = useLanguage();
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [expandedSections, setExpandedSections] = useState<Set<number>>(new Set([0]));
  const [expandedBicep, setExpandedBicep] = useState<Set<number>>(new Set([0]));
  const dialogRef = useModalFocus<HTMLDivElement>(isOpen);
  useEscapeKey(isOpen, onClose);

  if (!isOpen) return null;

  const handleCopy = async (text: string, index: number) => {
    await navigator.clipboard.writeText(text);
    setCopiedIndex(index);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const handleDownload = () => {
    if (!guide) return;
    downloadDeploymentGuide(guide);
  };

  const handleDownloadBicep = (template: BicepModule) => {
    downloadBicepTemplate(template);
  };

  const handleDownloadAllBicep = () => {
    if (!guide) return;
    downloadAllBicepTemplates(guide);
  };

  const toggleSection = (index: number) => {
    const newExpanded = new Set(expandedSections);
    if (newExpanded.has(index)) {
      newExpanded.delete(index);
    } else {
      newExpanded.add(index);
    }
    setExpandedSections(newExpanded);
  };

  const toggleBicep = (index: number) => {
    const newExpanded = new Set(expandedBicep);
    if (newExpanded.has(index)) {
      newExpanded.delete(index);
    } else {
      newExpanded.add(index);
    }
    setExpandedBicep(newExpanded);
  };

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal-content deployment-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Deployment Guide")}
        tabIndex={-1}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="modal-header">
          <h2>{t("📋 Deployment Guide")}</h2>
          <button className="modal-close" onClick={onClose} title={t("Close")} aria-label={t("Close")}>
            <X size={24} />
          </button>
        </div>

        {isLoading ? (
          <div className="modal-loading">
            <div className="spinner"></div>
            <p>{t("Generating comprehensive deployment guide...")}</p>
          </div>
        ) : guide ? (
          <>
            <div className="modal-body">
            {/* Title and Overview */}
            <div className="guide-header">
              <h3>{guide.title}</h3>
              <p className="guide-overview">{guide.overview}</p>
              <div className="guide-meta">
                <span className="meta-item">
                  {' '}{t("⏱️ Estimated Time:")}{' '}<strong>{guide.estimatedTime}</strong>
                </span>
                {guide.metrics && (
                  <span className="meta-item ai-metrics-inline">
                    <Clock size={14} />
                    {' '}{t("Generated in")}{' '}{(guide.metrics.elapsedTimeMs / 1000).toFixed(1)}{t("s")}{' '}<Zap size={14} style={{ marginLeft: '12px' }} />
                    {guide.metrics.promptTokens.toLocaleString()} {' '}{t("in →")}{' '}{guide.metrics.completionTokens.toLocaleString()} {' '}{t("out (")}{guide.metrics.totalTokens.toLocaleString()} {' '}{t("tokens)")}{' '}</span>
                )}
              </div>
            </div>

            {/* Prerequisites */}
            <div className="guide-section">
              <h4>{t("✅ Prerequisites")}</h4>
              <ul className="prerequisites-list">
                {guide.prerequisites.map((prereq, index) => (
                  <li key={index}>{prereq}</li>
                ))}
              </ul>
            </div>

            {/* Deployment Steps */}
            <div className="guide-section">
              <h4>{t("🚀 Deployment Steps")}</h4>
              <div className="steps-list">
                {guide.deploymentSteps.map((step, index) => (
                  <div key={index} className="step-card">
                    <button
                      type="button"
                      className="step-header"
                      onClick={() => toggleSection(index)}
                      aria-expanded={expandedSections.has(index)}
                      aria-controls={`deployment-step-${index}`}
                      id={`deployment-step-toggle-${index}`}
                    >
                      <span className="step-title">
                        <span className="step-number">{index + 1}</span>
                        <span>{step.title}</span>
                      </span>
                      {expandedSections.has(index) ? (
                        <ChevronUp size={20} />
                      ) : (
                        <ChevronDown size={20} />
                      )}
                    </button>

                    {expandedSections.has(index) && (
                      <div
                        className="step-content"
                        id={`deployment-step-${index}`}
                        role="region"
                        aria-labelledby={`deployment-step-toggle-${index}`}
                      >
                        <p className="step-description">{step.description}</p>
                        
                        {step.commands && step.commands.length > 0 && (
                          <div className="commands-section">
                            <div className="commands-header">
                              <span>{t("Commands")}</span>
                            </div>
                            {step.commands.map((cmd, cmdIndex) => (
                              <div key={cmdIndex} className="command-block">
                                <pre>{cmd}</pre>
                                <button
                                  type="button"
                                  className="copy-button"
                                  onClick={() => handleCopy(cmd, index * 100 + cmdIndex)}
                                  title={t("Copy to clipboard")}
                                  aria-label={t("Copy to clipboard")}
                                >
                                  {copiedIndex === index * 100 + cmdIndex ? (
                                    <Check size={16} />
                                  ) : (
                                    <Copy size={16} />
                                  )}
                                </button>
                              </div>
                            ))}
                          </div>
                        )}

                        {step.notes && (
                          <div className="step-notes">
                            <strong>{t("📝 Note:")}</strong> {step.notes}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>

            {/* Configuration */}
            {guide.configuration && guide.configuration.length > 0 && (
              <div className="guide-section">
                <h4>{t("⚙️ Configuration")}</h4>
                {guide.configuration.map((section, sectionIndex) => (
                  <div key={sectionIndex} className="configuration-section">
                    <h5>{section.section}</h5>
                    <div className="configuration-table">
                      <table>
                        <thead>
                          <tr>
                            <th>{t("Setting")}</th>
                            <th>{t("Value")}</th>
                            <th>{t("Description")}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {section.settings.map((config, index) => (
                            <tr key={index}>
                              <td className="config-key">{config.name}</td>
                              <td className="config-value">
                                <code>{config.value}</code>
                              </td>
                              <td className="config-description">{config.description}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* Post-Deployment */}
            {guide.postDeployment && guide.postDeployment.length > 0 && (
              <div className="guide-section">
                <h4>{t("✔️ Post-Deployment Validation")}</h4>
                <ul className="validation-list">
                  {guide.postDeployment.map((item, index) => (
                    <li key={index}>
                      <input type="checkbox" id={`validation-${index}`} />
                      <label htmlFor={`validation-${index}`}>{item}</label>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {/* References — Microsoft Learn grounding (Phase 1) */}
            {guide.groundingSources && guide.groundingSources.length > 0 && (
              <div className="guide-section grounding-section">
                <h4>{t("📚 Grounded with Microsoft Learn")}</h4>
                <p className="grounding-note">
                  {' '}{t("This guide was informed by the following official documentation:")}{' '}</p>
                <ul className="grounding-list">
                  {guide.groundingSources.map((src, index) => (
                    <li key={index}>
                      <a href={src.url} target="_blank" rel="noopener noreferrer">{src.title}</a>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            {/* Bicep Templates - Infrastructure as Code */}
            {guide.bicepTemplates && guide.bicepTemplates.length > 0 && (
              <div className="guide-section bicep-section">
                <div className="bicep-header">
                  <h4><FileCode size={20} /> {' '}{t("Infrastructure as Code (Bicep)")}</h4>
                  <button 
                    className="btn-download-all-bicep"
                    onClick={handleDownloadAllBicep}
                    title={t("Download all Bicep templates")}
                  >
                    <Package size={16} />
                    {' '}{t("Download All Templates")}{' '}</button>
                </div>
                <p className="bicep-description">
                  {' '}{t("Production-ready Bicep templates for automated infrastructure deployment. Deploy with:")}{' '}<code>{t("az deployment group create --resource-group <rg-name> --template-file main.bicep")}</code>
                </p>
                <div className="bicep-templates-list">
                  {guide.bicepTemplates.map((template, index) => (
                    <div key={index} className="bicep-template-card">
                      <div className="bicep-template-header">
                        <button
                          type="button"
                          className="bicep-template-toggle"
                          onClick={() => toggleBicep(index)}
                          aria-expanded={expandedBicep.has(index)}
                          aria-controls={`bicep-template-${index}`}
                          id={`bicep-template-toggle-${index}`}
                        >
                          <span className="bicep-template-info">
                            <FileCode size={18} className="bicep-icon" />
                            <span className="bicep-template-meta">
                              <span className="bicep-template-name">{template.name}</span>
                              <span className="bicep-template-filename">{template.filename}</span>
                            </span>
                          </span>
                          {expandedBicep.has(index) ? (
                            <ChevronUp size={20} />
                          ) : (
                            <ChevronDown size={20} />
                          )}
                        </button>
                        <div className="bicep-template-actions">
                          <button
                            type="button"
                            className="btn-download-bicep"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleDownloadBicep(template);
                            }}
                            title={localize(language, {
                              en: `Download ${template.filename}`,
                              ja: `${template.filename} をダウンロード`,
                            })}
                            aria-label={localize(language, {
                              en: `Download ${template.filename}`,
                              ja: `${template.filename} をダウンロード`,
                            })}
                          >
                            <Download size={14} />
                          </button>
                        </div>
                      </div>

                      {expandedBicep.has(index) && (
                        <div
                          className="bicep-template-content"
                          id={`bicep-template-${index}`}
                          role="region"
                          aria-labelledby={`bicep-template-toggle-${index}`}
                        >
                          <p className="bicep-template-description">{template.description}</p>
                          <div className="bicep-code-block">
                            <div className="bicep-code-header">
                              <span>{template.filename}</span>
                              <button
                                type="button"
                                className="copy-button"
                                onClick={() => handleCopy(template.content, 1000 + index)}
                                title={t("Copy to clipboard")}
                                aria-label={t("Copy to clipboard")}
                              >
                                {copiedIndex === 1000 + index ? (
                                  <Check size={14} />
                                ) : (
                                  <Copy size={14} />
                                )}
                              </button>
                            </div>
                            <pre className="bicep-code">{template.content}</pre>
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* Troubleshooting */}
            {guide.troubleshooting && guide.troubleshooting.length > 0 && (
              <div className="guide-section troubleshooting-section">
                <h4>{t("🔧 Troubleshooting")}</h4>
                <div className="troubleshooting-list">
                  {guide.troubleshooting.map((item, index) => (
                    <div key={index} className="troubleshooting-item">
                      <div className="troubleshooting-problem">
                        <strong>{t("Problem:")}</strong> {item.issue}
                      </div>
                      <div className="troubleshooting-solution">
                        <strong>{t("Solution:")}</strong> {item.solution}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            </div>

            {/* Actions - Fixed at bottom */}
            <div className="modal-actions">
              <button className="btn-secondary" onClick={handleDownload}>
                <Download size={18} />
                {' '}{t("Download Guide")}{' '}</button>
              <button className="btn-primary" onClick={onClose}>
                {' '}{t("Close")}{' '}</button>
            </div>
          </>
        ) : (
          <div className="modal-empty">
            <p>{t("No deployment guide available.")}</p>
          </div>
        )}
      </div>
    </div>
  );
};

export default DeploymentGuideModal;
