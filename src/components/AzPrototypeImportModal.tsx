// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useState, useCallback } from 'react';
import { Terminal, Upload, X, AlertCircle, Check } from 'lucide-react';
import { importFromAzPrototype, type ImportResult } from '../services/azPrototypeService';
import './AzPrototypeExportModal.css';
import './AzPrototypeImportModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import ModalScaffold from './ModalScaffold';

export interface AzPrototypeImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  /** Called when the user confirms the import. Receives the parsed architecture. */
  onImport: (result: ImportResult) => void;
}

type ImportStage = 'select' | 'preview' | 'error';

export default function AzPrototypeImportModal({
  isOpen,
  onClose,
  onImport,
}: AzPrototypeImportModalProps) {
  const { t, translate, language } = useLanguage();
  const [stage, setStage] = useState<ImportStage>('select');
  const [importResult, setImportResult] = useState<ImportResult | null>(null);
  const [errorMessage, setErrorMessage] = useState('');
  const [fileName, setFileName] = useState('');

  const reset = useCallback(() => {
    setStage('select');
    setImportResult(null);
    setErrorMessage('');
    setFileName('');
  }, []);

  const handleClose = useCallback(() => {
    reset();
    onClose();
  }, [onClose, reset]);

  const handleFileSelect = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    if (file.size > 10 * 1024 * 1024) {
      setErrorMessage(localize(language, {
        en: 'The manifest file is too large. The maximum size is 10 MB.',
        ja: 'マニフェスト ファイルが大きすぎます。最大サイズは10 MBです。',
      }));
      setStage('error');
      event.target.value = '';
      return;
    }
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const result = importFromAzPrototype(content);
        setImportResult(result);
        setStage('preview');
      } catch (err: any) {
        setErrorMessage(err.message ? translate(err.message) : localize(language, {
          en: 'Failed to parse the file.',
          ja: 'ファイルの解析に失敗しました。',
        }));
        setStage('error');
      }
    };
    reader.readAsText(file);
    // Reset the input so the same file can be re-selected
    event.target.value = '';
  }, [language, translate]);

  const handleConfirmImport = useCallback(() => {
    if (importResult) {
      onImport(importResult);
      handleClose();
    }
  }, [importResult, onImport, handleClose]);

  if (!isOpen) return null;

  return (
    <ModalScaffold
      isOpen={isOpen}
      onClose={handleClose}
      className="azp-modal azp-import-modal"
      overlayClassName="azp-modal-overlay"
      ariaLabel={t("Import from az prototype")}
    >
        <div className="modal-header azp-modal-header">
          <div className="azp-modal-title">
            <Terminal size={22} />
            <span>{t("Import from az prototype")}</span>
          </div>
          <button className="modal-close azp-modal-close" onClick={handleClose} aria-label={t("Close")}>
            <X size={20} />
          </button>
        </div>

        <div className="modal-body azp-modal-body">
          {stage === 'select' && (
            <>
              <p className="azp-modal-description">
                {' '}{t("Import an architecture manifest produced by")}{' '}<code>{t("az prototype design")}</code> {' '}{t("or exported from the Azure Diagram Builder. The architecture will be rendered as an interactive, editable diagram with official Azure icons, workflow animation, and multi-region cost estimation.")}{' '}</p>

              <label className="azp-import-dropzone azd-surface">
                <Upload size={32} />
                <span className="azp-import-dropzone-text">
                  {' '}{t("Click to select an")}{' '}<code>{t("az-prototype-manifest.json")}</code> {' '}{t("file")}{' '}</span>
                <span className="azp-import-dropzone-hint azd-field-hint">
                  {' '}{t("Accepts .json files (az prototype manifest or raw architecture JSON)")}{' '}</span>
                <input
                  type="file"
                  accept=".json"
                  onChange={handleFileSelect}
                  style={{ display: 'none' }}
                />
              </label>
            </>
          )}

          {stage === 'preview' && importResult && (
            <>
              <div className="azp-import-success azd-callout azd-callout--success">
                <Check size={18} />
                <span>{t("Successfully parsed")}{' '}<strong>{fileName}</strong></span>
              </div>

              <div className="azp-modal-stats">
                <div className="azp-stat azd-surface">
                  <span className="azp-stat-value">{importResult.stats.services}</span>
                  <span className="azp-stat-label">{t("Services")}</span>
                </div>
                <div className="azp-stat azd-surface">
                  <span className="azp-stat-value">{importResult.stats.connections}</span>
                  <span className="azp-stat-label">{t("Connections")}</span>
                </div>
                <div className="azp-stat azd-surface">
                  <span className="azp-stat-value">{importResult.stats.groups}</span>
                  <span className="azp-stat-label">{t("Groups")}</span>
                </div>
                <div className="azp-stat azd-surface">
                  <span className="azp-stat-value">{importResult.stats.workflowSteps}</span>
                  <span className="azp-stat-label">{t("Workflow Steps")}</span>
                </div>
              </div>

              <div className="azp-import-project-info azd-surface">
                <div className="azp-import-project-row">
                  <span className="azp-import-project-key">{t("Project")}</span>
                  <span className="azp-import-project-val">{importResult.projectInfo.name}</span>
                </div>
                <div className="azp-import-project-row">
                  <span className="azp-import-project-key">{t("Region")}</span>
                  <span className="azp-import-project-val">{importResult.projectInfo.location}</span>
                </div>
                <div className="azp-import-project-row">
                  <span className="azp-import-project-key">{t("IaC tool")}</span>
                  <span className="azp-import-project-val">{importResult.projectInfo.iacTool}</span>
                </div>
                {importResult.hasCostData && (
                  <div className="azp-import-project-row">
                    <span className="azp-import-project-key">{t("Cost data")}</span>
                    <span className="azp-import-project-val azp-import-project-val--green">{t("Included")}</span>
                  </div>
                )}
              </div>

              <div className="azp-import-services-preview">
                <div className="azp-import-services-title">{t("Services to import:")}</div>
                <div className="azp-import-services-list">
                  {importResult.architecture.services.map((svc) => (
                    <span key={svc.id} className="azp-import-service-chip">{svc.name}</span>
                  ))}
                </div>
              </div>
            </>
          )}

          {stage === 'error' && (
            <div className="azp-import-error azd-callout azd-callout--danger" role="alert">
              <AlertCircle size={20} />
              <div>
                <strong>{t("Import failed")}</strong>
                <p>{errorMessage}</p>
              </div>
              <button className="azd-button azd-button--secondary" onClick={reset}>
                {' '}{t("Try again")}{' '}</button>
            </div>
          )}
        </div>

        <div className="modal-actions azp-modal-footer">
          <button className="azd-button azd-button--secondary" onClick={handleClose}>
            {' '}{t("Cancel")}{' '}</button>
          {stage === 'preview' && (
            <button
              className="azd-button azd-button--primary"
              onClick={handleConfirmImport}
            >
              <Upload size={18} />
              {' '}{t("Import to Diagram")}{' '}</button>
          )}
        </div>
    </ModalScaffold>
  );
}
