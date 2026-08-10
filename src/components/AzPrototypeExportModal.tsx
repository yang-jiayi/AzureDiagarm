// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { useState } from 'react';
import { Terminal, Download, X } from 'lucide-react';
import './AzPrototypeExportModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import ModalScaffold from './ModalScaffold';

export interface AzPrototypeExportModalProps {
  isOpen: boolean;
  onClose: () => void;
  /** Called when the user confirms export with their chosen options. */
  onExport: (options: {
    projectName: string;
    location: string;
    iacTool: 'bicep' | 'terraform';
    includeCosts: boolean;
    includeWorkflow: boolean;
  }) => void;
  /** Current diagram service count (shown for context). */
  serviceCount: number;
  /** Current diagram connection count. */
  connectionCount: number;
  /** Current diagram group count. */
  groupCount: number;
  /** Whether cost data is available on the diagram. */
  hasCostData: boolean;
  /** Current architecture name from title block. */
  architectureName?: string;
}

const AZURE_REGIONS = [
  { value: 'eastus', label: 'East US' },
  { value: 'eastus2', label: 'East US 2' },
  { value: 'westus2', label: 'West US 2' },
  { value: 'westus3', label: 'West US 3' },
  { value: 'centralus', label: 'Central US' },
  { value: 'westeurope', label: 'West Europe' },
  { value: 'northeurope', label: 'North Europe' },
  { value: 'swedencentral', label: 'Sweden Central' },
  { value: 'uksouth', label: 'UK South' },
  { value: 'southeastasia', label: 'Southeast Asia' },
  { value: 'australiaeast', label: 'Australia East' },
  { value: 'canadacentral', label: 'Canada Central' },
  { value: 'brazilsouth', label: 'Brazil South' },
  { value: 'japaneast', label: 'Japan East' },
];

export default function AzPrototypeExportModal({
  isOpen,
  onClose,
  onExport,
  serviceCount,
  connectionCount,
  groupCount,
  hasCostData,
  architectureName,
}: AzPrototypeExportModalProps) {
  const { t } = useLanguage();
  const [projectName, setProjectName] = useState(
    () => (architectureName || 'my-prototype').toLowerCase().replace(/[^a-z0-9-]/g, '-').replace(/-+/g, '-').slice(0, 40),
  );
  const [location, setLocation] = useState('eastus');
  const [iacTool, setIacTool] = useState<'bicep' | 'terraform'>('bicep');
  const [includeCosts, setIncludeCosts] = useState(hasCostData);
  const [includeWorkflow, setIncludeWorkflow] = useState(true);

  if (!isOpen) return null;

  const handleExport = () => {
    onExport({ projectName, location, iacTool, includeCosts, includeWorkflow });
    onClose();
  };

  return (
    <ModalScaffold
      isOpen={isOpen}
      onClose={onClose}
      className="azp-modal"
      overlayClassName="azp-modal-overlay"
      ariaLabel={t("Export to az prototype")}
    >
        <div className="modal-header azp-modal-header">
          <div className="azp-modal-title">
            <Terminal size={22} />
            <span>{t("Export to az prototype")}</span>
          </div>
          <button className="modal-close azp-modal-close" onClick={onClose} aria-label={t("Close")}>
            <X size={20} />
          </button>
        </div>

        <div className="modal-body azp-modal-body">
          <p className="azp-modal-description">
            {' '}{t("Export this architecture as an")}{' '}<code>{t("az prototype")}</code> {' '}{t("manifest. Use the downloaded file with")}{' '}<code>{t("az prototype design --import")}</code> {' '}{t("to skip the discovery phase and go straight to IaC generation with full governance, security review, and staged deployment.")}{' '}</p>

          <div className="azp-modal-stats">
            <div className="azp-stat azd-surface">
              <span className="azp-stat-value">{serviceCount}</span>
              <span className="azp-stat-label">{t("Services")}</span>
            </div>
            <div className="azp-stat azd-surface">
              <span className="azp-stat-value">{connectionCount}</span>
              <span className="azp-stat-label">{t("Connections")}</span>
            </div>
            <div className="azp-stat azd-surface">
              <span className="azp-stat-value">{groupCount}</span>
              <span className="azp-stat-label">{t("Groups")}</span>
            </div>
          </div>

          <div className="azp-form">
            <div className="azp-form-group azd-field">
              <label htmlFor="azp-project-name">{t("Project name")}</label>
              <input
                id="azp-project-name"
                className="azd-control"
                type="text"
                value={projectName}
                onChange={(e) => setProjectName(e.target.value.toLowerCase().replace(/[^a-z0-9-]/g, '-'))}
                placeholder={t("my-prototype")}
                maxLength={40}
              />
              <span className="azp-form-hint azd-field-hint">{t("Used by")}{' '}<code>{t("az prototype init --name")}</code></span>
            </div>

            <div className="azp-form-group azd-field">
              <label htmlFor="azp-location">{t("Azure region")}</label>
              <select
                id="azp-location"
                className="azd-control"
                value={location}
                onChange={(e) => setLocation(e.target.value)}
              >
                {AZURE_REGIONS.map((r) => (
                  <option key={r.value} value={r.value}>{r.label}</option>
                ))}
              </select>
              <span className="azp-form-hint azd-field-hint">{t("Used by")}{' '}<code>{t("az prototype init --location")}</code></span>
            </div>

            <div className="azp-form-group azd-field">
              <label htmlFor="azp-iac-tool">{t("IaC tool")}</label>
              <select
                id="azp-iac-tool"
                className="azd-control"
                value={iacTool}
                onChange={(e) => setIacTool(e.target.value as 'bicep' | 'terraform')}
              >
                <option value="bicep">{t("Bicep")}</option>
                <option value="terraform">{t("Terraform")}</option>
              </select>
              <span className="azp-form-hint azd-field-hint">{t("Used by")}{' '}<code>{t("az prototype init --iac-tool")}</code></span>
            </div>

            <div className="azp-form-group azp-form-group--checkbox">
              <label>
                <input
                  type="checkbox"
                  checked={includeWorkflow}
                  onChange={(e) => setIncludeWorkflow(e.target.checked)}
                />
                {' '}{t("Include workflow steps")}{' '}</label>
            </div>

            <div className="azp-form-group azp-form-group--checkbox">
              <label>
                <input
                  type="checkbox"
                  checked={includeCosts}
                  onChange={(e) => setIncludeCosts(e.target.checked)}
                  disabled={!hasCostData}
                />
                {' '}{t("Include cost estimates")}{' '}{!hasCostData && <span className="azp-form-hint azd-field-hint"> {' '}{t("(no pricing data on diagram)")}</span>}
              </label>
            </div>
          </div>

          <div className="azp-cli-preview azd-surface">
            <div className="azp-cli-preview-title">{t("After downloading, run:")}</div>
            <pre className="azp-cli-code">
{`# Initialize the project
az prototype init --name ${projectName || 'my-prototype'} \\
  --location ${location} --iac-tool ${iacTool}

# Import the architecture (skip design phase)
az prototype design --import ./az-prototype-manifest.json

# Generate production IaC with governance
az prototype build

# Deploy with preflight checks
az prototype deploy`}
            </pre>
          </div>
        </div>

        <div className="modal-actions azp-modal-footer">
          <button className="azd-button azd-button--secondary" onClick={onClose}>
            {' '}{t("Cancel")}{' '}</button>
          <button
            className="azd-button azd-button--primary"
            onClick={handleExport}
            disabled={serviceCount === 0}
          >
            <Download size={18} />
            {' '}{t("Export Manifest")}{' '}</button>
        </div>
    </ModalScaffold>
  );
}
