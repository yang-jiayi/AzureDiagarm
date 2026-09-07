// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { forwardRef } from 'react';
import { ChevronDown, Sparkles, X } from 'lucide-react';
import { getReasoningEffortLabel, useModelSettings } from '../stores/modelSettingsStore';
import { useLanguage } from '../i18n/LanguageContext';
import AstraReasoningSettings from './AstraReasoningSettings';
import AIConnectionSelector, { connectionStateLabel } from './AIConnectionSelector';
import { getEffectiveAIModelInfo } from '../services/aiModelRuntime';
import { getBYOAIConnectionState, useBYOAISettings } from '../stores/byoAISettingsStore';
import { useRuntimeConfig } from '../services/runtimeConfig';
import './ModelSettingsPopover.css';

interface ModelSettingsPopoverProps {
  isOpen: boolean;
  onToggle: () => void;
  onConfigureConnections?: () => void;
}

const ModelSettingsPopover = forwardRef<HTMLDivElement, ModelSettingsPopoverProps>(
  ({ isOpen, onToggle, onConfigureConnections }, ref) => {
    const { t, language } = useLanguage();
    useModelSettings();
    const { storageError } = useBYOAISettings();
    const policy = useRuntimeConfig();
    const selected = getEffectiveAIModelInfo('architectureGeneration');
    const connection = selected.profileId ? getBYOAIConnectionState(selected.profileId) : null;
    const connectionStatus = policy.status === 'error' ? 'policy-unavailable'
      : policy.status !== 'ready' ? 'policy-checking' : !policy.bringYourOwnAI ? 'admin-disabled' : connection?.status;
    const effortLabel = t(getReasoningEffortLabel(selected.isReasoning ? selected.reasoningEffort : 'none'));
    const badge = storageError ? connectionStateLabel('storage-unavailable', language) : connection ? selected.ready
      ? `${connectionStateLabel('verified', language)} · ${effortLabel}`
      : connectionStateLabel(connectionStatus ?? 'missing-profile', language) : effortLabel;
    return (
      <div className="toolbar-dropdown" ref={ref}>
        <button
          type="button"
          onClick={onToggle}
          className="btn btn-secondary model-popover-trigger"
          title={t('AI model settings')}
          aria-haspopup="dialog"
          aria-expanded={isOpen}
        >
          <Sparkles size={14} />
          <span className="model-popover-label">{selected.displayName}</span>
          <span className="model-popover-reasoning" title={badge}>{badge}</span>
          <ChevronDown size={14} />
        </button>
        {isOpen && (
          <div className="toolbar-dropdown-menu toolbar-dropdown-menu--model-settings"
            role="dialog" aria-label={t('AI model settings')}>
            <div className="toolbar-dropdown-heading">
              <span>{t('AI model settings')}</span>
              <button type="button" className="msp-close-btn" onClick={onToggle}
                title={t('Close')} aria-label={t('Close')}><X size={13} /></button>
            </div>
            <AIConnectionSelector compact onConfigureConnections={onConfigureConnections} />
            <AstraReasoningSettings />
          </div>
        )}
      </div>
    );
  },
);

ModelSettingsPopover.displayName = 'ModelSettingsPopover';
export default ModelSettingsPopover;
