// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState } from 'react';
import { X, Camera } from 'lucide-react';
import './SaveSnapshotModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useEscapeKey } from '../hooks/useEscapeKey';

interface SaveSnapshotModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (notes: string) => void;
  diagramName: string;
  serviceCount: number;
}

const SaveSnapshotModal: React.FC<SaveSnapshotModalProps> = ({
  isOpen,
  onClose,
  onSave,
  diagramName,
  serviceCount
}) => {
  const { t, language } = useLanguage();
  useEscapeKey(isOpen, onClose);
  const [notes, setNotes] = useState('');
  const [isSaving, setIsSaving] = useState(false);

  const handleSave = async () => {
    setIsSaving(true);
    try {
      await onSave(notes);
      setNotes('');
      onClose();
    } catch (error) {
      console.error('Failed to save snapshot:', error);
      alert(t("Failed to save snapshot"));
    } finally {
      setIsSaving(false);
    }
  };

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal-content save-snapshot-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Save Snapshot")}
        tabIndex={-1}
        autoFocus
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') onClose();
        }}
      >
        <div className="modal-header">
          <h2>
            <Camera size={24} />
            {' '}{t("Save Snapshot")}{' '}</h2>
          <button className="modal-close" onClick={onClose} title={t("Close")} aria-label={t("Close")}>
            <X size={24} />
          </button>
        </div>

        <div className="modal-body">
          <div className="snapshot-info">
            <p className="snapshot-info-text">
              {' '}{t("Creating a snapshot of")}{' '}<strong>{diagramName}</strong>
            </p>
            <p className="snapshot-info-details">
              {serviceCount} {' '}{t("services •")}{' '}{new Date().toLocaleString()}
            </p>
          </div>

          <div className="form-group">
            <label htmlFor="snapshot-notes">
              {' '}{t("Notes (optional)")}{' '}<span className="label-hint">{t("Describe what makes this version special")}</span>
            </label>
            <textarea
              id="snapshot-notes"
              className="snapshot-notes"
              placeholder={t("e.g., Before adding authentication, Initial production setup, Working state before experiment...")}
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              rows={4}
              maxLength={500}
              disabled={isSaving}
            />
            <div className="character-count">
              {notes.length}{t("/500")}{' '}</div>
          </div>

          <div className="snapshot-hint">
            {language === 'ja'
              ? '💡 スナップショットはこのブラウザーに保存され、利用可能な場合はクラウドにも安全に保存されます。'
              : '💡 Snapshots are saved in this browser and securely copied to the cloud when available.'}
          </div>
        </div>

        <div className="modal-actions">
          <button 
            className="btn-secondary" 
            onClick={onClose}
            disabled={isSaving}
          >
            {' '}{t("Cancel")}{' '}</button>
          <button 
            className="btn-primary" 
            onClick={handleSave}
            disabled={isSaving}
          >
            {isSaving ? (
              <>
                <div className="spinner-small"></div>
                {' '}{t("Saving...")}{' '}</>
            ) : (
              <>
                <Camera size={18} />
                {' '}{t("Save Snapshot")}{' '}</>
            )}
          </button>
        </div>
      </div>
    </div>
  );
};

export default SaveSnapshotModal;
