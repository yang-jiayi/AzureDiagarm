// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useRef, useState } from 'react';
import { X, Camera } from 'lucide-react';
import './SaveSnapshotModal.css';
import { useLanguage } from '../i18n/LanguageContext';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { OperationGeneration } from '../utils/operationGeneration';

interface SaveSnapshotModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSave: (notes: string) => Promise<void>;
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
  const [notes, setNotes] = useState('');
  const [isSaving, setIsSaving] = useState(false);
  const isOpenRef = useRef(isOpen);
  const savingRef = useRef(false);
  const saveGenerationRef = useRef(new OperationGeneration());

  isOpenRef.current = isOpen;

  useEffect(() => {
    if (isOpen) return;
    saveGenerationRef.current.advance();
    savingRef.current = false;
    setIsSaving(false);
  }, [isOpen]);

  useEscapeKey(isOpen && !isSaving, onClose);

  const handleSave = async () => {
    if (savingRef.current) return;
    const generation = saveGenerationRef.current.advance();
    savingRef.current = true;
    setIsSaving(true);
    try {
      await onSave(notes);
      if (!isOpenRef.current || !saveGenerationRef.current.isCurrent(generation)) return;
      setNotes('');
      onClose();
    } catch (error) {
      if (!isOpenRef.current || !saveGenerationRef.current.isCurrent(generation)) return;
      console.error('Failed to save snapshot:', error);
      alert(t("Failed to save snapshot"));
    } finally {
      if (saveGenerationRef.current.isCurrent(generation)) {
        savingRef.current = false;
        setIsSaving(false);
      }
    }
  };

  if (!isOpen) return null;

  return (
    <div className="modal-overlay" onClick={() => {
      if (!savingRef.current) onClose();
    }}>
      <div
        className="modal-content save-snapshot-modal"
        role="dialog"
        aria-modal="true"
        aria-label={t("Save Snapshot")}
        tabIndex={-1}
        autoFocus
        onClick={(e) => e.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape' && !savingRef.current) onClose();
        }}
      >
        <div className="modal-header">
          <h2>
            <Camera size={24} />
            {' '}{t("Save Snapshot")}{' '}</h2>
          <button
            className="modal-close"
            onClick={onClose}
            title={t("Close")}
            aria-label={t("Close")}
            disabled={isSaving}
          >
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
