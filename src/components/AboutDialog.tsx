// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ExternalLink, Info, X } from 'lucide-react';
import packageMetadata from '../../package.json';
import { useEscapeKey } from '../hooks/useEscapeKey';
import { useModalFocus } from '../hooks/useModalFocus';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './AboutDialog.css';

const REPOSITORY_URL = 'https://github.com/yang-jiayi/AzureDiagarm';
const UPSTREAM_URL = 'https://github.com/Arturo-Quiroga-MSFT/azure-architecture-diagram-builder';

interface AboutDialogProps {
  isOpen: boolean;
  onClose: () => void;
}

export default function AboutDialog({ isOpen, onClose }: AboutDialogProps) {
  const { language } = useLanguage();
  const dialogRef = useModalFocus<HTMLElement>(isOpen);
  useEscapeKey(isOpen, onClose);

  if (!isOpen) return null;

  const text = (en: string, ja: string) => localize(language, { en, ja });
  const version = import.meta.env.VITE_APP_VERSION || packageMetadata.version;

  return (
    <div className="about-dialog-overlay" onClick={onClose}>
      <section
        ref={dialogRef}
        className="about-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="about-dialog-title"
        tabIndex={-1}
        onClick={(event) => event.stopPropagation()}
      >
        <header className="about-dialog-header">
          <div className="about-dialog-heading">
            <span className="about-dialog-icon" aria-hidden="true">
              <Info size={22} />
            </span>
            <div>
              <span className="about-dialog-eyebrow">{text('About this application', 'このアプリについて')}</span>
              <h2 id="about-dialog-title">Azure Architecture Diagram Builder</h2>
            </div>
          </div>
          <button
            type="button"
            className="about-dialog-close"
            onClick={onClose}
            aria-label={text('Close About dialog', 'このダイアログを閉じる')}
            title={text('Close', '閉じる')}
          >
            <X size={20} />
          </button>
        </header>

        <div className="about-dialog-body">
          <p className="about-dialog-summary">
            {text(
              'An independent, community-maintained fork for designing, reviewing, and sharing Azure architectures.',
              'Azure アーキテクチャの設計、レビュー、共有に対応した、独立したコミュニティ運営のフォークです。',
            )}
          </p>

          <dl className="about-dialog-facts">
            <div>
              <dt>{text('Version', 'バージョン')}</dt>
              <dd>{version}</dd>
            </div>
            <div>
              <dt>{text('Original creator', 'オリジナル作成者')}</dt>
              <dd>
                <a href={UPSTREAM_URL} target="_blank" rel="noreferrer">
                  Arturo Quiroga
                  <ExternalLink size={14} aria-hidden="true" />
                </a>
              </dd>
            </div>
            <div>
              <dt>{text('Latest customizations', '最新のカスタマイズ')}</dt>
              <dd>Swarm Data SE, Jiayi Yang</dd>
            </div>
            <div>
              <dt>{text('Repository', 'リポジトリ')}</dt>
              <dd>
                <a href={REPOSITORY_URL} target="_blank" rel="noreferrer">
                  yang-jiayi/AzureDiagarm
                  <ExternalLink size={14} aria-hidden="true" />
                </a>
              </dd>
            </div>
          </dl>

          <p className="about-dialog-note">
            {text(
              'Licensed under the MIT License. This community project is not an official Microsoft product and is not sponsored or endorsed by Microsoft. Microsoft and Azure names and icons remain subject to their respective trademark and brand guidelines.',
              'MIT License の下で提供されています。本コミュニティ プロジェクトは Microsoft の公式製品ではなく、Microsoft による後援または推奨を受けていません。Microsoft、Azure の名称とアイコンには、それぞれの商標・ブランド ガイドラインが適用されます。',
            )}
          </p>
        </div>
      </section>
    </div>
  );
}
