import { useRef, type KeyboardEvent, type ReactNode } from 'react';
import { PencilRuler, ShieldCheck, Download } from 'lucide-react';
import { useLanguage } from '../i18n/LanguageContext';
import { localize } from '../i18n/localization';
import './WorkspaceNavigation.css';

export type WorkspaceSection = 'create' | 'review' | 'export';

interface WorkspaceNavigationProps {
  section: WorkspaceSection;
  onSectionChange: (section: WorkspaceSection) => void;
  children?: ReactNode;
}

const sections = [
  { id: 'create', label: { en: 'Create', ja: '作成' }, icon: PencilRuler },
  { id: 'review', label: { en: 'Review', ja: 'レビュー' }, icon: ShieldCheck },
  { id: 'export', label: { en: 'Export', ja: '出力' }, icon: Download },
] as const;

export default function WorkspaceNavigation({ section, onSectionChange, children }: WorkspaceNavigationProps) {
  const { language } = useLanguage();
  const tabs = useRef<Array<HTMLButtonElement | null>>([]);

  const navigate = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    let next: number;
    switch (event.key) {
      case 'ArrowRight': next = (index + 1) % sections.length; break;
      case 'ArrowLeft': next = (index + sections.length - 1) % sections.length; break;
      case 'Home': next = 0; break;
      case 'End': next = sections.length - 1; break;
      default: return;
    }
    event.preventDefault();
    onSectionChange(sections[next].id);
    tabs.current[next]?.focus();
  };

  return (
    <div className="workspace-navigation">
      <div className="workspace-task-tabs" role="tablist"
        aria-label={localize(language, { en: 'Workspace tasks', ja: 'ワークスペースの操作' })}>
        {sections.map(({ id, label, icon: Icon }, index) => (
          <button key={id} ref={element => { tabs.current[index] = element; }}
            type="button" role="tab" id={`workspace-tab-${id}`}
            aria-selected={section === id} aria-controls="application-toolbar"
            tabIndex={section === id ? 0 : -1}
            className="workspace-task-tab" onClick={() => onSectionChange(id)}
            onKeyDown={event => navigate(event, index)}>
            <Icon size={16} aria-hidden="true" />
            <span>{localize(language, label)}</span>
          </button>
        ))}
      </div>
      {children && <div className="workspace-navigation-actions">{children}</div>}
    </div>
  );
}
