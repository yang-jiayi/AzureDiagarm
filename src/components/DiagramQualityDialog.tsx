import { useEffect, useMemo, useState } from 'react';
import {
  AlertTriangle,
  CheckCircle2,
  LocateFixed,
  ScanSearch,
  Sparkles,
  Wrench,
  X,
} from 'lucide-react';
import type { Edge, Node } from 'reactflow';
import {
  analyzeDiagramQuality,
  type DiagramQualityCategory,
  type DiagramQualityFinding,
} from '../utils/diagramQuality';
import ModalScaffold from './ModalScaffold';
import './DiagramQualityDialog.css';

interface DiagramQualityDialogProps {
  isOpen: boolean;
  nodes: Node[];
  edges: Edge[];
  language: 'en' | 'ja';
  onClose: () => void;
  onLocateFinding: (finding: DiagramQualityFinding) => void;
  onApplyFixes: (findingIds: string[]) => Promise<void>;
}

const categoryLabels: Record<DiagramQualityCategory, { en: string; ja: string }> = {
  crossing: { en: 'Crossed connections', ja: '接続線の交差' },
  overlap: { en: 'Overlapping items', ja: 'ノードの重なり' },
  orphan: { en: 'Disconnected services', ja: '未接続のサービス' },
  label: { en: 'Crowded labels', ja: '読みにくいラベル' },
  density: { en: 'Dense layout', ja: '過密なレイアウト' },
  'group-padding': { en: 'Group spacing', ja: 'グループの余白' },
  contrast: { en: 'Low contrast', ja: '低いコントラスト' },
  'unrecognized-service': { en: 'Unrecognized service', ja: '未認識のサービス' },
  'generic-edge': { en: 'Generic connection', ja: '汎用的な接続ラベル' },
  'dangling-edge': { en: 'Dangling connection', ja: '接続先の欠落' },
  'self-loop': { en: 'Self-referencing connection', ja: '自己参照の接続' },
  'duplicate-edge': { en: 'Duplicate connection', ja: '重複した接続' },
  'under-specified': { en: 'Under-specified architecture', ja: '詳細度の不足' },
  'over-crowded': { en: 'Over-crowded diagram', ja: '過密な図' },
  'empty-group': { en: 'Empty group', ja: '空のグループ' },
};

const findingMessages: Record<DiagramQualityCategory, { title: string; detail: string }> = {
  crossing: {
    title: '接続線が交差しています',
    detail: '接続関係を追いやすくするため、レイアウトの整理を推奨します。',
  },
  overlap: {
    title: 'ノードが重なっています',
    detail: 'サービス同士の視認性を高めるため、間隔を確保してください。',
  },
  orphan: {
    title: '未接続のサービスがあります',
    detail: '意図的に独立させているか、接続の追加が必要かを確認してください。',
  },
  label: {
    title: 'ラベルが窮屈です',
    detail: 'サービス名を読みやすくするため、ラベル領域を広げられます。',
  },
  density: {
    title: 'レイアウトが過密です',
    detail: 'サービス間の余白を増やし、構成を追いやすくできます。',
  },
  'group-padding': {
    title: 'グループ内の余白が不足しています',
    detail: '子ノードが境界に近すぎるため、グループの範囲を調整できます。',
  },
  contrast: {
    title: 'コントラストが不足しています',
    detail: '文字や境界線を読みやすくするため、安全な既定色に戻せます。',
  },
  'unrecognized-service': {
    title: 'サービスを認識できません',
    detail: '名称が Azure アイコンに一致しません。正しく表示されるよう、既知のサービス名に変更してください。',
  },
  'generic-edge': {
    title: '接続ラベルが汎用的です',
    detail: '接続のラベルが空または汎用的です。運ぶデータやプロトコルを明記してください。',
  },
  'dangling-edge': {
    title: '接続先が見つかりません',
    detail: '接続が存在しないノードを参照しているため、描画できません。',
  },
  'self-loop': {
    title: '自己参照の接続があります',
    detail: 'ノードが自分自身に接続しています。意図した接続かどうかを確認してください。',
  },
  'duplicate-edge': {
    title: '接続が重複しています',
    detail: '同じノード間の接続が複数回描かれています。重複を整理してください。',
  },
  'under-specified': {
    title: 'アーキテクチャの詳細が不足しています',
    detail: 'サービス数が少なめです。本番向けの構成では、通常さらに詳細が必要です。',
  },
  'over-crowded': {
    title: '図が過密です',
    detail: 'サービスが多く、図が読み取りにくくなっています。目的別のビューに分割することを検討してください。',
  },
  'empty-group': {
    title: 'グループが空です',
    detail: 'サービスを含まないグループがあり、空のラベル付きボックスとして表示されます。',
  },
};

function getFindingCopy(
  finding: DiagramQualityFinding,
  language: 'en' | 'ja',
): { title: string; detail: string } {
  return language === 'ja'
    ? findingMessages[finding.category]
    : { title: finding.title, detail: finding.detail };
}

export default function DiagramQualityDialog({
  isOpen,
  nodes,
  edges,
  language,
  onClose,
  onLocateFinding,
  onApplyFixes,
}: DiagramQualityDialogProps) {
  const report = useMemo(
    () => analyzeDiagramQuality(nodes, edges),
    [edges, nodes],
  );
  const [selectedFixIds, setSelectedFixIds] = useState<string[]>([]);
  const [isApplying, setIsApplying] = useState(false);

  useEffect(() => {
    if (!isOpen) return;
    setSelectedFixIds(
      report.findings
        .filter(finding => Boolean(finding.fixKind))
        .map(finding => finding.id),
    );
  }, [isOpen, report]);

  if (!isOpen) return null;

  const fixableFindings = report.findings.filter(finding => Boolean(finding.fixKind));
  const grade = report.score >= 90
    ? 'A'
    : report.score >= 75
      ? 'B'
      : report.score >= 60
        ? 'C'
        : 'D';
  const issueCountLabel = language === 'ja'
    ? `${report.findings.length} 件の改善候補`
    : `${report.findings.length} improvement${report.findings.length === 1 ? '' : 's'}`;

  const toggleFix = (id: string) => {
    setSelectedFixIds(current => (
      current.includes(id)
        ? current.filter(findingId => findingId !== id)
        : [...current, id]
    ));
  };

  const applyFixes = async () => {
    if (selectedFixIds.length === 0 || isApplying) return;
    setIsApplying(true);
    try {
      await onApplyFixes(selectedFixIds);
    } finally {
      setIsApplying(false);
    }
  };

  return (
    <ModalScaffold
      isOpen={isOpen}
      onClose={onClose}
      className="quality-doctor-dialog"
      overlayClassName="quality-doctor-overlay"
      ariaLabelledBy="quality-doctor-title"
      closeOnBackdrop={!isApplying}
      closeOnEscape={!isApplying}
    >
      <header className="modal-header quality-doctor-header">
          <div className="quality-doctor-heading">
            <span className="quality-doctor-heading-icon" aria-hidden="true">
              <ScanSearch size={22} />
            </span>
            <div>
              <h2 id="quality-doctor-title">
                {language === 'ja' ? 'ダイアグラム品質診断' : 'Diagram Quality Doctor'}
              </h2>
              <p>
                {language === 'ja'
                  ? '構成を変更せずに、読みやすさと保守性を確認します。'
                  : 'Check readability and maintainability without changing the architecture.'}
              </p>
            </div>
          </div>
          <button
            type="button"
            className="modal-close"
            onClick={onClose}
            disabled={isApplying}
            aria-label={language === 'ja' ? '閉じる' : 'Close'}
          >
            <X size={20} />
          </button>
      </header>

      <div className="modal-body quality-doctor-body">
          <section className="quality-doctor-score" aria-label={language === 'ja' ? '診断スコア' : 'Quality score'}>
            <div
              className={`quality-score-ring quality-score-${grade.toLowerCase()}`}
              role="img"
              aria-label={`${report.score} / 100`}
            >
              <strong>{report.score}</strong>
              <span>/ 100</span>
            </div>
            <div className="quality-score-copy">
              <span className="quality-score-grade">
                {language === 'ja' ? `品質グレード ${grade}` : `Quality grade ${grade}`}
              </span>
              <strong>
                {report.findings.length === 0
                  ? (language === 'ja' ? '良好な状態です' : 'Your diagram looks healthy')
                  : issueCountLabel}
              </strong>
              <span>
                {language === 'ja'
                  ? '安全に自動修正できる項目だけを選択できます。'
                  : 'Only non-destructive fixes can be selected for automatic repair.'}
              </span>
            </div>
            <div className="quality-score-summary">
              <span><AlertTriangle size={15} /> {report.counts.high} {language === 'ja' ? '要確認' : 'review'}</span>
              <span><Sparkles size={15} /> {report.counts.medium + report.counts.low} {language === 'ja' ? '提案' : 'suggestions'}</span>
            </div>
          </section>

          {report.findings.length === 0 ? (
            <div className="quality-doctor-empty">
              <CheckCircle2 size={40} aria-hidden="true" />
              <strong>{language === 'ja' ? '改善が必要な項目はありません' : 'No improvements needed'}</strong>
              <p>
                {language === 'ja'
                  ? '現在の図には、検出可能な品質上の問題はありません。'
                  : 'No detectable quality issues were found in the current diagram.'}
              </p>
            </div>
          ) : (
            <div className="quality-doctor-findings">
              {report.findings.map(finding => {
                const copy = getFindingCopy(finding, language);
                const isFixable = Boolean(finding.fixKind);
                const isSelected = selectedFixIds.includes(finding.id);
                const severityTone = finding.severity === 'high' ? 'warning' : 'suggestion';
                return (
                  <article
                    key={finding.id}
                    className={`quality-finding quality-finding-${severityTone}`}
                  >
                    <div className="quality-finding-marker" aria-hidden="true">
                      {finding.severity === 'high'
                        ? <AlertTriangle size={18} />
                        : <Sparkles size={18} />}
                    </div>
                    <div className="quality-finding-content">
                      <span className="quality-finding-category">
                        {categoryLabels[finding.category][language]}
                      </span>
                      <strong>{copy.title}</strong>
                      <p>{copy.detail}</p>
                      <div className="quality-finding-actions">
                        <button
                          type="button"
                          className="quality-locate-btn"
                          onClick={() => onLocateFinding(finding)}
                        >
                          <LocateFixed size={15} />
                          {language === 'ja' ? 'キャンバスで確認' : 'Locate on canvas'}
                        </button>
                        {isFixable && (
                          <label className="quality-fix-option">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => toggleFix(finding.id)}
                            />
                            <span>
                              {language === 'ja' ? '安全な修正に含める' : 'Include safe fix'}
                            </span>
                          </label>
                        )}
                      </div>
                    </div>
                  </article>
                );
              })}
            </div>
          )}
      </div>

      <footer className="modal-actions quality-doctor-footer">
          <span>
            {language === 'ja'
              ? `${fixableFindings.length} 件は自動修正に対応`
              : `${fixableFindings.length} item${fixableFindings.length === 1 ? '' : 's'} support safe fixes`}
          </span>
          <div className="quality-doctor-footer-actions">
            <button
              type="button"
              className="btn-secondary"
              onClick={onClose}
              disabled={isApplying}
            >
              {language === 'ja' ? '閉じる' : 'Close'}
            </button>
            <button
              type="button"
              className="btn-primary"
              onClick={applyFixes}
              disabled={selectedFixIds.length === 0 || isApplying}
            >
              <Wrench size={16} />
              {isApplying
                ? (language === 'ja' ? '修正中...' : 'Applying...')
                : (language === 'ja'
                    ? `選択した修正を適用 (${selectedFixIds.length})`
                    : `Apply selected fixes (${selectedFixIds.length})`)}
            </button>
          </div>
      </footer>
    </ModalScaffold>
  );
}
