// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type {
  CloudDiagramComment,
  CloudDiagramDocument,
  CloudReviewStatus,
} from '../services/cloudDiagramService';

const statusLabels: Record<CloudReviewStatus, { en: string; ja: string }> = {
  draft: { en: 'Draft', ja: '下書き' },
  in_review: { en: 'In review', ja: 'レビュー中' },
  changes_requested: { en: 'Changes requested', ja: '変更依頼' },
  approved: { en: 'Approved', ja: '承認済み' },
};

function safeInline(value: unknown): string {
  return String(value ?? '')
    .replace(/[\r\n]+/g, ' ')
    // The escape character must be escaped first, otherwise a trailing
    // backslash in the input consumes the escape this adds and the pipe still
    // breaks out of the Markdown table cell.
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .trim();
}

function quoteComment(comment: CloudDiagramComment): string[] {
  const messageLines = comment.message.replace(/\r\n?/g, '\n').split('\n');
  return messageLines.map(line => `> ${line || ' '}`);
}

export function buildCloudReviewReport(
  document: CloudDiagramDocument,
  language: 'en' | 'ja',
): string {
  const ja = language === 'ja';
  const review = document.review || { status: 'draft' as const };
  const openComments = document.comments.filter(comment => !comment.resolved);
  const resolvedComments = document.comments.filter(comment => comment.resolved);
  const lines = [
    `# ${ja ? 'クラウド レビューレポート' : 'Cloud Review Report'}`,
    '',
    `- **${ja ? 'ダイアグラム' : 'Diagram'}:** ${safeInline(document.diagramName)}`,
    `- **${ja ? 'レビュー状態' : 'Review status'}:** ${statusLabels[review.status][language]}`,
    `- **${ja ? 'リビジョン' : 'Revision'}:** ${document.revision}`,
    `- **${ja ? '生成日時' : 'Generated'}:** ${new Date().toISOString()}`,
    `- **${ja ? '未解決コメント' : 'Open comments'}:** ${openComments.length}`,
    `- **${ja ? '解決済みコメント' : 'Resolved comments'}:** ${resolvedComments.length}`,
  ];

  if (review.requestedAt) {
    lines.push(
      `- **${ja ? 'レビュー依頼' : 'Review requested'}:** ${safeInline(review.requestedAt)}`
      + (review.requestedByEmail ? ` — ${safeInline(review.requestedByEmail)}` : ''),
    );
  }
  if (review.decidedAt) {
    lines.push(
      `- **${ja ? '判定' : 'Decision'}:** ${safeInline(review.decidedAt)}`
      + (review.decidedByEmail ? ` — ${safeInline(review.decidedByEmail)}` : ''),
    );
  }
  if (review.decisionNote) {
    lines.push('', `## ${ja ? '判定メモ' : 'Decision note'}`, '', ...quoteComment({
      commentId: 'decision',
      message: review.decisionNote,
      authorEmail: review.decidedByEmail || '',
      createdAt: review.decidedAt || '',
    }));
  }

  const appendComments = (
    title: string,
    comments: CloudDiagramComment[],
  ) => {
    lines.push('', `## ${title}`, '');
    if (comments.length === 0) {
      lines.push(ja ? '_該当するコメントはありません。_' : '_No comments._');
      return;
    }
    comments.forEach((comment, index) => {
      const anchor = comment.anchor?.type === 'canvas'
        ? (ja ? 'キャンバス全体' : 'Whole canvas')
        : comment.anchor?.label || comment.anchor?.targetId || (ja ? 'アンカーなし' : 'No anchor');
      lines.push(
        `### ${index + 1}. ${safeInline(anchor)}`,
        '',
        `- **${ja ? '投稿者' : 'Author'}:** ${safeInline(comment.authorEmail)}`,
        `- **${ja ? '投稿日時' : 'Created'}:** ${safeInline(comment.createdAt)}`,
      );
      if (comment.resolvedAt) {
        lines.push(
          `- **${ja ? '解決日時' : 'Resolved'}:** ${safeInline(comment.resolvedAt)}`
          + (comment.resolvedByEmail ? ` — ${safeInline(comment.resolvedByEmail)}` : ''),
        );
      }
      lines.push('', ...quoteComment(comment), '');
    });
  };

  appendComments(ja ? '未解決コメント' : 'Open comments', openComments);
  appendComments(ja ? '解決済みコメント' : 'Resolved comments', resolvedComments);
  return `${lines.join('\n').trim()}\n`;
}
