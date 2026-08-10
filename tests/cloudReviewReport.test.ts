import assert from 'node:assert/strict';
import test from 'node:test';
import type { CloudDiagramDocument } from '../src/services/cloudDiagramService';
import { buildCloudReviewReport } from '../src/utils/cloudReviewReport';

function document(): CloudDiagramDocument {
  return {
    id: 'diagram-1',
    diagramName: 'Production | Architecture',
    payload: { nodes: [], edges: [] },
    owner: { id: 'owner', email: 'owner@example.com' },
    createdAt: '2026-08-10T00:00:00.000Z',
    updatedAt: '2026-08-10T01:00:00.000Z',
    revision: 4,
    access: 'owner',
    role: 'owner',
    etag: '"4"',
    review: {
      status: 'approved',
      requestedAt: '2026-08-10T00:30:00.000Z',
      decidedAt: '2026-08-10T01:00:00.000Z',
      decidedByEmail: 'reviewer@example.com',
      decisionNote: 'Approved\nfor release',
    },
    comments: [
      {
        commentId: 'comment-1',
        message: 'Check this boundary.',
        authorEmail: 'reviewer@example.com',
        createdAt: '2026-08-10T00:40:00.000Z',
        anchor: { type: 'node', targetId: 'node-1', label: 'App Service' },
      },
      {
        commentId: 'comment-2',
        message: 'Resolved note.',
        authorEmail: 'owner@example.com',
        createdAt: '2026-08-10T00:20:00.000Z',
        resolved: true,
        resolvedAt: '2026-08-10T00:50:00.000Z',
      },
    ],
  };
}

test('cloud review report includes status anchors and resolution sections', () => {
  const report = buildCloudReviewReport(document(), 'en');
  assert.match(report, /Review status:\*\* Approved/);
  assert.match(report, /Open comments:\*\* 1/);
  assert.match(report, /### 1\. App Service/);
  assert.match(report, /> Approved\n> for release/);
  assert.match(report, /## Resolved comments/);
  assert.match(report, /Production \\\| Architecture/);
});

test('cloud review report localizes headings in Japanese', () => {
  const report = buildCloudReviewReport(document(), 'ja');
  assert.match(report, /クラウド レビューレポート/);
  assert.match(report, /レビュー状態:\*\* 承認済み/);
  assert.match(report, /未解決コメント/);
});
