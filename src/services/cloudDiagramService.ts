// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { PricingScenario } from '../types/pricing';
import type { IaCBaseline } from './iacRoundTrip';

export type CloudDiagramAccess = 'owner' | 'shared';
export type CloudDiagramRole = 'owner' | 'viewer' | 'editor';

export interface CloudDiagramPayload {
  nodes: any[];
  edges: any[];
  architecturePrompt?: string;
  originalPrompt?: string;
  validationScore?: number;
  titleBlockData?: any;
  workflow?: any[];
  pricingScenarios?: PricingScenario[];
  iacBaseline?: IaCBaseline | null;
}

export interface CloudDiagramOwner {
  id: string;
  email: string;
}

export interface CloudDiagramComment {
  commentId: string;
  message: string;
  authorEmail: string;
  authorId?: string;
  createdAt: string;
}

export interface CloudDiagramShare {
  shareId: string;
  role: Exclude<CloudDiagramRole, 'owner'>;
  createdAt: string;
  createdByEmail?: string;
}

export interface CloudDiagramDocument {
  id: string;
  diagramName: string;
  payload: CloudDiagramPayload;
  owner: CloudDiagramOwner;
  createdAt: string;
  updatedAt: string;
  revision: number;
  comments: CloudDiagramComment[];
  shares?: CloudDiagramShare[];
  access: CloudDiagramAccess;
  role: CloudDiagramRole;
  etag: string;
}

export interface CloudDiagramSummary {
  id: string;
  diagramName: string;
  createdAt: string;
  updatedAt: string;
  revision: number;
  serviceCount: number;
  connectionCount: number;
  commentCount: number;
  shareCount: number;
  access: CloudDiagramAccess;
  role: CloudDiagramRole;
  etag?: string;
}

export interface CloudDiagramVersion {
  versionId: string;
  diagramId: string;
  diagramName: string;
  payload: CloudDiagramPayload;
  notes?: string;
  createdAt: string;
  createdByEmail?: string;
  sourceRevision: number;
}

export interface CloudShareResult {
  share: CloudDiagramShare;
  token: string;
  url: string;
}

export interface CloudDocumentContext {
  documentId: string;
  access: CloudDiagramAccess;
  role: CloudDiagramRole;
  shareToken?: string;
}

export class CloudDiagramApiError extends Error {
  status: number;
  code?: string;

  constructor(message: string, status: number, code?: string) {
    super(message);
    this.name = 'CloudDiagramApiError';
    this.status = status;
    this.code = code;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

async function readResponseBody(response: Response): Promise<unknown> {
  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    return response.json().catch(() => null);
  }
  const text = await response.text().catch(() => '');
  return text || null;
}

async function request(path: string, init?: RequestInit): Promise<{
  response: Response;
  body: unknown;
}> {
  const headers = new Headers(init?.headers);
  headers.set('Accept', 'application/json');
  if (init?.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const response = await fetch(path, {
    ...init,
    headers,
    credentials: 'same-origin',
  });
  const body = response.status === 204 ? null : await readResponseBody(response);
  if (!response.ok) {
    const record = isRecord(body) ? body : {};
    const message = typeof record.error === 'string'
      ? record.error
      : typeof record.message === 'string'
        ? record.message
        : `Cloud diagram request failed (${response.status})`;
    const code = typeof record.code === 'string' ? record.code : undefined;
    throw new CloudDiagramApiError(message, response.status, code);
  }
  return { response, body };
}

function withEtag<T>(body: unknown, response: Response): T {
  const value = isRecord(body) && isRecord(body.document) ? body.document : body;
  if (!isRecord(value)) {
    throw new CloudDiagramApiError('The cloud service returned an invalid document.', 502);
  }
  return {
    ...value,
    etag: response.headers.get('etag')
      || (typeof value.etag === 'string' ? value.etag : ''),
  } as unknown as T;
}

function unwrapArray<T>(body: unknown, key: string): T[] {
  if (Array.isArray(body)) return body as T[];
  if (isRecord(body) && Array.isArray(body[key])) return body[key] as T[];
  return [];
}

const documentPath = (documentId: string) =>
  `/api/diagrams/${encodeURIComponent(documentId)}`;
const sharedPath = (token: string) =>
  `/api/diagrams/shared/${encodeURIComponent(token)}`;

export async function listCloudDiagrams(): Promise<CloudDiagramSummary[]> {
  const { body } = await request('/api/diagrams');
  return unwrapArray<CloudDiagramSummary>(body, 'documents');
}

export async function createCloudDiagram(
  diagramName: string,
  payload: CloudDiagramPayload,
): Promise<CloudDiagramDocument> {
  const { response, body } = await request('/api/diagrams', {
    method: 'POST',
    body: JSON.stringify({ diagramName, payload }),
  });
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function getCloudDiagram(documentId: string): Promise<CloudDiagramDocument> {
  const { response, body } = await request(documentPath(documentId));
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function updateCloudDiagram(
  documentId: string,
  etag: string,
  diagramName: string,
  payload: CloudDiagramPayload,
): Promise<CloudDiagramDocument> {
  const { response, body } = await request(documentPath(documentId), {
    method: 'PUT',
    headers: { 'If-Match': etag },
    body: JSON.stringify({ diagramName, payload }),
  });
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function deleteCloudDiagram(documentId: string, etag: string): Promise<void> {
  await request(documentPath(documentId), {
    method: 'DELETE',
    headers: { 'If-Match': etag },
  });
}

export async function getSharedCloudDiagram(token: string): Promise<CloudDiagramDocument> {
  const { response, body } = await request(sharedPath(token));
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function updateSharedCloudDiagram(
  token: string,
  etag: string,
  diagramName: string,
  payload: CloudDiagramPayload,
): Promise<CloudDiagramDocument> {
  const { response, body } = await request(sharedPath(token), {
    method: 'PUT',
    headers: { 'If-Match': etag },
    body: JSON.stringify({ diagramName, payload }),
  });
  return withEtag<CloudDiagramDocument>(body, response);
}

function versionCollectionPath(context: CloudDocumentContext): string {
  if (context.access === 'shared' && context.shareToken) {
    return `${sharedPath(context.shareToken)}/versions`;
  }
  return `${documentPath(context.documentId)}/versions`;
}

export async function listCloudVersions(
  context: CloudDocumentContext,
): Promise<CloudDiagramVersion[]> {
  const { body } = await request(versionCollectionPath(context));
  return unwrapArray<CloudDiagramVersion>(body, 'versions');
}

export async function createCloudVersion(
  context: CloudDocumentContext,
  notes: string,
): Promise<CloudDiagramVersion> {
  const { body } = await request(versionCollectionPath(context), {
    method: 'POST',
    body: JSON.stringify({ notes }),
  });
  const value = isRecord(body) && isRecord(body.version) ? body.version : body;
  if (!isRecord(value)) {
    throw new CloudDiagramApiError('The cloud service returned an invalid version.', 502);
  }
  return value as unknown as CloudDiagramVersion;
}

export async function getCloudVersion(
  context: CloudDocumentContext,
  versionId: string,
): Promise<CloudDiagramVersion> {
  const { body } = await request(
    `${versionCollectionPath(context)}/${encodeURIComponent(versionId)}`,
  );
  const value = isRecord(body) && isRecord(body.version) ? body.version : body;
  if (!isRecord(value)) {
    throw new CloudDiagramApiError('The cloud service returned an invalid version.', 502);
  }
  return value as unknown as CloudDiagramVersion;
}

export async function addCloudComment(
  context: CloudDocumentContext,
  message: string,
): Promise<CloudDiagramDocument> {
  const path = context.access === 'shared' && context.shareToken
    ? `${sharedPath(context.shareToken)}/comments`
    : `${documentPath(context.documentId)}/comments`;
  const { response, body } = await request(path, {
    method: 'POST',
    body: JSON.stringify({ message }),
  });
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function listCloudShares(documentId: string): Promise<CloudDiagramShare[]> {
  const { body } = await request(`${documentPath(documentId)}/shares`);
  return unwrapArray<CloudDiagramShare>(body, 'shares');
}

export async function createCloudShare(
  documentId: string,
  role: Exclude<CloudDiagramRole, 'owner'>,
): Promise<CloudShareResult> {
  const { body } = await request(`${documentPath(documentId)}/shares`, {
    method: 'POST',
    body: JSON.stringify({ role }),
  });
  const value = isRecord(body) && isRecord(body.result) ? body.result : body;
  if (!isRecord(value) || typeof value.token !== 'string' || typeof value.url !== 'string') {
    throw new CloudDiagramApiError('The cloud service returned an invalid share link.', 502);
  }
  const share = isRecord(value.share) ? value.share : value;
  return {
    share: share as unknown as CloudDiagramShare,
    token: value.token,
    url: value.url,
  };
}

export async function revokeCloudShare(documentId: string, shareId: string): Promise<void> {
  await request(`${documentPath(documentId)}/shares/${encodeURIComponent(shareId)}`, {
    method: 'DELETE',
  });
}
