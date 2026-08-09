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

export type CloudDiagramVersionSummary = Omit<CloudDiagramVersion, 'payload'>;

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
  retryAfterMs?: number;

  constructor(message: string, status: number, code?: string, retryAfterMs?: number) {
    super(message);
    this.name = 'CloudDiagramApiError';
    this.status = status;
    this.code = code;
    this.retryAfterMs = retryAfterMs;
  }
}

interface CloudDiagramRequestOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}

const DEFAULT_REQUEST_TIMEOUT_MS = 30_000;
const MIN_RETRY_DELAY_MS = 1_000;
const MAX_RETRY_DELAY_MS = 5 * 60_000;

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

async function readResponseBody(response: Response): Promise<unknown> {
  const contentType = response.headers.get('content-type') || '';
  try {
    if (contentType.includes('application/json')) {
      return await response.json();
    }
    const text = await response.text();
    return text || null;
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') throw error;
    return null;
  }
}

function parseRetryAfter(value: string | null): number | undefined {
  if (!value) return undefined;
  const seconds = Number(value);
  if (Number.isFinite(seconds) && seconds >= 0) {
    return Math.min(MAX_RETRY_DELAY_MS, Math.max(MIN_RETRY_DELAY_MS, seconds * 1_000));
  }
  const retryAt = Date.parse(value);
  if (!Number.isFinite(retryAt)) return undefined;
  return Math.min(
    MAX_RETRY_DELAY_MS,
    Math.max(MIN_RETRY_DELAY_MS, retryAt - Date.now()),
  );
}

export function getCloudDiagramRetryDelay(
  error: unknown,
  fallbackMs: number,
): number {
  const requested = error instanceof CloudDiagramApiError
    ? error.retryAfterMs
    : undefined;
  return Math.min(
    MAX_RETRY_DELAY_MS,
    Math.max(MIN_RETRY_DELAY_MS, requested ?? fallbackMs),
  );
}

function isTransientCloudError(error: unknown): boolean {
  if (!(error instanceof CloudDiagramApiError)) return false;
  return (
    error.status === 0
    || error.status === 408
    || error.status === 425
    || error.status === 429
    || error.status >= 500
  );
}

function waitForRetry(delayMs: number, signal?: AbortSignal): Promise<void> {
  if (signal?.aborted) {
    return Promise.reject(new CloudDiagramApiError(
      'The cloud diagram request was cancelled.',
      499,
      'REQUEST_ABORTED',
    ));
  }
  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout>;
    const cleanup = () => {
      clearTimeout(timer);
      signal?.removeEventListener('abort', onAbort);
    };
    const onElapsed = () => {
      cleanup();
      resolve();
    };
    const onAbort = () => {
      cleanup();
      reject(new CloudDiagramApiError(
        'The cloud diagram request was cancelled.',
        499,
        'REQUEST_ABORTED',
      ));
    };
    timer = setTimeout(onElapsed, delayMs);
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

async function request(
  path: string,
  init?: RequestInit,
  options: CloudDiagramRequestOptions = {},
): Promise<{
  response: Response;
  body: unknown;
}> {
  const headers = new Headers(init?.headers);
  headers.set('Accept', 'application/json');
  if (init?.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const controller = new AbortController();
  let timedOut = false;
  const timeoutMs = options.timeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;
  const timeoutId = setTimeout(() => {
    timedOut = true;
    controller.abort();
  }, timeoutMs);
  const externalSignal = options.signal ?? init?.signal ?? undefined;
  const abortFromExternal = () => controller.abort();
  if (externalSignal?.aborted) {
    abortFromExternal();
  } else {
    externalSignal?.addEventListener('abort', abortFromExternal, { once: true });
  }

  try {
    const response = await fetch(path, {
      ...init,
      headers,
      credentials: 'same-origin',
      signal: controller.signal,
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
      throw new CloudDiagramApiError(
        message,
        response.status,
        code,
        parseRetryAfter(response.headers.get('retry-after')),
      );
    }
    return { response, body };
  } catch (error) {
    if (error instanceof CloudDiagramApiError) throw error;
    if (timedOut) {
      throw new CloudDiagramApiError(
        'The cloud diagram request timed out.',
        408,
        'REQUEST_TIMEOUT',
      );
    }
    if (externalSignal?.aborted) {
      throw new CloudDiagramApiError(
        'The cloud diagram request was cancelled.',
        499,
        'REQUEST_ABORTED',
      );
    }
    throw new CloudDiagramApiError(
      'The cloud diagram service could not be reached.',
      0,
      'NETWORK_ERROR',
    );
  } finally {
    clearTimeout(timeoutId);
    externalSignal?.removeEventListener('abort', abortFromExternal);
  }
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
  idempotencyKey?: string,
): Promise<CloudDiagramDocument> {
  const headers = idempotencyKey
    ? { 'Idempotency-Key': idempotencyKey }
    : undefined;
  const { response, body } = await request('/api/diagrams', {
    method: 'POST',
    headers,
    body: JSON.stringify({ diagramName, payload }),
  });
  return withEtag<CloudDiagramDocument>(body, response);
}

export async function getCloudDiagram(
  documentId: string,
  options?: CloudDiagramRequestOptions,
): Promise<CloudDiagramDocument> {
  const { response, body } = await request(documentPath(documentId), undefined, options);
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

export async function getSharedCloudDiagram(
  token: string,
  options?: CloudDiagramRequestOptions,
): Promise<CloudDiagramDocument> {
  const attempts = 3;
  for (let attempt = 0; attempt < attempts; attempt += 1) {
    try {
      const { response, body } = await request(sharedPath(token), undefined, options);
      return withEtag<CloudDiagramDocument>(body, response);
    } catch (error) {
      if (
        attempt === attempts - 1
        || !isTransientCloudError(error)
        || options?.signal?.aborted
      ) {
        throw error;
      }
      const fallbackMs = 500 * (2 ** attempt);
      await waitForRetry(getCloudDiagramRetryDelay(error, fallbackMs), options?.signal);
    }
  }
  throw new CloudDiagramApiError(
    'The shared cloud diagram could not be loaded.',
    503,
    'SHARED_LOAD_FAILED',
  );
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
): Promise<CloudDiagramVersionSummary[]> {
  const { body } = await request(versionCollectionPath(context));
  return unwrapArray<CloudDiagramVersionSummary>(body, 'versions');
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
