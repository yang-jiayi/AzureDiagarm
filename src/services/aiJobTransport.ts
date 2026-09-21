import { awaitWithAISignal, runtimeConfigCancellationError } from './runtimeConfig';

export const AI_JOB_CLIENT_TIMEOUT_MS = 16 * 60_000;
const HTTP_TIMEOUT_MS = 20_000;
const STATUSES = ['queued', 'running', 'cancelling', 'succeeded', 'failed', 'cancelled', 'expired'] as const;
type JobStatus = (typeof STATUSES)[number];

export interface AIJobProgress {
  id: string;
  status: JobStatus;
  elapsedMs: number;
  deadlineAt: number;
  pollAfterMs: number;
}

function parseProgress(value: unknown, id: string): AIJobProgress | null {
  if (!value || typeof value !== 'object' || !('job' in value)) return null;
  const job = value.job;
  if (!job || typeof job !== 'object'
    || !('id' in job) || job.id !== id
    || !('status' in job) || !STATUSES.some(status => status === job.status)
    || !('elapsedMs' in job) || typeof job.elapsedMs !== 'number' || !Number.isFinite(job.elapsedMs) || job.elapsedMs < 0
    || !('deadlineAt' in job) || typeof job.deadlineAt !== 'number' || !Number.isFinite(job.deadlineAt)
    || !('pollAfterMs' in job) || typeof job.pollAfterMs !== 'number' || !Number.isFinite(job.pollAfterMs)) return null;
  return {
    id, status: job.status as JobStatus, elapsedMs: job.elapsedMs,
    deadlineAt: job.deadlineAt, pollAfterMs: Math.min(5000, Math.max(1000, job.pollAfterMs)),
  };
}

function failure(id: string, code: string, status = 503, source = 'job'): Response {
  return new Response(JSON.stringify({ error: { source, code, requestId: id } }), {
    status, headers: { 'Content-Type': 'application/json', 'X-AzureDiagarm-Request-Id': id },
  });
}

function pause(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const cleanup = () => { clearTimeout(timer); signal?.removeEventListener('abort', abort); };
    const abort = () => { cleanup(); reject(runtimeConfigCancellationError()); };
    const timer = setTimeout(() => { cleanup(); resolve(); }, ms);
    if (signal?.aborted) abort();
    else signal?.addEventListener('abort', abort, { once: true });
  });
}

async function shortRequest(url: string, init: RequestInit, signal?: AbortSignal): Promise<Response> {
  const controller = new AbortController();
  const abort = () => controller.abort();
  const timer = setTimeout(abort, HTTP_TIMEOUT_MS);
  if (signal?.aborted) abort();
  else signal?.addEventListener('abort', abort, { once: true });
  try {
    const response = await awaitWithAISignal(fetch(url, {
      ...init, signal: controller.signal, credentials: 'same-origin', redirect: 'manual', cache: 'no-store',
    }), controller.signal);
    // Include the response body in the short HTTP deadline, not just headers.
    await awaitWithAISignal(response.clone().arrayBuffer(), controller.signal);
    return response;
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', abort);
  }
}

export async function fetchAIJob(
  init: RequestInit,
  options: { signal?: AbortSignal; onProgress?: (progress: AIJobProgress) => void; beforeSubmit?: () => void } = {},
): Promise<Response> {
  const { signal, onProgress } = options;
  if (signal?.aborted) throw runtimeConfigCancellationError();
  const id = crypto.randomUUID();
  const headers = new Headers(init.headers);
  headers.set('Prefer', 'respond-async');
  headers.set('Idempotency-Key', id);
  const url = `/api/openai/jobs/${id}`;
  const deadline = Date.now() + AI_JOB_CLIENT_TIMEOUT_MS;
  let accepted = false;
  let terminal = false;
  let submitted = false;
  let recoveryAttempted = false;
  let lastResponse: Response | undefined;
  const unavailable = () => failure(id, 'ai_job_status_unavailable', 504, 'client');
  const request = async (path: string, options: RequestInit, signal?: AbortSignal) => {
    const response = await shortRequest(path, options, signal);
    // Easy Auth redirects become opaque in a browser. Never follow them with
    // a prompt/key or hide expired authentication as an inference timeout.
    return response.type === 'opaqueredirect' || (response.status >= 300 && response.status < 400)
      ? failure(id, 'application_authentication_required', 401, 'application')
      : response;
  };
  const cancel = async () => {
    try {
      const response = await request(url, { method: 'DELETE', keepalive: true });
      if (!response.ok) console.warn('[ai-jobs] Cancellation could not be confirmed.', { requestId: id, status: response.status });
    } catch {
      console.warn('[ai-jobs] Cancellation could not reach the server.', { requestId: id });
    }
  };
  try {
    // Recover a lost acknowledgment by looking up this ID, never by sending
    // the prompt again: a pre-jobs server during rollback cannot deduplicate.
    options.beforeSubmit?.();
    submitted = true;
    try {
      lastResponse = await request('/api/openai', { ...init, headers }, signal);
      if (lastResponse.status >= 500) {
        const value: unknown = await lastResponse.clone().json().catch(() => null);
        if (value && typeof value === 'object' && 'error' in value
          && value.error && typeof value.error === 'object' && 'code' in value.error
          && typeof value.error.code === 'string' && value.error.code !== 'ai_jobs_unavailable') {
          terminal = true;
          return lastResponse;
        }
      }
    } catch {
      if (signal?.aborted) throw runtimeConfigCancellationError();
    }
    if (lastResponse && lastResponse.status !== 202 && lastResponse.status < 500) {
      // Compatibility with an already-open tab during a rolling release and
      // immediate, explicitly rejected submissions.
      terminal = true;
      return lastResponse;
    }
    accepted = lastResponse?.status === 202;
    while (true) {
      if (Date.now() >= deadline) {
        if (recoveryAttempted) return unavailable();
        recoveryAttempted = true;
        lastResponse = undefined;
      }
      let response: Response;
      try {
        response = accepted && lastResponse?.status === 202
          ? lastResponse : await request(url, { method: 'GET' }, signal);
        lastResponse = undefined;
      } catch {
        if (signal?.aborted) throw runtimeConfigCancellationError();
        if (recoveryAttempted) return unavailable();
        await pause(2000, signal);
        continue;
      }
      if (response.status >= 500) {
        if (recoveryAttempted) return unavailable();
        await pause(2000, signal);
        continue;
      }
      if (!response.ok) {
        terminal = terminal || [401, 403, 410].includes(response.status);
        return response;
      }
      let payload: unknown;
      try { payload = await response.json(); }
      catch { return failure(id, 'ai_job_invalid_response', 502); }
      const job = parseProgress(payload, id);
      if (!job) return failure(id, 'ai_job_invalid_response', 502);
      accepted = true;
      onProgress?.(job);
      if (!['queued', 'running', 'cancelling'].includes(job.status)) {
        terminal = true;
        try {
          const result = await request(`${url}/result`, { method: 'GET' }, signal);
          if (result.status === 202) {
            terminal = false;
            if (recoveryAttempted) return unavailable();
            lastResponse = result;
            continue;
          }
          // A 503 can be a terminal provider failure. The result endpoint's
          // job/source code distinguishes it from a temporary storage outage.
          const value: unknown = result.status === 503 ? await result.clone().json().catch(() => null) : null;
          if (value && typeof value === 'object' && 'error' in value
            && value.error && typeof value.error === 'object' && 'code' in value.error
            && value.error.code === 'ai_jobs_unavailable') {
            if (recoveryAttempted) return unavailable();
            await pause(2000, signal);
            continue;
          }
          return result;
        } catch {
          if (signal?.aborted) throw runtimeConfigCancellationError();
          if (recoveryAttempted) return unavailable();
          await pause(2000, signal);
          continue;
        }
      }
      // A suspended tab can resume after the clock deadline even though the
      // server finished on time. The final bounded read above is authoritative.
      if (recoveryAttempted) return unavailable();
      await pause(job.pollAfterMs, signal);
    }
  } finally {
    // The precomputed ID also cancels an accepted job whose 202 was lost.
    // The server records a tombstone if cancellation races submission.
    if (!terminal && submitted) await cancel();
  }
}
