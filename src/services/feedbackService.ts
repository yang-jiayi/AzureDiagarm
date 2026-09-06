// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { trackFeedback } from './telemetryService';

export interface FeedbackContext {
  diagramName?: string;
  serviceCount?: number;
  model?: string;
}

export interface FeedbackInput {
  rating: number;
  category: string;
  comment: string;
  context?: FeedbackContext;
  includeMetadata?: boolean;
  contact?: {
    consent: boolean;
    email?: string;
  };
}

export interface FeedbackReceipt {
  id: string;
  archiveSaved: boolean;
  expiresAt: string | null;
  emailDelivered: boolean;
  canDelete: boolean;
}

export interface FeedbackPolicy {
  archiveEnabled: boolean;
  retentionDays: number;
  emailEnabled: boolean;
  contactEnabled: boolean;
  contactRetentionDays: number;
  legacyRetentionEnabled: boolean;
}

export interface FeedbackSubmitResult extends FeedbackReceipt {
  persisted: boolean;
}

export function buildFeedbackPayload(input: FeedbackInput) {
  const context: { serviceCount?: number; model?: string; url?: string } = {};
  if (input.includeMetadata === true) {
    const count = input.context?.serviceCount;
    if (Number.isSafeInteger(count) && count! >= 0) context.serviceCount = Math.min(count!, 10_000);
    if (typeof input.context?.model === 'string') {
      context.model = input.context.model.replace(/[^a-zA-Z0-9._-]/g, '').slice(0, 64);
    }
    // Only the origin: no path, query, fragment, username or password.
    if (typeof window !== 'undefined') context.url = new URL(window.location.href).origin;
  }
  return {
    rating: input.rating,
    category: input.category.slice(0, 100),
    comment: (input.comment || '').trim().slice(0, 1000),
    includeMetadata: input.includeMetadata === true,
    contact: input.contact?.consent === true
      ? { consent: true, email: (input.contact.email || '').trim().toLowerCase() }
      : { consent: false },
    context,
  };
}

export async function getFeedbackPolicy(signal?: AbortSignal): Promise<FeedbackPolicy> {
  const response = await fetch('/api/feedback/policy', { credentials: 'same-origin', signal });
  if (!response.ok) throw new Error('Feedback policy unavailable.');
  return response.json();
}

export async function deleteFeedback(id: string): Promise<void> {
  const response = await fetch(`/api/feedback/${encodeURIComponent(id)}`, {
    method: 'DELETE', credentials: 'same-origin',
  });
  if (!response.ok) throw new Error(`Feedback deletion returned HTTP ${response.status}`);
}

/** sessionStorage key set once feedback has been given, to suppress re-prompting. */
export const FEEDBACK_DONE_KEY = 'aqdb_feedback_done';

/**
 * Submit feedback to both Application Insights (sentiment) and the durable
 * /api/feedback endpoint. Shared by the modal and the quick toast so
 * the telemetry shape and storage payload stay identical.
 *
 * Throws when durable storage fails so the UI can tell the user that the
 * feedback was not saved instead of showing a false success message.
 */
export async function submitFeedback(input: FeedbackInput): Promise<FeedbackSubmitResult> {
  const payload = buildFeedbackPayload(input);

  // Always record sentiment in App Insights, independent of durable storage.
  trackFeedback({
    rating: payload.rating,
    category: payload.category,
    hasComment: payload.comment.length > 0,
    commentLength: payload.comment.length,
  });

  const res = await fetch('/api/feedback', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    throw new Error(`Feedback storage returned HTTP ${res.status}`);
  }
  const receipt: FeedbackReceipt = await res.json();

  try {
    sessionStorage.setItem(FEEDBACK_DONE_KEY, '1');
  } catch {
    /* sessionStorage unavailable — ignore */
  }
  return { ...receipt, persisted: true };
}
