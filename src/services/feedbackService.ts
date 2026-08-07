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
  contact?: {
    consent: boolean;
    email?: string;
  };
}

export interface FeedbackSubmitResult {
  persisted: boolean;
}

/** sessionStorage key set once feedback has been given, to suppress re-prompting. */
export const FEEDBACK_DONE_KEY = 'aqdb_feedback_done';

/**
 * Submit feedback to both Application Insights (sentiment) and the durable
 * /api/feedback endpoint (Cosmos). Shared by the modal and the quick toast so
 * the telemetry shape and storage payload stay identical.
 *
 * Throws when durable storage fails so the UI can tell the user that the
 * feedback was not saved instead of showing a false success message.
 */
export async function submitFeedback(input: FeedbackInput): Promise<FeedbackSubmitResult> {
  const comment = (input.comment || '').trim();
  const contactEmail = input.contact?.consent ? (input.contact.email || '').trim() : '';

  // Always record sentiment in App Insights, independent of durable storage.
  trackFeedback({
    rating: input.rating,
    category: input.category,
    hasComment: comment.length > 0,
    commentLength: comment.length,
  });

  const payload = {
    rating: input.rating,
    category: input.category,
    comment,
    contact: {
      consent: input.contact?.consent === true,
      email: contactEmail,
    },
    context: {
      diagramName: input.context?.diagramName ?? '',
      serviceCount: input.context?.serviceCount ?? 0,
      model: input.context?.model ?? '',
      url: window.location.href,
      userAgent: navigator.userAgent,
    },
  };

  const res = await fetch('/api/feedback', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    throw new Error(`Feedback storage returned HTTP ${res.status}`);
  }

  try {
    sessionStorage.setItem(FEEDBACK_DONE_KEY, '1');
  } catch {
    /* sessionStorage unavailable — ignore */
  }

  return { persisted: true };
}
