# User Feedback Feature

The feedback modal sends ratings, categories, optional comments, and limited
diagram context to `POST /api/feedback`.

## Data flow

```text
Browser
  |-- User_Feedback telemetry --> Application Insights (no comment text)
  `-- POST /api/feedback -------> token server
                                    |-- Azure Communication Services Email
                                    |-- Azure Table Storage (optional archive)
                                    `-- Cosmos DB (compatibility fallback)
```

The UI reports success only after the server confirms delivery. Failed writes
remain retryable and are not copied into telemetry.

## Runtime configuration

Preferred direct-email settings:

```text
FEEDBACK_EMAIL_ENDPOINT=https://<resource>.communication.azure.com
FEEDBACK_EMAIL_SENDER=donotreply@<managed-domain>.azurecomm.net
FEEDBACK_EMAIL_RECIPIENT=<owner@example.com>
AZURE_CLIENT_ID=<managed-identity-client-id>
```

Optional queryable archive settings:

```text
AZURE_TABLES_ENDPOINT=https://<account>.table.core.windows.net
AZURE_TABLES_FEEDBACK_TABLE=feedback

AZURE_COSMOS_ENDPOINT=https://<account>.documents.azure.com:443/
COSMOS_DATABASE_ID=diagrams
COSMOS_FEEDBACK_CONTAINER_ID=feedback
```

Email is attempted first. When an archive is configured, the same submission is
also written to Azure Table Storage or Cosmos DB. A successful archive acts as
the fallback when email delivery is temporarily unavailable. When no backend is
configured, the API returns HTTP 503.

The API limits each client address to ten submissions per hour.

## Optional follow-up contact

Follow-up contact is disabled by default. It is shown and accepted only when
both of these settings are explicitly enabled:

```text
VITE_FEEDBACK_CONTACT_ENABLED=true
FEEDBACK_CONTACT_ENABLED=true
```

The server also requires all direct-email settings above. Contact consent is
unchecked by default, the email field is rendered only after opt-in, and a
valid address is required only when consent is enabled.

The address is included in the one-time feedback notification email. It is not
written to Application Insights, Azure Table Storage, or Cosmos DB. Archives
retain only consent timestamps, the 180-day consent boundary, and follow-up
status metadata. Operators must apply the same 180-day retention policy to the
recipient mailbox.

## Reading archived feedback

If a Table Storage or Cosmos archive is configured,
`GET /api/feedback/list?limit=50` reads the newest records. The endpoint is
disabled unless `FEEDBACK_ADMIN_TOKEN` is configured. Supply the token as:

```text
Authorization: Bearer <token>
X-Admin-Token: <token>
```

The endpoint returns HTTP 503 when the admin token is unset, HTTP 401 for an
invalid token, and applies a separate admin rate limit.

`server/read-feedback.js` provides a keyless Cosmos CLI for environments that
can reach the account:

```text
cd server
node read-feedback.js
node read-feedback.js --json
```

## Privacy

- Application Insights receives rating metadata only.
- Verbatim comments stay in the configured delivery or archive backend.
- Contact addresses are delivered only to the configured feedback recipient.
- Failed persistence never sends comments or contact details to telemetry.

## Files

| File | Purpose |
|---|---|
| `src/components/FeedbackModal.tsx` | Full feedback form and retryable error UI |
| `src/components/FeedbackToast.tsx` | Quick rating and persistence status |
| `src/services/feedbackService.ts` | Client submission and telemetry metadata |
| `server/token-server.js` | Delivery, archive, validation, and protected admin read |
| `server/read-feedback.js` | Keyless Cosmos read utility |
