# User Feedback Feature

The Feedback modal and quick-rating toast send ratings, categories, optional
comments, and limited diagram context to `POST /api/feedback`.

## Data flow

```text
Browser
  |-- User_Feedback telemetry --> Application Insights (no comment text)
  `-- POST /api/feedback -------> token server
                                    |-- Azure Communication Services Email
                                    |-- Azure Table Storage (optional archive)
                                    `-- Cosmos DB (compatibility fallback)
```

The UI reports success only after durable storage returns a successful response.
Failed writes remain retryable and are not copied into telemetry.

## Production storage

`azurediagarm.mssql.biz` sends each accepted submission directly to:

| Setting | Value |
|---|---|
| Recipient | Configured through `FEEDBACK_EMAIL_RECIPIENT` |
| Communication resource | Configured through `FEEDBACK_EMAIL_ENDPOINT` |
| Sender | Configured through `FEEDBACK_EMAIL_SENDER` |
| Authentication | Container App user-assigned managed identity |

This avoids a database, private endpoint, or VNet solely for low-volume feedback.
The API limits each client address to ten submissions per hour.

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

## Reading comments

Email-delivered feedback is read in the recipient mailbox. If a Table Storage
or Cosmos archive is configured, `GET /api/feedback/list?limit=50` reads the
newest archived records. It is disabled
unless `FEEDBACK_ADMIN_TOKEN` is configured on the server. Supply the token as
either:

```text
Authorization: Bearer <token>
X-Admin-Token: <token>
```

The endpoint returns HTTP 503 when the admin token is unset and HTTP 401 when
the presented token does not match.

## Privacy

Application Insights receives rating metadata only. Verbatim comments stay in
the durable feedback store and are never used as a telemetry fallback.

## Files

| File | Purpose |
|---|---|
| `src/components/FeedbackModal.tsx` | Full feedback form and retryable error UI |
| `src/components/FeedbackToast.tsx` | Quick rating and persistence status |
| `src/services/feedbackService.ts` | Client submission and telemetry metadata |
| `server/token-server.js` | Durable write and token-protected admin read |
