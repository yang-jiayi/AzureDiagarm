# Runtime controls and privacy

These are application/deployment contracts, not a claim about the configuration
of any currently running environment. No deployment is performed by the tests.

## Local versus public

`node server/token-server.js` defaults to `APP_DEPLOYMENT_MODE=local` outside
`NODE_ENV=production`. Local mode binds the API to `127.0.0.1`, uses a single
development budget, and permits loopback browser origins. For a local Docker
run, explicitly set `APP_DEPLOYMENT_MODE=local` and publish only on loopback
(`-p 127.0.0.1:8080:80`). Do not expose local mode to other users.

The production image defaults to **public** mode. Startup and the maintained
deployment scripts fail closed unless these controls are provided:

| Variable | Requirement/default |
| --- | --- |
| `APP_DEPLOYMENT_MODE` | `public` |
| `EASY_AUTH_ENABLED` | `true`, only after verifying platform Easy Auth |
| `ACCESS_CONTROL_ENABLED` | `true` |
| `ACCESS_ADMIN_EMAIL` | Administrator email for the existing access list |
| `AZURE_ACCESS_KEY_VAULT_RESOURCE_ID` or `AZURE_TABLES_ACCESS_ENDPOINT` | Existing access-list store accessible to the managed identity |
| `PUBLIC_URL` | Exact HTTPS origin; no credentials, path, query, or fragment |
| `FRONT_DOOR_ID` | Approved Front Door UUID; applied to nginx at runtime |
| `AZURE_OPENAI_ALLOWED_DEPLOYMENTS` / `AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS` | At least one nonempty deployment allowlist; each provider enforces its own list |
| `AZURE_IMPORT_ENABLED` | Must not be `true` in public mode |
| `AI_BUDGET_STORE` | `cosmos` or `table`; never memory in public mode |

The platform must validate authentication and overwrite `X-MS-CLIENT-PRINCIPAL-*`.
The Node listener must remain internal. Front Door origin ingress restrictions
are necessary: the header check by itself is not proof that traffic came from
Front Door. The production sync workflow checks platform authentication, ingress,
WAF and TLS before deployment and configures all runtime flags. It verifies
Cosmos item-level TTL but does **not** enable it automatically: archive-impact
review and any necessary storage configuration must precede publication.
The generic `deploy_aca.sh` reuses preconfigured Easy Auth/Front Door/access
stores; it verifies single-tenant authentication/audiences, rejects excluded API
paths, and checks that a direct origin request with a spoofed Front Door ID is
denied. The obsolete `azure-dev.yml` and `azd-prepackage.sh` paths stay deleted.
Reference Bicep defaults `easyAuthVerified=false`; it does not assert that an
authentication topology has been provisioned or verified.

Protected API requests repeat the existing access-list check in Node, not just
nginx. Public mutations require the exact `Origin` in `PUBLIC_URL`.
Access-list administration remains administrator-only. Existing per-IP OpenAI,
utility, feedback and administrator rate limits remain in place. The OpenAI
hourly limiter uses the existing shared Table Storage counter when
`AZURE_TABLES_ENDPOINT` is set (mandatory in the production workflow); it fails
closed on storage errors and is awaited before identity token reservation.
Utility, feedback and administrator limits remain process-local. Blob-backed
diagram routes, server-gated BYO providers, Speech-only STS tokens, readiness,
and graceful shutdown retain the production behavior.

## Managed GPT-6 Astra deployment

The production build consumes `VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA`, supplied
by the GitHub variable `AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA`. Set its value to
`gpt-6-astra` only after the actual Azure deployment has been provisioned.
The same value participates in runtime allowlists and deployment-drift hashing;
the generic deployment script and `openAiDeploymentGpt6Astra` Bicep parameter
also propagate it. An explicit generic-script `AZURE_OPENAI_ALLOWED_DEPLOYMENTS`
override must include Astra to permit it.

This is a separate deployment, never an alias for GPT-5.6. The proxy forwards
the allowlisted deployment as the upstream model. Existing Sol/Terra/Luna
configuration and revision rollback remain available; BYO configuration is
unchanged. Browser default/recommendation and saved-selection migration are
owned by the model settings integration, and only activate when Astra is
configured. Runtime wiring does not assert model pricing or reasoning support.

## Shared AI budgets

| Variable | Default |
| --- | --- |
| `AI_DAILY_TOKEN_BUDGET` | `250000` input + output tokens per identity per UTC day |
| `AI_MAX_CONCURRENT_REQUESTS` | `2` |
| `AI_BUDGET_STORE` | Public: `table` when a Table endpoint exists, otherwise `cosmos`; local: `memory` |
| `AZURE_COSMOS_ENDPOINT` / `COSMOS_DATABASE_ID` | Existing Cosmos endpoint / `diagrams` |
| `COSMOS_BUDGET_CONTAINER_ID` | `COSMOS_FEEDBACK_CONTAINER_ID`, otherwise `feedback` |
| `AZURE_TABLES_BUDGET_ENDPOINT` | Falls back to `AZURE_TABLES_ENDPOINT` |
| `AZURE_TABLES_BUDGET_TABLE` | `aibudgets` |

Cosmos containers must use partition key `/id` and have TTL enabled
(`defaultTtl=-1` permits per-document TTL). The existing SDK uses the application's
managed identity. Grant data read/create/replace/delete permissions for this
container; existing feedback-container Data Contributor permissions cover the
default shared-container option. Table mode needs Table data permissions and
creates the configured budget table if absent. There is no public memory fallback.
Cosmos multi-region **multi-write must not be enabled** for the budget account:
use one write region, so `If-Match` is serialized at that region.

One document per hashed authenticated ID holds daily usage and expiring request
leases. Conditional ETags and conflict retries make reservation and settlement
atomic across replicas. Tokens are reserved before contacting the AI endpoint:
UTF-8 text/JSON bytes plus protocol overhead, a conservative 131072-token allowance
per image (rather than treating base64 transport as text), and the capped output
tokens. This is deliberately conservative, not a tokenizer or a monetary estimate;
multi-image requests may require a larger daily quota. Images must be inline (the
12 MB request limit still applies). Stored input references, remote image URLs,
remote tools, audio/files, background
requests are not supported by this complete-response proxy. Streaming is forced
off before dispatch. Responses, reasoning/non-reasoning Chat Completions, BYO
endpoints and Foundry Anthropic Messages all reserve the capped output allowance;
Chat Completions is restricted to one choice. Anthropic inline images and cache
read/creation input tokens are included. BYO credentials never enter budget
documents or application logs, and upstream redirects remain blocked.

Trusted upstream usage reconciles the reservation. Unknown outcomes, cancellation,
transport failure and crashed replicas retain the reserved token charge rather
than granting an exploitable refund. Concurrency is released in `finally`; if
storage is unavailable or a process crashes, its lease expires after 315 seconds
(the upstream timeout is 210 seconds, below nginx's 220 and Front Door's 240).
An expired lease cannot refund usage.
Midnight resets token usage, not active concurrency. Budget documents expire
after three inactive days (Cosmos TTL; periodic conditional deletion for Tables).

`GET /api/ai/budget` is guarded like the AI API and returns:

```json
{"available":true,"limitTokens":250000,"usedTokens":5000,"reservedTokens":4000,"remainingTokens":245000,"concurrentRequests":1,"concurrentLimit":2,"resetAt":"2026-09-06T00:00:00.000Z","mode":"public"}
```

Storage/authentication failures are not a zero/full balance. The client shows
“AI budget unavailable.” Quota/concurrency failures use HTTP 429 with actionable
messages and `Retry-After`; storage contention/outage uses 503.

Client integration: default import `AIBudgetStatus` from
`src/components/AIBudgetStatus.tsx`, render `<AIBudgetStatus />` inside the language
provider. No props are required. It refreshes every 30 seconds/on focus/manually.

## Feedback

The feedback modal and quick rating share `buildFeedbackPayload`.
`includeMetadata` defaults to false. With explicit opt-in, only service count,
sanitized model and URL **origin** are accepted. Both client and server exclude
diagram names, user agents, credentials, paths, queries and fragments. The modal
previews the same payload it submits. Comments are user-entered and may contain
sensitive information; the UI warns against including it.

`POST /api/feedback` stores a pseudonymous owner hash (not an email or raw subject)
for authorized deletion and returns the feedback ID, archive expiry and separate
email-delivery status. Table archives retain Cosmos fallback, including lookup
and owner deletion of fallback records. Archive failures are not hidden by
successful email.

Follow-up contact requires both `VITE_FEEDBACK_CONTACT_ENABLED=true` in the
browser build and `FEEDBACK_CONTACT_ENABLED=true` with configured email delivery
on the server. The preview includes the normalized address only after explicit
consent. Addresses are delivered **only by email**, never persisted in Table or
Cosmos archives or sent to analytics. Archives retain consent timestamps, the
180-day follow-up consent expiry, and follow-up status. That consent period is
not a promise of deleting mailbox copies after 180 days.

* `GET /api/feedback/policy`: current archive/retention/email configuration.
* `GET /api/feedback/list`: public administrator identity required; legacy bearer
  `FEEDBACK_ADMIN_TOKEN` is accepted only in local mode.
* `DELETE /api/feedback/:id`: signed-in owner or administrator, with same-origin
  enforcement, deletes from the application archive. Unknown/non-owned IDs
  return 404. The modal exposes deletion for its saved receipt.

`FEEDBACK_RETENTION_DAYS` defaults to **30** for newly saved feedback. New Cosmos
feedback gets a real item TTL; its container is validated before use. New Table
records have explicit expiry and are swept at startup and every 15 minutes
while the server runs. Cosmos TTL deletion is asynchronous.

**Legacy rollout is off by default.** `FEEDBACK_LEGACY_RETENTION_ENABLED=true`
additionally sweeps pre-expiry Table/Cosmos feedback using its creation time
(and deletes undated legacy records). Enabling it can irreversibly remove live
historical feedback: review aggregate impact and obtain the release owner's
retention decision first. The default leaves those older records untouched,
without weakening expiry or owner deletion for new submissions. Reads do not
delete data and omit explicitly expired records. Table cleanup may be delayed
while replicas are stopped or storage is unavailable; failures are logged and
retried, rather than making liveness depend on a cleanup pass.

Email-only feedback cannot be deleted using the archive endpoint (409). Sent
emails, mailbox retention, storage backups and separately collected aggregate
analytics follow their own policies; archive deletion does **not** promise deletion
of those copies. The UI and notification email explain these limits.

Application Insights still records basic sentiment/usage, never comment or prompt
text. Blanket request/response header capture and cross-origin correlation are
disabled; automatic URL telemetry is reduced to origins while the newer
capability-token redaction and deferred SDK/event buffering remain in place.

## Local verification

Run `npm --prefix server ci` when dependencies are missing, then
`npm --prefix server test`. Tests use fake ETag stores plus an isolated local
server: no Azure calls, cloud credentials, or environment files are needed.
