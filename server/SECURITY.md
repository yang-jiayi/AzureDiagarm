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
| `AZURE_OPENAI_ENDPOINT` | Azure OpenAI HTTPS account origin for GPT-6.1 Sol |
| `AZURE_OPENAI_DEPLOYMENT_GPT61SOL` | Explicit GPT-6.1 Sol deployment alias |
| `AZURE_OPENAI_ALLOWED_DEPLOYMENTS` | Exactly the same single GPT-6.1 Sol alias |
| `AZURE_FOUNDRY_ENDPOINT` | Separate Microsoft Foundry HTTPS account origin for Claude Opus 5.5 |
| `AZURE_FOUNDRY_DEPLOYMENT_OPUS55` | Explicit Claude Opus 5.5 deployment alias |
| `AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS` | Exactly the same single Claude Opus 5.5 alias |
| `ALLOW_BYO_AI_ENDPOINTS` | `false` by default; only explicit `true` authorizes user-key BYO connections |
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
diagram routes, Speech-only STS tokens, readiness,
and graceful shutdown retain the production behavior.

## Dual managed model deployments

The managed catalog has two isolated routes: GPT-6.1 Sol uses Azure OpenAI
Responses, and Claude Opus 5.5 uses the Microsoft Foundry Anthropic Messages API
(Anthropic API version `2026-10-21`). Each provider has its own server-side
endpoint, deployment alias and exact singleton allowlist. The build receives
only public routing values as `VITE_*`; no credential is built into the image.
GPT-6.1 Sol prefers managed identity and can use the existing optional
server-side Azure OpenAI key fallback; Claude requires managed identity and
rejects `AZURE_FOUNDRY_API_KEY`. An alias is not proof of model identity.

The production release requires both `AZURE_OPENAI_RESOURCE_ID` and
`AZURE_FOUNDRY_RESOURCE_ID` as deployment-time variables. Before image
publication and revision update, read-only ARM checks verify each resource ID,
expected account kinds (`OpenAI` for GPT-6.1 Sol, `AIServices` for Claude Opus
5.5) and endpoint ownership, then require successful provisioning of
`OpenAI/gpt-6.1-sol`, version `2026-09-29`, and
`Anthropic/claude-opus-5-5`, version `2`. Lookup failures or mismatches block
deployment. This verification does not create, rename, delete, or grant access
to Azure model resources.

Managed requests must match their selected provider's exact deployment and API
format. The endpoint and credentials are fixed server-side; enabling BYO does
not extend either allowlist or authorize endpoint/credential overrides for a
managed request. Invalid requests fail before credential acquisition, budget
reservation or dispatch. Release cleanup removes the retired Astra deployment
alias and other obsolete provider settings from the active app environment, but
does not remove the generic model allowlists, BYO policy, account resources, or
historical model provenance. Existing Azure Astra resources are retained.

Claude's Anthropic Messages route uses the `2026-10-21` API version. Structured
diagram output is requested as a strict JSON object envelope containing the
`azureDiagramJson` string; the client unwraps that field before normal result
validation. Adaptive thinking is enabled only for requests that explicitly
select it. Both managed models share the 32,000-token output ceiling, and
Claude's thinking tokens count toward that ceiling.

The desired deployment hash includes both model endpoints, aliases and account
resource IDs. If a candidate revision fails verification, rollback copies the
last healthy revision as a complete unit: its image, AI configuration,
environment/secret references and `DESIRED_CONFIG_HASH` stay together. The
rollback does not graft new model settings onto an older image, reread mutable
model settings, or require another ARM lookup. Thus a historical revision may
retain Astra settings for its own recovery history; the active dual-model
revision removes Astra execution settings. No model resource is deleted.

Unconfigured local mode permits manual diagrams; managed AI returns a clear 503.
Public/self-host startup can still permit BYO-only mode when explicitly opted in
and managed endpoint, alias and allowlist settings are absent. Partial/invalid
managed configuration fails clearly. The maintained production release target
requires both managed routes and their ARM identity preflight. This does not
relax authentication, origin isolation, access-list or shared-budget startup
requirements. No live configuration change or publication is implied.

## Bring-your-own AI connections

Administrators enable the feature using `ALLOW_BYO_AI_ENDPOINTS=true`; missing
configuration defaults to false and unrecognized values fail startup closed.
`GET /api/runtime-config` returns `features.bringYourOwnAI` as a boolean and
describes the managed catalog with `ai.defaultModel: "gpt-6.1-sol"`,
`ai.configured`, and public descriptors for GPT-6.1 Sol and Claude Opus 5.5.
A stale browser capability cannot authorize a connection on a disabled server.

A BYO request uses the same authenticated `/api/openai` endpoint:

```json
{
  "apiFormat": "chat-completions",
  "deployment": "your-model-or-deployment",
  "body": {
    "model": "your-model-or-deployment",
    "messages": [{"role": "user", "content": "Reply OK"}],
    "max_tokens": 8
  },
  "byo": {
    "provider": "azure-openai",
    "endpoint": "https://your-resource.openai.azure.com",
    "apiKey": "<user-supplied key>"
  }
}
```

Only `azure-openai` and official `openai` are supported. Azure origins must use
trusted `.openai.azure.*`, `.cognitiveservices.azure.*` or `.services.ai.azure.*`
resource hosts in the supported commercial, US government or China clouds.
OpenAI permits only `https://api.openai.com`. HTTPS origins may have a single
trailing slash, but no ports (including explicit `:443`), user information, API
paths, queries, fragments, IP addresses or localhost. Arbitrary OpenAI-compatible
hosts and upstream redirects are prohibited.

BYO model IDs must be explicit, match `body.model`, and contain 1–128 ASCII
letters/digits or `._:-`, with at least one letter/digit. They are user choices,
not additional managed deployments. Profiles select model capabilities and
either `responses` or `chat-completions`; the proxy does not silently change the
model, provider, API format, reasoning settings or JSON-output format. Azure uses
`/openai/v1/responses` or `/openai/v1/chat/completions`; OpenAI uses the corresponding
`/v1` path.

BYO uses **only the supplied user key**, never managed identity or the server API
key. Malformed, null or disabled BYO cannot fall through to managed execution.
Keys are forwarded only as upstream authentication headers and are not placed
in budget documents or logged. Diagnostics use stable codes/messages and an
application-generated request ID; raw upstream errors, custom request IDs,
endpoints, keys and prompt bodies are not logged or reflected as diagnostics.

Connection testing is an explicit user action issuing a small, bounded,
cancellable **normal proxy request**, not a privileged test route. It obeys the
same access list, origin checks, rate limits, token budget and concurrency limit
as generation. It may consume tokens. There is no automatic request, alternate
model retry, credential fallback or test bypass.

## Asynchronous generation

Generation submits an owner-scoped job through `POST /api/openai` with
`Prefer: respond-async` and a UUID-v4 idempotency key. Protected
`/api/openai/jobs/:id` status, result and cancellation routes use the same
access control; cross-replica ownership is derived from the authenticated
principal. The browser does not replay an ambiguous submission, including
during rollback to a legacy server.

Request bodies, images and provider keys are not persisted as job inputs.
Private results expire from the API after one hour; idempotency records last
24 hours. Cleanup includes older Blob versions and honors existing soft-delete
recovery protection. This is not an assurance of immediate physical erasure.
Public mode requires the existing Blob backend and never falls back to memory.

Workers run for at most 15 minutes and renew budget leases while active.
Shutdown, worker loss, expired ownership or failed budget renewal interrupts
work explicitly without replaying uncertain inference or refunding unknown
usage. See [the job lifecycle and retention contract](../DOCS/ASYNC-AI-JOBS.md).

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
off and storage is disabled before dispatch. GPT-6.1 Sol Responses keeps its 32768-token
output ceiling, including 32000-token architecture requests. Managed Sol accepts
only `low`, `medium`, `high`, `xhigh` or `max` when an effort is explicitly
supplied; unsupported values are rejected before credential acquisition,
budget reservation or dispatch. BYO Responses and
Chat Completions also cap and reserve the chosen output field (`max_output_tokens`,
`max_completion_tokens` or `max_tokens`) without translating formats. Chat's
`prompt_tokens`/`completion_tokens` usage and inline vision content are metered
alongside Responses input/output usage. Multiple generations per request and
legacy function/tool modes are not supported. Custom credentials never enter
budget documents, and upstream redirects remain blocked.

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
