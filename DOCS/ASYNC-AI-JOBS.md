# Asynchronous AI Generation

Generation is submitted as an authenticated background job. The browser uses
separate short HTTP requests for status and result retrieval instead of keeping
one generation request open beyond the former 210-second server deadline.
Managed GPT-6 Astra and explicitly selected BYO Responses/Chat Completions
connections use the same job lifecycle. Connection tests remain synchronous.

## User Experience

The generator displays the actual phase (shared component manifest, topology,
or blueprint), accepted/running/result-ready state and elapsed time. These
states are not a percentage estimate of the model's internal reasoning.
Keep the generation view open; closing or cancelling it requests cancellation.
The existing diagram and completed partial outputs remain protected.

Each job can execute for up to **15 minutes**. Individual browser exchanges
have a **20-second** deadline. Polling can recover transient network/storage
failures within a bounded **16-minute** client lifetime; caller cancellation
guards allow one additional minute for cleanup. MAX reasoning, the chosen
provider, prompt and output cap are not reduced to hide slow responses.
The worker uses a bounded native HTTPS request with TCP keepalive probes,
rather than fetch's independent five-minute response-headers deadline.

## Protocol

All paths remain behind the existing application access control and origin
guard. Ownership comes from the authenticated principal, never a body field.

| Request | Meaning |
| --- | --- |
| `POST /api/openai` with `Prefer: respond-async` and a UUID-v4 `Idempotency-Key` | Accept the existing AI request envelope; return HTTP 202 and job metadata |
| `GET /api/openai/jobs/:id` | Read state, elapsed time and polling interval |
| `GET /api/openai/jobs/:id/result` | Return 202 while active; otherwise return the original successful response or a sanitized structured failure |
| `DELETE /api/openai/jobs/:id` | Request cancellation, including when the submission acknowledgment was lost |

An ambiguous submission is recovered by looking up its **precomputed ID**,
without retransmitting the prompt or BYO key. This also prevents duplicate
inference if a rollback temporarily routes a new browser to a legacy server
without job support. Poll failures never submit new inference.
Cancellation arriving before submission creates a rate-limited tombstone, so a
delayed POST cannot start cancelled work. Retrying a finished request with the
same ID does not generate again. A different body with that ID is rejected.

## Storage, Accounting and Failure Boundaries

The existing private diagram Blob container holds owner-hashed `ai-jobs/`
records. Local development without Blob storage uses an explicit in-memory
backend; public deployment never falls back to process-local storage.

Request bodies, input images and provider credentials are **not persisted as
job inputs**. A BYO key remains only in the browser tab and originating worker
while needed; status/result/cancel requests carry no BYO key. Generated results
can contain sensitive content derived from the user's brief and remain private
to that owner. Successful outputs that echo the submitted BYO key are rejected.
Provider-side storage remains disabled (`store=false`); provider background
mode and its separate retention policy are not enabled.

Results are no longer retrievable through the API after one hour. Minimal
idempotency records remain for 24 hours. Bounded, non-overlapping sweeps delete
expired current blobs and superseded Blob versions under the job prefix;
they never remove diagram documents or their versions. Cleanup is best-effort
during storage outages; expiry is also enforced on reads.

Deletion follows the existing storage account's recovery policy: production
has Blob versioning and **30-day soft deletion**, so deleted copies can remain
in protected storage during that recovery window. A one-hour API expiry is
not a promise of one-hour physical erasure. This release does not disable
versioning, reduce backup protection or change account-wide retention.
Records use conditional ETag writes across replicas.

Shared AI budgets still precharge token reservations and enforce concurrency.
Active jobs renew reservations every five seconds. Cancellation or uncertain
provider usage releases concurrency **without manufacturing a token refund**.
Only validated actual usage reconciles the token charge.

This is **not a durable replay queue**: execution and the BYO key belong to one
live worker. Other replicas can retrieve results and request cancellation, but
a worker crash, deployment or scale-in can interrupt unfinished work. Graceful
shutdown records interruption; a lost worker is detected after its 60-second
lease expires. Such work is explicitly failed, never automatically replayed or
reported as successful. A page reload does not restore the generator session;
review any known job ID before explicitly submitting new work.

## Operational Checks

Run `node --test` in `server`, the targeted client job/cancellation tests, and
the browser async-generation/PNG checks. On Windows,
`node scripts\verify-ai-jobs-long.cjs` uses an ephemeral, locally trusted TLS
test endpoint to withhold the provider response for 325 real seconds. It
checks short submission/polls, cross-replica result retrieval, the former
210/225-second limits, the five-minute headers boundary and budget renewal.
It makes no real model call and must not be described as proof of upstream
model availability. Use existing authorized production access separately for
live-provider verification; do not grant roles or change quotas to create proof.
