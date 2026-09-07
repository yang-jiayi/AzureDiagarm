# Bring Your Own AI Connections

The default managed connection uses GPT-6 Astra. BYO is a separate, explicit
choice for a user's own Azure OpenAI or official OpenAI model. It does not
reenable retired managed model deployments or multi-model comparison.

## Administrator Policy

Set `ALLOW_BYO_AI_ENDPOINTS=true` on the application server only when BYO access
is approved. The default is disabled. The capability endpoint communicates this
policy to the browser; the proxy enforces it independently on every request.
Public authentication, same-origin checks, rate limits, and shared AI budgets
still apply, including to connection tests.

For this production fork, managed Astra identity verification remains part of
release preflight. BYO does not grant access to another Azure resource, provision
a deployment, increase quotas, or borrow the application's managed identity.

## Connection Profiles

Open the AI connection settings to create, name, edit, or remove a profile. Up
to ten profiles can be retained. Configure the provider, endpoint, deployment or
model identifier, API format, vision/reasoning capabilities, and output limit.
Use the actual model capabilities; a friendly Azure deployment alias is not
proof that its model supports vision or a particular reasoning level.

| Provider | Endpoint | Model identifier |
| --- | --- | --- |
| Azure OpenAI | An approved Azure resource HTTPS origin, such as `https://your-resource.openai.azure.com/` | Your deployment alias |
| OpenAI | `https://api.openai.com/` | Your model ID |

Use an origin, not a full REST path. URL credentials, query strings, fragments,
ports, local/IP endpoints, and arbitrary OpenAI-compatible hosts are rejected.
The proxy selects the provider's fixed v1 Responses or Chat Completions route
and does not follow redirects.

## Test, Then Select

Enter the API key and run **Test connection**. This is an explicit, cancellable
request to the selected provider, not a background fallback. It uses the same
protected proxy and application budget as other AI requests and may consume
provider tokens. A completed test confirms that request path, not the quality
or availability of every future architecture-generation request. A text-only
test does not certify vision support, architecture JSON compatibility, or every
reasoning/output-limit combination.

Select the verified profile to use it. Generation, image analysis, blueprint and
Both modes, chat, validation, IaC analysis, and deployment guidance honor the
selected connection. Unsupported capabilities or invalid structured output are
reported; the application does not switch providers or lower generation settings
to hide a failure.

Choose **managed GPT-6 Astra** explicitly to switch back. A missing key, disabled
server policy, invalid profile, or provider failure never triggers that switch
automatically.

## Key and Verification Safety

Only public profile configuration is saved. API keys and connection verification
stay in memory in the current browser tab. Reloading requires key re-entry and
a new test; there is no credential stored in localStorage, sessionStorage,
downloaded profiles, diagram history, telemetry, or server logs.

Changing a key or request-affecting profile setting invalidates verification.
A response from an earlier or cancelled test cannot verify the new settings.
Requests capture their connection context: changing profiles cannot redirect
an already submitted operation or relabel its result.

Existing single-connection BYO preferences migrate to a profile, without
overwriting managed preferences or historical diagrams/reviews. Migration
does not restore a key or invent successful verification.

## Troubleshooting

| State | Action |
| --- | --- |
| Administrator disabled | Ask the application administrator to approve and enable BYO; do not bypass the proxy |
| Key required | Re-enter the key in this tab and test the connection |
| Unverified or settings changed | Test the current profile before selecting it |
| Authentication rejected | Confirm the user's key and resource/model access; the application identity is not a fallback |
| Model or API format rejected | Confirm the deployment/model and supported format/capabilities |
| Quota, rate limit, or application budget exhausted | Follow the displayed retry/reset guidance; switching connection does not bypass the application budget |
| Incomplete response or provider failure | Review the request and provider state; keep the request ID for support without sharing the key |
