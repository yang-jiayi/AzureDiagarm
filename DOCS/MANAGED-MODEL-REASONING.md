# Managed Model Reasoning Settings

The AI settings popover lets you select GPT-6.1 Sol or Claude Opus 5.5, with
independent model and reasoning overrides for architecture generation,
validation, deployment guidance, and blueprint diagrams. GPT-6.1 Sol is the
default selection. Production deployment is subject to procurement, provisioning,
and release verification; these instructions do not imply that either managed
deployment is live.

## Supported Efforts

| Managed model | Default | Supported efforts | API |
|---|---|---|---|
| GPT-6.1 Sol | `low` | `low`, `medium`, `high`, `xhigh`, `max` | Azure OpenAI Responses |
| Claude Opus 5.5 | `medium` | `low`, `medium`, `high`, `xhigh`, `max` | Azure AI Foundry Anthropic Messages with adaptive thinking |

Claude's adaptive thinking is always the reasoning mechanism; the selected level
controls its effort. Neither managed model supports `none`. Existing Sol `none`
preferences normalize to `low`; explicit unsupported requests are rejected, not
silently downgraded during dispatch. When a per-feature override
selects a different model, the effort is normalized to that model's supported
choices. BYO profiles instead use their configured capabilities, effort, and
output limit.

## Adjusting Settings

Use the global controls for the selected default, then optionally set a model
and effort override per feature. Managed preferences persist in browser storage.
Retired managed model selections migrate to GPT-6.1 Sol; compatible efforts are
preserved, and unsupported values are normalized. Reasoning is controlled by the
settings store, not `VITE_REASONING_EFFORT`.

Higher efforts can affect latency and token consumption, but the application
does not guarantee timing, cost, or output quality for any level. Both managed
models share a 32,000-token output ceiling. For Claude, thinking and visible
response tokens both count toward that ceiling. Review generated architecture
and deployment guidance before use.

## Request Fidelity

The browser sends requests through `/api/openai`. GPT-6.1 Sol is bound to the
approved Azure OpenAI deployment and calls `/openai/v1/responses`. Its reasoning
uses the Responses API shape:

```json
{
  "reasoning": { "effort": "max" },
  "max_output_tokens": 32000
}
```

Claude Opus 5.5 uses the Foundry Anthropic Messages API with adaptive thinking;
the application does not send a `none` effort for Claude. For both models, the
32,000-token output budget includes internal reasoning/thinking and visible
response tokens. Eligible rate-limit retries retain the submitted prompt,
deployment, reasoning effort, and output limit. The application does not
silently lower the requested effort or switch models or API formats to recover
an error.

Provider failures, timeouts, budget exhaustion, and incomplete responses are
reported explicitly. Correct configuration and request shape do not by
themselves demonstrate upstream inference availability.

Generation now uses [asynchronous jobs](ASYNC-AI-JOBS.md), with separate status
and result requests and a 15-minute processing limit per job. This removes the
former 210-second synchronous deadline without lowering MAX quality settings.

## References

- [Application setup](../README.md#-getting-started)
- [Server runtime policy](../server/SECURITY.md)
- [Azure OpenAI documentation](https://learn.microsoft.com/azure/ai-services/openai/)
- [Azure AI Foundry documentation](https://learn.microsoft.com/azure/ai-foundry/)
