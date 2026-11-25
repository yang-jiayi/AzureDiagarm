# Managed Model Reasoning Settings

The AI settings popover lets you select GPT-6.2 Sol (the default) or GPT-6.2
Astra, with independent model and reasoning overrides for architecture
generation, validation, deployment guidance, and blueprint diagrams. Both models
use Azure OpenAI Responses deployments in the app-specific OpenAI account.

## Supported Efforts

| Managed model | Default | Supported efforts | API |
|---|---|---|---|
| GPT-6.2 Astra | `low` | `low`, `medium`, `high`, `xhigh`, `max` | Azure OpenAI Responses |
| GPT-6.2 Sol | `low` | `low`, `medium`, `high`, `xhigh`, `max` | Azure OpenAI Responses |

Neither model accepts `none`. Existing preferences for retired models migrate
to GPT-6.2 Sol, except the former GPT-6 Astra preference, which migrates to
GPT-6.2 Astra. Unsupported saved effort values normalize to `low`; explicit
unsupported requests are rejected rather than silently lowered during
dispatch. BYO profiles instead use their configured capabilities, effort, and
output limit.

## Adjusting Settings

Use the global controls for the selected default, then optionally set a model
and effort override per feature. Managed preferences persist in browser storage.
Compatible reasoning preferences are preserved during migration. Historical
diagram and review metadata is not rewritten. Reasoning is controlled by the
settings store, not `VITE_REASONING_EFFORT`.

Higher efforts can affect latency and token consumption, but the application
does not guarantee timing, cost, or output quality for any level. Both managed
models use a 128,000-token output ceiling at `max`, and 32,000 at other efforts.
This includes internal reasoning, not just visible text. Review generated architecture and
deployment guidance before use.

## Request Fidelity

The browser sends requests through `/api/openai`. Each model is bound to its
approved Azure OpenAI deployment and calls `/openai/v1/responses`. Reasoning
uses the Responses API shape:

```json
{
  "reasoning": { "effort": "low" },
  "max_output_tokens": 32000
}
```

Eligible rate-limit retries retain the submitted prompt, deployment, reasoning
effort, and output limit. The application does not silently lower the requested
effort or switch models or API formats to recover an error.

Provider failures, timeouts, budget exhaustion, and incomplete responses are
reported explicitly. Correct configuration and request shape do not by
themselves demonstrate upstream inference availability.

Generation uses [asynchronous jobs](ASYNC-AI-JOBS.md), with separate status and
result requests and a 15-minute processing limit per job. This removes the
former 210-second synchronous deadline without lowering the selected effort.

## References

- [Application setup](../README.md#-getting-started)
- [Server runtime policy](../server/SECURITY.md)
- [Azure OpenAI documentation](https://learn.microsoft.com/azure/ai-services/openai/)
