# GPT-6 Astra Reasoning Configuration

## Model Policy

GPT-6 Astra is the only managed model in this application. The configured
deployment must actually host Astra; renaming a different model is not a
migration. If Astra is unavailable, the application reports the error instead
of substituting another model or custom endpoint. A verified BYO connection can
be selected explicitly when administrator-enabled; see the
[BYO connection guide](BYO-AI-CONNECTIONS.md).

Deployment verification requires the approved model version `2026-09-03`,
matching `infra/gpt6-astra.bicep`. A missing or different ARM-reported version
blocks release; a future version upgrade requires an explicit policy review.

## Adjusting Reasoning

Open the AI settings popover to choose a global reasoning effort. The configured
choices are `none`, `low`, `medium`, `high`, `xhigh`, and `max`; the application
default is `low`. Generation, validation, deployment guidance, and blueprints
can each override the effort while managed Astra is selected. A selected BYO
profile instead uses its own capability, reasoning, and output-limit settings;
the managed controls do not silently override those choices.

Preferences persist in browser storage. Legacy model selections migrate to
Astra without changing supported reasoning preferences, including `max`.
Reasoning is controlled by the settings store, not `VITE_REASONING_EFFORT`.

Higher reasoning settings can affect latency and token consumption, but the
application does not guarantee timing, cost percentages, or output quality for
any level. Review generated architecture and deployment guidance before use.

## Request Fidelity

The browser sends requests through `/api/openai`. The server binds them to the
approved Astra deployment and calls Azure OpenAI's `/openai/v1/responses`.
Reasoning uses the Responses API shape:

```json
{
  "reasoning": { "effort": "max" },
  "max_output_tokens": 32000
}
```

The model's configured architecture output budget is 32,000 tokens, including
reasoning and response tokens. Eligible rate-limit retries retain the submitted
prompt, deployment, reasoning effort, and output limit. The application does not
silently reduce `max` to `xhigh` or switch to Chat Completions to recover an error.

Provider failures, timeouts, budget exhaustion, and incomplete responses are
reported explicitly. Correct configuration and request shape do not by
themselves demonstrate upstream inference availability.

## References

- [Application setup](../README.md#-getting-started)
- [Server runtime policy](../server/SECURITY.md)
- [Azure OpenAI documentation](https://learn.microsoft.com/azure/ai-services/openai/)
