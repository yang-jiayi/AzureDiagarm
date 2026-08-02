import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildApiUrl,
  buildRequestBody,
  parseApiResponse,
} from '../src/services/apiHelper.ts';

test('buildRequestBody converts text and image input to Anthropic Messages format', () => {
  const body = buildRequestBody({
    deployment: 'claude-opus-5',
    messages: [
      { role: 'system', content: 'Return JSON only.' },
      {
        role: 'user',
        content: [
          { type: 'input_text', text: 'Analyze this diagram.' },
          { type: 'input_image', image_url: 'data:image/png;base64,QUJD' },
        ],
      },
    ],
    maxTokens: 12000,
    apiFormat: 'anthropic-messages',
    isReasoning: true,
    reasoningEffort: 'high',
  });

  assert.equal(
    buildApiUrl('https://example.services.ai.azure.com', 'claude-opus-5', 'anthropic-messages'),
    'https://example.services.ai.azure.com/anthropic/v1/messages',
  );
  assert.equal(body.model, 'claude-opus-5');
  assert.equal(body.max_tokens, 12000);
  assert.deepEqual(body.thinking, { type: 'adaptive' });
  assert.deepEqual(body.output_config, { effort: 'high' });
  assert.deepEqual(body.system, [{ type: 'text', text: 'Return JSON only.' }]);
  assert.deepEqual(body.messages[0], {
    role: 'user',
    content: [
      { type: 'text', text: 'Analyze this diagram.' },
      {
        type: 'image',
        source: { type: 'base64', media_type: 'image/png', data: 'QUJD' },
      },
    ],
  });
});

test('parseApiResponse extracts Anthropic text and token usage', () => {
  const parsed = parseApiResponse({
    content: [
      { type: 'thinking', thinking: 'internal' },
      { type: 'text', text: '{"services":[]}' },
    ],
    usage: { input_tokens: 120, output_tokens: 30 },
  }, 'anthropic-messages');

  assert.deepEqual(parsed, {
    content: '{"services":[]}',
    promptTokens: 120,
    completionTokens: 30,
    totalTokens: 150,
  });
});
