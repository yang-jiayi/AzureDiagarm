// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const https = require('node:https');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { createRequire } = require('node:module');
const serverRequire = createRequire(path.resolve(__dirname, '..', 'server', 'package.json'));
const express = serverRequire('express');
const { createAIJobs, createMemoryJobBackend } = require('../server/ai-jobs');
const { createOpenAIProxyRouter } = require('../server/openai-proxy');
const { fetchLongAIResponse } = require('../server/ai-http');
const { createBudgetManager, MemoryBudgetStore } = require('../server/ai-budget');

const DURATION_MS = 325_000;
const logger = { info() {}, error() {}, warn() {} };
const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
const listen = server => new Promise(resolve => server.listen(0, '127.0.0.1', () => resolve(server)));
const close = server => new Promise(resolve => { server.close(resolve); server.closeAllConnections(); });

async function main() {
  // The private key is generated in memory for this test process only. No
  // certificate-store changes, committed keys or disabled TLS checks.
  const certificate = JSON.parse(execFileSync('pwsh', ['-NoProfile', '-NonInteractive', '-Command', `
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
      'CN=localhost', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256,
      [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $san = [System.Security.Cryptography.X509Certificates.SubjectAlternativeNameBuilder]::new()
    $san.AddIpAddress([System.Net.IPAddress]::Parse('127.0.0.1'))
    $request.CertificateExtensions.Add($san.Build())
    $cert = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-1), [DateTimeOffset]::UtcNow.AddHours(1))
    @{
      pfx = [Convert]::ToBase64String($cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, ''))
      cert = [Convert]::ToBase64String($cert.RawData)
    } | ConvertTo-Json -Compress
    $cert.Dispose()
    $rsa.Dispose()
  `], { encoding: 'utf8', timeout: 20000, windowsHide: true }));
  const ca = `-----BEGIN CERTIFICATE-----\n${certificate.cert.match(/.{1,64}/g).join('\n')}\n-----END CERTIFICATE-----\n`;
  let upstreamCalls = 0;
  let upstreamAborted = false;
  const upstream = https.createServer({ pfx: Buffer.from(certificate.pfx, 'base64'), passphrase: '' }, async (req, res) => {
    upstreamCalls++;
    const chunks = [];
    for await (const chunk of req) chunks.push(chunk);
    const body = JSON.parse(Buffer.concat(chunks).toString());
    assert.deepEqual(body.reasoning, { effort: 'max' });
    assert.equal(body.max_output_tokens, 32000);
    assert.equal(body.store, false);
    assert.equal(body.stream, false);
    const timer = setTimeout(() => {
      res.setHeader('Content-Type', 'application/json');
      res.end(JSON.stringify({ output_text: 'controlled local HTTPS completion', usage: { total_tokens: 42 } }));
    }, DURATION_MS);
    res.on('close', () => {
      clearTimeout(timer);
      if (!res.writableEnded) upstreamAborted = true;
    });
  });
  await listen(upstream);
  const upstreamUrl = `https://127.0.0.1:${upstream.address().port}/`;
  const originalRequest = https.request;
  https.request = (url, options, callback) => {
    assert.equal(url, upstreamUrl);
    return originalRequest(url, { ...options, ca }, callback);
  };
  const backend = createMemoryJobBackend();
  const store = new MemoryBudgetStore();
  const budget = createBudgetManager({ store, dailyTokens: 250000, concurrency: 2 });
  let renewals = 0;
  const renew = budget.renew;
  budget.renew = async (...args) => { await renew(...args); renewals++; };
  const replicas = [];
  try {
    for (let index = 0; index < 2; index++) {
      const jobs = createAIJobs({ backend, mode: 'local', logger });
      const app = express();
      app.use(express.json());
      app.use('/api/openai', createOpenAIProxyRouter({
        jobs, budget, mode: 'local', logger,
        endpoint: 'https://example.openai.azure.com/', astraDeployment: 'gpt-6-astra',
        allowedDeployments: new Set(['gpt-6-astra']), apiKey: 'local-test-only-key',
        fetchImpl: (_url, options) => fetchLongAIResponse(upstreamUrl, options),
      }));
      const server = await listen(require('node:http').createServer(app));
      replicas.push({ jobs, server, url: `http://127.0.0.1:${server.address().port}/api/openai` });
    }
    const id = crypto.randomUUID();
    const startedAt = Date.now();
    const accepted = await fetch(replicas[0].url, {
      method: 'POST', headers: { 'Content-Type': 'application/json', Prefer: 'respond-async', 'Idempotency-Key': id },
      signal: AbortSignal.timeout(10000),
      body: JSON.stringify({
        apiFormat: 'responses', deployment: 'gpt-6-astra',
        body: { model: 'gpt-6-astra', input: 'Controlled local transport proof, not a model request.', reasoning: { effort: 'max' }, max_output_tokens: 32000 },
      }),
    });
    const submissionMs = Date.now() - startedAt;
    assert.equal(accepted.status, 202);
    assert.ok(submissionMs < 5000);
    let polls = 0;
    let longestPollMs = 0;
    let crossed210 = false;
    let crossed315 = false;
    while (Date.now() - startedAt < DURATION_MS + 30000) {
      const pollingAt = Date.now();
      const response = await fetch(`${replicas[1].url}/jobs/${id}`, { signal: AbortSignal.timeout(10000) });
      const status = await response.json();
      const elapsedMs = Date.now() - startedAt;
      polls++;
      longestPollMs = Math.max(longestPollMs, Date.now() - pollingAt);
      assert.equal(response.status, 200);
      if (status.job.status === 'succeeded') break;
      assert.ok(['queued', 'running'].includes(status.job.status), `Unexpected terminal state: ${status.job.status}`);
      if (elapsedMs > 210000) crossed210 = true;
      if (elapsedMs > 315000) {
        crossed315 = true;
        assert.equal((await budget.status('local-development')).concurrentRequests, 1);
      }
      await wait(2000);
    }
    const response = await fetch(`${replicas[1].url}/jobs/${id}/result`, { signal: AbortSignal.timeout(10000) });
    const result = await response.json();
    assert.equal(response.status, 200);
    assert.equal(result.output_text, 'controlled local HTTPS completion');
    assert.equal(upstreamCalls, 1);
    assert.equal(upstreamAborted, false);
    assert.ok(crossed210 && crossed315);
    assert.ok(renewals >= 60);
    assert.ok(longestPollMs < 5000);
    const finalBudget = await budget.status('local-development');
    assert.equal(finalBudget.concurrentRequests, 0);
    assert.equal(finalBudget.usedTokens, 42);
    console.log(JSON.stringify({
      proof: 'controlled-local-https-not-live-inference',
      elapsedMs: Date.now() - startedAt, submissionMs, polls, longestPollMs,
      upstreamCalls, renewals, crossed210, crossed315, tlsVerified: true,
      resultFromDifferentReplica: true, maxReasoningPreserved: true, outputLimit: 32000,
    }, null, 2));
  } finally {
    for (const replica of replicas) await replica.jobs.close();
    for (const replica of replicas) await close(replica.server);
    https.request = originalRequest;
    await close(upstream);
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
