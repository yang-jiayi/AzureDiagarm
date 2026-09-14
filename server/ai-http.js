// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
const https = require('node:https');

const MAX_RESPONSE_BYTES = 4 * 1024 * 1024;

function awaitWithSignal(work, signal) {
  if (!signal) return work;
  return new Promise((resolve, reject) => {
    const abort = () => { cleanup(); reject(signal.reason); };
    const cleanup = () => signal.removeEventListener('abort', abort);
    work.then(value => {
      cleanup();
      if (signal.aborted) reject(signal.reason);
      else resolve(value);
    }, error => { cleanup(); reject(signal.aborted ? signal.reason : error); });
    if (signal.aborted) abort();
    else signal.addEventListener('abort', abort, { once: true });
  });
}

// Native HTTPS avoids fetch's independent five-minute headers deadline.
// The executor owns the complete request deadline and cancellation signal.
function fetchLongAIResponse(url, options) {
  return new Promise((resolve, reject) => {
    const request = https.request(url, {
      method: options.method,
      headers: { ...options.headers, 'Accept-Encoding': 'identity', 'Content-Length': Buffer.byteLength(options.body) },
      signal: options.signal,
      agent: new https.Agent({ keepAlive: false }),
    }, response => {
      const status = response.statusCode;
      if (status >= 300 && status < 400) {
        response.destroy();
        reject(new Error('AI endpoint redirects are not allowed.'));
        return;
      }
      const headers = new Headers();
      for (const [name, values] of Object.entries(response.headers)) {
        for (const value of Array.isArray(values) ? values : [values]) {
          if (value !== undefined) headers.append(name, value);
        }
      }
      const body = (async () => {
        const chunks = [];
        let length = 0;
        for await (const chunk of response) {
          length += chunk.length;
          if (length > MAX_RESPONSE_BYTES) {
            response.destroy();
            throw new Error('AI response exceeded its size limit.');
          }
          chunks.push(chunk);
        }
        return Buffer.concat(chunks, length).toString('utf8');
      })();
      // A cancelled caller may never ask for text after receiving the headers.
      body.catch(() => {});
      resolve({ status, ok: status >= 200 && status < 300, headers, text: () => body });
    });
    // TCP probes keep a quiet inference connection alive through egress NAT;
    // HTTP connection pooling remains disabled for this one-shot request.
    request.once('socket', socket => socket.setKeepAlive(true, 60_000));
    request.once('error', reject);
    request.end(options.body);
  });
}

module.exports = { awaitWithSignal, fetchLongAIResponse, MAX_RESPONSE_BYTES };
