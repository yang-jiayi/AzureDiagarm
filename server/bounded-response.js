function contentLengthExceedsLimit(response, maxBytes) {
  const rawLength = response.headers.get('content-length');
  if (!rawLength || !/^\d+$/.test(rawLength)) return false;
  const length = Number(rawLength);
  return Number.isSafeInteger(length) && length > maxBytes;
}

async function cancelReader(reader) {
  try {
    await reader.cancel();
  } catch (error) {
    throw new Error('Response exceeded its size limit and cancellation failed.', { cause: error });
  }
}

async function cancelErrorResponseBody(response, operation) {
  if (!response.body) return;
  try {
    await response.body.cancel();
  } catch {
    console.error(
      `[${operation}] failed to cancel upstream error body (upstream_status=${response.status})`,
    );
  }
}

async function readBoundedText(response, maxBytes) {
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 0) {
    throw new TypeError('maxBytes must be a non-negative safe integer.');
  }
  if (contentLengthExceedsLimit(response, maxBytes)) {
    try {
      await response.body?.cancel();
    } catch (error) {
      throw new Error('Response exceeded its size limit and cancellation failed.', { cause: error });
    }
    throw new Error('Response exceeded its size limit.');
  }

  const reader = response.body?.getReader();
  if (!reader) return '';

  const chunks = [];
  let length = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      length += value.byteLength;
      if (length > maxBytes) {
        await cancelReader(reader);
        throw new Error('Response exceeded its size limit.');
      }
      chunks.push(Buffer.from(value));
    }
  } finally {
    reader.releaseLock();
  }
  return Buffer.concat(chunks, length).toString('utf8');
}

module.exports = { cancelErrorResponseBody, readBoundedText };
