import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import test from 'node:test';

import { verifyImageSizeSecurityPatch } from '../scripts/patch-image-size.mjs';

function writeAscii(buffer: Uint8Array, offset: number, value: string): void {
  for (let index = 0; index < value.length; index += 1) {
    buffer[offset + index] = value.charCodeAt(index);
  }
}

function writeUint32(buffer: Uint8Array, offset: number, value: number): void {
  new DataView(buffer.buffer).setUint32(offset, value, false);
}

function probeImageSize(input: Uint8Array): string {
  const source = `
    import { imageSize } from 'image-size';
    const input = Uint8Array.from(Buffer.from(process.env.IMAGE_SIZE_PROBE, 'base64'));
    try {
      imageSize(input);
      console.log('unexpected-success');
    } catch (error) {
      console.log(error instanceof Error ? error.message : String(error));
    }
  `;
  // The timeout is a hang guard, not a performance budget: the vulnerability
  // class here is a parser that loops forever, and a loop that never ends is
  // caught just as surely at thirty seconds as at one. One second was not,
  // though, enough to cover Node's own startup and module resolution on a
  // loaded machine — the probe took 1.1s and this security gate went red for
  // no reason, which is the fastest way to teach people to ignore it.
  const result = spawnSync(process.execPath, ['--input-type=module', '--eval', source], {
    cwd: process.cwd(),
    encoding: 'utf8',
    env: {
      ...process.env,
      IMAGE_SIZE_PROBE: Buffer.from(input).toString('base64'),
    },
    timeout: 30_000,
  });

  assert.notEqual(
    (result.error as NodeJS.ErrnoException | undefined)?.code,
    'ETIMEDOUT',
    'image-size did not terminate — the parser is looping on this input',
  );
  assert.equal(result.error, undefined, result.error?.message);
  assert.equal(result.status, 0, result.stderr);
  return result.stdout.trim();
}

test('image-size parser safeguards are installed from the expected package version', () => {
  assert.deepEqual(verifyImageSizeSecurityPatch(), []);
});

test('patched image-size rejects a zero-length ICNS entry', () => {
  const input = new Uint8Array(16);
  writeAscii(input, 0, 'icns');
  writeUint32(input, 4, input.length);
  writeAscii(input, 8, 'ic07');
  writeUint32(input, 12, 0);

  assert.match(probeImageSize(input), /Invalid ICNS image entry length/);
});

test('patched image-size terminates on a zero-sized JPEG XL box', () => {
  const input = new Uint8Array(36);
  writeUint32(input, 0, 12);
  writeAscii(input, 4, 'JXL ');
  input.set([0x0d, 0x0a, 0x87, 0x0a], 8);

  writeUint32(input, 12, 12);
  writeAscii(input, 16, 'ftyp');
  writeAscii(input, 20, 'jxl ');

  writeUint32(input, 24, 0);
  writeAscii(input, 28, 'jxlp');

  assert.match(probeImageSize(input), /Reached end of input|Invalid JPEG XL|No codestream found/);
});
