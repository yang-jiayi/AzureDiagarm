import { test } from 'node:test';
import assert from 'node:assert/strict';
import { stripXmlForbidden, hasXmlForbidden } from '../src/utils/xmlText';

test('strips the C0 controls XML 1.0 forbids, leaving a word gap', () => {
  for (const code of [0x00, 0x01, 0x08, 0x0b, 0x0c, 0x0e, 0x1f]) {
    const ch = String.fromCharCode(code);
    assert.equal(
      stripXmlForbidden(`Payments${ch}gateway`),
      'Payments gateway',
      `U+${code.toString(16).padStart(4, '0')} survived`,
    );
  }
});

test('keeps the three whitespace controls XML does allow', () => {
  const text = 'a\tb\nc\rd';
  assert.equal(stripXmlForbidden(text), text);
  assert.equal(hasXmlForbidden(text), false);
});

test('keeps astral characters whole', () => {
  // A valid surrogate pair must never be read as two broken halves.
  for (const text of ['ops \u{1F680} team', '\u{1F600}', '\u{10000}\u{10FFFD}']) {
    assert.equal(stripXmlForbidden(text), text);
    assert.equal(hasXmlForbidden(text), false);
  }
});

test('removes lone surrogates, which have no UTF-8 encoding at all', () => {
  assert.equal(stripXmlForbidden('a\uD800b'), 'a b');
  assert.equal(stripXmlForbidden('a\uDC00b'), 'a b');
  assert.equal(stripXmlForbidden('\uDC00lead'), ' lead');
  assert.equal(stripXmlForbidden('trail\uD800'), 'trail ');
});

test('removes every lone surrogate in a run, not just the first', () => {
  // The first implementation matched a low surrogate together with the
  // character before it, so a second low surrogate immediately after had no
  // preceding character left to match against and escaped into the file.
  assert.equal(stripXmlForbidden('a\uDC00\uDC00b'), 'a  b');
  assert.equal(stripXmlForbidden('a\uD800\uD800b'), 'a  b');
  assert.equal(stripXmlForbidden('\uDC00\uDC00\uDC00'), '   ');
  assert.equal(hasXmlForbidden('a\uDC00\uDC00b'), true);
});

test('a high surrogate followed by a broken pair keeps the pair', () => {
  assert.equal(stripXmlForbidden(`\uD800\u{1F680}`), ' \u{1F680}');
});

test('removes the BMP non-characters', () => {
  assert.equal(stripXmlForbidden('a\uFFFEb\uFFFFc'), 'a b c');
});

test('leaves ordinary text and the XML metacharacters untouched', () => {
  // Escaping is a separate job; this must not double-handle it.
  const text = 'Contoso & Sons <prod> "west" \'2\' — 日本語';
  assert.equal(stripXmlForbidden(text), text);
  assert.equal(hasXmlForbidden(text), false);
});

test('handles the empty string', () => {
  assert.equal(stripXmlForbidden(''), '');
  assert.equal(hasXmlForbidden(''), false);
});

test('hasXmlForbidden is repeatable', () => {
  // A global regex carries lastIndex between calls; a predicate that answered
  // differently the second time would make the audit rule flap.
  const text = 'Payments\u000bgateway';
  for (let i = 0; i < 5; i += 1) {
    assert.equal(hasXmlForbidden(text), true);
    assert.equal(hasXmlForbidden('clean text'), false);
  }
});

test('the strip is idempotent', () => {
  const once = stripXmlForbidden('a\u000bb\uD800c\u{1F680}');
  assert.equal(stripXmlForbidden(once), once);
  assert.equal(hasXmlForbidden(once), false);
});
