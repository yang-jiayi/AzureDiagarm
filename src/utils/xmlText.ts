// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * The scanner that finds everything XML cannot carry.
 *
 * The first branch matches a *well-formed* surrogate pair and exists only to
 * consume it, so that the emoji in someone's label is never mistaken for two
 * broken halves. Order matters: alternation is first-match, so a valid pair is
 * always claimed before the lone-surrogate branch can see it.
 *
 * The second branch is the C0 controls XML 1.0 forbids outright — every one
 * below U+0020 except tab, newline and carriage return — plus the two
 * permanently-unassigned non-characters at the end of the BMP. These are not a
 * matter of escaping: `&#11;` is just as illegal as a raw U+000B, so there is
 * no encoding that lets them through, and the only correct handling is removal.
 * An OPC part containing one is rejected by every conforming parser, which
 * means PowerPoint and Visio refuse to open the document at all: the export
 * succeeds, the user sends the deck, and the recipient gets "the file is
 * corrupt".
 *
 * The third branch is any surrogate left over once the pairs are accounted for.
 * A lone D800–DFFF is not a character and has no UTF-8 encoding, so it corrupts
 * the part's byte stream rather than merely its grammar.
 */
const XML_SCAN = /[\uD800-\uDBFF][\uDC00-\uDFFF]|[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDFFF]/g;

/** A two-unit match is a legitimate astral character, not damage. */
function isValidPair(match: string): boolean {
  return match.length === 2;
}

/**
 * Remove every code point XML cannot carry, so a label can never make an
 * exported document unopenable.
 *
 * A vertical tab is Word and PowerPoint's *manual line break*, so it arrives by
 * copy-paste from either of them without anyone typing anything unusual; it is
 * also a legal JSON escape, so it survives an IaC or prototype import intact.
 * Replaced with a space rather than deleted, because in every case that puts
 * one in a label it was separating two words that would otherwise run together.
 */
export function stripXmlForbidden(value: string): string {
  if (!value) return '';
  return value.replace(XML_SCAN, (match) => (isValidPair(match) ? match : ' '));
}

/**
 * True when a string carries a code point XML 1.0 cannot represent.
 *
 * Defined in terms of the strip so the two can never disagree about what
 * counts — a predicate that drifted from the fix it guards would be worse than
 * no predicate at all.
 */
export function hasXmlForbidden(value: string): boolean {
  return stripXmlForbidden(value) !== value;
}
