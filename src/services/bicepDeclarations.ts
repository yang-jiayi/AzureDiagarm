// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface BicepDeclaration {
  logicalName: string;
  providerType: string;
  resourceName: string | null;
  kind: string | null;
  conditional: boolean;
  loop: boolean;
  notes: string[];
}

interface Token {
  kind: 'identifier' | 'string' | 'symbol';
  value: string;
  literal?: string | null;
  line: number;
  endLine: number;
}

const MAX_SOURCE_LENGTH = 2_000_000;
const MAX_TOKENS = 200_000;
const MAX_DEPTH = 64;
const MAX_DECLARATIONS = 5_000;

function stringLiteral(raw: string): string | null {
  if (raw.includes('${')) return null;
  if (raw.startsWith("'''")) return raw.slice(3, -3).trim() || null;
  let valid = true;
  const value = raw.slice(1, -1).replace(/\\(u\{[0-9a-f]+\}|.)/gi, (_match, escape: string) => {
    if (escape.startsWith('u{')) {
      const code = Number.parseInt(escape.slice(2, -1), 16);
      if (code <= 0x10ffff) return String.fromCodePoint(code);
    } else {
      const escaped: Record<string, string> = { n: '\n', r: '\r', t: '\t', '\\': '\\', "'": "'", '$': '$' };
      if (escape in escaped) return escaped[escape];
    }
    valid = false;
    return '';
  });
  return valid ? value.trim() || null : null;
}

/** Lexical inspection only: no condition, function, import or loop is executed. */
export function scanBicepDeclarations(source: string, filename: string): {
  declarations: BicepDeclaration[];
  warnings: string[];
} {
  const warnings: string[] = [];
  const warn = (line: number, message: string) => warnings.push(`${filename}:${line}: ${message}`);
  const text = source.slice(0, MAX_SOURCE_LENGTH);
  if (source.length > text.length) warn(1, 'Bicep source length limit reached; baseline is incomplete.');
  const tokens: Token[] = [];
  const pairs = new Map<number, number>();
  const stack: number[] = [];
  const closing: Record<string, string> = { '{': '}', '[': ']', '(': ')' };
  let line = 1;
  let cursor = 0;
  let stopped = false;

  const blockCommentEnd = (start: number): number => {
    const end = text.indexOf('*/', start + 2);
    if (end >= 0) return end + 2;
    warn(line, 'Unterminated Bicep block comment; baseline is incomplete.');
    return text.length;
  };
  const quotedEnd = (start: number, depth: number): number => {
    if (depth > MAX_DEPTH) {
      warn(line, 'Bicep string nesting limit reached; baseline is incomplete.');
      return text.length;
    }
    if (text.startsWith("'''", start)) {
      const end = text.indexOf("'''", start + 3);
      if (end >= 0) return end + 3;
      warn(line, 'Unterminated Bicep multiline string; baseline is incomplete.');
      return text.length;
    }
    let i = start + 1;
    while (i < text.length) {
      if (text[i] === '\\') { i += 2; continue; }
      if (text[i] === "'") return i + 1;
      if (text.startsWith('${', i)) {
        let braces = 1;
        i += 2;
        while (i < text.length && braces > 0) {
          if (text[i] === "'") { i = quotedEnd(i, depth + 1); continue; }
          if (text.startsWith('/*', i)) { i = blockCommentEnd(i); continue; }
          if (text.startsWith('//', i)) {
            const end = text.indexOf('\n', i + 2);
            i = end < 0 ? text.length : end;
            continue;
          }
          if (text[i] === '{') braces++;
          if (text[i] === '}') braces--;
          if (braces + depth > MAX_DEPTH) {
            warn(line, 'Bicep interpolation nesting limit reached; baseline is incomplete.');
            return text.length;
          }
          i++;
        }
        continue;
      }
      i++;
    }
    warn(line, 'Unterminated Bicep string or interpolation; baseline is incomplete.');
    return text.length;
  };
  const advance = (end: number) => {
    while (cursor < end) {
      if (text[cursor] === '\n') line++;
      cursor++;
    }
  };

  while (cursor < text.length) {
    if (tokens.length >= MAX_TOKENS || stack.length > MAX_DEPTH) {
      warn(line, 'Bicep lexical size or nesting limit reached; baseline is incomplete.');
      stopped = true;
      break;
    }
    const ch = text[cursor];
    if (/\s/.test(ch)) { advance(cursor + 1); continue; }
    if (text.startsWith('//', cursor)) {
      const end = text.indexOf('\n', cursor + 2);
      advance(end < 0 ? text.length : end);
      continue;
    }
    if (text.startsWith('/*', cursor)) { advance(blockCommentEnd(cursor)); continue; }
    const start = cursor;
    const startLine = line;
    let kind: Token['kind'] = 'symbol';
    if (ch === "'") {
      kind = 'string';
      advance(quotedEnd(cursor, 0));
    } else if (/[A-Za-z_]/.test(ch)) {
      kind = 'identifier';
      let end = cursor + 1;
      while (end < text.length && /[\w]/.test(text[end])) end++;
      advance(end);
    } else {
      advance(cursor + 1);
    }
    const value = text.slice(start, cursor);
    const index = tokens.length;
    tokens.push({ kind, value, line: startLine, endLine: line, ...(kind === 'string' ? { literal: stringLiteral(value) } : {}) });
    if (kind !== 'symbol') continue;
    if (closing[value]) stack.push(index);
    else if (value === '}' || value === ']' || value === ')') {
      const open = stack[stack.length - 1];
      if (open === undefined || closing[tokens[open].value] !== value) {
        warn(startLine, 'Mismatched Bicep delimiter; baseline is incomplete.');
      } else {
        stack.pop();
        pairs.set(open, index);
      }
    }
  }
  if (!stopped && stack.length) warn(tokens[stack[0]].line, 'Unterminated Bicep block; baseline is incomplete.');

  interface Envelope { open: number; close: number; end: number; conditional: boolean; loop: boolean }
  const envelope = (start: number, limit: number, depth = 0): Envelope | null => {
    if (depth > MAX_DEPTH || start >= limit) return null;
    const token = tokens[start];
    const end = pairs.get(start);
    if (token.value === '{' && end !== undefined && end < limit) {
      return { open: start, close: end, end, conditional: false, loop: false };
    }
    if (token.value === 'if' && tokens[start + 1]?.value === '(') {
      const conditionEnd = pairs.get(start + 1);
      if (conditionEnd === undefined || conditionEnd >= limit) return null;
      const body = envelope(conditionEnd + 1, limit, depth + 1);
      return body ? { ...body, conditional: true } : null;
    }
    if (token.value === '[' && tokens[start + 1]?.value === 'for' && end !== undefined && end < limit) {
      let ternaries = 0;
      for (let i = start + 2; i < end; i++) {
        const nestedEnd = pairs.get(i);
        if (nestedEnd !== undefined) { i = nestedEnd; continue; }
        if (tokens[i].value === '?') {
          // Neither ?? nor .? consumes the separator before the loop body.
          if (tokens[i + 1]?.value === '?') { i++; continue; }
          if (tokens[i - 1]?.value !== '.') ternaries++;
        }
        if (tokens[i].value !== ':') continue;
        if (ternaries > 0) { ternaries--; continue; }
        const body = envelope(i + 1, end, depth + 1);
        return body ? { ...body, end, loop: true } : null;
      }
    }
    return null;
  };
  const property = (body: Envelope, name: string): string | null => {
    for (let i = body.open + 1; i < body.close; i++) {
      const token = tokens[i];
      if ((token.value === name || token.literal === name) && tokens[i + 1]?.value === ':') {
        const value = tokens[i + 2];
        const next = tokens[i + 3];
        if (value?.kind !== 'string') return null;
        const complete = !next || i + 3 === body.close || next.value === ','
          || (next.line > value.endLine && !['+', '?', '[', '.', '('].includes(next.value));
        return complete ? value.literal ?? null : null;
      }
      const end = pairs.get(i);
      if (end !== undefined) i = end;
    }
    return null;
  };
  const declarations: BicepDeclaration[] = [];
  const visit = (start: number, end: number, parent?: BicepDeclaration): void => {
    for (let i = start; i < end; i++) {
      if (declarations.length >= MAX_DECLARATIONS) {
        warn(tokens[i].line, 'Bicep declaration limit reached; baseline is incomplete.');
        return;
      }
      const token = tokens[i];
      const lineStart = i === start || tokens[i - 1].endLine < token.line;
      if (lineStart && token.kind === 'identifier' && (token.value === 'resource' || token.value === 'module')) {
        const name = tokens[i + 1];
        const type = tokens[i + 2];
        if (name?.kind !== 'identifier' || type?.kind !== 'string') {
          if (name?.value !== ':') warn(token.line, `Unsupported Bicep ${token.value} declaration; baseline is incomplete.`);
          continue;
        }
        let value = i + 3;
        if (tokens[value]?.value === 'existing') value++;
        const body = tokens[value]?.value === '=' ? envelope(value + 1, end) : null;
        if (token.value === 'module') {
          warn(token.line, `Bicep module "${name.value}" is not expanded; baseline is incomplete.`);
        } else {
          const declaredType = (type.literal ?? type.value.slice(1, -1)).split('@')[0];
          const providerType = parent && !/^[^/]+\.[^/]+\//.test(declaredType)
            ? `${parent.providerType}/${declaredType}` : declaredType;
          const conditional = !!(body?.conditional || parent?.conditional);
          const loop = !!(body?.loop || parent?.loop);
          const notes: string[] = [];
          if (conditional) notes.push('The Bicep condition is not evaluated.');
          if (loop) notes.push('One loop declaration is represented; instance multiplicity is not evaluated.');
          if (body?.loop) warn(token.line, `Bicep loop "${name.value}" instance multiplicity is not evaluated; baseline is incomplete.`);
          if (!body) {
            notes.push('The resource body could not be traversed.');
            warn(token.line, `Unsupported or unterminated Bicep resource body "${name.value}"; baseline is incomplete.`);
          }
          if (!type.literal) warn(token.line, `Bicep resource type "${name.value}" is not a literal; baseline is incomplete.`);
          const declaration: BicepDeclaration = {
            logicalName: parent ? `${parent.logicalName}::${name.value}` : name.value,
            providerType,
            resourceName: body ? property(body, 'name') : null,
            kind: body ? property(body, 'kind') : null,
            conditional,
            loop,
            notes,
          };
          declarations.push(declaration);
          if (body) visit(body.open + 1, body.close, declaration);
        }
        if (body) { i = body.end; continue; }
      }
      const blockEnd = pairs.get(i);
      if (blockEnd !== undefined) i = blockEnd;
    }
  };
  visit(0, tokens.length);
  return { declarations, warnings: [...new Set(warnings)] };
}
