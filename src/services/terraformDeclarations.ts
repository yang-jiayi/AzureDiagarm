// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export interface TerraformDeclaration {
  logicalName: string;
  providerType: string;
  resourceName: string | null;
  multiple: boolean;
  notes: string[];
}

interface Token {
  kind: 'identifier' | 'string' | 'heredoc' | 'symbol';
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
  let valid = !/[\r\n]/.test(raw);
  const value = raw.slice(1, -1).replace(/\\(u[0-9a-fA-F]{4}|U[0-9a-fA-F]{8}|.)/g, (_match, escape: string) => {
    if (escape.startsWith('u') || escape.startsWith('U')) {
      const code = Number.parseInt(escape.slice(1), 16);
      if (Number.isFinite(code) && code <= 0x10ffff) return String.fromCodePoint(code);
    } else {
      const escaped: Record<string, string> = { n: '\n', r: '\r', t: '\t', '\\': '\\', '"': '"' };
      if (escape in escaped) return escaped[escape];
    }
    valid = false;
    return '';
  }).replace(/\$\$\{/g, '${').replace(/%%\{/g, '%{');
  return valid ? value.trim() || null : null;
}

/** HCL is inspected, never evaluated. Templates and heredocs are opaque to declaration discovery. */
export function scanTerraformDeclarations(source: string, filename: string): {
  declarations: TerraformDeclaration[];
  warnings: string[];
} {
  const warnings: string[] = [];
  const warn = (line: number, message: string) => {
    if (warnings.length < 100) warnings.push(`${filename}:${line}: ${message}; baseline is incomplete.`);
  };
  const text = source.slice(0, MAX_SOURCE_LENGTH);
  if (source.length > text.length) warn(1, 'Terraform source length limit reached');
  const tokens: Token[] = [];
  const pairs = new Map<number, number>();
  const stack: number[] = [];
  const closing: Record<string, string> = { '{': '}', '[': ']', '(': ')' };
  let cursor = 0;
  let line = 1;
  let stopped = false;

  const lineEnd = (start: number): number => {
    const end = text.indexOf('\n', start);
    return end < 0 ? text.length : end;
  };
  const commentEnd = (start: number): number => {
    const end = text.indexOf('*/', start + 2);
    if (end >= 0) return end + 2;
    warn(line, 'Unterminated Terraform block comment');
    return text.length;
  };
  const heredocEnd = (start: number): number => {
    const headerEnd = lineEnd(start);
    const header = /^<<(-?)([A-Za-z_][\w-]*)[ \t]*\r?$/.exec(text.slice(start, headerEnd));
    if (!header || headerEnd === text.length) {
      warn(line, 'Unsupported or unterminated Terraform heredoc');
      return text.length;
    }
    let next = headerEnd + 1;
    while (next < text.length) {
      const end = lineEnd(next);
      const raw = text.slice(next, end);
      const candidate = header[1] ? raw.trim() : raw.trimEnd();
      if (candidate === header[2]) return end;
      next = end + 1;
    }
    warn(line, 'Unterminated Terraform heredoc');
    return text.length;
  };
  const templateEnd = (start: number, depth: number): number => {
    let braces = 1;
    let i = start;
    while (i < text.length) {
      if (braces + depth > MAX_DEPTH) {
        warn(line, 'Terraform template nesting limit reached');
        return text.length;
      }
      if (text[i] === '"') { i = quotedEnd(i, depth + 1).end; continue; }
      if (text.startsWith('<<', i)) { i = heredocEnd(i); continue; }
      if (text.startsWith('/*', i)) { i = commentEnd(i); continue; }
      if (text[i] === '#' || text.startsWith('//', i)) { i = lineEnd(i); continue; }
      if (text[i] === '{') braces++;
      if (text[i] === '}') {
        braces--;
        if (braces === 0) return i + 1;
      }
      i++;
    }
    warn(line, 'Unterminated Terraform template expression');
    return text.length;
  };
  const quotedEnd = (start: number, depth = 0): { end: number; literal: string | null } => {
    if (depth > MAX_DEPTH) {
      warn(line, 'Terraform string nesting limit reached');
      return { end: text.length, literal: null };
    }
    let i = start + 1;
    let template = false;
    let newline = false;
    while (i < text.length) {
      if (text[i] === '\\') { i += 2; continue; }
      if (text[i] === '"') {
        const raw = text.slice(start, i + 1);
        if (newline) warn(line, 'Unsupported newline in Terraform quoted string');
        return { end: i + 1, literal: template ? null : stringLiteral(raw) };
      }
      if (text.startsWith('$${', i) || text.startsWith('%%{', i)) { i += 3; continue; }
      if (text.startsWith('${', i) || text.startsWith('%{', i)) {
        template = true;
        i = templateEnd(i + 2, depth + 1);
        continue;
      }
      if (text[i] === '\r' || text[i] === '\n') newline = true;
      i++;
    }
    warn(line, 'Unterminated Terraform quoted string');
    return { end: text.length, literal: null };
  };
  const advance = (end: number): void => {
    while (cursor < Math.min(end, text.length)) {
      if (text[cursor] === '\n') line++;
      cursor++;
    }
  };

  while (cursor < text.length) {
    if (tokens.length >= MAX_TOKENS || stack.length > MAX_DEPTH) {
      warn(line, 'Terraform lexical size or nesting limit reached');
      stopped = true;
      break;
    }
    const ch = text[cursor];
    if (/\s/.test(ch)) { advance(cursor + 1); continue; }
    if (ch === '#' || text.startsWith('//', cursor)) { advance(lineEnd(cursor)); continue; }
    if (text.startsWith('/*', cursor)) { advance(commentEnd(cursor)); continue; }
    const start = cursor;
    const startLine = line;
    let kind: Token['kind'] = 'symbol';
    let literal: string | null = null;
    if (ch === '"') {
      kind = 'string';
      const quoted = quotedEnd(cursor);
      literal = quoted.literal;
      advance(quoted.end);
    } else if (text.startsWith('<<', cursor)) {
      kind = 'heredoc';
      advance(heredocEnd(cursor));
    } else if (/[A-Za-z_]/.test(ch)) {
      kind = 'identifier';
      let end = cursor + 1;
      while (end < text.length && /[\w-]/.test(text[end])) end++;
      advance(end);
    } else {
      advance(cursor + 1);
    }
    const value = text.slice(start, cursor);
    const index = tokens.length;
    tokens.push({ kind, value, literal, line: startLine, endLine: line });
    if (kind !== 'symbol') continue;
    if (closing[value]) stack.push(index);
    else if (value === '}' || value === ']' || value === ')') {
      const open = stack[stack.length - 1];
      if (open === undefined || closing[tokens[open].value] !== value) {
        warn(startLine, 'Mismatched Terraform delimiter');
      } else {
        stack.pop();
        pairs.set(open, index);
      }
    }
  }
  if (!stopped && stack.length) warn(tokens[stack[0]].line, 'Unterminated Terraform block');

  const blockLabel = (token?: Token): string | null =>
    token?.kind === 'identifier' ? token.value : token?.kind === 'string' ? token.literal ?? null : null;
  const properties = (open: number, close: number): Map<string, string | null> => {
    const result = new Map<string, string | null>();
    for (let i = open + 1; i < close; i++) {
      const token = tokens[i];
      if (token.kind === 'identifier' && tokens[i + 1]?.value === '=') {
        const value = tokens[i + 2];
        const next = tokens[i + 3];
        const complete = i + 3 === close
          || (next?.line > value?.endLine && next.kind === 'identifier');
        result.set(token.value, value?.kind === 'string' && complete ? value.literal ?? null : null);
      }
      const end = pairs.get(i);
      if (end !== undefined) i = end;
    }
    return result;
  };
  const declarations: TerraformDeclaration[] = [];
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    if (token.kind === 'identifier' && token.value === 'module') {
      warn(token.line, `Terraform module "${blockLabel(tokens[i + 1]) ?? '?'}" is not expanded`);
    }
    if (token.kind === 'identifier' && token.value === 'resource') {
      if (declarations.length >= MAX_DECLARATIONS) {
        warn(token.line, 'Terraform declaration limit reached');
        break;
      }
      const type = blockLabel(tokens[i + 1]);
      const name = blockLabel(tokens[i + 2]);
      const open = i + 3;
      if (!type || !name || tokens[open]?.value !== '{') {
        warn(token.line, 'Unsupported Terraform resource declaration');
        continue;
      }
      const close = pairs.get(open);
      const attrs = close === undefined ? new Map<string, string | null>() : properties(open, close);
      const multiple = attrs.has('count') || attrs.has('for_each');
      const notes: string[] = [];
      if (multiple) {
        notes.push('One Terraform declaration is represented; instance multiplicity is not evaluated.');
        warn(token.line, `Terraform resource "${type}.${name}" instance multiplicity is not evaluated`);
      }
      if (close === undefined) {
        notes.push('The Terraform resource body could not be traversed.');
        warn(token.line, `Unsupported or unterminated Terraform resource body "${type}.${name}"`);
      }
      declarations.push({
        logicalName: `${type}.${name}`, providerType: type, resourceName: attrs.get('name') ?? null, multiple, notes,
      });
      i = close ?? tokens.length;
      continue;
    }
    const end = pairs.get(i);
    if (end !== undefined) i = end;
    else if (token.kind === 'symbol' && closing[token.value]) break;
  }
  return { declarations, warnings: [...new Set(warnings)] };
}
