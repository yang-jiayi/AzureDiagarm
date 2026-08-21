/**
 * Swap the body of a generated table literal in place.
 *
 * The advance tables are a few thousand numbers across two files. Retyping them
 * by hand is how a transcription error gets into a model whose whole job is to
 * be exact, so the numbers only ever move by machine.
 */
import { readFileSync, writeFileSync } from 'node:fs';

const [, , file, spec] = process.argv;
if (!file || !spec) throw new Error('usage: applyTable.ts <file> <name=sourceFile:GENERATED_NAME>');

const [target, source] = spec.split('=');
const [sourceFile, generated] = source.split(':');
const generatedText = readFileSync(sourceFile, 'utf8');

/**
 * Find the line that opens `const <name> ... = [` or `= {`, by name.
 *
 * The name is compared, never interpolated into a pattern. Building the
 * pattern from an argument meant a table called `A.B` would have matched
 * `AxB` — this tool rewrites source files in place from the command line, so
 * a near-miss does not fail, it silently overwrites the wrong constant with
 * the right numbers.
 *
 * Returns the offset just past the opening bracket, which is where the body
 * the caller wants begins.
 */
function openerOf(text: string, name: string): { end: number; open: string } | null {
  let offset = 0;
  for (const raw of text.split('\n')) {
    const line = raw.endsWith('\r') ? raw.slice(0, -1) : raw;
    const m = /^const ([A-Za-z_$][\w$]*)[^=]*= ([[{])$/.exec(line);
    if (m && m[1] === name) return { end: offset + line.length, open: m[2] };
    offset += raw.length + 1;
  }
  return null;
}

/** The body between the opening bracket of `name` and its matching close. */
function body(text: string, name: string, open: string, close: string): string {
  const head = openerOf(text, name);
  if (!head) throw new Error(`cannot find "${name}" in the generated output`);
  if (head.open !== open) throw new Error(`"${name}" opens with "${head.open}", not "${open}"`);
  const end = text.indexOf(`\n${close};`, head.end);
  if (end < 0) throw new Error(`cannot find the end of "${name}"`);
  return text.slice(head.end, end);
}

const original = readFileSync(file, 'utf8');
const opener = openerOf(original, target);
if (!opener) throw new Error(`cannot find "${target}" in ${file}`);
const open = opener.open;
const close = open === '[' ? ']' : '}';
const from = opener.end;
const end = original.indexOf(`\n${close};`, from);
if (end < 0) throw new Error(`cannot find the end of "${target}" in ${file}`);

const replacement = body(generatedText, generated, open, close);
const updated = original.slice(0, from) + replacement + original.slice(end);
writeFileSync(file, updated);
process.stdout.write(`${target}: ${end - from} chars -> ${replacement.length} chars\n`);
