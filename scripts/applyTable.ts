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

/** The body between the opening bracket of `name` and its matching close. */
function body(text: string, name: string, open: string, close: string): string {
  const head = new RegExp(`^const ${name}[^=]*= \\${open}$`, 'm').exec(text);
  if (!head) throw new Error(`cannot find "${name}" in the generated output`);
  const from = head.index + head[0].length;
  const end = text.indexOf(`\n${close};`, from);
  if (end < 0) throw new Error(`cannot find the end of "${name}"`);
  return text.slice(from, end);
}

const original = readFileSync(file, 'utf8');
const opener = new RegExp(`^const ${target}[^=]*= ([\\[{])$`, 'm').exec(original);
if (!opener) throw new Error(`cannot find "${target}" in ${file}`);
const open = opener[1];
const close = open === '[' ? ']' : '}';
const from = opener.index + opener[0].length;
const end = original.indexOf(`\n${close};`, from);
if (end < 0) throw new Error(`cannot find the end of "${target}" in ${file}`);

const replacement = body(generatedText, generated, open, close);
const updated = original.slice(0, from) + replacement + original.slice(end);
writeFileSync(file, updated);
process.stdout.write(`${target}: ${end - from} chars -> ${replacement.length} chars\n`);
