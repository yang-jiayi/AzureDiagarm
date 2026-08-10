#!/usr/bin/env node
/**
 * Materializes the Docker build context implied by .dockerignore into a temp
 * directory so the image build can be rehearsed without a Docker daemon.
 *
 * Docker's ignore semantics: patterns are evaluated top-to-bottom and the LAST
 * matching pattern wins; a leading '!' re-includes. Because .dockerignore here
 * is allowlist-style, an unlisted file silently disappears from the image —
 * exactly the failure mode that broke the MCP build stage.
 *
 * Usage: node scripts/emulate-docker-context.mjs <destination>
 */
import { cpSync, mkdirSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const destination = resolve(process.argv[2] ?? '');
if (!process.argv[2]) {
  console.error('usage: node scripts/emulate-docker-context.mjs <destination>');
  process.exit(1);
}

const patterns = readFileSync(join(repoRoot, '.dockerignore'), 'utf8')
  .split(/\r?\n/)
  .map((line) => line.trim())
  .filter((line) => line && !line.startsWith('#'))
  .map((line) => {
    const negated = line.startsWith('!');
    const body = negated ? line.slice(1) : line;
    return { negated, matcher: compile(body) };
  });

function compile(pattern) {
  // Translate a Docker ignore pattern to a RegExp over '/'-separated paths.
  let source = '';
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index];
    if (char === '*') {
      if (pattern[index + 1] === '*') {
        index += 1;
        if (pattern[index + 1] === '/') {
          index += 1;
          source += '(?:.*/)?';
        } else {
          source += '.*';
        }
      } else {
        source += '[^/]*';
      }
    } else if (char === '?') {
      source += '[^/]';
    } else {
      source += char.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    }
  }
  // A directory pattern also matches everything beneath it.
  return new RegExp(`^${source}(?:/.*)?$`);
}

function isIncluded(relativePath) {
  let included = true;
  for (const { negated, matcher } of patterns) {
    if (matcher.test(relativePath)) included = negated;
  }
  return included;
}

let copied = 0;
function walk(directory) {
  for (const entry of readdirSync(directory)) {
    if (entry === '.git') continue;
    const absolute = join(directory, entry);
    const relativePath = relative(repoRoot, absolute).split(sep).join('/');
    const stats = statSync(absolute);
    if (stats.isDirectory()) {
      walk(absolute);
    } else if (isIncluded(relativePath)) {
      const target = join(destination, relativePath);
      mkdirSync(dirname(target), { recursive: true });
      cpSync(absolute, target);
      copied += 1;
    }
  }
}

mkdirSync(destination, { recursive: true });
walk(repoRoot);
console.log(`Copied ${copied} files into ${destination}`);
