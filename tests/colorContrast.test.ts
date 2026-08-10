import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

/**
 * Guards the WCAG contrast floors of the shared design tokens.
 *
 * These pairings previously regressed silently: the success confirm button rendered white on
 * #10b981 (2.54:1) and every form control was outlined in #cbd5e1 on white (1.48:1), so
 * low-vision users could not locate inputs. Encoding the ratios here means a token tweak that
 * breaks them fails CI instead of shipping.
 */

const tokensPath = join(process.cwd(), 'src', 'styles', 'design-tokens.css');
const css = readFileSync(tokensPath, 'utf8');

/** Splits design-tokens.css into its `:root` (light) and `body.dark-mode` blocks. */
function readTokenScopes(source: string): { light: Map<string, string>; dark: Map<string, string> } {
  const light = new Map<string, string>();
  const dark = new Map<string, string>();
  const darkIndex = source.indexOf('body.dark-mode');
  assert.ok(darkIndex > 0, 'design-tokens.css must declare a body.dark-mode block');

  const collect = (chunk: string, target: Map<string, string>) => {
    for (const match of chunk.matchAll(/(--azd-[a-z0-9-]+)\s*:\s*([^;]+);/gi)) {
      target.set(match[1], match[2].trim());
    }
  };
  collect(source.slice(0, darkIndex), light);
  collect(source.slice(darkIndex), dark);
  // Dark mode only overrides a subset; unspecified tokens inherit the light value.
  for (const [key, value] of light) {
    if (!dark.has(key)) dark.set(key, value);
  }
  return { light, dark };
}

function parseHex(value: string): [number, number, number] {
  const hex = value.trim().replace(/^#/, '');
  const full = hex.length === 3 ? hex.split('').map((c) => c + c).join('') : hex;
  assert.match(full, /^[0-9a-f]{6}$/i, `expected a hex colour, received "${value}"`);
  return [
    parseInt(full.slice(0, 2), 16),
    parseInt(full.slice(2, 4), 16),
    parseInt(full.slice(4, 6), 16),
  ];
}

function relativeLuminance(hex: string): number {
  const channels = parseHex(hex).map((channel) => {
    const srgb = channel / 255;
    return srgb <= 0.03928 ? srgb / 12.92 : ((srgb + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

function contrastRatio(foreground: string, background: string): number {
  const a = relativeLuminance(foreground);
  const b = relativeLuminance(background);
  const [lighter, darker] = a > b ? [a, b] : [b, a];
  return (lighter + 0.05) / (darker + 0.05);
}

const { light, dark } = readTokenScopes(css);

function token(scope: Map<string, string>, name: string): string {
  const value = scope.get(name);
  assert.ok(value, `design-tokens.css must define ${name}`);
  return value as string;
}

/** WCAG 1.4.3 — normal-size text needs 4.5:1. */
const TEXT_PAIRS: Array<{ label: string; fg: string; bg: string }> = [
  { label: 'success confirm button', fg: '--azd-color-on-success', bg: '--azd-color-success-button' },
  { label: 'body text on the app surface', fg: '--azd-color-text', bg: '--azd-color-surface-app' },
  { label: 'body text on elevated surfaces', fg: '--azd-color-text', bg: '--azd-color-surface-elevated' },
  { label: 'secondary text on panels', fg: '--azd-color-text-secondary', bg: '--azd-color-surface-panel' },
  { label: 'muted text on panels', fg: '--azd-color-text-muted', bg: '--azd-color-surface-panel' },
  { label: 'subtle text on sunken surfaces', fg: '--azd-color-text-subtle', bg: '--azd-color-surface-subtle' },
  { label: 'warning text on the warning surface', fg: '--azd-color-warning', bg: '--azd-color-warning-soft' },
  { label: 'danger text on the danger surface', fg: '--azd-color-danger-strong', bg: '--azd-color-danger-soft' },
  { label: 'control text on control background', fg: '--azd-color-text', bg: '--azd-color-control-background' },
  { label: 'brand text on elevated surfaces', fg: '--azd-color-brand-text', bg: '--azd-color-surface-elevated' },
  { label: 'brand text on panels', fg: '--azd-color-brand-text', bg: '--azd-color-surface-panel' },
  { label: 'brand text on the soft brand surface', fg: '--azd-color-brand-text', bg: '--azd-color-brand-soft' },
];

/** WCAG 1.4.11 — UI component boundaries and states need 3:1. */
const UI_PAIRS: Array<{ label: string; fg: string; bg: string }> = [
  { label: 'form control border', fg: '--azd-color-control-border', bg: '--azd-color-control-background' },
  { label: 'form control border on elevated surfaces', fg: '--azd-color-control-border', bg: '--azd-color-surface-elevated' },
  { label: 'strong border on the app surface', fg: '--azd-color-border-strong', bg: '--azd-color-surface-app' },
  { label: 'success button fill on the app surface', fg: '--azd-color-success-button', bg: '--azd-color-surface-app' },
  { label: 'focus ring on control background', fg: '--azd-color-focus', bg: '--azd-color-control-background' },
];

for (const [themeName, scope] of [['light', light], ['dark', dark]] as const) {
  test(`${themeName} theme text tokens meet the WCAG 4.5:1 minimum`, () => {
    for (const pair of TEXT_PAIRS) {
      const ratio = contrastRatio(token(scope, pair.fg), token(scope, pair.bg));
      assert.ok(
        ratio >= 4.5,
        `${themeName}: ${pair.label} (${pair.fg} on ${pair.bg}) is ${ratio.toFixed(2)}:1, below 4.5:1`,
      );
    }
  });

  test(`${themeName} theme UI boundary tokens meet the WCAG 3:1 minimum`, () => {
    for (const pair of UI_PAIRS) {
      const ratio = contrastRatio(token(scope, pair.fg), token(scope, pair.bg));
      assert.ok(
        ratio >= 3,
        `${themeName}: ${pair.label} (${pair.fg} on ${pair.bg}) is ${ratio.toFixed(2)}:1, below 3:1`,
      );
    }
  });
}

/** Every stylesheet under src/, recursively. */
function collectCssFiles(dir: string, found: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectCssFiles(full, found);
    else if (entry.name.endsWith('.css')) found.push(full);
  }
  return found;
}

test('brand-coloured text uses the dark-aware brand text token', () => {
  const offenders: string[] = [];
  for (const file of collectCssFiles(join(process.cwd(), 'src'))) {
    const contents = readFileSync(file, 'utf8');
    // --azd-color-brand-strong has no dark override because it doubles as a fill,
    // so using it for text renders ~1.8:1 on dark surfaces.
    if (/(?:^|[;{\s])color:\s*var\(--azd-color-brand-strong\)/m.test(contents)) {
      offenders.push(file.split(/[\\/]/).pop() as string);
    }
  }
  assert.deepEqual(
    offenders,
    [],
    `these stylesheets must use var(--azd-color-brand-text) for text: ${offenders.join(', ')}`,
  );
});

test('theming is keyed off body.dark-mode, never the OS colour scheme', () => {
  const offenders: string[] = [];
  for (const file of collectCssFiles(join(process.cwd(), 'src'))) {
    // App.tsx toggles body.dark-mode, so a prefers-color-scheme block silently
    // desyncs from the in-app theme switch (e.g. dark OS + light app).
    if (/@media\s*\([^)]*prefers-color-scheme/.test(readFileSync(file, 'utf8'))) {
      offenders.push(file.split(/[\\/]/).pop() as string);
    }
  }
  assert.deepEqual(
    offenders,
    [],
    `these stylesheets must use body.dark-mode / body:not(.dark-mode): ${offenders.join(', ')}`,
  );
});

test('the success confirm button never reuses the low-contrast emphasis fill', () => {
  const primitives = readFileSync(join(process.cwd(), 'src', 'styles', 'modal-primitives.css'), 'utf8');
  const rule = primitives.match(/(?:^|\})\s*\.modal-actions \.btn-success\s*\{([^}]*)\}/);
  assert.ok(rule, 'modal-primitives.css must style .modal-actions .btn-success');
  assert.match(rule[1], /background:\s*var\(--azd-color-success-button\)/);
  assert.match(rule[1], /color:\s*var\(--azd-color-on-success\)/);
  assert.doesNotMatch(
    rule[1],
    /--azd-color-success-emphasis/,
    'success-emphasis is a large-fill accent (2.54:1 against white) and must not carry button text',
  );
});
