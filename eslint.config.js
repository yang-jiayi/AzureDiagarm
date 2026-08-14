import tseslint from '@typescript-eslint/eslint-plugin';
import tsParser from '@typescript-eslint/parser';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';

export default [
  {
    ignores: ['dist/**', 'node_modules/**', '.data/**', 'generated_diagrams/**'],
  },
  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        ecmaVersion: 2022,
        sourceType: 'module',
        ecmaFeatures: { jsx: true },
      },
    },
    plugins: {
      '@typescript-eslint': tseslint,
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
    },
    rules: {
      // Keep lint lightweight while enforcing React hook correctness.
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_', ignoreRestSiblings: true }],

      '@typescript-eslint/no-explicit-any': 'off',
      'react-hooks/rules-of-hooks': 'error',
      'react-hooks/exhaustive-deps': 'error',
      'react-refresh/only-export-components': 'off',
    },
  },
  {
    // Everything under `src/` is bundled for the browser.
    //
    // `vite.config.ts` has no `define` block, so Vite does not shim Node's
    // globals: a `process.env` read in this tree type-checks (the repo has
    // `@types/node` in scope for `scripts/`), bundles without complaint, and
    // passes both the unit tests and the export audit, because both of those
    // run under Node where `process` exists. It then throws
    // `ReferenceError: process is not defined` in the shipped SPA, at runtime,
    // on the user's first click. A debug probe of exactly this shape reached
    // the planner on the PowerPoint export path and survived all six gates.
    //
    // Use `import.meta.env.VITE_*` for anything the browser needs to read.
    files: ['src/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-globals': ['error',
        { name: 'process', message: 'Node globals are not shimmed in the browser bundle. Use import.meta.env.VITE_* instead.' },
        { name: 'Buffer', message: 'Node globals are not shimmed in the browser bundle.' },
        { name: '__dirname', message: 'Node globals are not shimmed in the browser bundle.' },
        { name: '__filename', message: 'Node globals are not shimmed in the browser bundle.' },
        { name: 'require', message: 'Node globals are not shimmed in the browser bundle. Use a static or dynamic import.' },
      ],
    },
  },
  {
    // `no-restricted-globals` above does not cover `console`, and `console` IS
    // defined in the browser - so a debug probe left in an exporter does not
    // throw, it just prints, on every export, in every user's console, forever.
    // One of exactly that shape passed all six gates during this work: type
    // check, script type check, unit tests, the export audit, lint and build.
    //
    // Scoped to the two exporters rather than to `src/**` on purpose. Twelve
    // other services carry deliberate diagnostic logging that predates this
    // rule; widening the scope would either turn lint red on untouched files or
    // force a sweep of disable comments, and neither buys anything here. These
    // two files are where instrumentation is written while chasing a geometry
    // bug, which is the only place the mistake has actually been made.
    files: ['src/services/pptxExporter.ts', 'src/services/visioVsdxExporter.ts'],
    rules: {
      'no-console': ['error', { allow: ['warn', 'error'] }],
    },
  },
];
