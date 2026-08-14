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
];
