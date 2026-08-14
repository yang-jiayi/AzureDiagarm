/**
 * Types for the JS build script `patch-image-size.mjs`.
 *
 * The script stays plain `.mjs` because `postinstall` runs it with bare node,
 * but `tests/dependencySecurity.test.ts` imports it, so `tsconfig.scripts.json`
 * needs a shape to check that call against.
 */

/** Repo-relative paths of any vendored file the security patch no longer covers. */
export function verifyImageSizeSecurityPatch(): string[];
export function applyImageSizeSecurityPatch(): void;
