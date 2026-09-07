export const baselinePath: string;
export const upstreamRepository: string;
type RunGit = (...args: string[]) => string;
export function readApprovedBaseline(runGit?: RunGit): string;
export function recordSelectedBaseline(
  commit: string,
  runGit?: RunGit,
  write?: (path: string, content: string, encoding: 'utf8') => void,
): void;
