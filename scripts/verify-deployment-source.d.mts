export interface DeploymentSource {
  ref?: string;
  repository?: string;
  expectedCommit?: string;
  checkedOutCommit?: string;
  token?: string;
}

export function verifyDeploymentSource(
  source: DeploymentSource,
  fetchMain?: typeof fetch,
): Promise<string>;
