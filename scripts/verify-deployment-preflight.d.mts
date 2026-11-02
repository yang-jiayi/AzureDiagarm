export interface AccessStore {
  type: 'key-vault' | 'tables';
  value: string;
}

export interface DeploymentRuntimeEnvironment {
  [name: string]: string | undefined;
}

export interface PolicyAssignment {
  name?: string;
  id?: string;
  scope?: string;
  policyDefinitionId?: string;
  properties?: {
    scope?: string;
    policyDefinitionId?: string;
  };
}

export interface PolicySetDefinition {
  id?: string;
  policyDefinitions?: Array<{
    policyDefinitionReferenceId?: string;
  }>;
  properties?: {
    policyDefinitions?: Array<{
      policyDefinitionReferenceId?: string;
    }>;
  };
}

export interface PolicyExemption {
  id?: string;
  policyAssignmentId?: string;
  scope?: string;
  exemptionCategory?: string;
  expiresOn?: string;
  policyDefinitionReferenceIds?: string[];
  properties?: {
    policyAssignmentId?: string;
    exemptionCategory?: string;
    expiresOn?: string;
    policyDefinitionReferenceIds?: string[];
  };
}

export function verifyAccessStore(env: DeploymentRuntimeEnvironment): AccessStore;
export function verifyRuntimeConfiguration(env: DeploymentRuntimeEnvironment): {
  accessStore: AccessStore;
  budgetEndpoint: string;
};
export function verifyPreservedRuntimeConfiguration(
  env: DeploymentRuntimeEnvironment,
  currentEnv: Array<{ name?: string; value?: string; secretRef?: string }>,
): void;
export function verifyNoReplicaReduction(
  currentMaxReplicas: string | number,
  desiredMaxReplicas: string | number,
): void;
export function verifyAfdCustomDomain(domains: unknown, host: string | undefined): string;
export function policySetNameForAssignments(assignments: unknown): string;
export function verifyStoragePolicyExemption(input: {
  assignments: unknown;
  policySet: PolicySetDefinition;
  exemption: PolicyExemption;
  storageScope: string | undefined;
  now?: number;
}): string[];
