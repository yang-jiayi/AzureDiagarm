# Security Policy

## Supported versions

Security fixes are applied to the latest commit on `main` and to the current
production deployment at <https://azurediagarm.mssql.biz>.

## Reporting a vulnerability

Do not disclose vulnerabilities through a public issue, discussion, or pull
request. Use GitHub's private vulnerability reporting form:

<https://github.com/yang-jiayi/AzureDiagarm/security/advisories/new>

Include:

- the affected component and commit or deployment;
- reproduction steps and required configuration;
- the expected and observed behavior;
- the potential impact; and
- a proof of concept when it can be shared safely.

Do not include live credentials, access tokens, customer data, or private
architecture content. The maintainer will acknowledge the report through the
private advisory and coordinate remediation and disclosure there.

## Deployment security

The public source repository does not grant access to the production Azure
environment. Production uses managed identity, GitHub Actions OIDC, isolated
Container Apps ingress through Azure Front Door, and repository secrets for
sensitive deployment values.
