# Security Guidance

- Follow [`sharepoint-audit-agent/docs/SECURITY.md`](../sharepoint-audit-agent/docs/SECURITY.md) for operational policies.
- Secrets: never commit certificates/keys. Load the PFX password through the `PFX_PASS` environment variable and rely on the orchestrator’s secure-string handling.
- Least privilege: grant Sites.Selected at **Read** scope unless remediation demands Write, and revoke grants after each audit when mandated by policy.
- Data handling: reports include user emails and access patterns; store the `./runs` directory inside an encrypted, access-controlled location with retention policies.
- Review trail: log who executed the audits and which scope was used to satisfy compliance reviews.
