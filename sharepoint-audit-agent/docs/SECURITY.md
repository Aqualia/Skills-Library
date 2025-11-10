# Security
- Never commit PFX/keys. Use env var `PFX_PASS`.
- Prefer least privilege: Sites.Selected per site; remove grants post-audit if required.
- Reports contain PII (emails/access). Store in a secure location and apply retention policies.
