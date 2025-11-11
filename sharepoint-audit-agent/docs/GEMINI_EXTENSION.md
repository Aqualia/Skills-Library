# Gemini CLI Extension

## Install
```
gemini extensions install https://github.com/Aqualia/Skills-Library --ref main
```
The repository now surfaces `gemini-extension.json` at its root so Gemini CLI discovers the manifest automatically.

## Usage
```
/extensions list          # confirm sharepoint-audit is available
/audit \
  tenant_id=<TENANT_GUID> \
  app_id=<CLIENT_ID> \
  pfx_path=./certs/audit-agent.pfx \
  site_url=https://<tenant>.sharepoint.com/sites/<SiteName> \
  internal_domains="aqualia.ie aqualia.onmicrosoft.com"
```
- Set `PFX_PASS` in your shell before invoking `/audit`.
- Ensure PowerShell 7.4+, Python 3.10+, and the `markdown` Python package are installed locally.
- The extension quotes every argument; provide literal paths (no shell expansion needed).
- Reports are written to `./runs/<timestamp>/site-*/`.

## Safety
- Commands execute locally on your workstation; the extension never uploads secrets.
- Sites.Selected grants default to **Read** scope. Pass `sites_selected_permission=Write` only when remediation requires write access.
