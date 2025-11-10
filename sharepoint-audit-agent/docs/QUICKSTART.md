# Quick Start

## Prereqs
- PowerShell 7.4+ and Python 3.10+
- App-only Entra ID app with **Sites.Selected** and **Directory.Read.All** (admin consent)
- PFX cert + password in env var `PFX_PASS`

## Install modules
```powershell
pwsh ./sharepoint-audit-agent/agent/powershell/Install-Modules.ps1
```

## Single-site run

```bash
export PFX_PASS='your-pfx-password'
python ./sharepoint-audit-agent/agent/python/audit_agent.py \
  --tenant-id <TENANT_GUID> \
  --app-id <CLIENT_ID> \
  --pfx-path ./certs/audit-agent.pfx \
  --script-path "./sharepoint-audit-agent/agent/powershell/Site-Audit.ps1" \
  --site-url https://<tenant>.sharepoint.com/sites/<SiteName> \
  --internal-domains aqualia.ie aqualia.onmicrosoft.com \
  --output ./runs
```

## CSV run

`samples/sites.csv`:

```
SiteUrl
https://<tenant>.sharepoint.com/sites/Finance
https://<tenant>.sharepoint.com/sites/HR
```

Run:

```bash
export PFX_PASS='your-pfx-password'
python ./sharepoint-audit-agent/agent/python/audit_agent.py \
  --tenant-id <TENANT_GUID> \
  --app-id <CLIENT_ID> \
  --pfx-path ./certs/audit-agent.pfx \
  --script-path "./sharepoint-audit-agent/agent/powershell/Site-Audit.ps1" \
  --csv ./sharepoint-audit-agent/samples/sites.csv \
  --output ./runs
```
