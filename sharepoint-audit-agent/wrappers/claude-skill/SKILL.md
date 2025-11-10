---
name: sharepoint-audit
description: >
  Guide and run a SharePoint audit locally. Collect inputs, call PowerShell with certificate
  auth via wrapper, parse audit.json, and render Markdown/HTML. Use only local shell commands.
---

# SharePoint Audit Skill

## When to use
- User asks to audit a SharePoint site or a CSV of sites and wants a local, guided flow.

## What to do
1) Ask for: Tenant ID, App (Client) ID, PFX path, internal domains, site URL or CSV.
2) Run:
   - `pwsh ./sharepoint-audit-agent/agent/powershell/Install-Modules.ps1`
   - `python ./sharepoint-audit-agent/agent/python/audit_agent.py …`
3) On success, point to `./runs/<timestamp>/site-*/report.html`.

## Rules
- Only run local commands. Do not fetch from the internet.
- Never echo secrets. Read PFX password from env var.
