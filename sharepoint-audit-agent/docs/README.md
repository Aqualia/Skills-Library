# SharePoint Audit Agent (MVP)

This skill packages a hardened PowerShell scan with an orchestration layer tailored for Claude Code skills and Gemini CLI extensions.

## Components
- **PowerShell core (`agent/powershell`)** – Certificate-authenticated wrapper (`run_audit.ps1`) and reference scan (`Original-Site-Audit.ps1`). Requires PowerShell 7.4+, PnP.PowerShell, and ImportExcel.
- **Python orchestrator (`agent/python/audit_agent.py`)** – Grants Sites.Selected (Read by default), runs the wrapper, and renders Markdown/HTML summaries for each target site.
- **Wrappers (`wrappers/*`)** – LLM manifests that guide users to collect inputs and run local commands only.

## Supported IDE Integrations
- **Claude Code Skill** – Auto-discovered via `wrappers/claude-skill/SKILL.md`. The manifest instructs Claude to gather tenant/app/site inputs, call the PowerShell installer, and run the Python orchestrator while keeping secrets out of transcripts.
- **Gemini CLI Extension** – Defined in `wrappers/gemini-extension/gemini-extension.json`. Adds an `/audit` command that shells out to the orchestrator with safe argument quoting.

## Operational Flow
```
Claude/Gemini prompt
   ↓ (collect inputs, set env vars)
Local shell
   ↓
Python orchestrator  ── grants Sites.Selected (Read) and spawns wrapper
   ↓
PowerShell wrapper    ── connects via certificate, executes original audit script,
                         extracts/normalizes JSON
   ↓
Python analyzer       ── generates Markdown + HTML findings in ./runs/<timestamp>/site-*
```

## Quick Start
Follow [`docs/QUICKSTART.md`](QUICKSTART.md) for environment setup, single-site, and CSV-driven runs.

## Security
See [`docs/SECURITY.md`](SECURITY.md) for guidance on least privilege, report handling, and secret storage. The orchestrator enforces Read-only Sites.Selected unless `--sites-selected-permission Write` is explicitly provided.
