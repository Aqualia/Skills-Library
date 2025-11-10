# Aqualia Skills Library

Aqualia's open skills catalog brings world-class SharePoint security automation to conversational IDEs such as Claude Code and Gemini CLI. Each skill pairs a hardened local agent with an LLM wrapper so regulated enterprises can keep privileged workloads on-device while granting guided UX to millions of builders.

## Current Portfolio

### SharePoint Audit Agent (MVP)
- Certificates + Sites.Selected app-only access to enumerate web/list/item permissions safely.
- PowerShell wrapper preserves legacy scripts while emitting normalized JSON for downstream analysis.
- Python orchestrator handles consent grants, batching, markdown/HTML rendering, and LLM-friendly outputs.
- Claude Skill (`sharepoint-audit-agent/wrappers/claude-skill/SKILL.md`) and Gemini Extension (`sharepoint-audit-agent/wrappers/gemini-extension/gemini-extension.json`) expose the workflow inside IDEs with zero server dependencies.

Read the detailed quickstart at [`sharepoint-audit-agent/docs/QUICKSTART.md`](sharepoint-audit-agent/docs/QUICKSTART.md).

## Repository Layout

```
sharepoint-audit-agent/
├── agent/          # PowerShell + Python runtime bits
├── docs/           # Architecture, security, skill/extension guides
├── samples/        # Input CSV templates
└── wrappers/       # Claude Code skill + Gemini CLI extension manifests
Security/           # Additional operational guidance
```

## Quality & Tooling
- **Runtime requirements**: PowerShell 7.4+ (pwsh) and Python 3.10+; ensure `markdown>=3.6` is installed for HTML reports.
- **CI**: `.github/workflows/ci.yml` runs Python compilation, PSScriptAnalyzer, and doc existence checks on every push/PR.
- **Local validation**: `python3 -m py_compile sharepoint-audit-agent/agent/python/audit_agent.py` and `pwsh -Command "Install-Module PSScriptAnalyzer; Invoke-ScriptAnalyzer sharepoint-audit-agent/agent/powershell"`.

## Security Expectations
- Never commit certificates or secrets; load the PFX password from `PFX_PASS` (see [`sharepoint-audit-agent/docs/SECURITY.md`](sharepoint-audit-agent/docs/SECURITY.md)).
- Grant Sites.Selected at **Read** scope by default; escalate to Write only when remediation absolutely requires it.
- Treat generated reports as PII-bearing artifacts and store them in governed locations.

## Contributing
1. Fork/branch from `main`.
2. Implement changes with clear docs and comments; keep scripts PowerShell 7+ compatible.
3. Run the CI tasks locally (Python compile, PSScriptAnalyzer) and ensure `QUICKSTART.md` reflects any CLI changes.
4. Submit a PR using the provided template; include testing evidence.

Questions? Open an issue via `.github/ISSUE_TEMPLATE/bug.yml` and include sanitized logs.
