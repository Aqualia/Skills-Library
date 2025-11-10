#!/usr/bin/env python3
from __future__ import annotations
import argparse, json, os, re, subprocess, sys, textwrap
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

def which_prog(cands):
    from shutil import which
    for c in cands:
        p = which(c)
        if p: return p
    return None

POWERSHELL = which_prog(["pwsh","powershell"]) or "pwsh"

class CmdError(RuntimeError): pass

def run(cmd, timeout=None):
    proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out, err = proc.communicate(timeout=timeout)
    if proc.returncode != 0:
        raise CmdError(f"Command failed ({proc.returncode}): {' '.join(cmd)}\nSTDERR:\n{err}\nSTDOUT:\n{out}")
    return out

def ts(): return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

@dataclass
class RunContext:
    base_out: Path
    run_dir: Path
    logs_dir: Path
    @staticmethod
    def create(base: Path) -> "RunContext":
        run_dir = base / ts()
        logs = run_dir / "logs"
        logs.mkdir(parents=True, exist_ok=True)
        return RunContext(base, run_dir, logs)

def derive_admin_url(site_url: str) -> str:
    import re
    m = re.match(r"https://([a-zA-Z0-9-]+)\.sharepoint\.com/", site_url)
    if not m: raise ValueError(f"Unsupported SharePoint URL: {site_url}")
    tenant = m.group(1)
    return f"https://{tenant}-admin.sharepoint.com"

def grant_sites_selected(site_url: str, tenant: str, app_id: str, pfx_path: Path, pfx_pass: str, display_name="Audit Agent"):
    admin_url = derive_admin_url(site_url)
    ps = textwrap.dedent(f"""
        $sec = ConvertTo-SecureString -String '{pfx_pass.replace("'", "''")}' -AsPlainText -Force
        Connect-PnPOnline -Url '{admin_url}' -Tenant '{tenant}' -ClientId '{app_id}' -CertificatePath '{pfx_path}' -CertificatePassword $sec
        try {{
            $existing = Get-PnPAzureADAppSitePermission -Site '{site_url}' -ErrorAction SilentlyContinue | Where-Object {{$_.AppId -eq '{app_id}'}}
            if ($existing) {{
                if ($existing.Permissions -notcontains 'Write') {{
                    Grant-PnPAzureADAppSitePermission -AppId '{app_id}' -DisplayName '{display_name}' -Site '{site_url}' -Permissions Write | Out-Null
                }}
            }} else {{
                Grant-PnPAzureADAppSitePermission -AppId '{app_id}' -DisplayName '{display_name}' -Site '{site_url}' -Permissions Write | Out-Null
            }}
            Write-Host "Granted Sites.Selected:Write on {site_url}"
        }} finally {{ Disconnect-PnPOnline -ErrorAction SilentlyContinue }}
    """)
    return run([POWERSHELL, "-NoProfile", "-Command", ps])

def run_audit_script(script_path: Path, site_url: str, tenant: str, app_id: str, pfx_path: Path, pfx_pass: str,
                     internal_domains: list[str], out_dir: Path, max_items: int = 50000, batch_size: int = 200,
                     time_budget_minutes: int = 60):
    emit_json = out_dir / "audit.json"
    internal = ' '.join([f"'{d}'" for d in internal_domains]) if internal_domains else ""
    sec_pass = pfx_pass.replace("'", "''")
    ps = textwrap.dedent(f"""
        $sec = ConvertTo-SecureString -String '{sec_pass}' -AsPlainText -Force
        & '{script_path}' `
          -Url '{site_url}' -Tenant '{tenant}' -ClientId '{app_id}' `
          -CertificatePath '{pfx_path}' -CertificatePassword $sec `
          -AutoConfirm -EmitJsonPath '{emit_json}' -MaxItemsToScan {max_items} -BatchSize {batch_size} `
          -InternalDomains {internal}
        if ($LASTEXITCODE) {{ exit $LASTEXITCODE }}
    """)
    run([POWERSHELL, "-NoProfile", "-Command", ps], timeout=time_budget_minutes*60)
    return {"json": emit_json}

DEFAULT_THRESHOLDS = {
    "critical": {"anyone_or_everyone": True, "external_owner": True},
    "high": {"direct_web_perms": True, "unique_items_gt": 250},
    "medium": {"external_item_identities_gte": 10, "group_without_owner": True},
}

CSS = "body{font-family:system-ui,Segoe UI,Roboto,Helvetica,Arial,sans-serif;max-width:960px;margin:2rem auto;padding:0 1rem} h1{font-size:1.8rem} h2{font-size:1.3rem;margin-top:2rem} code,pre{background:#f6f8fa;border:1px solid #eaecef;border-radius:6px;padding:.2rem .4rem}"

def analyze(json_path: Path, thresholds: dict, internal_domains_supplied: bool) -> tuple[str,str]:
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    metrics = data.get("metrics", {})
    counts = {
        "unique_items": metrics.get("itemsWithUniquePermissions") or 0,
        "external_identities": metrics.get("externalUsers") or 0,
        "direct_web_assignments": metrics.get("webDirectAssignments") or 0,
        "groups_without_owner": metrics.get("orphanedGroups") or 0,
        "anyone_or_everyone": bool(metrics.get("anyoneOrEveryoneAtWeb")),
        "external_owner": bool(metrics.get("externalOwnerPresent")),
    }
    findings=[]
    if thresholds["critical"]["anyone_or_everyone"] and counts["anyone_or_everyone"]:
        findings.append(("Critical","'Anyone/Everyone' access detected at web/site scope."))
    if thresholds["critical"]["external_owner"] and counts["external_owner"]:
        findings.append(("Critical","Guest/external user with Owner role detected."))
    if thresholds["high"]["direct_web_perms"] and counts["direct_web_assignments"]:
        findings.append(("High", f"Direct user permissions at web scope: {counts['direct_web_assignments']}") )
    if counts["unique_items"] > thresholds["high"]["unique_items_gt"]:
        findings.append(("High", f"Items with unique permissions: {counts['unique_items']}") )
    if counts["external_identities"] >= thresholds["medium"]["external_item_identities_gte"]:
        findings.append(("Medium", f"External identities with item-level access: {counts['external_identities']}") )
    if thresholds["medium"]["group_without_owner"] and counts["groups_without_owner"]:
        findings.append(("Medium", f"SharePoint groups without owners: {counts['groups_without_owner']}") )

    lines = [
        "# SharePoint Audit — Findings & Recommendations","",
        f"_Generated: {datetime.now().isoformat(timespec='seconds')}_","",
        "## Summary","",
        f"- Items with unique permissions: **{counts['unique_items']}**",
        f"- External identities (item-level): **{counts['external_identities']}**",
        f"- Direct web assignments: **{counts['direct_web_assignments']}**",
        f"- Groups without owners: **{counts['groups_without_owner']}**",
        f"- Anyone/Everyone at web/site: **{counts['anyone_or_everyone']}**",
        f"- External Owner present: **{counts['external_owner']}**","",
        "## Risk Ratings","",
    ]
    if findings:
        for lvl, msg in findings: lines.append(f"- **{lvl}** — {msg}")
    else:
        lines.append("- No risks met the configured thresholds.")
    lines += ["","## Recommendations (PnP Snippets)","",
              "- Review anonymous sharing & site sharing settings.",
              "- Remove direct web permissions where unjustified.",
              "- Reduce item-level unique permissions where possible.",
              "- Ensure each SharePoint group has an owner.","","---",
              "_PII notice: contains user emails and access data. Handle per policy._"]
    md = "\n".join(lines)
    try:
        import markdown
        html_body = markdown.markdown(md, extensions=["tables","fenced_code"])
    except Exception:
        html_body = "<pre>"+md.replace("&","&amp;").replace("<","&lt;")+"</pre>"
    html = f"<!doctype html><html><head><meta charset='utf-8'><title>SharePoint Audit</title><style>{CSS}</style></head><body>{html_body}</body></html>"
    return md, html

def main():
    ap = argparse.ArgumentParser(description="SharePoint Audit Agent — MVP")
    ap.add_argument("--tenant-id", required=True)
    ap.add_argument("--app-id", required=True)
    ap.add_argument("--pfx-path", required=True)
    ap.add_argument("--pfx-pass-env", default="PFX_PASS")
    ap.add_argument("--script-path", required=True)
    m = ap.add_mutually_exclusive_group(required=True)
    m.add_argument("--site-url"); m.add_argument("--csv")
    ap.add_argument("--internal-domains", nargs="*", default=[])
    ap.add_argument("--output", required=True)
    ap.add_argument("--max-items", type=int, default=50000)
    ap.add_argument("--batch-size", type=int, default=200)
    ap.add_argument("--time-budget-minutes", type=int, default=60)
    args = ap.parse_args()

    pfx_pass = os.environ.get(args.pfx_pass_env)
    if not pfx_pass: sys.exit("ERROR: PFX password env var not set")

    out_root = Path(args.output).expanduser().resolve(); out_root.mkdir(parents=True, exist_ok=True)
    run_dir = RunContext.create(out_root).run_dir

    # Collect sites
    sites=[]
    if args.site_url: sites=[args.site_url]
    else:
        import csv
        with open(args.csv, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("SiteUrl"): sites.append(row["SiteUrl"].strip())
    if not sites: sys.exit("No sites provided.")

    for site in sites:
        safe = re.sub(r"[^a-zA-Z0-9._-]+","_", site.strip("/"))
        site_dir = run_dir / f"site-{safe}"; site_dir.mkdir(parents=True, exist_ok=True)

        try:
            print(f"[grant] {site}")
            grant_sites_selected(site, args.tenant_id, args.app_id, Path(args.pfx_path), pfx_pass)
        except Exception as e:
            print(f"Grant failed for {site}: {e}", file=sys.stderr); continue

        try:
            print(f"[audit] {site}")
            arts = run_audit_script(Path(args.script_path), site, args.tenant_id, args.app_id, Path(args.pfx_path),
                                    pfx_pass, args.internal_domains, site_dir, args.max_items, args.batch_size, args.time_budget_minutes)
        except Exception as e:
            print(f"Audit failed for {site}: {e}", file=sys.stderr); continue

        try:
            print(f"[analyze] {site}")
            md, html = analyze(arts["json"], DEFAULT_THRESHOLDS, bool(args.internal_domains))
            (site_dir/"report.md").write_text(md, encoding="utf-8")
            (site_dir/"report.html").write_text(html, encoding="utf-8")
        except Exception as e:
            print(f"Analyze failed for {site}: {e}", file=sys.stderr); continue

    print(f"\nRun complete → {run_dir}")

if __name__ == "__main__":
    main()
