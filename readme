# Microsoft 365 Migration Pre-Flight Readiness Check

A PowerShell tool that validates a Microsoft 365 tenant **before** a tenant-to-tenant migration and produces a clean green / amber / red readiness report — so failed credential verification and half the common cutover problems are caught up front, not mid-migration.

Run it on both the **source** and **destination** tenants. In a few minutes you get a shareable HTML report (printable to PDF) and a CSV of per-user provisioning status.

---

## Why this exists

Most tenant-to-tenant migration failures trace back to a handful of things that were never checked beforehand: MFA still enforced on the migration service account, EWS turned off, mailboxes or OneDrive not provisioned at the destination, or the domain not yet verified. This script checks all of them in one pass and tells you exactly what to fix.

## What it checks

**Modern auth & organisation**
- Modern authentication enabled
- EWS enabled at the organisation level (and any EWS application access policy)
- Security defaults state (MFA enforcement)
- Conditional Access policies that enforce MFA

**Service account** (the account the migration tool will use)
- Account exists and is enabled
- Per-user MFA state — must be **off**, or the migration tool cannot authenticate
- ApplicationImpersonation rights
- EWS enabled on the account's mailbox

**EWS & throttling**
- Current EWS throttling policy values (with guidance)

**Destination provisioning** (per user, from a CSV or a sample)
- Account enabled
- Mailbox provisioned
- OneDrive provisioned
- Teams provisioned / licensed

**Domain / cutover**
- Verified domains on the tenant

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- Modules (installed automatically if missing):
  - `Microsoft.Graph.Authentication`
  - `ExchangeOnlineManagement`
- An admin sign-in (Global Reader is enough for most checks; some Exchange checks need a role that can read organisation and mailbox config)
- Microsoft Graph scopes (requested at sign-in): `User.Read.All`, `Directory.Read.All`, `Policy.Read.All`, `Organization.Read.All`, `UserAuthenticationMethod.Read.All`

## Usage

```powershell
# Simplest — the script prompts for the migration service account UPN
.\Test-MigrationReadiness.ps1

# Provide the service account and a list of destination users to check
.\Test-MigrationReadiness.ps1 -ServiceAccount migration@contoso.com -UsersCsv .\dest-users.csv

# Skip Exchange Online checks (Graph-only run)
.\Test-MigrationReadiness.ps1 -SkipExchange
```

**CSV format** — a single column named `DestUPN` (or `UPN`):

```csv
DestUPN
ali.khan@contoso.com
sara.malik@contoso.com
```

If no CSV is supplied, the script samples a set of users (default 15, change with `-SampleSize`).

## Output

- `MigrationReadiness.html` — the readiness report. Open it, then **Print -> Save as PDF** to hand to a client or attach to a runbook.
- `Readiness-Users.csv` — per-user provisioning detail (enabled, mailbox, OneDrive, Teams).

Each check shows a status (Pass / Warn / Fail / Info) with the value found and a short note on what to do about it.

## Parameters

| Parameter | Description | Default |
|---|---|---|
| `-ServiceAccount` | UPN of the migration service account | prompted |
| `-UsersCsv` | Path to a CSV of destination users to check | none (samples users) |
| `-SampleSize` | Users to sample when no CSV is given | 15 |
| `-SkipExchange` | Skip Exchange Online checks | off |

## Honest notes

- The script is **read-only** — it never changes any setting; it only reports.
- Exchange Online throttling cannot be disabled by a script. The report shows the current policy and explains how to handle throttling (impersonation, retries, or an increased EWS policy via admin-centre diagnostics).
- Per-user MFA state is read from the Graph beta endpoint; if your tenant blocks that read, the check reports "unknown" and asks you to verify manually.
- Run it on **both** tenants — source readiness (EWS, service account) and destination readiness (provisioning, domain) matter for different reasons.

## Roadmap

- SharePoint site-level provisioning checks
- Optional export to `.xlsx`
- Batch mode for multiple client tenants

## Author

**Nabeel Mukhtar Hussain** — Microsoft 365 & Cloud Migration Specialist
Helping enterprises move to the cloud — without the chaos.

- Website: [nabeelmukhtar.xyz](https://nabeelmukhtar.xyz)
- LinkedIn: [nabeelmukhtarhussain](https://www.linkedin.com/in/nabeelmukhtarhussain/)
- Email: nabeelmukhtar00@gmail.com

## License

Released under the MIT License — free to use and adapt. See `LICENSE`.

---

*Part of a broader Microsoft 365 migration toolkit. If this saved you a painful cutover, a star on the repo is appreciated.*
