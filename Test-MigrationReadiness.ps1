<#
.SYNOPSIS
    Microsoft 365 Migration Pre-Flight Readiness Check.
    Runs BEFORE a tenant-to-tenant migration and reports, as a green/amber/red
    checklist, whether everything is ready — so failed credential verification and
    half the common problems are caught up front.

.CHECKS
    Modern auth & org       : modern auth on, EWS enabled at org, security defaults
    Service account         : exists & enabled, MFA state (must be off for the tool),
                              ApplicationImpersonation, EWS on its mailbox
    EWS & throttling        : current EWS throttling policy values
    Destination provisioning: per user — account enabled, mailbox, OneDrive, Teams
    Domain / cutover        : verified domains

.HOW TO RUN
    Run this on EACH tenant (source and destination) so both are ready.
    .\Test-MigrationReadiness.ps1
    .\Test-MigrationReadiness.ps1 -ServiceAccount migration@contoso.com -UsersCsv .\dest-users.csv
    (CSV needs a column named DestUPN or UPN. Without it, a sample of users is checked.)

.REQUIREMENTS
    Modules: Microsoft.Graph.Authentication, ExchangeOnlineManagement
    Sign in as an admin. Graph scopes are requested automatically.

.OUTPUT
    MigrationReadiness.html   (open, then Print -> Save as PDF)   +   Readiness-Users.csv
#>

param(
    [string]$ServiceAccount,
    [string]$UsersCsv,
    [int]$SampleSize = 15,
    [switch]$SkipExchange
)

$ErrorActionPreference = "Stop"
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) { Write-Host "Installing Microsoft.Graph.Authentication..." -ForegroundColor Cyan; Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force }
Import-Module Microsoft.Graph.Authentication

if (-not $ServiceAccount) { $ServiceAccount = Read-Host "Enter the migration SERVICE ACCOUNT UPN (the account the migration tool will use)" }

$checks = @()   # each: cat, name, status (Pass/Warn/Fail/Info/Skip), value, note
function Add-Check($cat,$name,$status,$value,$note){ $script:checks += [PSCustomObject]@{ cat=$cat; name=$name; status=$status; value=$value; note=$note } }

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All","Policy.Read.All","Organization.Read.All","UserAuthenticationMethod.Read.All" -NoWelcome
$org = (Get-MgContext).TenantId
$tenantName = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization").value[0].displayName

# ---- Exchange Online ----
$exo = $false
if (-not $SkipExchange) {
    try {
        if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { Write-Host "Installing ExchangeOnlineManagement..." -ForegroundColor Cyan; Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force }
        Import-Module ExchangeOnlineManagement
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false
        $exo = $true
    } catch { Write-Host "Exchange Online connect failed — EWS/throttling/mailbox checks will be skipped: $($_.Exception.Message)" -ForegroundColor DarkYellow }
}

# ============ MODERN AUTH & ORG ============
if ($exo) {
    try {
        $oc = Get-OrganizationConfig
        Add-Check "Modern auth & org" "Modern authentication enabled" ($(if($oc.OAuth2ClientProfileEnabled){"Pass"}else{"Fail"})) ("{0}" -f $oc.OAuth2ClientProfileEnabled) "Required for OAuth-based migration."
        $ews = $oc.EwsEnabled
        Add-Check "Modern auth & org" "EWS enabled at org level" ($(if($ews -eq $false){"Fail"}else{"Pass"})) ($(if($ews -eq $null){"Default (enabled)"}else{"$ews"})) "MigrationWiz reads/writes mail over EWS."
        Add-Check "Modern auth & org" "EWS application access policy" "Info" ("{0}" -f $(if($oc.EwsApplicationAccessPolicy){$oc.EwsApplicationAccessPolicy}else{"None (all allowed)"})) "If set to allow-list, the migration app must be included."
    } catch { Add-Check "Modern auth & org" "Organization config" "Skip" "n/a" "Could not read Get-OrganizationConfig." }
}
try {
    $sd = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy"
    Add-Check "Modern auth & org" "Security defaults" ($(if($sd.isEnabled){"Warn"}else{"Pass"})) ($(if($sd.isEnabled){"Enabled"}else{"Disabled"})) $(if($sd.isEnabled){"Security defaults enforce MFA for all — the migration account must be exempt."}else{"Off — good; MFA not force-applied to all."})
} catch { Add-Check "Modern auth & org" "Security defaults" "Skip" "n/a" "Needs Policy.Read.All." }
try {
    $ca = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies").value
    $mfaPol = @($ca | Where-Object { $_.state -eq "enabled" -and ($_.grantControls.builtInControls -contains "mfa") })
    Add-Check "Modern auth & org" "Conditional Access MFA policies" ($(if($mfaPol.Count -gt 0){"Warn"}else{"Info"})) ("{0} enabled MFA policies" -f $mfaPol.Count) $(if($mfaPol.Count -gt 0){"Confirm the migration account is EXCLUDED from these."}else{"No MFA-enforcing CA policies found."})
} catch { Add-Check "Modern auth & org" "Conditional Access policies" "Skip" "n/a" "Needs Policy.Read.All." }

# ============ SERVICE ACCOUNT ============
$saId = $null
try {
    $sa = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$ServiceAccount`?`$select=id,displayName,userPrincipalName,accountEnabled"
    $saId = $sa.id
    Add-Check "Service account" "Account exists" "Pass" $sa.userPrincipalName ""
    Add-Check "Service account" "Account enabled" ($(if($sa.accountEnabled){"Pass"}else{"Fail"})) ("{0}" -f $sa.accountEnabled) $(if(-not $sa.accountEnabled){"Enable the account before migration."}else{""})
} catch { Add-Check "Service account" "Account exists" "Fail" $ServiceAccount "Account not found in this tenant — check the UPN." }

if ($saId) {
    # MFA per-user state
    try {
        $req = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/users/$saId/authentication/requirements"
        $mfa = "$($req.perUserMfaState)"
        Add-Check "Service account" "Per-user MFA state" ($(if($mfa -eq "disabled"){"Pass"}else{"Fail"})) $mfa $(if($mfa -ne "disabled"){"MFA must be OFF for the migration account, or the tool cannot authenticate. Disable per-user MFA / exclude from CA."}else{"Good — no MFA on this account."})
    } catch { Add-Check "Service account" "Per-user MFA state" "Warn" "Unknown" "Could not read MFA state — verify manually in the admin center that MFA is off for this account." }

    if ($exo) {
        # impersonation
        try {
            $imp = Get-ManagementRoleAssignment -Role "ApplicationImpersonation" -GetEffectiveUsers -ErrorAction Stop | Where-Object { $_.EffectiveUserName -eq $ServiceAccount -or $_.RoleAssigneeName -eq $ServiceAccount }
            Add-Check "Service account" "ApplicationImpersonation rights" ($(if($imp){"Pass"}else{"Warn"})) ($(if($imp){"Assigned"}else{"Not found"})) $(if(-not $imp){"Grant impersonation to the migration account (or configure per BitTitan modern-auth app permissions)."}else{""})
        } catch { Add-Check "Service account" "ApplicationImpersonation rights" "Skip" "n/a" "Could not query role assignment." }
        # EWS on SA mailbox
        try {
            $cas = Get-CASMailbox -Identity $ServiceAccount -ErrorAction Stop
            Add-Check "Service account" "EWS enabled on its mailbox" ($(if($cas.EwsEnabled -eq $false){"Fail"}else{"Pass"})) ($(if($cas.EwsEnabled -eq $null){"Default (enabled)"}else{"$($cas.EwsEnabled)"})) $(if($cas.EwsEnabled -eq $false){"Enable EWS on this mailbox."}else{""})
        } catch { Add-Check "Service account" "EWS on its mailbox" "Skip" "n/a" "Mailbox not found or not readable." }
    }
}

# ============ EWS & THROTTLING ============
if ($exo) {
    try {
        $tp = Get-ThrottlingPolicy | Select-Object -First 1
        $val = "EwsMaxConcurrency=$($tp.EwsMaxConcurrency); EwsMaxBurst=$($tp.EwsMaxBurst); EwsCutoffBalance=$($tp.EwsCutoffBalance)"
        Add-Check "EWS & throttling" "Throttling policy (EWS)" "Info" $val "Exchange Online throttling can't be disabled by script. Expect pacing; impersonation + retries handle it. For large jobs, request an increased EWS policy via admin-center diagnostics."
    } catch { Add-Check "EWS & throttling" "Throttling policy" "Skip" "n/a" "Could not read throttling policy." }
}

# ============ DESTINATION PROVISIONING (per user) ============
$targets = @()
if ($UsersCsv -and (Test-Path $UsersCsv)) {
    $csv = Import-Csv $UsersCsv
    $col = if ($csv[0].PSObject.Properties.Name -contains "DestUPN") { "DestUPN" } elseif ($csv[0].PSObject.Properties.Name -contains "UPN") { "UPN" } else { $csv[0].PSObject.Properties.Name[0] }
    $targets = $csv | ForEach-Object { $_.$col } | Where-Object { $_ }
} else {
    Write-Host "No -UsersCsv given; sampling $SampleSize users for provisioning checks..." -ForegroundColor DarkYellow
    $samp = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users?`$top=$SampleSize&`$select=userPrincipalName,accountEnabled").value
    $targets = $samp | ForEach-Object { $_.userPrincipalName }
}

$userRows = @(); $mbOk=0;$odOk=0;$tmOk=0;$enOk=0; $n=0
foreach ($upn in $targets) {
    $n++
    $enabled=$false;$mailbox="No";$onedrive="No";$teams="No"
    try {
        $u = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn`?`$select=accountEnabled,assignedPlans,displayName"
        $enabled = [bool]$u.accountEnabled; if($enabled){$enOk++}
        $teamsPlan = @($u.assignedPlans | Where-Object { "$($_.service)" -match "Teams|TeamspaceAPI" -and $_.capabilityStatus -eq "Enabled" })
        if ($teamsPlan.Count -gt 0) { $teams="Yes"; $tmOk++ }
    } catch {}
    # OneDrive
    try { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$upn/drive" | Out-Null; $onedrive="Yes"; $odOk++ } catch {}
    # Mailbox
    if ($exo) {
        try { Get-EXOMailbox -Identity $upn -ErrorAction Stop | Out-Null; $mailbox="Yes"; $mbOk++ } catch { $mailbox="No" }
    } else { $mailbox="?" }
    $userRows += [PSCustomObject]@{ User=$upn; Enabled=$enabled; Mailbox=$mailbox; OneDrive=$onedrive; Teams=$teams }
}
$userRows | Export-Csv ".\Readiness-Users.csv" -NoTypeInformation -Encoding UTF8

$provNote = "$mbOk/$n mailboxes, $odOk/$n OneDrive, $tmOk/$n Teams, $enOk/$n enabled"
Add-Check "Destination provisioning" "Mailboxes provisioned" ($(if($exo){ if($mbOk -eq $n){"Pass"}elseif($mbOk -gt 0){"Warn"}else{"Fail"} }else{"Skip"})) "$mbOk / $n" "Users must have a provisioned mailbox at the destination."
Add-Check "Destination provisioning" "OneDrive provisioned" ($(if($odOk -eq $n){"Pass"}elseif($odOk -gt 0){"Warn"}else{"Fail"})) "$odOk / $n" "OneDrive must be provisioned (user has signed in once, or pre-provisioned)."
Add-Check "Destination provisioning" "Teams provisioned/licensed" ($(if($tmOk -eq $n){"Pass"}elseif($tmOk -gt 0){"Warn"}else{"Fail"})) "$tmOk / $n" "Teams plan must be enabled on the license."
Add-Check "Destination provisioning" "Accounts enabled" ($(if($enOk -eq $n){"Pass"}elseif($enOk -gt 0){"Warn"}else{"Fail"})) "$enOk / $n" ""

# ============ DOMAIN / CUTOVER ============
try {
    $doms = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains").value
    $verified = @($doms | Where-Object { $_.isVerified }).Count
    Add-Check "Domain / cutover" "Verified domains" "Info" ("{0} verified of {1}" -f $verified,$doms.Count) "For cutover, the target domain must be verified on the destination tenant."
} catch { Add-Check "Domain / cutover" "Domains" "Skip" "n/a" "" }

# ---- summarise ----
$pass = @($checks | Where-Object status -eq "Pass").Count
$warn = @($checks | Where-Object status -eq "Warn").Count
$fail = @($checks | Where-Object status -eq "Fail").Count
$groups = $checks | Group-Object cat | ForEach-Object { @{ name=$_.Name; checks=@($_.Group | ForEach-Object { @{ name=$_.name; status=$_.status; value=$_.value; note=$_.note } }) } }

$data = [ordered]@{
    tenant=$tenantName; generated=(Get-Date -Format "yyyy-MM-dd HH:mm"); serviceAccount=$ServiceAccount
    summary=@{ pass=$pass; warn=$warn; fail=$fail }
    groups=$groups
    users=@($userRows | Select-Object -First 40 | ForEach-Object { @{ User=$_.User; Enabled=$_.Enabled; Mailbox=$_.Mailbox; OneDrive=$_.OneDrive; Teams=$_.Teams } })
    exo=$exo
}
$json = $data | ConvertTo-Json -Depth 7 -Compress

$html = @'
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Migration Pre-Flight Readiness</title>
<link href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,700;12..96,800&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#0A0E0D;--surface:#111917;--surface-2:#0E1513;--line:#1E2A26;--emerald:#10B981;--bright:#34D399;--amber:#fbbf24;--red:#f87171;--blue:#60a5fa;--text:#E9EFEC;--muted:#8FA29A;--dim:#5d716a;--mono:'JetBrains Mono',monospace;--disp:'Bricolage Grotesque',sans-serif;--body:'Inter',sans-serif;}
*{box-sizing:border-box;margin:0;padding:0}body{background:var(--bg);color:var(--text);font-family:var(--body);line-height:1.55}
.wrap{max-width:960px;margin:0 auto;padding:0 20px}
header{border-bottom:1px solid var(--line)}nav{display:flex;align-items:center;justify-content:space-between;height:58px}
.brand{font-family:var(--disp);font-weight:800;font-size:16px}.brand span{color:var(--bright)}
.print{background:var(--emerald);color:#04140f;border:none;border-radius:8px;padding:9px 16px;font-weight:600;font-size:13px;cursor:pointer}
.hero{padding:24px 0 6px}.eyebrow{font-family:var(--mono);font-size:11.5px;letter-spacing:.16em;text-transform:uppercase;color:var(--emerald)}
.hero h1{font-family:var(--disp);font-weight:800;font-size:26px;margin:9px 0 5px}.hero p{color:var(--muted);font-size:12.5px;font-family:var(--mono)}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:16px 0}
.card{background:var(--surface);border:1px solid var(--line);border-radius:13px;padding:18px;text-align:center}.card .n{font-family:var(--disp);font-weight:800;font-size:32px}.card .l{font-size:12px;color:var(--muted);margin-top:3px}
.card.p .n{color:var(--bright)}.card.w .n{color:var(--amber)}.card.f .n{color:var(--red)}
.group{background:var(--surface);border:1px solid var(--line);border-radius:15px;padding:8px 20px 14px;margin-bottom:14px}
.group h3{font-family:var(--disp);font-size:15px;margin:14px 0 6px}
.chk{display:flex;align-items:flex-start;gap:12px;padding:11px 0;border-bottom:1px solid var(--line)}
.chk:last-child{border-bottom:none}
.dot{width:22px;height:22px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;margin-top:1px}
.dot.Pass{background:rgba(52,211,153,.15);color:var(--bright)}.dot.Warn{background:rgba(251,191,36,.15);color:var(--amber)}.dot.Fail{background:rgba(248,113,113,.15);color:var(--red)}.dot.Info{background:rgba(96,165,250,.15);color:var(--blue)}.dot.Skip{background:rgba(143,162,154,.15);color:var(--dim)}
.chk .mid{flex:1}.chk .nm{font-size:13.5px;font-weight:600;color:var(--text)}.chk .nt{font-size:12px;color:var(--muted);margin-top:2px}
.chk .val{font-family:var(--mono);font-size:11px;color:var(--dim);white-space:nowrap;text-align:right;max-width:210px;overflow-wrap:anywhere}
table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:6px}th{text-align:left;font-family:var(--mono);font-size:10px;text-transform:uppercase;color:var(--dim);padding:8px 10px;border-bottom:1px solid var(--line)}td{padding:8px 10px;border-bottom:1px solid var(--line);color:var(--muted)}td b{color:var(--text);font-weight:500}
.yes{color:var(--bright)}.no{color:var(--red)}.q{color:var(--dim)}
footer{border-top:1px solid var(--line);padding:20px 0;text-align:center;font-size:12px;color:var(--dim)}
@media print{body{background:#fff;color:#111}.print{display:none}.card,.group{border-color:#ccc;background:#fff}.brand span,.eyebrow{color:#0a7a55}th,td,.chk .nm{color:#222}}
</style></head><body>
<header><nav class="wrap"><div class="brand">Nabeel <span>Mukhtar</span> · Migration Pre-Flight</div><button class="print" onclick="window.print()">🖨 Save as PDF</button></nav></header>
<div class="wrap">
<section class="hero"><div class="eyebrow">Pre-migration readiness</div><h1 id="tenant"></h1><p id="meta"></p></section>
<div class="grid">
<div class="card p"><div class="n" id="p">0</div><div class="l">Passed</div></div>
<div class="card w"><div class="n" id="w">0</div><div class="l">Warnings</div></div>
<div class="card f"><div class="n" id="f">0</div><div class="l">Failed</div></div>
</div>
<div id="groups"></div>
<div class="group"><h3>Destination provisioning — sampled users</h3><table><thead><tr><th>User</th><th>Enabled</th><th>Mailbox</th><th>OneDrive</th><th>Teams</th></tr></thead><tbody id="users"></tbody></table></div>
</div>
<footer class="wrap">Migration Pre-Flight Readiness · Nabeel Mukhtar Hussain · run on both source & destination</footer>
<script>
const D=/*__DATA__*/;const $=i=>document.getElementById(i);
$("tenant").textContent=D.tenant+" — Readiness Check";
$("meta").textContent="Generated "+D.generated+" · service account: "+D.serviceAccount+(D.exo?"":" · (Exchange checks skipped)");
$("p").textContent=D.summary.pass;$("w").textContent=D.summary.warn;$("f").textContent=D.summary.fail;
const ic={Pass:"✓",Warn:"!",Fail:"✕",Info:"i",Skip:"–"};
$("groups").innerHTML=D.groups.map(g=>`<div class="group"><h3>${g.name}</h3>${g.checks.map(c=>`<div class="chk"><div class="dot ${c.status}">${ic[c.status]||'?'}</div><div class="mid"><div class="nm">${c.name}</div>${c.note?`<div class="nt">${c.note}</div>`:''}</div><div class="val">${c.value||''}</div></div>`).join("")}</div>`).join("");
$("users").innerHTML=D.users.length?D.users.map(u=>`<tr><td><b>${u.User}</b></td><td class="${u.Enabled?'yes':'no'}">${u.Enabled?'Yes':'No'}</td><td class="${u.Mailbox=='Yes'?'yes':(u.Mailbox=='No'?'no':'q')}">${u.Mailbox}</td><td class="${u.OneDrive=='Yes'?'yes':'no'}">${u.OneDrive}</td><td class="${u.Teams=='Yes'?'yes':'no'}">${u.Teams}</td></tr>`).join(""):`<tr><td colspan="5" style="color:var(--dim)">No users checked.</td></tr>`;
</script></body></html>
'@
$html = $html.Replace('/*__DATA__*/', $json)
$out = ".\MigrationReadiness.html"; $html | Out-File $out -Encoding utf8

Write-Host "`n===== MIGRATION READINESS =====" -ForegroundColor Cyan
Write-Host ("Passed {0} | Warnings {1} | Failed {2}" -f $pass,$warn,$fail)
Write-Host ("Provisioning: {0}" -f $provNote)
Write-Host "Report: $out  (open, then Print -> Save as PDF)" -ForegroundColor Green
Write-Host "CSV: Readiness-Users.csv" -ForegroundColor Green
try { Invoke-Item $out } catch {}
if ($exo) { Disconnect-ExchangeOnline -Confirm:$false | Out-Null }
Disconnect-MgGraph | Out-Null
