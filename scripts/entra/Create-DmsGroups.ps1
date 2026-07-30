<#
.SYNOPSIS
    Creates the DMS security groups for the Group Head Office (GHO) segment —
    Group Finance (GF) and Group Legal, Risk & Compliance (GLRC) units.

.DESCRIPTION
    For each unit it creates up to 3 mail-disabled security groups following the
    CTO naming model:  DMS_GHO_<dept>_<unit>[_UPL|_APR]
        - base (no suffix) = viewer / read
        - _UPL             = uploader
        - _APR             = approver

    Idempotent: skips any group whose displayName already exists. After running,
    it writes CreatedGroups.csv (DisplayName, Id, UnitTermGuid, Role) next to this
    script so you can cross-check / bulk-map them in the Group Map Builder.

.PREREQUISITES
    Install-Module Microsoft.Graph -Scope CurrentUser      # one time
    You must be a Groups Administrator (or higher) in the target tenant.

.NOTES
    Unit term GUIDs are taken from DMS Folder Map.csv. Role is intentionally NOT
    baked into the directory — it is inferred from the name suffix by the app.
    Review $Units before running. Nothing is created until you confirm the prompt.
#>

param(
    [switch]$WhatIf,           # dry-run: list what would be created, create nothing
    [switch]$IncludeUploaders = $true,
    [switch]$IncludeApprovers = $true
)

# ── Unit catalogue (GHO segment) ─────────────────────────────────────────────
# Dept + short code + full label + unit term GUID (from DMS Folder Map.csv).
$Units = @(
    # Group Finance (GF)
    @{ Dept = 'GF';   Code = 'CORU';    Name = 'Compliance & Operational Risk (CORU)'; Guid = 'b60ac538-20bf-423b-a59d-05c2a0471011' },
    @{ Dept = 'GF';   Code = 'FLSDGRE'; Name = 'Finance Land & SDGRE';                  Guid = 'ea62fa60-ec55-4f93-a11d-18f125a07235' },
    @{ Dept = 'GF';   Code = 'FINOPS';  Name = 'Finance Operations';                    Guid = 'f36c790e-faf8-4260-a6c1-12d27408cd80' },
    @{ Dept = 'GF';   Code = 'FUM';     Name = 'Finance Upstream Malaysia';             Guid = '2a45aaf0-81c0-4804-b8fb-5f2e20328317' },
    @{ Dept = 'GF';   Code = 'GSP';     Name = 'Global Strategic Procurement';          Guid = '3384e0fd-87a4-439d-9bb8-c304ba0ddf6d' },
    @{ Dept = 'GF';   Code = 'GITP';    Name = 'Group IT & Projects';                   Guid = '8e58ad71-a9d4-4c2d-a636-3336799dd649' },
    @{ Dept = 'GF';   Code = 'SPRM';    Name = 'Strategy & Price Risk Management';      Guid = '3117ace3-55dc-4efc-a1d6-89b03932cf56' },
    @{ Dept = 'GF';   Code = 'TAX';     Name = 'Tax';                                   Guid = '12694570-ccae-42cc-9406-e78ea827991e' },
    @{ Dept = 'GF';   Code = 'TREAS';   Name = 'Treasury';                              Guid = 'b4ff0cdc-f8a4-4248-8bd9-af9b7eddd888' },
    # Group Legal, Risk & Compliance (GLRC)
    @{ Dept = 'GLRC'; Code = 'GCO';     Name = 'Group Compliance (GCO)';                Guid = '23786464-ee86-4d97-85a8-35449a88b1d7' },
    @{ Dept = 'GLRC'; Code = 'GL';      Name = 'Group Legal';                           Guid = '5eae1f72-5e69-4b10-ad9c-24accde38e17' },
    @{ Dept = 'GLRC'; Code = 'GR';      Name = 'Group Risk';                            Guid = '9df36014-e32d-4c26-aa74-a54262149aa1' }
)

# ── Roles to create per unit ─────────────────────────────────────────────────
$Roles = @( @{ Suffix = '';     Label = 'viewer (read)' } )
if ($IncludeUploaders) { $Roles += @{ Suffix = '_UPL'; Label = 'uploader' } }
if ($IncludeApprovers) { $Roles += @{ Suffix = '_APR'; Label = 'approver' } }

# ── Build the full plan ──────────────────────────────────────────────────────
$Plan = foreach ($u in $Units) {
    foreach ($r in $Roles) {
        $display = "DMS_GHO_$($u.Dept)_$($u.Code)$($r.Suffix)"
        [pscustomobject]@{
            DisplayName  = $display
            MailNickname = $display                 # valid: no spaces/&, underscores OK
            Description  = "GHO > $($u.Dept) > $($u.Name) - $($r.Label). Unit term: $($u.Guid)"
            UnitTermGuid = $u.Guid
            Role         = if ($r.Suffix) { $r.Suffix.TrimStart('_') } else { 'MEMBER' }
        }
    }
}

Write-Host ("Planned groups: {0} ({1} units x {2} roles)" -f $Plan.Count, $Units.Count, $Roles.Count) -ForegroundColor Cyan
$Plan | Format-Table DisplayName, Role -AutoSize

if ($WhatIf) { Write-Host "`n-WhatIf set. Nothing created." -ForegroundColor Yellow; return }

$confirm = Read-Host "`nCreate these $($Plan.Count) security groups in the CURRENT tenant? (y/N)"
if ($confirm -ne 'y') { Write-Host "Aborted." -ForegroundColor Yellow; return }

# ── Connect & create ─────────────────────────────────────────────────────────
Import-Module Microsoft.Graph.Groups -ErrorAction Stop
Connect-MgGraph -Scopes 'Group.ReadWrite.All' -NoWelcome

$results = foreach ($g in $Plan) {
    $existing = Get-MgGroup -Filter "displayName eq '$($g.DisplayName)'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "SKIP (exists): $($g.DisplayName)" -ForegroundColor DarkGray
        [pscustomobject]@{ DisplayName = $g.DisplayName; Id = $existing.Id; UnitTermGuid = $g.UnitTermGuid; Role = $g.Role; Status = 'existing' }
    }
    else {
        $new = New-MgGroup -DisplayName $g.DisplayName -MailNickname $g.MailNickname `
            -Description $g.Description -SecurityEnabled -MailEnabled:$false -ErrorAction Stop
        Write-Host "CREATED: $($g.DisplayName)" -ForegroundColor Green
        [pscustomobject]@{ DisplayName = $g.DisplayName; Id = $new.Id; UnitTermGuid = $g.UnitTermGuid; Role = $g.Role; Status = 'created' }
    }
}

$out = Join-Path $PSScriptRoot 'CreatedGroups.csv'
$results | Export-Csv -Path $out -NoTypeInformation -Encoding UTF8
Write-Host "`nDone. Wrote $out" -ForegroundColor Cyan
