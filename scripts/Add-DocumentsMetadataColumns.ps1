<#
.SYNOPSIS
  Adds the 8 missing metadata columns to the DMS Documents library so it mirrors
  Staging, plus an "Approval Status" column defaulted to "Approved".

.DESCRIPTION
  Root cause fix for blank metadata on approved/bulk-uploaded files in Documents:
  the library was missing 8 of the 11 Staging metadata columns, so tag writes had
  nowhere to land. See docs/superpowers/specs/2026-07-24-documents-metadata-parity-design.md.

  Idempotent: each column is skipped if it already exists, so the script is safe
  to re-run. Internal names are set EXPLICITLY to match Staging exactly — in
  particular Department_x0020_Type (NOT Document_x0020_Type), the frozen-internal-name
  gotcha from CLAUDE.md.

.PREREQUISITES
  Install-Module PnP.PowerShell -Scope CurrentUser
  You need Full Control (Site Owner) or SharePoint Admin on the site.

.EXAMPLE
  ./Add-DocumentsMetadataColumns.ps1 -SiteUrl "https://dcidigitalcom.sharepoint.com/sites/SPFX-Sandbox-Testing-Ground"
#>

param(
  [Parameter(Mandatory = $true)]
  [string] $SiteUrl,

  # Documents library display title.
  [string] $List = "Documents",

  # Document Type term set (must match Staging's binding). From CLAUDE.md.
  [string] $DocumentTypeTermSetId = "0540e66e-7cb3-47ac-b0ef-4e3069387394",

  # Backfill Approval Status = "Approved" on existing items (default only applies to NEW items).
  [switch] $BackfillApprovalStatus
)

$ErrorActionPreference = "Stop"

Connect-PnPOnline -Url $SiteUrl -Interactive

# --- helper: add a field only if its internal name is not already present -----
function Add-IfMissing {
  param(
    [string] $InternalName,
    [scriptblock] $Create
  )
  $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
  if ($null -ne $existing) {
    Write-Host "  = $InternalName already exists - skipped" -ForegroundColor DarkGray
    return
  }
  & $Create
  Write-Host "  + $InternalName created" -ForegroundColor Green
}

Write-Host "Adding DMS metadata columns to '$List' on $SiteUrl" -ForegroundColor Cyan

# --- 6 plain-text level columns (label + Tid pairs) ---------------------------
Add-IfMissing "Business_x0020_Segment" { Add-PnPField -List $List -DisplayName "Business Segment"     -InternalName "Business_x0020_Segment" -Type Text | Out-Null }
Add-IfMissing "BusinessSegmentTid"     { Add-PnPField -List $List -DisplayName "Business Segment Tid" -InternalName "BusinessSegmentTid"     -Type Text | Out-Null }
Add-IfMissing "Department"             { Add-PnPField -List $List -DisplayName "Department"           -InternalName "Department"             -Type Text | Out-Null }
Add-IfMissing "DepartmentTid"          { Add-PnPField -List $List -DisplayName "Department Tid"       -InternalName "DepartmentTid"          -Type Text | Out-Null }
Add-IfMissing "Unit"                   { Add-PnPField -List $List -DisplayName "Unit"                 -InternalName "Unit"                   -Type Text | Out-Null }
Add-IfMissing "UnitTid"                { Add-PnPField -List $List -DisplayName "Unit Tid"             -InternalName "UnitTid"                -Type Text | Out-Null }

# --- Document Date (DateTime) -------------------------------------------------
Add-IfMissing "DocumentDate"           { Add-PnPField -List $List -DisplayName "Document Date"        -InternalName "DocumentDate"           -Type DateTime | Out-Null }

# --- Document Type (Taxonomy, same term set as Staging) -----------------------
# InternalName is Department_x0020_Type by design - DO NOT change to Document_x0020_Type.
Add-IfMissing "Department_x0020_Type" {
  Add-PnPTaxonomyField -List $List -DisplayName "Document Type" `
    -InternalName "Department_x0020_Type" -TermSetId $DocumentTypeTermSetId | Out-Null
}

# --- Approval Status (Choice, default "Approved") -----------------------------
Add-IfMissing "Approval_x0020_Status" {
  Add-PnPField -List $List -DisplayName "Approval Status" -InternalName "Approval_x0020_Status" `
    -Type Choice -Choices @("Approved") | Out-Null
  Set-PnPField -List $List -Identity "Approval_x0020_Status" -Values @{ DefaultValue = "Approved" } | Out-Null
}

# --- optional one-time backfill of existing items -----------------------------
if ($BackfillApprovalStatus) {
  Write-Host "Backfilling Approval Status = Approved on existing items..." -ForegroundColor Cyan
  $items = Get-PnPListItem -List $List -PageSize 500 -Fields "ID", "Approval_x0020_Status", "FSObjType"
  $n = 0
  foreach ($it in $items) {
    if ($it["FSObjType"] -eq 1) { continue } # skip folders
    if ([string]::IsNullOrWhiteSpace($it["Approval_x0020_Status"])) {
      Set-PnPListItem -List $List -Identity $it.Id -Values @{ "Approval_x0020_Status" = "Approved" } | Out-Null
      $n++
    }
  }
  Write-Host "  backfilled $n item(s)" -ForegroundColor Green
}

Write-Host "Done. Verify with the /fields query in the parity spec (all 11 present, types match Staging)." -ForegroundColor Cyan
Disconnect-PnPOnline
