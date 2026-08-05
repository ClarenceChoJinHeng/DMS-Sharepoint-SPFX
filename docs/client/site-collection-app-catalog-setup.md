# Enabling the App Catalog on the CRS Site — Guide for SDG's SharePoint Administrator

**What we are asking for:** one command, run once, that creates an empty document library on the
CRS site. It takes about five minutes end to end, most of which is signing in.

**Who can do it:** an account holding the **SharePoint Administrator** or **Global Administrator**
role. This is the only part of the deployment that cannot be delegated to Trinergy — everything
afterwards is done by Clarence using the site permissions he already has.

---

## What this does, and what it does not do

**It does:** create one empty document library called *Apps for SharePoint* at
`https://sdguthrie.sharepoint.com/sites/CRS/AppCatalog`.

**It does not:**

- grant anyone any new permissions
- touch the tenant App Catalog at `/sites/AppCatalog`
- affect any other site collection in the tenant
- require you to receive, handle, open, or deploy the solution package
- cause downtime, or change any existing content, list, or permission

It is fully reversible with a single command (see [Undoing it](#undoing-it)).

**Why we are asking for this rather than the tenant App Catalog:** deploying via the tenant catalog
would require placing a copy of the solution package in the tenant-wide catalog, and would require
granting Trinergy tenant-wide upload rights. A site collection app catalog is the narrowest possible
option — it scopes everything to the single CRS site.

---

## Before you start

You do **not** need to install PowerShell. Windows PowerShell 5.1 is already on every Windows 10,
Windows 11, and Windows Server 2016 or later machine.

You will need internet access and your admin sign-in credentials.

---

## Step 1 — Open Windows PowerShell (not PowerShell 7)

1. Press the **Windows key**.
2. Type `Windows PowerShell`.
3. Click **Windows PowerShell** in the results.

> **Important:** do not use **PowerShell 7**, **pwsh**, or **Windows Terminal** if it opens
> PowerShell 7. The SharePoint module only works in Windows PowerShell 5.1. This is the single most
> common reason these commands fail.

Confirm you are in the right place by running:

```powershell
$PSVersionTable.PSVersion
```

You should see `Major 5` and `Minor 1`. If you see `Major 7`, close the window and go back to step 1.

---

## Step 2 — Check whether the SharePoint module is already installed

Most tenant administrators already have it. Run:

```powershell
Get-Module -ListAvailable Microsoft.Online.SharePoint.PowerShell
```

- **If it returns a version number** — skip step 3 entirely and go to step 4.
- **If it returns nothing** — do step 3.

---

## Step 3 — Install the SharePoint module (only if step 2 found nothing)

```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser -Force
```

This installs for your user account only and does not require administrator elevation.

You may see one or two prompts the first time:

- **"NuGet provider is required to continue"** → type `Y` and press Enter.
- **"You are installing the modules from an untrusted repository"** → type `A` (Yes to All) and
  press Enter. This is the standard Microsoft PowerShell Gallery warning.

It takes roughly 30 seconds.

---

## Step 4 — Sign in to the SharePoint admin endpoint

```powershell
Connect-SPOService -Url https://sdguthrie-admin.sharepoint.com
```

A sign-in window will appear. Sign in with your SharePoint Administrator or Global Administrator
account. Multi-factor authentication is supported and will prompt as normal.

When it succeeds you are returned to the prompt with **no message**. That is correct.

> Note the URL contains **`sdguthrie-admin`**, not `sdguthrie`. Using the plain site URL here is the
> second most common mistake.

---

## Step 5 — Create the app catalog

```powershell
Add-SPOSiteCollectionAppCatalog -Site https://sdguthrie.sharepoint.com/sites/CRS
```

> Note this URL is the **plain site URL** — `sdguthrie.sharepoint.com`, with no `-admin`. The two
> commands deliberately use different addresses.

**Success produces no output at all.** The command is silent when it works. If you see no red error
text and get your prompt back, it worked.

---

## Step 6 — Verify

Open this address in a browser:

```
https://sdguthrie.sharepoint.com/sites/CRS/AppCatalog/Forms/AllItems.aspx
```

You should see an empty document library titled **Apps for SharePoint**. Empty is correct — the
package has not been uploaded yet.

---

## Step 7 — Let Clarence know

That completes your part. Clarence will upload and deploy the solution himself, and will do all
future updates the same way. You will not need to run anything again.

---

## If something goes wrong

| Message you see | What it means | Fix |
|---|---|---|
| `The term 'Add-SPOSiteCollectionAppCatalog' is not recognized` | Module not installed, or you are in PowerShell 7 | Check step 1 (`$PSVersionTable`), then step 3 |
| `The term 'Connect-SPOService' is not recognized` | Module not installed | Do step 3 |
| `Current user is not a tenant administrator` | The signed-in account lacks the required role | Use a SharePoint Administrator or Global Administrator account |
| `Cannot contact web site ... or the web site does not support SharePoint Online credentials` | The plain site URL was used in step 4 instead of the `-admin` URL | Re-run step 4 with `sdguthrie-admin.sharepoint.com` |
| `File skipped ... cannot be loaded because running scripts is disabled` | Execution policy is Restricted (common on Windows Server) | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned` then retry. `-Scope Process` applies only to this window and reverts when you close it |
| `Unable to resolve package source` / `Install-Module` fails to reach the gallery | TLS 1.2 not enabled (older Windows Server) | Run `[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12` then retry step 3 |
| An error saying the app catalog already exists | Someone has already done this | Nothing to do — go to step 6 and verify |

---

## Undoing it

```powershell
Remove-SPOSiteCollectionAppCatalog -Site https://sdguthrie.sharepoint.com/sites/CRS
```

---

## Questions your security review may raise

**"Does this grant Trinergy tenant-wide access?"**
No. It creates one library on one site. Clarence's access remains scoped to the CRS site collection,
which he already has through the *Guthrie Central Repository System Owners* group. No new permission
is granted by this command.

**"Site Collection Administrators on the App Catalog can approve installations for apps with
Tenant-Scoped rights" — this warning appears on the site. Is that a risk?**
Not for this solution. The package requests **no** API permissions at all
(`webApiPermissionRequests` is empty), so the deployment dialog presents nothing to approve. This
can be confirmed visually at deployment time.

**"Isn't the SharePoint Add-in model being retired?"**
That retirement notice concerns the legacy **SharePoint Add-in** model. This solution is built on
the **SharePoint Framework (SPFx)**, which is Microsoft's current and supported model and is not
affected.

**"Can this be done through the admin center UI instead?"**
No. Microsoft provides no UI for enabling a site collection app catalog; PowerShell is the only
supported method.

**"Can it be run from Azure Cloud Shell to avoid installing anything?"**
No. Cloud Shell runs PowerShell 7 on Linux, and the SharePoint Online Management Shell module
requires Windows PowerShell. It must be run from a Windows machine.
