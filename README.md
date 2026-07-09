# ScanFabricWS — Fabric Item Scanner

Find out which Power BI / Fabric workspaces are **safe to migrate** across regions, and which ones are **blocked** because they contain Fabric items.

**Why this exists:** moving workspaces from a **P license (Southeast Asia)** to an **F license / Fabric capacity (East Asia)** is **not** allowed when a workspace holds non–Power BI Fabric items ([Microsoft docs](https://learn.microsoft.com/en-us/fabric/admin/portal-workspace-capacity-reassignment)). These scripts find those blockers *before* you migrate.

## How it works — two scripts, run in order

### Step 1 — `listworkspace.ps1`  (fast, run this first)

A quick scan of workspace **metadata only**. It lists every workspace in your tenant into **`workspaces.csv`** and gives each one a first verdict:

- **`PurePowerBI`** — the workspace has **no capacity**, so it *cannot* contain Fabric items. It's safe to migrate. ✅ *(final answer)*
- **`Pending`** — the workspace is on a capacity, so it *might* contain Fabric items. Step 2 will check it.

This runs in **seconds** and won't hit any rate limit.

### Step 2 — `listfabricitems.ps1`  (slow, optional)

Run this **only if** you have the time and want to know exactly which Fabric items must be removed before migrating. It looks up the **`Pending`** workspaces, scans how many Fabric items each one has, and writes the results to **`fullset.csv`** and **`fabricitem.csv`**.

It runs **very slowly on purpose** — it sleeps about **20 seconds between every API call** so the Fabric admin API doesn't reject it for hitting the throttle limit (~200 calls/hour). The upside: it's **resumable** — if it's interrupted or throttled, just run it again and it continues where it left off.

```powershell
# Sign in once (a browser device-code prompt appears)
Connect-AzAccount -UseDeviceAuthentication

# Step 1 — fast
./listworkspace.ps1

# Step 2 — slow & optional; re-run until nothing is 'Pending'
./listfabricitems.ps1
```

**Handy options for Step 2:**

- `-IntervalSeconds 30` — wait longer between calls (safer if other people share the admin quota).
- `-Fresh` — ignore saved progress and scan everything again.
- `-OutputFolder "C:\path"` — read/write the CSVs somewhere else (both scripts accept this).

## What you get

| File | From | What it contains |
|------|------|------------------|
| `workspaces.csv` | Step 1 | Every workspace, marked `PurePowerBI` or `Pending`. |
| `fullset.csv` | Step 2 | Every workspace with its final verdict and Fabric-item details. |
| `fabricitem.csv` | Step 2 | **Only the blocked workspaces** (`HasFabricItems`) — your clean-up list. |
| `items_partial.csv`, `capacities_done.csv` | Step 2 | Progress checkpoints so Step 2 can resume. Safe to delete (or use `-Fresh`). |

### What the `Classification` column means

| Value | Meaning | Migration |
|-------|---------|-----------|
| `PurePowerBI` | Only Power BI content (or empty) | ✅ Safe to migrate |
| `HasFabricItems` | Contains one or more Fabric items | ❌ Blocked — remove them first |
| `Pending` | On a capacity but not scanned yet (run Step 2) | ⏳ Unknown |

`fullset.csv` columns: `WorkspaceId`, `WorkspaceName`, `Type`, `State`, `Classification`, `CapacityId`, `FabricItemCount`, `fabricitems`.

## What counts as a "Fabric item"?

Anything that is **not** one of these five Power BI types is treated as a Fabric item (a migration blocker):

> `Report` · `Dashboard` · `SemanticModel` · `PaginatedReport` · `Datamart`

Using this short "allowlist" is future-proof: any new Fabric item type Microsoft adds later is flagged automatically, with no code change. (Types come from the official [Fabric `ItemType` list](https://learn.microsoft.com/en-us/rest/api/fabric/core/items/list-items).)

> **One small caveat:** a workspace with no capacity is assumed to have no Fabric items. That's true except in the rare case where a capacity was removed while old Fabric items were left behind dormant. For a pre-migration check this is a safe, accepted assumption.

## Before you start

**1. Install the `Az.Accounts` PowerShell module** (needs PowerShell 5.1+ or 7+):

```powershell
Install-Module Az.Accounts -Scope CurrentUser
```

Any version **2.x–5.x** works. If the install is blocked by policy, run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
```

**2. Sign in as an administrator.** You must be a **Fabric Administrator**, **Power BI Administrator**, or **Global Administrator**. The scripts use read-only **admin** APIs to see every workspace, so you do **not** need to be a member of each one — but a non-admin account won't work.

> Signing in as a normal **user** (the device-code prompt) is all you need. You do **not** have to enable the *"Service principals can access read-only admin APIs"* tenant setting — that only matters if you change the scripts to log in as a service principal instead of a user.

## After the scan — what to do

1. Open **`fabricitem.csv`** — every workspace here is **blocked**. The `fabricitems` column shows which item types to remove (or recreate in the destination region after migrating).
2. Anything marked **`PurePowerBI`** in `fullset.csv` is safe to migrate now.
3. Anything still **`Pending`** hasn't been checked yet — run Step 2 to resolve it.
4. After cleaning up a workspace, re-run Step 2 until `fabricitem.csv` is empty.

---

> **Note:** an older all-in-one script, `scanfabricitems.ps1`, still does both steps in a single run (no pause/resume). It's fine for small tenants but can hit throttling on large ones, so the two-step flow above is recommended.
