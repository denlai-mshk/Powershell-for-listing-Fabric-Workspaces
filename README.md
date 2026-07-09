# ScanFabricWS — Fabric Item Scanner

## What these scripts do

This repo inventories **every Power BI / Fabric workspace in your tenant** (using read-only **admin** APIs) and flags those containing **Fabric items** that block cross-region workspace migration.

Context: this repo supports migrating Power BI workspaces from a **P license (Southeast Asia)** to an **F license / Fabric capacity (East Asia)**. Microsoft does **not** support cross-region migration when a workspace contains non–Power BI Fabric items ([Microsoft Learn reference](https://learn.microsoft.com/en-us/fabric/admin/portal-workspace-capacity-reassignment)). These scripts inventory those blockers *before* you attempt the migration.

The work is split into **two self-contained scripts**, so the fast/safe part is decoupled from the slow/rate-limited part. Each script is standalone — the shared auth/REST/pagination helpers are duplicated inside both, so there is no common file to install or keep on the path:

| File | Role |
|------|------|
| `listworkspace.ps1` | **Step 1 — fast.** Lists every workspace → `workspaces.csv`. No item calls. |
| `listfabricitems.ps1` | **Step 2 — slow, paced, resumable.** Scans items for capacity-backed workspaces → `fullset.csv` + `fabricitem.csv`. |

### Why two scripts?

The `/v1/admin/items` endpoint is rate-limited (~**200 requests/hour**), while `/v1/admin/workspaces` is not and uses a **separate** quota. Splitting the work lets you:

- Run **Step 1** in seconds to get the full inventory and immediately mark every no-capacity workspace `PurePowerBI` — no throttling risk.
- Run **Step 2** slowly and **resumably** to resolve only the capacity-backed workspaces, staying under the item rate limit.

> A legacy all-in-one script, `scanfabricitems.ps1`, still exists and does both steps in a single pass (no resume). It's convenient for small tenants but can hit `/admin/items` throttling on large ones — prefer the two-step workflow below.

## Step 1 — `listworkspace.ps1` (fast inventory)

Lists every workspace via the Fabric admin endpoint (falling back to the Power BI admin endpoint) and writes `workspaces.csv`. Because it makes **no item calls**, it's fast and cannot trip `/admin/items` throttling.

It assigns a preliminary `Classification` based purely on capacity:

| Classification | Meaning | Next step |
|----------------|---------|-----------|
| `PurePowerBI` | Workspace has **no** capacity → cannot host Fabric items → **final, migration-safe** | none |
| `Pending` | Workspace **is** capacity-backed → *might* contain Fabric items | resolved by Step 2 |

`workspaces.csv` columns: `WorkspaceId`, `WorkspaceName`, `Type`, `State`, `Classification`, `CapacityId`.

```powershell
Connect-AzAccount
./listworkspace.ps1
# optional: choose an output folder
./listworkspace.ps1 -OutputFolder "C:\CodeProject\ScanFabricWS\out"
```

## Step 2 — `listfabricitems.ps1` (paced, resumable item scan)

Reads `workspaces.csv`, takes only the capacity-backed (`Pending`) workspaces, and fetches their items via `/v1/admin/items?capacityId=...`, **iterating distinct capacities** (dozens) rather than workspaces (thousands). It resolves each `Pending` workspace to `PurePowerBI` or `HasFabricItems`, then writes `fullset.csv` + `fabricitem.csv`.

Because `/admin/items` is throttled, this script is built to be gentle and durable rather than fast:

- **Paced:** sleeps `-IntervalSeconds` (default **20s**, ≈180 req/hr) before **every** `/admin/items` request — including each continuation-token page, which is where throttling actually bites. There is deliberately **no tight retry loop**; the interval is the primary 429 defence.
- **Fault-isolated:** each capacity is wrapped in `try/catch`. A throttled or failed capacity is logged and left unfinished — the run continues instead of aborting.
- **Resumable:** finished capacities are committed to checkpoint files, so re-running **skips completed capacities** and retries only the unfinished ones. You can `Ctrl-C` and continue later. Use `-Fresh` to discard checkpoints and start over.

```powershell
Connect-AzAccount
./listfabricitems.ps1
# tune the pace (slower = safer if the quota is shared):
./listfabricitems.ps1 -IntervalSeconds 30
# start over, ignoring saved progress:
./listfabricitems.ps1 -Fresh
```

### Parameters

| Parameter | Script | Default | Purpose |
|-----------|--------|---------|---------|
| `-OutputFolder` | both | script directory | Where CSVs and checkpoints are read/written. |
| `-IntervalSeconds` | `listfabricitems.ps1` | `20` | Seconds to wait before every `/admin/items` request. |
| `-Fresh` | `listfabricitems.ps1` | off | Discard checkpoints and re-scan every capacity. |

### Files produced

| File | By | Contents |
|------|----|----------|
| `workspaces.csv` | Step 1 | Full inventory; `Classification` is `PurePowerBI` or `Pending`. |
| `fullset.csv` | Step 2 | All workspaces with final `Classification`, `FabricItemCount`, `fabricitems`. Unscanned capacity-backed rows stay `Pending`. |
| `fabricitem.csv` | Step 2 | Only `HasFabricItems` workspaces — your migration blockers. |
| `items_partial.csv` | Step 2 | **Checkpoint:** raw items collected so far (`CapacityId`, `WorkspaceId`, `ItemId`, `ItemType`). |
| `capacities_done.csv` | Step 2 | **Checkpoint:** capacities fully scanned (the resume marker). |

> The `items_partial.csv` and `capacities_done.csv` files are checkpoints that make Step 2 resumable. Delete them (or run with `-Fresh`) to force a clean re-scan.

### `Classification` values

| Classification | Meaning | Migration |
|----------------|---------|-----------|
| `PurePowerBI` | Only Power BI items, or empty | ✅ Safe |
| `HasFabricItems` | Contains one or more Fabric items | ❌ Blocked |
| `Pending` | Capacity-backed but not yet item-scanned (Step 2 not run, or interrupted for this capacity) | ⏳ Unknown — re-run Step 2 |

### Columns in `fullset.csv`

| Column | Description |
|--------|-------------|
| `WorkspaceId` | Workspace GUID. |
| `WorkspaceName` | Workspace display name. |
| `Type` | Workspace type (e.g. `Workspace`, `PersonalGroup`). |
| `State` | Workspace state (`Active`, `Deleted`, `Orphaned`, `Removing`). |
| `Classification` | `PurePowerBI`, `HasFabricItems`, or `Pending`. |
| `CapacityId` | Backing capacity GUID; **empty means the workspace is on no capacity**, so it cannot host Fabric items. |
| `FabricItemCount` | Number of Fabric item instances found (`0` for pure Power BI; blank while `Pending`). |
| `fabricitems` | Semicolon-separated **distinct** Fabric item types found (empty = migration-safe). |

## How items are classified — the Power BI "allowlist"

Rather than maintaining a list of Fabric item types (which Microsoft keeps growing), the scripts use a small, stable **allowlist of Power BI item types**. A workspace is `PurePowerBI` only if **every** item it contains is one of these — anything else counts as a Fabric item (a migration blocker):

- Report
- Dashboard
- SemanticModel
- PaginatedReport
- Datamart

This is **future-proof**: any new or unknown Fabric item type Microsoft adds (e.g. `DigitalTwinBuilder`, `GraphModel`) is automatically treated as a blocker without a code change. Item types are per the official [Fabric Core REST API `ItemType` enumeration](https://learn.microsoft.com/en-us/rest/api/fabric/core/items/list-items).

## How the scan works — capacity-aware, admin-only

The scripts are **admin-only** and never use per-workspace (user-scoped) endpoints, so they do not need membership on each workspace. Classification is capacity-aware for performance at tenant scale (tens of thousands of workspaces):

1. **List workspaces** (Step 1) via the Fabric admin endpoint (`/v1/admin/workspaces`), falling back to the Power BI admin endpoint if that is unavailable.
2. **Capacity fast-path:** Fabric items can only exist on a Fabric-capable capacity, so any workspace with **no `capacityId`** is classified `PurePowerBI` immediately — with no item lookup at all.
3. **Targeted item scan** (Step 2): for the capacity-backed minority, items are fetched **per distinct capacity** via `/v1/admin/items?capacityId=...` — not a tenant-wide sweep. The payload therefore scales with items-on-capacities, not total tenant items, and the per-capacity loop keeps request counts low against the throttled endpoint.

> **Trade-off:** the capacity fast-path assumes a workspace with no `capacityId` cannot hold Fabric items — true except in the rare case where a capacity was deleted/unassigned while dormant Fabric items lingered. Such a workspace could report no `capacityId` yet still contain items. For a migration pre-check this is an acceptable, documented edge case.

## Prerequisites

### 1. PowerShell module installation

- **PowerShell 5.1+** or **PowerShell 7+**.
- **`Az.Accounts` module** — provides `Connect-AzAccount` and `Get-AzAccessToken`, used for authentication and token acquisition. The scripts verify it is installed and stop with an install hint if it is missing.

  ```powershell
  # Check whether it is already installed
  Get-Module -ListAvailable -Name Az.Accounts

  # Install (current user, no admin rights needed)
  Install-Module Az.Accounts -Scope CurrentUser

  # Or install the full Az bundle if you prefer
  Install-Module Az -Scope CurrentUser

  # Optional: update to the latest version
  Update-Module Az.Accounts
  ```

  > `Az.Accounts` **v2.x, v3.x, v4.x, or v5.x** are all supported. The scripts automatically handle the `SecureString` token returned by v5.x, so no code change is needed after upgrading.

  If `Install-Module` fails with an execution-policy or repository-trust error:

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
  ```

### 2. Permissions required — administrator role

This tool is **admin-only**. The signed-in account must hold one of:

- **Fabric Administrator**, or
- **Power BI Administrator**, or
- **Global Administrator** (a superset of the above).

Specifically, the account must be able to call the read-only admin APIs `/v1/admin/workspaces` and `/v1/admin/items`, and obtain a token for the Fabric API resource `https://api.fabric.microsoft.com`.

> **Why admin?** The scripts list every workspace in the tenant and enumerate their items through admin APIs, so they do **not** require membership on each workspace. A non-admin account will fail against the admin endpoints and the script will stop.
>
> **Service principals:** the scripts authenticate as an interactive **user** by default, so the *"Service principals can access read-only admin APIs"* tenant setting is **not** required. It is only relevant if you adapt the scripts to sign in as a service principal.

## How to run — end to end

```powershell
# 1. Sign in (once per session). Device code flow: open the shown URL and enter the code.
Connect-AzAccount -UseDeviceAuthentication

# 2. Step 1 — fast inventory (seconds)
./listworkspace.ps1

# 3. Step 2 — paced item scan (minutes to hours; resumable). Re-run until nothing is 'Pending'.
./listfabricitems.ps1
```

Both scripts write to the output folder (default: the script directory). If you are not already signed in, the scripts prompt with device code authentication automatically.

> **Resuming:** if Step 2 is interrupted or a capacity is throttled, just run `./listfabricitems.ps1` again — it skips capacities already recorded in `capacities_done.csv` and retries only the rest.

## How to use the outcome

1. Open **`fabricitem.csv`** — this is your prioritized worklist. Every workspace listed here is classified `HasFabricItems` and **cannot be migrated cross-region as-is**.
2. For each blocking workspace, review the `fabricitems` column to see which item types are present, then either:
   - Remove / migrate those Fabric items out of the workspace, or
   - Recreate them in the destination region after migration.
3. Use **`fullset.csv`** as the master inventory: rows with `Classification = PurePowerBI` are safe to migrate immediately. Rows still marked **`Pending`** have not been item-scanned yet — re-run Step 2 to resolve them before trusting the inventory as complete.
4. Re-run Step 2 after remediation to confirm `fabricitem.csv` is empty (or shrinking) before starting the P→F migration.
