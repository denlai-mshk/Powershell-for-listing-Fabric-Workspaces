# ScanFabricWS — Fabric Item Scanner

## What this script does

`scanfabricitems.ps1` scans **every Power BI / Fabric workspace** your identity can access and detects **Fabric items** that block cross-region workspace migration.

Context: this repo supports migrating Power BI workspaces from a **P license (Southeast Asia)** to an **F license / Fabric capacity (East Asia)**. Microsoft does **not** support cross-region migration when a workspace contains non–Power BI Fabric items ([Microsoft Learn reference](https://learn.microsoft.com/en-us/fabric/admin/portal-workspace-capacity-reassignment)). This script inventories those blockers *before* you attempt the migration.

The script produces two files:

| File | Contents |
|------|----------|
| `fullset.csv` | **All** workspaces, with a `fabricitems` column listing any Fabric item types found (semicolon-separated; empty = migration-safe). |
| `fabricitem.csv` | **Only** workspaces that contain Fabric items — i.e. your migration blockers. |

**Item types treated as Fabric items** (migration blockers) are validated against the official [Fabric Core REST API `ItemType` enumeration](https://learn.microsoft.com/en-us/rest/api/fabric/core/items/list-items):

- Lakehouse
- Warehouse
- WarehouseSnapshot
- Notebook
- DataPipeline
- CopyJob
- Dataflow
- Eventhouse
- Eventstream
- KQLDatabase
- KQLQueryset
- KQLDashboard
- MirroredDatabase
- MirroredWarehouse
- MirroredAzureDatabricksCatalog
- SQLDatabase
- SQLEndpoint
- Environment
- MLModel
- MLExperiment
- SparkJobDefinition
- Reflex
- GraphQLApi
- MountedDataFactory
- VariableLibrary
- ApacheAirflowJob
- UserDataFunction
- DataAgent

Power BI artifacts are **ignored** (not migration blockers):

- Report
- SemanticModel
- Dashboard
- PaginatedReport
- Datamart

## Prerequisites

### 1. PowerShell module installation

- **PowerShell 5.1+** or **PowerShell 7+**.
- **`Az.Accounts` module** — provides `Connect-AzAccount` and `Get-AzAccessToken`, used for authentication and token acquisition. The script verifies it is installed and stops with an install hint if it is missing.

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

  > `Az.Accounts` **v2.x, v3.x, v4.x, or v5.x** are all supported. The script automatically handles the `SecureString` token returned by v5.x, so no code change is needed after upgrading.

  If `Install-Module` fails with an execution-policy or repository-trust error:

  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
  ```

### 2. Permissions required

1. **An Entra ID (Azure AD) account** signed in via `Connect-AzAccount`.
2. **Fabric / Power BI access** so the account can enumerate workspaces and their items:
   - Membership (Admin / Member / Contributor) on the workspaces you want scanned, **or**
   - **Fabric Administrator** / **Power BI Administrator** role for full-tenant visibility.
3. The identity must be able to obtain a token for the Fabric API resource `https://api.fabric.microsoft.com`.

> Note: workspaces where the account has no role will not appear in the results. Use an admin account for a complete tenant scan.

## How to run

```powershell
# 1. Sign in (once per session). Device code flow: open the shown URL and enter the code.
Connect-AzAccount -UseDeviceAuthentication

# 2. Run the scanner
./scanfabricitems.ps1

# Optional: choose an output folder
./scanfabricitems.ps1 -OutputFolder "C:\CodeProject\ScanFabricWS\out"
```

Output `fullset.csv` and `fabricitem.csv` are written to the output folder (defaults to the script directory).

> If you are not already signed in, the script itself will prompt with device code authentication automatically.

## How to use the outcome

1. Open **`fabricitem.csv`** — this is your prioritized worklist. Every workspace listed here has Fabric items and **cannot be migrated cross-region as-is**.
2. For each blocking workspace, review the `fabricitems` column to see which item types are present, then either:
   - Remove / migrate those Fabric items out of the workspace, or
   - Recreate them in the destination region after migration.
3. Use **`fullset.csv`** as the master inventory: rows with an **empty `fabricitems`** column are safe to migrate immediately.
4. Re-run the script after remediation to confirm `fabricitem.csv` is empty (or shrinking) before starting the P→F migration.
