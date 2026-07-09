<#
.SYNOPSIS
    Part 2 of 2: slow, paced, resumable Fabric-item scan for capacity-backed
    workspaces (resolves the 'Pending' rows from listworkspace.ps1).

.DESCRIPTION
    Reads workspaces.csv (produced by listworkspace.ps1), takes only the
    capacity-backed workspaces, and fetches their items via the Fabric admin
    endpoint GET /v1/admin/items?capacityId=... -- iterating DISTINCT CAPACITIES
    (dozens), not workspaces (thousands).

    Because /admin/items is rate-limited (~200 requests/hour), this script is
    built to be gentle and durable rather than fast:

      * PACING: it sleeps -IntervalSeconds (default 20s) before every
        /admin/items request -- including each continuation-token page, which is
        where throttling actually bites. There is deliberately NO tight retry
        loop; the interval is the primary 429 defence.

      * FAULT ISOLATION: each capacity is wrapped in try/catch. If one capacity
        fails (e.g. a shared-quota 429), it is logged and left "not done"; the
        script moves on instead of aborting the whole run.

      * RESUMABLE: completed capacities are committed to capacities_done.csv and
        their items appended to items_partial.csv. Re-running SKIPS completed
        capacities, so you can Ctrl-C and resume later (or just re-run to retry
        the failed ones). Use -Fresh to discard checkpoints and start over.

    Classification uses the Power BI allowlist (the $PowerBIItemTypes list below):
    any item that is NOT a known Power BI type is a Fabric item / migration blocker.

    Outputs (rebuilt every run from current knowledge):
      - fullset.csv    : all workspaces. Capacity-backed workspaces whose capacity
                         has not been scanned yet stay 'Pending'.
      - fabricitem.csv : only workspaces classified 'HasFabricItems'.

.PARAMETER OutputFolder
    Folder containing workspaces.csv and where outputs/checkpoints are written.
    Defaults to the script's own directory.

.PARAMETER IntervalSeconds
    Seconds to wait before every /admin/items request. Default 20 (~180/hour,
    just under the documented 200/hour limit). Increase if you share the quota.

.PARAMETER Fresh
    Discard existing checkpoints (items_partial.csv / capacities_done.csv) and
    scan every capacity from scratch.

.EXAMPLE
    Connect-AzAccount
    ./listfabricitems.ps1

.EXAMPLE
    ./listfabricitems.ps1 -IntervalSeconds 30

.NOTES
    Admin-only, read-only APIs. Requires Az.Accounts and a signed-in account with
    Fabric Administrator or Power BI Administrator rights. See README.md.
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = $PSScriptRoot,
    [ValidateRange(0, 3600)][int]$IntervalSeconds = 20,
    [switch]$Fresh
)

$ErrorActionPreference = 'Stop'

# ===========================================================================
# Shared helpers (INLINED, self-contained). This block is intentionally
# DUPLICATED verbatim in listworkspace.ps1 and listfabricitems.ps1 so each
# script runs standalone with no external dependency. If you change it here,
# make the same change in the other script.
# ===========================================================================

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$FabricResourceUrl  = 'https://api.fabric.microsoft.com'
$FabricApiBase      = 'https://api.fabric.microsoft.com/v1'
$PowerBIResourceUrl = 'https://analysis.windows.net/powerbi/api'
$PowerBIApiBase     = 'https://api.powerbi.com/v1.0/myorg'

# Known Power BI item types (the "allowlist"). A workspace is considered
# "pure Power BI" (migration-safe) only if EVERY item it contains is in this
# list -- or it has no items at all. Any other item type, including new or
# unknown Fabric item types Microsoft may add over time, is treated as a
# migration blocker. This allowlist is future-proof: the Power BI item set is
# small and stable, while the Fabric item set keeps growing.
# Item types per the official Fabric Core REST API ItemType enumeration:
# https://learn.microsoft.com/en-us/rest/api/fabric/core/items/list-items
$PowerBIItemTypes = @(
    'Report',
    'Dashboard',
    'SemanticModel',
    'PaginatedReport',
    'Datamart'
)

# ---------------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------------
function Assert-RequiredModules {
    # Ensure the Az.Accounts module (which provides Connect-AzAccount /
    # Get-AzAccessToken) is available before we try to authenticate.
    $module = Get-Module -ListAvailable -Name Az.Accounts |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $module) {
        throw "Required module 'Az.Accounts' is not installed. Install it with: Install-Module Az.Accounts -Scope CurrentUser"
    }

    if (-not (Get-Module -Name Az.Accounts)) {
        Import-Module Az.Accounts -ErrorAction Stop
    }

    Write-Host ("Using Az.Accounts version {0}." -f $module.Version) -ForegroundColor DarkGray
}

function ConvertFrom-TokenValue {
    param([Parameter(Mandatory)]$Token)

    # Az.Accounts 5.x returns the access token as a SecureString by default,
    # while older versions return a plain String. Normalize to plain text.
    if ($Token -is [System.Security.SecureString]) {
        return [System.Net.NetworkCredential]::new('', $Token).Password
    }
    return [string]$Token
}

function Get-ApiHeaders {
    param(
        [Parameter(Mandatory)][string]$ResourceUrl,
        [string]$ResourceLabel = $ResourceUrl
    )

    Assert-RequiredModules

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "No Azure context found. Signing in with device code..." -ForegroundColor Yellow
        Write-Host "Open the displayed URL in a browser and enter the code to authenticate." -ForegroundColor Yellow
        Connect-AzAccount -UseDeviceAuthentication | Out-Null
    }

    # Request the token as a SecureString when the parameter is supported
    # (Az.Accounts 5.x); fall back gracefully on older versions.
    $tokenParams = @{ ResourceUrl = $ResourceUrl }
    if ((Get-Command Get-AzAccessToken).Parameters.ContainsKey('AsSecureString')) {
        $tokenParams['AsSecureString'] = $true
    }

    $tokenObject = Get-AzAccessToken @tokenParams
    $token = ConvertFrom-TokenValue -Token $tokenObject.Token
    if (-not $token) {
        throw "Failed to acquire an access token for $ResourceLabel ($ResourceUrl)."
    }

    return @{
        Authorization  = "Bearer $token"
        'Content-Type' = 'application/json'
    }
}

function Get-HttpStatusCodeFromError {
    param([Parameter(Mandatory)]$ErrorRecord)

    if ($ErrorRecord.Exception.Response) {
        try {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
        catch {
            return $null
        }
    }

    return $null
}

function Test-CanFallbackFromFabricAdminError {
    param([Parameter(Mandatory)]$ErrorRecord)

    $status = Get-HttpStatusCodeFromError -ErrorRecord $ErrorRecord
    if ($status -in 401, 403, 404) {
        return $true
    }

    $message = [string]$ErrorRecord.Exception.Message
    return ($message -match 'forbidden|unauthorized|not found|notfound')
}

# ---------------------------------------------------------------------------
# REST helpers with pagination and retry (429 / 5xx)
# ---------------------------------------------------------------------------
function Invoke-RestGetWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$MaxRetries = 5
    )

    $attempt = 0

    while ($true) {
        try {
            return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get
        }
        catch {
            $attempt++
            $status = Get-HttpStatusCodeFromError -ErrorRecord $_

            if (($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $attempt -le $MaxRetries) {
                $delay = [Math]::Min(60, [Math]::Pow(2, $attempt))
                Write-Warning "Request to $Uri failed with status $status. Retry $attempt/$MaxRetries in $delay s..."
                Start-Sleep -Seconds $delay
                continue
            }

            throw
        }
    }
}

function Get-FabricContinuationUri {
    param(
        [Parameter(Mandatory)][string]$CurrentUri,
        [Parameter(Mandatory)]$Response
    )

    if ($Response.PSObject.Properties.Name -contains 'continuationUri' -and $Response.continuationUri) {
        return [string]$Response.continuationUri
    }

    if ($Response.PSObject.Properties.Name -contains 'continuationToken' -and $Response.continuationToken) {
        $baseUri = $CurrentUri -replace '([?&])continuationToken=[^&]*', ''
        $baseUri = $baseUri -replace '\?&', '?'
        $baseUri = $baseUri -replace '[?&]$', ''

        $separator = if ($baseUri.Contains('?')) { '&' } else { '?' }
        $encodedToken = [System.Uri]::EscapeDataString([string]$Response.continuationToken)
        return "$baseUri${separator}continuationToken=$encodedToken"
    }

    return $null
}

function Get-CollectionFromResponse {
    param(
        [Parameter(Mandatory)]$Response,
        [string[]]$CollectionPropertyNames = @('value', 'workspaces', 'itemEntities', 'items', 'data')
    )

    if ($null -eq $Response) {
        return @()
    }

    if ($Response -is [System.Array]) {
        return @($Response)
    }

    foreach ($propertyName in $CollectionPropertyNames) {
        if ($Response.PSObject.Properties.Name -contains $propertyName) {
            $propertyValue = $Response.$propertyName
            if ($null -eq $propertyValue) {
                return @()
            }

            return @($propertyValue)
        }
    }

    # Allow single-object responses only when they look like an entity.
    if (($Response.PSObject.Properties.Name -contains 'id') -or ($Response.PSObject.Properties.Name -contains 'Id')) {
        return @($Response)
    }

    return @()
}

function Invoke-FabricGet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$MaxRetries = 5
    )

    $results = @()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-RestGetWithRetry -Uri $nextUri -Headers $Headers -MaxRetries $MaxRetries

        $pageItems = Get-CollectionFromResponse -Response $response -CollectionPropertyNames @('value', 'workspaces', 'itemEntities', 'items', 'data')
        if ($pageItems.Count -gt 0) {
            $results += $pageItems
        }

        # Fabric pagination: continuationUri / continuationToken
        $nextUri = Get-FabricContinuationUri -CurrentUri $nextUri -Response $response
    }

    return $results
}

function Invoke-PowerBIAdminGetWorkspaces {
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$MaxRetries = 5,
        [int]$PageSize = 5000
    )

    $results = @()
    $skip = 0

    while ($true) {
        $uri = "$PowerBIApiBase/admin/workspaces?`$top=$PageSize&`$skip=$skip"
        $response = Invoke-RestGetWithRetry -Uri $uri -Headers $Headers -MaxRetries $MaxRetries

        $currentPage = Get-CollectionFromResponse -Response $response -CollectionPropertyNames @('value', 'workspaces', 'data')

        $results += $currentPage

        if ($currentPage.Count -lt $PageSize) {
            break
        }

        $skip += $PageSize
    }

    return $results
}

function ConvertTo-NormalizedWorkspace {
    <#
    .SYNOPSIS
        Normalises a raw workspace record from either admin listing endpoint into
        a consistent object (id / name / type / state / capacityId).

    .DESCRIPTION
        The Fabric and Power BI admin endpoints return slightly different field
        shapes (displayName vs name, capacityId vs dedicatedCapacityId), so every
        field is resolved defensively. Returns $null for records with no id.
    #>
    param([Parameter(Mandatory)]$Workspace)

    $ws = $Workspace

    $wsId = $ws.id
    if (-not $wsId -and ($ws.PSObject.Properties.Name -contains 'Id')) {
        $wsId = $ws.Id
    }
    if (-not $wsId) {
        return $null
    }

    $wsName = $ws.displayName
    if (-not $wsName -and ($ws.PSObject.Properties.Name -contains 'name')) {
        $wsName = $ws.name
    }
    if (-not $wsName) {
        $wsName = "(Unnamed workspace $wsId)"
    }

    $wsType = $ws.type
    if (-not $wsType -and ($ws.PSObject.Properties.Name -contains 'workspaceType')) {
        $wsType = $ws.workspaceType
    }

    $wsCapacityId = $ws.capacityId
    if (-not $wsCapacityId -and ($ws.PSObject.Properties.Name -contains 'dedicatedCapacityId')) {
        $wsCapacityId = $ws.dedicatedCapacityId
    }

    # Workspace state (Active / Deleted / Orphaned / Removing). Present on both
    # admin listings; guarded in case a source omits it.
    $wsState = $null
    if ($ws.PSObject.Properties.Name -contains 'state') {
        $wsState = $ws.state
    }

    return [pscustomobject]@{
        WorkspaceId   = [string]$wsId
        WorkspaceKey  = ([string]$wsId).ToLowerInvariant()
        WorkspaceName = $wsName
        Type          = $wsType
        State         = $wsState
        CapacityId    = $wsCapacityId
    }
}

# ---------------------------------------------------------------------------
# Paced, no-retry item fetch (local to this script).
# The interval is applied before EVERY /admin/items request across the whole
# run -- pages and capacities alike -- except the very first request.
# ---------------------------------------------------------------------------
$script:ItemsCallCount = 0

function Invoke-PacedItemsGet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$IntervalSeconds = 20
    )

    if ($script:ItemsCallCount -gt 0 -and $IntervalSeconds -gt 0) {
        Write-Host ("      pacing {0}s before next /admin/items call..." -f $IntervalSeconds) -ForegroundColor DarkGray
        Start-Sleep -Seconds $IntervalSeconds
    }
    $script:ItemsCallCount++

    # No retry loop by design: any failure (e.g. 429) propagates to the
    # per-capacity try/catch, which logs and leaves the capacity for a re-run.
    return Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get
}

function Get-FabricItemsForCapacityPaced {
    param(
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][string]$CapacityId,
        [int]$IntervalSeconds = 20
    )

    $items   = New-Object System.Collections.Generic.List[object]
    $encoded = [System.Uri]::EscapeDataString($CapacityId)
    $nextUri = "$FabricApiBase/admin/items?capacityId=$encoded"
    $page    = 0

    while ($nextUri) {
        $page++
        $response = Invoke-PacedItemsGet -Uri $nextUri -Headers $Headers -IntervalSeconds $IntervalSeconds

        $pageItems = Get-CollectionFromResponse -Response $response -CollectionPropertyNames @('itemEntities', 'value', 'items', 'data')
        foreach ($it in $pageItems) { $items.Add($it) }

        Write-Host ("      page {0}: {1} item(s)" -f $page, @($pageItems).Count) -ForegroundColor DarkGray
        $nextUri = Get-FabricContinuationUri -CurrentUri $nextUri -Response $response
    }

    return $items
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$workspacesPath   = Join-Path $OutputFolder 'workspaces.csv'
$itemsPartialPath = Join-Path $OutputFolder 'items_partial.csv'
$doneCapsPath     = Join-Path $OutputFolder 'capacities_done.csv'
$fullSetPath      = Join-Path $OutputFolder 'fullset.csv'
$fabricItemPath   = Join-Path $OutputFolder 'fabricitem.csv'

if (-not (Test-Path -Path $workspacesPath)) {
    throw "Input '$workspacesPath' not found. Run ./listworkspace.ps1 first to produce it."
}

if ($Fresh) {
    Remove-Item -Path $itemsPartialPath, $doneCapsPath -ErrorAction SilentlyContinue
    Write-Host "Fresh run: cleared previous checkpoint files (items_partial.csv, capacities_done.csv)." -ForegroundColor Yellow
}

# Load the workspace inventory and pick out the capacity-backed ones.
$wsRows = @(Import-Csv -Path $workspacesPath)
$capacityBacked = @($wsRows | Where-Object { $_.CapacityId })
$capacityIds    = @($capacityBacked | Select-Object -ExpandProperty CapacityId -Unique)

Write-Host ("Loaded {0} workspace(s): {1} capacity-backed across {2} distinct capacity(ies)." -f `
    $wsRows.Count, $capacityBacked.Count, $capacityIds.Count) -ForegroundColor Green

# Resume ledger: which capacities are already completed.
$doneCaps = New-Object 'System.Collections.Generic.HashSet[string]'
if (Test-Path -Path $doneCapsPath) {
    foreach ($r in Import-Csv -Path $doneCapsPath) {
        if ($r.CapacityId) { [void]$doneCaps.Add(([string]$r.CapacityId).ToLowerInvariant()) }
    }
    Write-Host ("Resume: {0} capacity(ies) already completed and will be skipped." -f $doneCaps.Count) -ForegroundColor Yellow
}

$remaining = @($capacityIds | Where-Object { -not $doneCaps.Contains(([string]$_).ToLowerInvariant()) })

if ($capacityIds.Count -eq 0) {
    Write-Host "No capacity-backed workspaces; nothing to scan." -ForegroundColor Green
}
elseif ($remaining.Count -eq 0) {
    Write-Host "All capacities already scanned; rebuilding outputs from checkpoints." -ForegroundColor Green
}
else {
    $estMinutes = [Math]::Ceiling(($remaining.Count * $IntervalSeconds) / 60.0)
    Write-Host ("Scanning {0} remaining capacity(ies) at {1}s/request (rough floor ~{2} min, more if capacities paginate)..." -f `
        $remaining.Count, $IntervalSeconds, $estMinutes) -ForegroundColor Cyan

    Write-Host "Authenticating to Fabric API..." -ForegroundColor Cyan
    $fabricHeaders = Get-ApiHeaders -ResourceUrl $FabricResourceUrl -ResourceLabel 'Fabric API'

    $capIndex = 0
    $failed   = New-Object System.Collections.Generic.List[string]

    foreach ($capacityId in $capacityIds) {
        $capIndex++
        $capKey = ([string]$capacityId).ToLowerInvariant()

        if ($doneCaps.Contains($capKey)) {
            Write-Host ("[Capacity {0}/{1}] {2} -- already done, skipping." -f $capIndex, $capacityIds.Count, $capacityId) -ForegroundColor DarkGray
            continue
        }

        Write-Host ("[Capacity {0}/{1}] Fetching items on {2}..." -f $capIndex, $capacityIds.Count, $capacityId) -ForegroundColor Cyan
        try {
            $items = Get-FabricItemsForCapacityPaced -Headers $fabricHeaders -CapacityId $capacityId -IntervalSeconds $IntervalSeconds

            # Flatten to (CapacityId, WorkspaceId, ItemId, ItemType) rows for the
            # checkpoint. ItemId is stored so the rebuild step can dedupe (see below).
            $itemRows = New-Object System.Collections.Generic.List[object]
            foreach ($it in $items) {
                $wsId = $it.workspaceId
                if (-not $wsId -and ($it.PSObject.Properties.Name -contains 'workspaceObjectId')) {
                    $wsId = $it.workspaceObjectId
                }
                if (-not $wsId -or -not $it.type) { continue }

                $itemId = $it.id
                if (-not $itemId -and ($it.PSObject.Properties.Name -contains 'objectId')) {
                    $itemId = $it.objectId
                }

                $itemRows.Add([pscustomobject]@{
                    CapacityId  = [string]$capacityId
                    WorkspaceId = [string]$wsId
                    ItemId      = [string]$itemId
                    ItemType    = [string]$it.type
                })
            }

            # Commit: append this capacity's items, THEN mark the capacity done.
            # (Order matters: the done-marker is the commit point for resume.)
            if ($itemRows.Count -gt 0) {
                $itemRows | Export-Csv -Path $itemsPartialPath -NoTypeInformation -Encoding UTF8 -Append
            }
            [pscustomobject]@{ CapacityId = [string]$capacityId } |
                Export-Csv -Path $doneCapsPath -NoTypeInformation -Encoding UTF8 -Append
            [void]$doneCaps.Add($capKey)

            Write-Host ("  -> {0} item instance(s) recorded; capacity marked done." -f $itemRows.Count) -ForegroundColor Green
        }
        catch {
            $status = Get-HttpStatusCodeFromError -ErrorRecord $_
            Write-Warning ("  Capacity {0} failed (status {1}): {2}" -f $capacityId, $status, $_.Exception.Message)
            Write-Warning "  Left as NOT done -- re-run the script to retry just this capacity."
            $failed.Add([string]$capacityId)
            continue
        }
    }

    if ($failed.Count -gt 0) {
        Write-Warning ("{0} capacity(ies) failed this run: {1}" -f $failed.Count, ($failed -join ', '))
    }
}

# ---------------------------------------------------------------------------
# Build the item lookup from the checkpoint and classify every workspace.
# ---------------------------------------------------------------------------
$itemLookup = @{}
if (Test-Path -Path $itemsPartialPath) {
    # Idempotency guard: a crash in the tiny window between "append items" and
    # "mark capacity done" could re-append a capacity's rows on the resume run.
    # De-dupe by item id so FabricItemCount can't be double-counted.
    $seenItems = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in Import-Csv -Path $itemsPartialPath) {
        if (-not $r.WorkspaceId) { continue }
        if (($r.PSObject.Properties.Name -contains 'ItemId') -and $r.ItemId) {
            if (-not $seenItems.Add(([string]$r.ItemId).ToLowerInvariant())) { continue }
        }
        $key = ([string]$r.WorkspaceId).ToLowerInvariant()
        if (-not $itemLookup.ContainsKey($key)) {
            $itemLookup[$key] = New-Object System.Collections.Generic.List[string]
        }
        if ($r.ItemType) { $itemLookup[$key].Add([string]$r.ItemType) }
    }
}

$rows = New-Object System.Collections.Generic.List[object]

foreach ($ws in $wsRows) {
    $wsKey = ([string]$ws.WorkspaceId).ToLowerInvariant()

    if (-not $ws.CapacityId) {
        # No capacity => provably pure Power BI (final).
        $classification  = 'PurePowerBI'
        $fabricItemCount = 0
        $fabricList      = ''
    }
    elseif ($doneCaps.Contains(([string]$ws.CapacityId).ToLowerInvariant())) {
        # Capacity scanned => classify from the item index using the allowlist.
        $typesForWs      = if ($itemLookup.ContainsKey($wsKey)) { $itemLookup[$wsKey] } else { @() }
        $fabricInstances = @($typesForWs | Where-Object { $_ -and ($PowerBIItemTypes -notcontains $_) })
        $fabricItemCount = $fabricInstances.Count
        $matchedTypes    = @($fabricInstances | Select-Object -Unique | Sort-Object)
        $classification  = if ($matchedTypes.Count) { 'HasFabricItems' } else { 'PurePowerBI' }
        $fabricList      = ($matchedTypes -join ';')
    }
    else {
        # Capacity-backed but its capacity hasn't been scanned yet (not done /
        # failed / interrupted). Leave unresolved; a future run will fill it in.
        $classification  = 'Pending'
        $fabricItemCount = ''
        $fabricList      = ''
    }

    $rows.Add([pscustomobject]@{
        WorkspaceId     = $ws.WorkspaceId
        WorkspaceName   = $ws.WorkspaceName
        Type            = $ws.Type
        State           = $ws.State
        Classification  = $classification
        CapacityId      = $ws.CapacityId
        FabricItemCount = $fabricItemCount
        fabricitems     = $fabricList
    })
}

$rows | Export-Csv -Path $fullSetPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fullSetPath) -ForegroundColor Green

$blockers = @($rows | Where-Object { $_.Classification -eq 'HasFabricItems' })
$blockers | Export-Csv -Path $fabricItemPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fabricItemPath) -ForegroundColor Green

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$pureCount    = @($rows | Where-Object { $_.Classification -eq 'PurePowerBI' }).Count
$pendingCount = @($rows | Where-Object { $_.Classification -eq 'Pending' }).Count

Write-Host ""
Write-Host "==================== Summary ====================" -ForegroundColor Cyan
Write-Host ("Total workspaces          : {0}" -f $rows.Count)
Write-Host ("Capacities scanned/total  : {0}/{1}" -f $doneCaps.Count, $capacityIds.Count)
Write-Host ("Has Fabric items          : {0}" -f $blockers.Count)
Write-Host ("Pure Power BI (safe)      : {0}" -f $pureCount)
Write-Host ("Pending (not yet scanned) : {0}" -f $pendingCount)
Write-Host "================================================" -ForegroundColor Cyan
if ($pendingCount -gt 0) {
    Write-Host ("{0} workspace(s) still Pending. Re-run ./listfabricitems.ps1 to resolve them (throttled or failed capacities will be retried)." -f $pendingCount) -ForegroundColor Yellow
}
else {
    Write-Host "All workspaces resolved. fabricitem.csv is your migration-blocker worklist." -ForegroundColor Green
}
