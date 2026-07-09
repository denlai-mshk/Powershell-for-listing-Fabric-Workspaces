<#
.SYNOPSIS
    Part 1 of 2: fast tenant-wide workspace inventory (no item scan).

.DESCRIPTION
    Lists every workspace in the tenant via the read-only admin endpoint
    (/v1/admin/workspaces, falling back to the Power BI admin endpoint) and writes
    a lightweight inventory to workspaces.csv. NO item calls are made here, so this
    runs fast and uses a completely separate rate-limit bucket from /admin/items --
    it will not trip the 429 throttling that item enumeration can.

    Classification here is preliminary and based purely on capacity:
      - PurePowerBI : workspace has NO capacityId. Fabric items can only live on a
                      Fabric-capable capacity, so this is a FINAL, safe verdict.
      - Pending     : workspace IS capacity-backed. It *might* contain Fabric items;
                      only listfabricitems.ps1 (the slow, paced item scan) can
                      resolve it to PurePowerBI or HasFabricItems.

    Run listfabricitems.ps1 afterwards to resolve the 'Pending' rows.

.PARAMETER OutputFolder
    Folder where workspaces.csv is written. Defaults to the script's own directory.

.EXAMPLE
    Connect-AzAccount
    ./listworkspace.ps1

.NOTES
    Admin-only, read-only APIs. Requires Az.Accounts and a signed-in account with
    Fabric Administrator or Power BI Administrator rights. See README.md.
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = $PSScriptRoot
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
# Main
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$workspacesPath = Join-Path $OutputFolder 'workspaces.csv'

Write-Host "Authenticating to Fabric API..." -ForegroundColor Cyan
$fabricHeaders = Get-ApiHeaders -ResourceUrl $FabricResourceUrl -ResourceLabel 'Fabric API'

$workspaceSource = $null
$fallbackMessage = $null
$workspaces = @()

try {
    Write-Host "Listing workspaces from Fabric admin endpoint..." -ForegroundColor Cyan
    $workspaceSource = 'FabricAdmin'
    $workspaces = Invoke-FabricGet -Uri "$FabricApiBase/admin/workspaces" -Headers $fabricHeaders
}
catch {
    if (-not (Test-CanFallbackFromFabricAdminError -ErrorRecord $_)) {
        throw
    }

    $fallbackMessage = "Fabric admin workspace listing failed ({0}). Falling back to Power BI admin endpoint." -f $_.Exception.Message
    Write-Warning $fallbackMessage

    Write-Host "Authenticating to Power BI API for fallback..." -ForegroundColor Yellow
    $powerBIHeaders = Get-ApiHeaders -ResourceUrl $PowerBIResourceUrl -ResourceLabel 'Power BI API'
    Write-Host "Listing workspaces from Power BI admin endpoint..." -ForegroundColor Yellow
    $workspaceSource = 'PowerBIAdmin'
    $workspaces = Invoke-PowerBIAdminGetWorkspaces -Headers $powerBIHeaders
}

$workspaces = @($workspaces) | Where-Object { $null -ne $_ }
Write-Host ("Found {0} workspace(s)." -f $workspaces.Count) -ForegroundColor Green

# Normalise + assign the preliminary (capacity-only) classification.
$rows = New-Object System.Collections.Generic.List[object]

foreach ($ws in $workspaces) {
    $n = ConvertTo-NormalizedWorkspace -Workspace $ws
    if (-not $n) {
        Write-Warning "Skipping workspace record with missing id."
        continue
    }

    # No capacity => provably pure Power BI (final). Capacity-backed => Pending,
    # to be resolved by the item scan in listfabricitems.ps1.
    $classification = if ($n.CapacityId) { 'Pending' } else { 'PurePowerBI' }

    $rows.Add([pscustomobject]@{
        WorkspaceId    = $n.WorkspaceId
        WorkspaceName  = $n.WorkspaceName
        Type           = $n.Type
        State          = $n.State
        Classification = $classification
        CapacityId     = $n.CapacityId
    })
}

$rows | Export-Csv -Path $workspacesPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $workspacesPath) -ForegroundColor Green

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$pendingRows  = @($rows | Where-Object { $_.Classification -eq 'Pending' })
$pureRows     = @($rows | Where-Object { $_.Classification -eq 'PurePowerBI' })
$capacityIds  = @($pendingRows | Select-Object -ExpandProperty CapacityId -Unique)

Write-Host ""
Write-Host "==================== Summary ====================" -ForegroundColor Cyan
Write-Host ("Workspace source          : {0}" -f $workspaceSource)
Write-Host ("Total workspaces          : {0}" -f $rows.Count)
Write-Host ("PurePowerBI (no capacity) : {0}  <- final, migration-safe" -f $pureRows.Count)
Write-Host ("Pending (capacity-backed) : {0}  across {1} distinct capacity(ies)" -f $pendingRows.Count, $capacityIds.Count)
if ($fallbackMessage) {
    Write-Host ("Listing fallback used     : Power BI admin endpoint") -ForegroundColor Yellow
}
Write-Host "================================================" -ForegroundColor Cyan
if ($pendingRows.Count -gt 0) {
    Write-Host ("Next: run ./listfabricitems.ps1 to resolve the {0} 'Pending' workspace(s)." -f $pendingRows.Count) -ForegroundColor Yellow
}
else {
    Write-Host "No capacity-backed workspaces -- every workspace is already resolved as PurePowerBI." -ForegroundColor Green
}
