<#
.SYNOPSIS
    Scans all accessible Power BI / Fabric workspaces and flags those containing
    Fabric items that block cross-region workspace migration (P -> F license move).

.DESCRIPTION
    Uses the Microsoft Fabric REST API to enumerate every workspace the signed-in
    identity can access, then inspects each workspace's items. Any item whose type
    matches the known Fabric item types (migration blockers) is recorded in a new
    'fabricitems' column.

    Outputs:
      - fullset.csv    : all workspaces, with the 'fabricitems' column
      - fabricitem.csv : only workspaces that contain Fabric items (migration blockers)

.PARAMETER OutputFolder
    Folder where the CSV files are written. Defaults to the script's own directory.

.EXAMPLE
    Connect-AzAccount
    ./scanfabricitems.ps1

.EXAMPLE
    ./scanfabricitems.ps1 -OutputFolder "C:\CodeProject\ScanFabricWS\out"

.NOTES
    Requires the Az PowerShell module and a signed-in Entra ID account
    (Connect-AzAccount). See README.md for required permissions.
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
$FabricResourceUrl = 'https://api.fabric.microsoft.com'
$FabricApiBase     = 'https://api.fabric.microsoft.com/v1'

# Item types considered "Fabric items" (cross-region migration blockers).
# Validated against the official Fabric Core REST API ItemType enumeration:
# https://learn.microsoft.com/en-us/rest/api/fabric/core/items/list-items
# Power BI artifacts (Dashboard, Report, SemanticModel, PaginatedReport, Datamart)
# are intentionally excluded. "Additional item types may be added over time."
$FabricItemTypes = @(
    'Lakehouse',
    'Warehouse',
    'WarehouseSnapshot',
    'Notebook',
    'DataPipeline',
    'CopyJob',
    'Dataflow',
    'Eventhouse',
    'Eventstream',
    'KQLDatabase',
    'KQLQueryset',
    'KQLDashboard',
    'MirroredDatabase',
    'MirroredWarehouse',
    'MirroredAzureDatabricksCatalog',
    'SQLDatabase',
    'SQLEndpoint',
    'Environment',
    'MLModel',
    'MLExperiment',
    'SparkJobDefinition',
    'Reflex',
    'GraphQLApi',
    'MountedDataFactory',
    'VariableLibrary',
    'ApacheAirflowJob',
    'UserDataFunction',
    'DataAgent'
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

function Get-FabricHeaders {
    Assert-RequiredModules

    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "No Azure context found. Signing in with device code..." -ForegroundColor Yellow
        Write-Host "Open the displayed URL in a browser and enter the code to authenticate." -ForegroundColor Yellow
        Connect-AzAccount -UseDeviceAuthentication | Out-Null
    }

    # Request the token as a SecureString when the parameter is supported
    # (Az.Accounts 5.x); fall back gracefully on older versions.
    $tokenParams = @{ ResourceUrl = $FabricResourceUrl }
    if ((Get-Command Get-AzAccessToken).Parameters.ContainsKey('AsSecureString')) {
        $tokenParams['AsSecureString'] = $true
    }

    $tokenObject = Get-AzAccessToken @tokenParams
    $token = ConvertFrom-TokenValue -Token $tokenObject.Token
    if (-not $token) {
        throw "Failed to acquire an access token for $FabricResourceUrl."
    }

    return @{
        Authorization  = "Bearer $token"
        'Content-Type' = 'application/json'
    }
}

# ---------------------------------------------------------------------------
# REST helper with pagination and retry (429 / 5xx)
# ---------------------------------------------------------------------------
function Invoke-FabricGet {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [int]$MaxRetries = 5
    )

    $results = @()
    $nextUri = $Uri

    while ($nextUri) {
        $attempt = 0
        $response = $null

        while ($true) {
            try {
                $response = Invoke-RestMethod -Uri $nextUri -Headers $Headers -Method Get
                break
            }
            catch {
                $attempt++
                $status = $null
                if ($_.Exception.Response) {
                    $status = [int]$_.Exception.Response.StatusCode
                }

                if (($status -eq 429 -or ($status -ge 500 -and $status -lt 600)) -and $attempt -le $MaxRetries) {
                    $delay = [Math]::Min(60, [Math]::Pow(2, $attempt))
                    Write-Warning "Request to $nextUri failed with status $status. Retry $attempt/$MaxRetries in $delay s..."
                    Start-Sleep -Seconds $delay
                    continue
                }
                throw
            }
        }

        if ($null -ne $response.value) {
            $results += $response.value
        }
        elseif ($null -ne $response) {
            $results += $response
        }

        # Fabric pagination: continuationUri / continuationToken
        if ($response.PSObject.Properties.Name -contains 'continuationUri' -and $response.continuationUri) {
            $nextUri = $response.continuationUri
        }
        else {
            $nextUri = $null
        }
    }

    return $results
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$fullSetPath    = Join-Path $OutputFolder 'fullset.csv'
$fabricItemPath = Join-Path $OutputFolder 'fabricitem.csv'

Write-Host "Authenticating to Fabric API..." -ForegroundColor Cyan
$headers = Get-FabricHeaders

Write-Host "Listing workspaces..." -ForegroundColor Cyan
$workspaces = Invoke-FabricGet -Uri "$FabricApiBase/workspaces" -Headers $headers
Write-Host ("Found {0} workspace(s)." -f $workspaces.Count) -ForegroundColor Green

$rows = New-Object System.Collections.Generic.List[object]
$counter = 0

foreach ($ws in $workspaces) {
    $counter++
    $wsId   = $ws.id
    $wsName = $ws.displayName
    Write-Host ("[{0}/{1}] Scanning '{2}'..." -f $counter, $workspaces.Count, $wsName)

    $matchedTypes = @()
    $scanError    = $null

    try {
        $items = Invoke-FabricGet -Uri "$FabricApiBase/workspaces/$wsId/items" -Headers $headers
        $matchedTypes = $items |
            Where-Object { $FabricItemTypes -contains $_.type } |
            Select-Object -ExpandProperty type -Unique |
            Sort-Object
    }
    catch {
        $scanError = $_.Exception.Message
        Write-Warning ("  Could not scan '{0}': {1}" -f $wsName, $scanError)
    }

    $fabricItemsValue = if ($scanError) {
        "ERROR: $scanError"
    }
    else {
        ($matchedTypes -join ';')
    }

    $rows.Add([pscustomobject]@{
        WorkspaceId   = $wsId
        WorkspaceName = $wsName
        Type          = $ws.type
        CapacityId    = $ws.capacityId
        fabricitems   = $fabricItemsValue
    })
}

# Export full set
$rows | Export-Csv -Path $fullSetPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fullSetPath) -ForegroundColor Green

# Export only workspaces containing Fabric items (non-empty, non-error)
$blockers = $rows | Where-Object {
    $_.fabricitems -and -not $_.fabricitems.StartsWith('ERROR:')
}
$blockers | Export-Csv -Path $fabricItemPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fabricItemPath) -ForegroundColor Green

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$errorCount = ($rows | Where-Object { $_.fabricitems -like 'ERROR:*' }).Count
Write-Host ""
Write-Host "==================== Summary ====================" -ForegroundColor Cyan
Write-Host ("Total workspaces scanned : {0}" -f $rows.Count)
Write-Host ("With Fabric items        : {0}" -f $blockers.Count)
Write-Host ("Migration-safe           : {0}" -f ($rows.Count - $blockers.Count - $errorCount))
if ($errorCount -gt 0) {
    Write-Host ("Scan errors              : {0} (see 'ERROR:' rows in fullset.csv)" -f $errorCount) -ForegroundColor Yellow
}
Write-Host "================================================" -ForegroundColor Cyan
