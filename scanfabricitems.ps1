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
$FabricResourceUrl  = 'https://api.fabric.microsoft.com'
$FabricApiBase      = 'https://api.fabric.microsoft.com/v1'
$PowerBIResourceUrl = 'https://analysis.windows.net/powerbi/api'
$PowerBIApiBase     = 'https://api.powerbi.com/v1.0/myorg'

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
# REST helper with pagination and retry (429 / 5xx)
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

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$fullSetPath    = Join-Path $OutputFolder 'fullset.csv'
$fabricItemPath = Join-Path $OutputFolder 'fabricitem.csv'

Write-Host "Authenticating to Fabric API..." -ForegroundColor Cyan
$fabricHeaders = Get-ApiHeaders -ResourceUrl $FabricResourceUrl -ResourceLabel 'Fabric API'

$workspaceSource = $null
$workspaceListFallbackMessage = $null
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

    $workspaceListFallbackMessage = "Fabric admin workspace listing failed ({0}). Falling back to Power BI admin endpoint." -f $_.Exception.Message
    Write-Warning $workspaceListFallbackMessage

    try {
        Write-Host "Authenticating to Power BI API for fallback..." -ForegroundColor Yellow
        $powerBIHeaders = Get-ApiHeaders -ResourceUrl $PowerBIResourceUrl -ResourceLabel 'Power BI API'
        Write-Host "Listing workspaces from Power BI admin endpoint..." -ForegroundColor Yellow
        $workspaceSource = 'PowerBIAdmin'
        $workspaces = Invoke-PowerBIAdminGetWorkspaces -Headers $powerBIHeaders
    }
    catch {
        Write-Warning ("Power BI admin fallback also failed ({0}). Falling back to Fabric user endpoint." -f $_.Exception.Message)
        Write-Host "Listing workspaces from Fabric user endpoint..." -ForegroundColor Yellow
        $workspaceSource = 'FabricUser'
        $workspaces = Invoke-FabricGet -Uri "$FabricApiBase/workspaces" -Headers $fabricHeaders
    }
}

$workspaces = @($workspaces) | Where-Object { $null -ne $_ }

Write-Host ("Found {0} workspace(s)." -f $workspaces.Count) -ForegroundColor Green

$rows = New-Object System.Collections.Generic.List[object]
$counter = 0

foreach ($ws in $workspaces) {
    $counter++
    $wsId   = $ws.id
    if (-not $wsId -and ($ws.PSObject.Properties.Name -contains 'Id')) {
        $wsId = $ws.Id
    }

    if (-not $wsId) {
        Write-Warning ("[{0}/{1}] Skipping workspace record with missing id." -f $counter, $workspaces.Count)
        continue
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

    Write-Host ("[{0}/{1}] Scanning '{2}'..." -f $counter, $workspaces.Count, $wsName)

    $matchedTypes = @()
    $scanError    = $null
    $scanSkippedReason = $null

    try {
        # Fabric item listing is workspace-role scoped.
        $itemsUri = "$FabricApiBase/workspaces/$wsId/items"

        $items = Invoke-FabricGet -Uri $itemsUri -Headers $fabricHeaders
        $matchedTypes = $items |
            Where-Object { $FabricItemTypes -contains $_.type } |
            Select-Object -ExpandProperty type -Unique |
            Sort-Object
    }
    catch {
        $status = Get-HttpStatusCodeFromError -ErrorRecord $_
        if ($status -in 401, 403, 404) {
            $scanSkippedReason = 'Workspace listed successfully, but Fabric item scan requires workspace membership (viewer or above), or the item endpoint is unavailable for this workspace.'
            Write-Warning ("  Skipped item scan for '{0}': {1}" -f $wsName, $scanSkippedReason)
        }
        else {
            $scanError = $_.Exception.Message
            Write-Warning ("  Could not scan '{0}': {1}" -f $wsName, $scanError)
        }
    }

    $fabricItemsValue = if ($scanError) {
        "ERROR: $scanError"
    }
    elseif ($scanSkippedReason) {
        "SKIPPED: $scanSkippedReason"
    }
    else {
        ($matchedTypes -join ';')
    }

    $rows.Add([pscustomobject]@{
        WorkspaceId   = $wsId
        WorkspaceName = $wsName
        Type          = $wsType
        CapacityId    = $wsCapacityId
        Source        = $workspaceSource
        fabricitems   = $fabricItemsValue
    })
}

# Export full set
$rows | Export-Csv -Path $fullSetPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fullSetPath) -ForegroundColor Green

# Export only workspaces containing Fabric items (non-empty, non-error)
$blockers = $rows | Where-Object {
    $_.fabricitems -and -not $_.fabricitems.StartsWith('ERROR:') -and -not $_.fabricitems.StartsWith('SKIPPED:')
}
$blockers | Export-Csv -Path $fabricItemPath -NoTypeInformation -Encoding UTF8
Write-Host ("Wrote {0}" -f $fabricItemPath) -ForegroundColor Green

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
$errorCount = ($rows | Where-Object { $_.fabricitems -like 'ERROR:*' }).Count
$skippedCount = ($rows | Where-Object { $_.fabricitems -like 'SKIPPED:*' }).Count
Write-Host ""
Write-Host "==================== Summary ====================" -ForegroundColor Cyan
Write-Host ("Workspace source        : {0}" -f $workspaceSource)
Write-Host ("Total workspaces scanned : {0}" -f $rows.Count)
Write-Host ("With Fabric items        : {0}" -f $blockers.Count)
Write-Host ("Migration-safe           : {0}" -f ($rows.Count - $blockers.Count - $errorCount - $skippedCount))
if ($skippedCount -gt 0) {
    Write-Host ("Scan skipped             : {0} (see 'SKIPPED:' rows in fullset.csv)" -f $skippedCount) -ForegroundColor Yellow
}
if ($errorCount -gt 0) {
    Write-Host ("Scan errors              : {0} (see 'ERROR:' rows in fullset.csv)" -f $errorCount) -ForegroundColor Yellow
}
if ($workspaceListFallbackMessage) {
    Write-Host ("Listing fallback used    : Yes" ) -ForegroundColor Yellow
}
Write-Host "================================================" -ForegroundColor Cyan
