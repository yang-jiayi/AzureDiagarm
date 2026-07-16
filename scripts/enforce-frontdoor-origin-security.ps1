[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $true)]
    [string]$ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string]$ContainerApp,

    [Parameter(Mandatory = $true)]
    [string]$FrontDoorProfile,

    [string]$DiagnosticSettingName = 'azurediagarm-local-logs'
)

$ErrorActionPreference = 'Stop'

function Invoke-AzCli {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $output = & az @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI failed: az $($Arguments -join ' ')"
    }

    return ($output -join "`n")
}

function ConvertTo-Utf8JsonFile {
    param(
        [Parameter(Mandatory = $true)]$Value,
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$AsArray
    )

    $json = if ($AsArray) {
        $Value | ConvertTo-Json -Depth 16 -Compress -AsArray
    }
    else {
        $Value | ConvertTo-Json -Depth 16 -Compress
    }
    [System.IO.File]::WriteAllText(
        $Path,
        $json,
        [System.Text.UTF8Encoding]::new($false)
    )
}

Invoke-AzCli @('account', 'set', '--subscription', $SubscriptionId) | Out-Null

$appJson = Invoke-AzCli @(
    'containerapp', 'show',
    '--resource-group', $ResourceGroup,
    '--name', $ContainerApp,
    '--output', 'json'
)
$app = $appJson | ConvertFrom-Json
$location = [string]$app.location
$appId = [string]$app.id

if ([string]::IsNullOrWhiteSpace($location) -or [string]::IsNullOrWhiteSpace($appId)) {
    throw 'Unable to resolve the Container App location or resource ID.'
}

$azureLocations = (
    Invoke-AzCli @('account', 'list-locations', '--output', 'json')
) | ConvertFrom-Json
$locationRecord = $azureLocations |
    Where-Object { $_.name -eq $location -or $_.displayName -eq $location } |
    Select-Object -First 1
$serviceTagLocation = [string]$locationRecord.name

if ([string]::IsNullOrWhiteSpace($serviceTagLocation)) {
    throw "Unable to resolve the canonical Azure location name for $location."
}

$serviceTagsUrl = (
    "https://management.azure.com/subscriptions/$SubscriptionId" +
    "/providers/Microsoft.Network/locations/$serviceTagLocation/serviceTags" +
    '?api-version=2024-05-01'
)
$frontDoorTagJson = Invoke-AzCli @(
    'rest',
    '--method', 'get',
    '--url', $serviceTagsUrl,
    '--query', "values[?name == 'AzureFrontDoor.Backend'] | [0]",
    '--output', 'json'
)
$frontDoorTag = $frontDoorTagJson | ConvertFrom-Json

if ($null -eq $frontDoorTag) {
    throw "AzureFrontDoor.Backend was not present in the $serviceTagLocation service tag response."
}

$frontDoorIpv4Ranges = @(
    $frontDoorTag.properties.addressPrefixes |
        Where-Object { $_ -notmatch ':' }
) | Sort-Object -Unique

if ($frontDoorIpv4Ranges.Count -eq 0) {
    throw 'No IPv4 ranges were resolved for AzureFrontDoor.Backend.'
}

$ipv4Ranges = @(
    $frontDoorIpv4Ranges
    '168.63.129.16/32'
    '169.254.169.254/32'
) | Sort-Object -Unique

$rules = for ($index = 0; $index -lt $ipv4Ranges.Count; $index++) {
    [ordered]@{
        name = 'afd-backend-{0:D3}' -f $index
        description = 'AzureFrontDoor.Backend service tag'
        ipAddressRange = $ipv4Ranges[$index]
        action = 'Allow'
    }
}

$apiVersion = '2025-01-01'
$resourceUrl = "https://management.azure.com${appId}?api-version=$apiVersion"
$resource = (Invoke-AzCli @('rest', '--method', 'get', '--url', $resourceUrl)) | ConvertFrom-Json
$ingress = $resource.properties.configuration.ingress

$patch = @{
    properties = @{
        configuration = @{
            ingress = @{
                external = [bool]$ingress.external
                targetPort = [int]$ingress.targetPort
                transport = [string]$ingress.transport
                allowInsecure = [bool]$ingress.allowInsecure
                ipSecurityRestrictions = $rules
            }
        }
    }
}

$patchFile = [System.IO.Path]::GetTempFileName()
try {
    ConvertTo-Utf8JsonFile -Value $patch -Path $patchFile
    Invoke-AzCli @(
        'rest',
        '--method', 'patch',
        '--url', $resourceUrl,
        '--headers', 'Content-Type=application/json',
        '--body', "@$patchFile",
        '--output', 'none'
    ) | Out-Null
}
finally {
    Remove-Item $patchFile -Force -ErrorAction SilentlyContinue
}

$restrictionsVerified = $false
for ($attempt = 1; $attempt -le 30; $attempt++) {
    $actualRulesJson = Invoke-AzCli @(
        'containerapp', 'ingress', 'access-restriction', 'list',
        '--resource-group', $ResourceGroup,
        '--name', $ContainerApp,
        '--output', 'json'
    )
    $actualRules = @($actualRulesJson | ConvertFrom-Json)
    $actualRanges = @(
        $actualRules |
            Where-Object { $_.action -eq 'Allow' } |
            ForEach-Object { $_.ipAddressRange }
    ) | Sort-Object -Unique

    if ($actualRanges.Count -eq $ipv4Ranges.Count) {
        $rangeDifference = Compare-Object -ReferenceObject $ipv4Ranges -DifferenceObject $actualRanges
        if (-not $rangeDifference) {
            $restrictionsVerified = $true
            break
        }
    }

    Start-Sleep -Seconds 5
}

if (-not $restrictionsVerified) {
    throw 'Container Apps ingress restrictions do not match AzureFrontDoor.Backend.'
}

$frontDoorId = Invoke-AzCli @(
    'afd', 'profile', 'show',
    '--resource-group', $ResourceGroup,
    '--profile-name', $FrontDoorProfile,
    '--query', 'id',
    '--output', 'tsv'
)
$workspaceIds = @(
    (Invoke-AzCli @(
        'monitor', 'log-analytics', 'workspace', 'list',
        '--resource-group', $ResourceGroup,
        '--query', '[].id',
        '--output', 'tsv'
    )) -split "`r?`n" |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
)

if ($workspaceIds.Count -ne 1) {
    throw "Expected exactly one Log Analytics workspace in $ResourceGroup; found $($workspaceIds.Count)."
}

$logs = @(
    @{ category = 'FrontDoorAccessLog'; enabled = $true }
    @{ category = 'FrontDoorHealthProbeLog'; enabled = $true }
    @{ category = 'FrontDoorWebApplicationFirewallLog'; enabled = $true }
)
$metrics = @(
    @{ category = 'AllMetrics'; enabled = $true }
)
$logsFile = [System.IO.Path]::GetTempFileName()
$metricsFile = [System.IO.Path]::GetTempFileName()

try {
    ConvertTo-Utf8JsonFile -Value $logs -Path $logsFile -AsArray
    ConvertTo-Utf8JsonFile -Value $metrics -Path $metricsFile -AsArray
    Invoke-AzCli @(
        'monitor', 'diagnostic-settings', 'create',
        '--name', $DiagnosticSettingName,
        '--resource', $frontDoorId.Trim(),
        '--workspace', $workspaceIds[0].Trim(),
        '--logs', "@$logsFile",
        '--metrics', "@$metricsFile",
        '--output', 'none'
    ) | Out-Null
}
finally {
    Remove-Item $logsFile, $metricsFile -Force -ErrorAction SilentlyContinue
}

$diagnosticSetting = (
    Invoke-AzCli @(
        'monitor', 'diagnostic-settings', 'show',
        '--name', $DiagnosticSettingName,
        '--resource', $frontDoorId.Trim(),
        '--output', 'json'
    )
) | ConvertFrom-Json
$enabledLogCategories = @(
    $diagnosticSetting.logs |
        Where-Object { $_.enabled } |
        ForEach-Object { $_.category }
)

foreach ($category in $logs.category) {
    if ($category -notin $enabledLogCategories) {
        throw "Front Door diagnostic category $category is not enabled."
    }
}

Write-Host "Enforced $($ipv4Ranges.Count) Front Door origin ranges and local Front Door diagnostics."
