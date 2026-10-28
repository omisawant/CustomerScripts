<#
.SYNOPSIS
Exports Azure quota limits and current utilization to an Excel workbook.

.DESCRIPTION
Uses each Azure service's provider-native usage API to collect current usage
and limits for every accessible Azure subscription. Microsoft.Quota does not
need to be registered. The workbook contains a summary, one worksheet per
resource provider, and an errors worksheet. Only Azure regions located in the
United States are queried.

Quota availability varies by resource provider and region. Providers without
a documented native usage API are recorded once per subscription in the
Errors worksheet. Supply -QuotaProvider to override the default provider list.

.PARAMETER OutputPath
Path of the .xlsx workbook to create.

.PARAMETER SubscriptionId
Optional list of subscription IDs to include. The default is every enabled
subscription visible to the signed-in account.

.PARAMETER Location
Optional list of U.S. Azure location names, such as eastus and westus3. The
default is every U.S. location returned by Get-AzLocation for each
subscription. Non-U.S. and unknown locations are rejected.

.PARAMETER QuotaProvider
Resource provider namespaces to query. Defaults to Compute, Network, Machine
Learning, Storage, Container Instances, Container Apps, App Service, AI
Search, Azure Database for PostgreSQL, and Azure Kubernetes Service.

.PARAMETER WarningThresholdPercent
Utilization percentage at which a quota row is marked Warning.

.PARAMETER InstallMissingModules
Installs missing Az and ImportExcel modules for the current user.

.PARAMETER Force
Overwrites OutputPath if it already exists.

.EXAMPLE
Connect-AzAccount
.\Export-AzureQuotaUtilization.ps1

.EXAMPLE
.\Export-AzureQuotaUtilization.ps1 -Location eastus,westus3 `
    -OutputPath C:\Reports\AzureQuota.xlsx -InstallMissingModules -Force
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string] $OutputPath = (Join-Path -Path (Get-Location) -ChildPath (
        'AzureQuotaUtilization_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    )),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]] $SubscriptionId,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]] $Location,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string[]] $QuotaProvider = @(
        'Microsoft.Web',
        'Microsoft.Automation',
        'Microsoft.App',
        'Microsoft.ContainerInstance',
        'Microsoft.DBforMySQL',
        'Microsoft.HDInsight',
        'Microsoft.ContainerService',
        'Microsoft.LabServices',
        'Microsoft.DBforPostgreSQL',
        'Microsoft.AVS',
        'Microsoft.Compute',
        'Microsoft.ClassicCompute',
        'Microsoft.DevCenter',
        'Microsoft.StorageCache',
        'Microsoft.MachineLearningServices',
        'Microsoft.DevOpsInfrastructure',
        'Microsoft.Fabric',
        'Microsoft.Network',
        'Microsoft.Search',
        'Microsoft.Storage',
        'Microsoft.ClassicStorage'
    ),

    [Parameter()]
    [ValidateRange(0, 100)]
    [double] $WarningThresholdPercent = 80,

    [Parameter()]
    [switch] $InstallMissingModules,

    [Parameter()]
    [switch] $Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Import-Dependency {
    param(
        [Parameter(Mandatory)]
        [string] $Name
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if (-not $InstallMissingModules) {
            throw "Required module '$Name' is not installed. Rerun with -InstallMissingModules or run: Install-Module $Name -Scope CurrentUser"
        }

        Write-Verbose "Installing module '$Name' for the current user."
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Get-FirstPropertyValue {
    param(
        [Parameter()]
        [AllowNull()]
        [object] $InputObject,

        [Parameter(Mandatory)]
        [string[]] $Path
    )

    foreach ($candidatePath in $Path) {
        $value = $InputObject
        $found = $true

        foreach ($segment in ($candidatePath -split '\.')) {
            if ($null -eq $value) {
                $found = $false
                break
            }

            $property = $value.PSObject.Properties[$segment]
            if ($null -eq $property) {
                $found = $false
                break
            }

            $value = $property.Value
        }

        if ($found -and $null -ne $value) {
            return $value
        }
    }

    return $null
}

function ConvertTo-NullableDouble {
    param(
        [Parameter()]
        [AllowNull()]
        [object] $Value
    )

    if ($null -eq $Value) {
        return $null
    }

    $number = 0.0
    if ([double]::TryParse(
            [string] $Value,
            [System.Globalization.NumberStyles]::Any,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref] $number
        )) {
        return $number
    }

    return $null
}

function Invoke-QuotaCommand {
    param(
        [Parameter(Mandatory)]
        [scriptblock] $ScriptBlock,

        [Parameter()]
        [ValidateRange(1, 10)]
        [int] $MaximumAttempts = 4
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $message = $_.Exception.Message
            $isTransient = $message -match '(429|Too Many Requests|throttl|timeout|temporar|5\d\d)'

            if (-not $isTransient -or $attempt -eq $MaximumAttempts) {
                throw
            }

            Start-Sleep -Seconds ([math]::Pow(2, $attempt))
        }
    }
}

function Get-NativeUsagePath {
    param(
        [Parameter(Mandatory)]
        [string] $SubscriptionId,

        [Parameter(Mandatory)]
        [string] $Provider,

        [Parameter(Mandatory)]
        [string] $AzureLocation
    )

    $basePath = "/subscriptions/$SubscriptionId/providers/$Provider/locations/$AzureLocation"
    switch ($Provider) {
        'Microsoft.Compute' {
            return "$basePath/usages?api-version=2026-04-01"
        }
        'Microsoft.Network' {
            return "$basePath/usages?api-version=2026-01-01"
        }
        'Microsoft.Storage' {
            return "$basePath/usages?api-version=2025-06-01"
        }
        'Microsoft.ContainerInstance' {
            return "$basePath/usages?api-version=2026-07-01"
        }
        'Microsoft.App' {
            return "$basePath/usages?api-version=2025-07-01"
        }
        'Microsoft.Web' {
            return "$basePath/usages?api-version=2026-07-15"
        }
        'Microsoft.Search' {
            return "$basePath/usages?api-version=2025-05-01"
        }
        'Microsoft.DBforPostgreSQL' {
            return "$basePath/resourceType/flexibleServers/usages?api-version=2026-10-01"
        }
        'Microsoft.MachineLearningServices' {
            return "$basePath/usages?api-version=2026-05-01"
        }
        'Microsoft.HDInsight' {
            return "$basePath/usages?api-version=2021-06-01"
        }
        'Microsoft.LabServices' {
            return "$basePath/usages?api-version=2023-06-07"
        }
        'Microsoft.DevCenter' {
            return "$basePath/usages?api-version=2025-02-01"
        }
        'Microsoft.StorageCache' {
            return "$basePath/usages?api-version=2026-08-01"
        }
        'Microsoft.DevOpsInfrastructure' {
            return "$basePath/usages?api-version=2026-06-02"
        }
        default {
            return $null
        }
    }
}

function Get-AzRestCollection {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $items = [System.Collections.Generic.List[object]]::new()
    $nextRequest = $Path

    while ($nextRequest) {
        $request = $nextRequest
        $response = Invoke-QuotaCommand -ScriptBlock {
            if ([uri]::IsWellFormedUriString($request, [System.UriKind]::Absolute)) {
                Invoke-AzRestMethod -Method GET -Uri $request -ErrorAction Stop
            }
            else {
                Invoke-AzRestMethod -Method GET -Path $request -ErrorAction Stop
            }
        }

        if ([int] $response.StatusCode -lt 200 -or [int] $response.StatusCode -ge 300) {
            throw "Azure Resource Manager returned HTTP $($response.StatusCode): $($response.Content)"
        }

        if ([string]::IsNullOrWhiteSpace([string] $response.Content)) {
            break
        }

        $body = $response.Content | ConvertFrom-Json
        foreach ($item in @($body.value)) {
            if ($null -ne $item) {
                $items.Add($item)
            }
        }

        $nextRequest = Get-FirstPropertyValue -InputObject $body -Path @(
            'nextLink',
            'NextLink'
        )
    }

    return $items.ToArray()
}

function ConvertTo-QuotaRow {
    param(
        [Parameter(Mandatory)]
        [object] $Subscription,

        [Parameter(Mandatory)]
        [string] $Provider,

        [Parameter(Mandatory)]
        [string] $AzureLocation,

        [Parameter(Mandatory)]
        [object] $Usage,

        [Parameter(Mandatory)]
        [double] $WarningThreshold
    )

    $resourceName = Get-FirstPropertyValue -InputObject $Usage -Path @(
        'name.value',
        'Name.Value',
        'name',
        'Name'
    )
    $localizedName = Get-FirstPropertyValue -InputObject $Usage -Path @(
        'name.localizedValue',
        'Name.LocalizedValue'
    )
    $currentUsage = ConvertTo-NullableDouble -Value (
        Get-FirstPropertyValue -InputObject $Usage -Path @(
            'currentValue',
            'CurrentValue'
        )
    )
    $limit = ConvertTo-NullableDouble -Value (
        Get-FirstPropertyValue -InputObject $Usage -Path @(
            'limit',
            'Limit'
        )
    )
    $unit = Get-FirstPropertyValue -InputObject $Usage -Path @('unit', 'Unit')

    $remaining = $null
    $utilizationPercent = $null
    if ($null -ne $limit -and $limit -ge 0 -and $null -ne $currentUsage) {
        $remaining = $limit - $currentUsage
        if ($limit -gt 0) {
            $utilizationPercent = [math]::Round(
                ($currentUsage / $limit) * 100,
                2
            )
        }
        elseif ($currentUsage -eq 0) {
            $utilizationPercent = 0
        }
    }

    if ($limit -eq -1) {
        $status = 'Unlimited'
    }
    elseif ($null -eq $limit) {
        $status = 'Usage only'
    }
    elseif ($null -eq $currentUsage) {
        $status = 'Limit only'
    }
    elseif ($limit -gt 0 -and $currentUsage -ge $limit) {
        $status = 'Exhausted'
    }
    elseif ($null -ne $utilizationPercent -and
        $utilizationPercent -ge $WarningThreshold) {
        $status = 'Warning'
    }
    else {
        $status = 'OK'
    }

    return [pscustomobject] @{
        SubscriptionName   = $Subscription.Name
        SubscriptionId     = $Subscription.Id
        TenantId           = $Subscription.TenantId
        Provider           = $Provider
        Location           = $AzureLocation
        QuotaName          = [string] $resourceName
        QuotaDisplayName   = $localizedName
        Unit               = $unit
        CurrentUsage       = $currentUsage
        Limit              = $limit
        Remaining          = $remaining
        UtilizationPercent = $utilizationPercent
        Status             = $status
        IsQuotaApplicable  = $true
        DataSource         = 'Provider-native usage API'
        CollectedAtUtc     = [datetime]::UtcNow
    }
}

function Get-WorksheetName {
    param(
        [Parameter(Mandatory)]
        [string] $Provider,

        [Parameter(Mandatory)]
        [hashtable] $UsedNames
    )

    $name = $Provider -replace '^Microsoft\.', ''
    $name = $name -replace '[:\\/\?\*\[\]]', '_'
    if ($name.Length -gt 31) {
        $name = $name.Substring(0, 31)
    }

    $candidate = $name
    $suffix = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $suffixText = "_$suffix"
        $maximumBaseLength = 31 - $suffixText.Length
        $baseName = $name.Substring(0, [math]::Min($name.Length, $maximumBaseLength))
        $candidate = "$baseName$suffixText"
        $suffix++
    }

    $UsedNames[$candidate] = $true
    return $candidate
}

function Add-CollectionError {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]] $ErrorList,

        [Parameter()]
        [AllowNull()]
        [object] $Subscription,

        [Parameter()]
        [AllowEmptyString()]
        [string] $Provider,

        [Parameter()]
        [AllowEmptyString()]
        [string] $AzureLocation,

        [Parameter(Mandatory)]
        [string] $Operation,

        [Parameter(Mandatory)]
        [string] $Message
    )

    $ErrorList.Add([pscustomobject] @{
            SubscriptionName = if ($null -ne $Subscription) { $Subscription.Name } else { $null }
            SubscriptionId   = if ($null -ne $Subscription) { $Subscription.Id } else { $null }
            TenantId         = if ($null -ne $Subscription) { $Subscription.TenantId } else { $null }
            Provider         = $Provider
            Location         = $AzureLocation
            Operation        = $Operation
            Message          = $Message
            CollectedAtUtc   = [datetime]::UtcNow
        })
}

foreach ($moduleName in @('Az.Accounts', 'Az.Resources', 'ImportExcel')) {
    Import-Dependency -Name $moduleName
}

$context = Get-AzContext
if ($null -eq $context -or $null -eq $context.Account) {
    throw 'No Azure session was found. Run Connect-AzAccount before this script.'
}

if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath $OutputPath
}

if ([System.IO.Path]::GetExtension($OutputPath) -ne '.xlsx') {
    $OutputPath = "$OutputPath.xlsx"
}

$outputDirectory = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path -LiteralPath $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

if (Test-Path -LiteralPath $OutputPath) {
    if (-not $Force) {
        throw "Output file already exists: $OutputPath. Use -Force to overwrite it."
    }

    Remove-Item -LiteralPath $OutputPath -Force
}

$allSubscriptions = @(
    Get-AzSubscription |
        Where-Object { $_.State -eq 'Enabled' } |
        Sort-Object -Property Name, Id
)

if ($SubscriptionId) {
    $requestedIds = @{}
    foreach ($id in $SubscriptionId) {
        $requestedIds[$id] = $true
    }

    $subscriptions = @($allSubscriptions | Where-Object { $requestedIds.ContainsKey($_.Id) })
    $missingIds = @($SubscriptionId | Where-Object { $_ -notin $subscriptions.Id })
    if ($missingIds.Count -gt 0) {
        throw "The following enabled subscriptions were not found or are not accessible: $($missingIds -join ', ')"
    }
}
else {
    $subscriptions = $allSubscriptions
}

if ($subscriptions.Count -eq 0) {
    throw 'No enabled Azure subscriptions are accessible to the signed-in account.'
}

$quotaRowsByProvider = @{}
foreach ($provider in $QuotaProvider) {
    $quotaRowsByProvider[$provider] = [System.Collections.Generic.List[object]]::new()
}

$summaryRows = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[object]]::new()
$progressTimer = [System.Diagnostics.Stopwatch]::StartNew()
$subscriptionIndex = 0

foreach ($subscription in $subscriptions) {
    $subscriptionIndex++
    $subscriptionBasePercent = [math]::Floor(
        (($subscriptionIndex - 1) / $subscriptions.Count) * 95
    )
    Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
        -Status "Subscription $subscriptionIndex/$($subscriptions.Count): $($subscription.Name)" `
        -CurrentOperation 'Selecting subscription context' `
        -PercentComplete $subscriptionBasePercent

    try {
        Set-AzContext -SubscriptionId $subscription.Id -TenantId $subscription.TenantId | Out-Null
    }
    catch {
        Add-CollectionError -ErrorList $errors -Subscription $subscription `
            -Provider '' -AzureLocation '' -Operation 'Set-AzContext' `
            -Message $_.Exception.Message
        continue
    }

    try {
        $unitedStatesLocations = @(
            Get-AzLocation |
                Where-Object {
                    $_.GeographyGroup -eq 'US' -or
                    $_.PhysicalLocation -match 'United States'
                } |
                Select-Object -ExpandProperty Location |
                Sort-Object -Unique
        )
    }
    catch {
        Add-CollectionError -ErrorList $errors -Subscription $subscription `
            -Provider '' -AzureLocation '' -Operation 'Get-AzLocation' `
            -Message $_.Exception.Message
        continue
    }

    if ($unitedStatesLocations.Count -eq 0) {
        Add-CollectionError -ErrorList $errors -Subscription $subscription `
            -Provider '' -AzureLocation '' -Operation 'Filter US locations' `
            -Message 'Get-AzLocation returned no Azure regions in the United States.'
        continue
    }

    if ($Location) {
        $invalidLocations = @(
            $Location |
                Where-Object { $_ -notin $unitedStatesLocations } |
                Sort-Object -Unique
        )
        if ($invalidLocations.Count -gt 0) {
            throw "The following locations are not recognized as U.S. Azure regions: $($invalidLocations -join ', ')"
        }

        $subscriptionLocations = @($Location | Sort-Object -Unique)
    }
    else {
        $subscriptionLocations = $unitedStatesLocations
    }

    $providerIndex = 0
    foreach ($provider in $QuotaProvider) {
        $providerIndex++
        $providerStartCount = $quotaRowsByProvider[$provider].Count
        $providerStartErrorCount = $errors.Count
        $scopesQueried = 0
        $scopesWithData = 0

        if ($provider -eq 'Microsoft.Automation') {
            try {
                $automationAccounts = @(
                    Get-AzResource -ResourceType 'Microsoft.Automation/automationAccounts' `
                        -ErrorAction Stop
                )
            }
            catch {
                Add-CollectionError -ErrorList $errors -Subscription $subscription `
                    -Provider $provider -AzureLocation '' `
                    -Operation 'List Automation accounts' -Message $_.Exception.Message
                $automationAccounts = @()
            }

            $accountIndex = 0
            foreach ($account in $automationAccounts) {
                $accountIndex++
                $scopesQueried++
                Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
                    -Status "Subscription $subscriptionIndex/$($subscriptions.Count) | Provider $providerIndex/$($QuotaProvider.Count)" `
                    -CurrentOperation "Automation account $accountIndex/$($automationAccounts.Count): $($account.Name)" `
                    -PercentComplete $subscriptionBasePercent

                try {
                    $usages = @(
                        Get-AzRestCollection -Path (
                            "$($account.ResourceId)/usages?api-version=2024-10-23"
                        )
                    )
                    if ($usages.Count -gt 0) {
                        $scopesWithData++
                    }

                    foreach ($usage in $usages) {
                        $quotaRowsByProvider[$provider].Add(
                            (ConvertTo-QuotaRow -Subscription $subscription `
                                -Provider $provider -AzureLocation $account.Location `
                                -Usage $usage `
                                -WarningThreshold $WarningThresholdPercent)
                        )
                    }
                }
                catch {
                    Add-CollectionError -ErrorList $errors -Subscription $subscription `
                        -Provider $provider -AzureLocation $account.Location `
                        -Operation "Get native usage for $($account.Name)" `
                        -Message $_.Exception.Message
                }
            }
        }
        else {
            $samplePath = Get-NativeUsagePath -SubscriptionId $subscription.Id `
                -Provider $provider -AzureLocation $subscriptionLocations[0]

            if ($null -eq $samplePath) {
                $unsupportedMessage = switch ($provider) {
                    'Microsoft.ContainerService' {
                        'AKS has no provider-native quota listing API. AKS consumes quotas reported in the Microsoft.Compute and Microsoft.Network worksheets.'
                    }
                    'Microsoft.AVS' {
                        'AVS exposes quota availability, but not current usage and limit values in a native usage-list API.'
                    }
                    default {
                        "Provider '$provider' has no documented provider-native usage and limit listing API."
                    }
                }

                Add-CollectionError -ErrorList $errors -Subscription $subscription `
                    -Provider $provider -AzureLocation '' `
                    -Operation 'Native quota API not available' `
                    -Message $unsupportedMessage
            }
            else {
                $locationIndex = 0
                foreach ($azureLocation in $subscriptionLocations) {
                    $locationIndex++
                    $scopesQueried++
                    $locationFraction = (
                        $locationIndex - 1
                    ) / [math]::Max(1, $subscriptionLocations.Count)
                    $subscriptionFraction = (
                        ($providerIndex - 1) + $locationFraction
                    ) / $QuotaProvider.Count
                    $overallFraction = (
                        ($subscriptionIndex - 1) + $subscriptionFraction
                    ) / $subscriptions.Count
                    $progressPercent = [math]::Min(
                        94,
                        [math]::Floor($overallFraction * 95)
                    )
                    $progressStatus = (
                        "Subscription {0}/{1} | Provider {2}/{3} | Location {4}/{5} | Elapsed {6}" -f
                        $subscriptionIndex,
                        $subscriptions.Count,
                        $providerIndex,
                        $QuotaProvider.Count,
                        $locationIndex,
                        $subscriptionLocations.Count,
                        $progressTimer.Elapsed.ToString('hh\:mm\:ss')
                    )
                    Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
                        -Status $progressStatus `
                        -CurrentOperation "$provider in $azureLocation" `
                        -PercentComplete $progressPercent

                    $path = Get-NativeUsagePath -SubscriptionId $subscription.Id `
                        -Provider $provider -AzureLocation $azureLocation
                    try {
                        $usages = @(Get-AzRestCollection -Path $path)
                        if ($usages.Count -gt 0) {
                            $scopesWithData++
                        }

                        foreach ($usage in $usages) {
                            $quotaRowsByProvider[$provider].Add(
                                (ConvertTo-QuotaRow -Subscription $subscription `
                                    -Provider $provider -AzureLocation $azureLocation `
                                    -Usage $usage `
                                    -WarningThreshold $WarningThresholdPercent)
                            )
                        }
                    }
                    catch {
                        Add-CollectionError -ErrorList $errors -Subscription $subscription `
                            -Provider $provider -AzureLocation $azureLocation `
                            -Operation 'Get provider-native usage' `
                            -Message $_.Exception.Message

                        if ($_.Exception.Message -match 'MissingSubscriptionRegistration') {
                            break
                        }
                    }
                }
            }
        }

        $providerRows = @(
            $quotaRowsByProvider[$provider].ToArray() |
                Select-Object -Skip $providerStartCount
        )
        $warningCount = @(
            $providerRows | Where-Object { $_.Status -in @('Warning', 'Exhausted') }
        ).Count
        $utilizationValues = @(
            $providerRows |
                Where-Object { $null -ne $_.UtilizationPercent } |
                Select-Object -ExpandProperty UtilizationPercent
        )
        $maximumUtilization = if ($utilizationValues.Count -gt 0) {
            ($utilizationValues | Measure-Object -Maximum).Maximum
        }
        else {
            $null
        }

        $summaryRows.Add([pscustomobject] @{
                SubscriptionName         = $subscription.Name
                SubscriptionId           = $subscription.Id
                TenantId                 = $subscription.TenantId
                Provider                 = $provider
                ScopesQueried             = $scopesQueried
                ScopesWithData            = $scopesWithData
                QuotaRecords             = $providerRows.Count
                WarningOrExhausted       = $warningCount
                MaximumUtilizationPercent = $maximumUtilization
                CollectionErrors         = $errors.Count - $providerStartErrorCount
                CollectedAtUtc           = [datetime]::UtcNow
            })
    }
}

Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
    -Status "Collection complete | Elapsed $($progressTimer.Elapsed.ToString('hh\:mm\:ss'))" `
    -CurrentOperation 'Writing Summary worksheet' -PercentComplete 95

$summaryRows.ToArray() |
    Sort-Object -Property SubscriptionName, Provider |
    Export-Excel -Path $OutputPath -WorksheetName 'Summary' -TableName 'QuotaSummary' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

$usedWorksheetNames = @{ Summary = $true; Errors = $true }
$tableIndex = 1
foreach ($provider in $QuotaProvider) {
    $exportPercent = 95 + [math]::Floor(
        (($tableIndex - 1) / [math]::Max(1, $QuotaProvider.Count)) * 4
    )
    Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
        -Status "Writing provider worksheet $tableIndex/$($QuotaProvider.Count)" `
        -CurrentOperation $provider -PercentComplete $exportPercent

    $worksheetName = Get-WorksheetName -Provider $provider -UsedNames $usedWorksheetNames
    $rows = @(
        $quotaRowsByProvider[$provider].ToArray() |
            Sort-Object -Property SubscriptionName, Location, QuotaDisplayName, QuotaName
    )

    if ($rows.Count -eq 0) {
        $rows = @([pscustomobject] @{
                Provider = $provider
                Status   = 'No quota data returned; review the Errors worksheet.'
            })
    }

    $rows |
        Export-Excel -Path $OutputPath -WorksheetName $worksheetName `
            -TableName "QuotaData$tableIndex" -Append -AutoSize -AutoFilter `
            -FreezeTopRow -BoldTopRow
    $tableIndex++
}

if ($errors.Count -eq 0) {
    $errorRows = @([pscustomobject] @{
            Operation = 'Collection'
            Message   = 'No errors were reported.'
        })
}
else {
    $errorRows = @(
        $errors.ToArray() |
            Sort-Object -Property SubscriptionName, Provider, Location, Operation
    )
}

Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' `
    -Status 'Finalizing workbook' -CurrentOperation 'Writing Errors worksheet' `
    -PercentComplete 99

$errorRows |
    Export-Excel -Path $OutputPath -WorksheetName 'Errors' -TableName 'CollectionErrors' `
        -Append -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

$progressTimer.Stop()
Write-Progress -Id 0 -Activity 'Collecting Azure quota utilization' -Completed
Write-Host "Azure quota utilization workbook created: $OutputPath"
