<#
.SYNOPSIS
Exports Azure quota limits and current utilization to an Excel workbook.

.DESCRIPTION
Uses the Azure Quota API cmdlets to collect quota limits (Get-AzQuota) and
current usage (Get-AzQuotaUsage) for every accessible Azure subscription.
The workbook contains a summary, one worksheet per quota resource provider,
and an errors worksheet. Only Azure regions located in the United States are
queried.

Quota availability varies by resource provider and region. Providers that do
not expose data through the centralized Azure Quota API are recorded in the
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
        'Microsoft.Compute',
        'Microsoft.Network',
        'Microsoft.MachineLearningServices',
        'Microsoft.Storage',
        'Microsoft.ContainerInstance',
        'Microsoft.App',
        'Microsoft.Web',
        'Microsoft.Search',
        'Microsoft.DBforPostgreSQL',
        'Microsoft.ContainerService'
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

foreach ($moduleName in @('Az.Accounts', 'Az.Resources', 'Az.Quota', 'ImportExcel')) {
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
        $locationsWithData = 0
        $locationIndex = 0

        foreach ($azureLocation in $subscriptionLocations) {
            $locationIndex++
            $locationFraction = if ($subscriptionLocations.Count -gt 0) {
                ($locationIndex - 1) / $subscriptionLocations.Count
            }
            else {
                0
            }
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

            $scope = "subscriptions/$($subscription.Id)/providers/$provider/locations/$azureLocation"
            $quotas = @()
            $usages = @()

            try {
                $quotas = @(Invoke-QuotaCommand -ScriptBlock {
                        Get-AzQuota -Scope $scope -ErrorAction Stop
                    })
            }
            catch {
                Add-CollectionError -ErrorList $errors -Subscription $subscription `
                    -Provider $provider -AzureLocation $azureLocation `
                    -Operation 'Get-AzQuota' -Message $_.Exception.Message
            }

            try {
                $usages = @(Invoke-QuotaCommand -ScriptBlock {
                        Get-AzQuotaUsage -Scope $scope -ErrorAction Stop
                    })
            }
            catch {
                Add-CollectionError -ErrorList $errors -Subscription $subscription `
                    -Provider $provider -AzureLocation $azureLocation `
                    -Operation 'Get-AzQuotaUsage' -Message $_.Exception.Message
            }

            $quotaByName = @{}
            foreach ($quota in $quotas) {
                $quotaName = Get-FirstPropertyValue -InputObject $quota -Path @('Name', 'Name.Value')
                if ($null -ne $quotaName) {
                    $quotaByName[[string] $quotaName] = $quota
                }
            }

            $usageByName = @{}
            foreach ($usage in $usages) {
                $usageName = Get-FirstPropertyValue -InputObject $usage -Path @('Name', 'Name.Value')
                if ($null -ne $usageName) {
                    $usageByName[[string] $usageName] = $usage
                }
            }

            $resourceNames = @(
                @($quotaByName.Keys) + @($usageByName.Keys) |
                    Sort-Object -Unique
            )

            if ($resourceNames.Count -gt 0) {
                $locationsWithData++
            }

            foreach ($resourceName in $resourceNames) {
                $quota = $quotaByName[$resourceName]
                $usage = $usageByName[$resourceName]

                $limit = ConvertTo-NullableDouble -Value (
                    Get-FirstPropertyValue -InputObject $quota -Path @('LimitValue', 'Limit.Value')
                )
                $currentUsage = ConvertTo-NullableDouble -Value (
                    Get-FirstPropertyValue -InputObject $usage -Path @(
                        'UsageValue',
                        'Usage.Value',
                        'Usages.Value'
                    )
                )

                $remaining = $null
                $utilizationPercent = $null
                if ($null -ne $limit -and $null -ne $currentUsage) {
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

                if ($null -eq $limit) {
                    $status = 'Usage only'
                }
                elseif ($null -eq $currentUsage) {
                    $status = 'Limit only'
                }
                elseif ($limit -gt 0 -and $currentUsage -ge $limit) {
                    $status = 'Exhausted'
                }
                elseif ($null -ne $utilizationPercent -and
                    $utilizationPercent -ge $WarningThresholdPercent) {
                    $status = 'Warning'
                }
                else {
                    $status = 'OK'
                }

                $localizedName = Get-FirstPropertyValue -InputObject $quota -Path @(
                    'NameLocalizedValue',
                    'Name.LocalizedValue'
                )
                if ($null -eq $localizedName) {
                    $localizedName = Get-FirstPropertyValue -InputObject $usage -Path @(
                        'NameLocalizedValue',
                        'Name.LocalizedValue'
                    )
                }

                $unit = Get-FirstPropertyValue -InputObject $quota -Path @('Unit')
                if ($null -eq $unit) {
                    $unit = Get-FirstPropertyValue -InputObject $usage -Path @('Unit')
                }

                $isQuotaApplicable = Get-FirstPropertyValue -InputObject $quota -Path @(
                    'IsQuotaApplicable'
                )
                if ($null -eq $isQuotaApplicable) {
                    $isQuotaApplicable = Get-FirstPropertyValue -InputObject $usage -Path @(
                        'IsQuotaApplicable'
                    )
                }

                $quotaRowsByProvider[$provider].Add([pscustomobject] @{
                        SubscriptionName   = $subscription.Name
                        SubscriptionId     = $subscription.Id
                        TenantId           = $subscription.TenantId
                        Provider           = $provider
                        Location           = $azureLocation
                        QuotaName          = [string] $resourceName
                        QuotaDisplayName   = $localizedName
                        Unit               = $unit
                        CurrentUsage       = $currentUsage
                        Limit              = $limit
                        Remaining          = $remaining
                        UtilizationPercent = $utilizationPercent
                        Status             = $status
                        IsQuotaApplicable  = $isQuotaApplicable
                        CollectedAtUtc     = [datetime]::UtcNow
                    })
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
                LocationsQueried         = $subscriptionLocations.Count
                LocationsWithData        = $locationsWithData
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
