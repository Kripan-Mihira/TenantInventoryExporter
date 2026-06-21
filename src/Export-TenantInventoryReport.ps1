#Requires -Version 7.2
<#
.SYNOPSIS
    Exports read-only Microsoft 365 tenant inventory reports to CSV files.

.DESCRIPTION
    Collects Microsoft Entra ID, Intune, Exchange Online, SharePoint Online, OneDrive,
    Microsoft 365 Group, distribution group, mailbox, domain, and user profile-photo
    information. The script is intended for delegated, interactive administration.

    This script does not make changes to the tenant. It creates local files only.

.NOTES
    File Name  : Export-TenantInventoryReport.ps1
    Requires   : PowerShell 7.2+ on Windows
    Modules    : Microsoft.Graph, ExchangeOnlineManagement,
                 Microsoft.Online.SharePoint.PowerShell
    Important  : The SharePoint Online Management Shell is imported through Windows
                 PowerShell compatibility when this script is run in PowerShell 7.
                 The optional Excel workbook requires Microsoft Excel for Windows.

    The AzureAD module is intentionally not used because it is retired. Microsoft Graph
    PowerShell replaces AzureAD cmdlets in this implementation.

.EXAMPLE
    .\Export-TenantInventoryReport.ps1 -OutputPath 'D:\TenantReports'

.EXAMPLE
    .\Export-TenantInventoryReport.ps1 `
        -TenantDomain '<tenant>.onmicrosoft.com' `
        -SharePointAdminUrl 'https://<tenant>-admin.sharepoint.com' `
        -OutputPath '<OutputPath>' `
        -SkipUserPhotos `
        -SkipMasterWorkbook `
        -Verbose

.EXAMPLE
    # Displays intended file operations without connecting to services or writing files.
    .\Export-TenantInventoryReport.ps1 -OutputPath '<OutputPath>' -WhatIf
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path $PWD -ChildPath 'TenantReports'),

    [Parameter()]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9.-]*\.onmicrosoft\.com$')]
    [string]$TenantDomain,

    [Parameter()]
    [ValidatePattern('^https://.+-admin\.sharepoint\.(com|us|de|cn|us)$')]
    [string]$SharePointAdminUrl,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [switch]$SkipUserPhotos,

    [Parameter()]
    [switch]$SkipMasterWorkbook,

    [Parameter()]
    [switch]$SkipManagerLookup,

    [Parameter()]
    [switch]$NoProgress,

    [Parameter()]
    [switch]$EnableTranscript,

    [Parameter()]
    [switch]$StopOnReportError,

    [Parameter()]
    [ValidateRange(1, 10)]
    [int]$MaxRetryAttempts = 4,

    [Parameter()]
    [ValidateRange(1, 60)]
    [int]$InitialRetryDelaySeconds = 2
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:TenantReportState = @{
    LogPath          = $null
    ReportDirectory  = $null
    PhotoDirectory   = $null
    Cache            = @{}
    Manifest         = [System.Collections.Generic.List[object]]::new()
    ConnectedGraph   = $false
    ConnectedExchange = $false
    ConnectedSharePoint = $false
    TranscriptStarted = $false
    NoProgress       = $false
    MaxRetryAttempts = 4
    InitialRetryDelaySeconds = 2
}

function Write-TenantReportLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')]
        [string]$Level,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Message
    )

    $timestamp = (Get-Date).ToString('o')
    $entry = '{0} [{1}] {2}' -f $timestamp, $Level, $Message

    if ($script:TenantReportState.LogPath) {
        try {
            Add-Content -LiteralPath $script:TenantReportState.LogPath -Value $entry -Encoding utf8 -ErrorAction Stop
        }
        catch {
            # Logging must not hide the underlying report error.
            [Console]::Error.WriteLine("Unable to write to log file: $($_.Exception.Message)")
        }
    }

    switch ($Level) {
        'DEBUG' { Write-Verbose $entry }
        'INFO'  { Write-Verbose $entry }
        'WARN'  { Write-Warning $Message }
        'ERROR' { [Console]::Error.WriteLine($entry) }
    }
}

function Test-TenantReportCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter()]
        [string]$ModuleHint
    )

    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        $hint = if ($ModuleHint) { " Install or import module '$ModuleHint'." } else { '' }
        throw "Required command '$Name' is unavailable.$hint"
    }
}

function Import-TenantReportModules {
    [CmdletBinding()]
    param()

    $graphCommands = @(
        'Connect-MgGraph',
        'Disconnect-MgGraph',
        'Get-MgDomain',
        'Get-MgUser',
        'Get-MgGroup',
        'Get-MgDevice',
        'Get-MgDeviceManagementManagedDevice'
    )

    foreach ($command in $graphCommands) {
        Test-TenantReportCommand -Name $command -ModuleHint 'Microsoft.Graph'
    }

    Test-TenantReportCommand -Name 'Connect-ExchangeOnline' -ModuleHint 'ExchangeOnlineManagement'
    Test-TenantReportCommand -Name 'Disconnect-ExchangeOnline' -ModuleHint 'ExchangeOnlineManagement'
    Test-TenantReportCommand -Name 'Get-Mailbox' -ModuleHint 'ExchangeOnlineManagement'

    if (-not $IsWindows) {
        throw 'SharePoint Online Management Shell and Excel COM workbook generation require Windows. Run this full script on Windows.'
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        try {
            Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -ErrorAction Stop
        }
        catch {
            throw "Unable to import Microsoft.Online.SharePoint.PowerShell through Windows PowerShell compatibility. $($_.Exception.Message)"
        }
    }
    else {
        Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop
    }

    Test-TenantReportCommand -Name 'Connect-SPOService' -ModuleHint 'Microsoft.Online.SharePoint.PowerShell'
    Test-TenantReportCommand -Name 'Get-SPOSite' -ModuleHint 'Microsoft.Online.SharePoint.PowerShell'
}

function Initialize-TenantReportDirectories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseOutputPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantShortName,

        [Parameter()]
        [string]$RequestedLogPath
    )

    $tenantRoot = Join-Path -Path $BaseOutputPath -ChildPath $TenantShortName
    $reportDirectory = Join-Path -Path $tenantRoot -ChildPath 'Reports'
    $photoDirectory = Join-Path -Path $tenantRoot -ChildPath 'Photos'

    foreach ($directory in @($tenantRoot, $reportDirectory, $photoDirectory)) {
        if (-not (Test-Path -LiteralPath $directory)) {
            $null = New-Item -Path $directory -ItemType Directory -Force
        }
    }

    $script:TenantReportState.ReportDirectory = $reportDirectory
    $script:TenantReportState.PhotoDirectory = $photoDirectory
    $script:TenantReportState.LogPath = if ($RequestedLogPath) {
        $RequestedLogPath
    }
    else {
        Join-Path -Path $reportDirectory -ChildPath 'TenantInventoryReport.log'
    }

    $logDirectory = Split-Path -Path $script:TenantReportState.LogPath -Parent
    if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
        $null = New-Item -Path $logDirectory -ItemType Directory -Force
    }

    if (-not (Test-Path -LiteralPath $script:TenantReportState.LogPath)) {
        $null = New-Item -Path $script:TenantReportState.LogPath -ItemType File -Force
    }
}

function Test-IsTransientTenantReportError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $message = $ErrorRecord.Exception.Message
    return $message -match '(?i)(429|throttl|too many requests|503|504|temporar|timeout|timed out|connection.*reset|service unavailable)'
}

function Invoke-TenantReportReadOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Operation,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Description
    )

    $attempt = 0

    while ($attempt -lt $script:TenantReportState.MaxRetryAttempts) {
        $attempt++

        try {
            return & $Operation
        }
        catch {
            if (-not (Test-IsTransientTenantReportError -ErrorRecord $_) -or $attempt -ge $script:TenantReportState.MaxRetryAttempts) {
                throw
            }

            $delaySeconds = [Math]::Min(
                60,
                $script:TenantReportState.InitialRetryDelaySeconds * [Math]::Pow(2, $attempt - 1)
            )

            Write-TenantReportLog -Level 'WARN' -Message (
                "Transient error during '{0}'. Attempt {1}/{2}; retrying in {3} seconds. {4}" -f
                $Description,
                $attempt,
                $script:TenantReportState.MaxRetryAttempts,
                $delaySeconds,
                $_.Exception.Message
            )
            Start-Sleep -Seconds $delaySeconds
        }
    }
}

function Write-TenantReportProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Activity,

        [Parameter(Mandatory)]
        [int]$Current,

        [Parameter(Mandatory)]
        [int]$Total,

        [Parameter()]
        [string]$Status
    )

    if ($script:TenantReportState.NoProgress -or $Total -le 0) {
        return
    }

    $percent = [Math]::Min(100, [Math]::Round(($Current / $Total) * 100, 0))
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $percent
}

function Get-SafeTenantReportFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Fallback = 'Unknown'
    )

    $invalidCharacters = [System.IO.Path]::GetInvalidFileNameChars()
    $safeValue = $Value

    foreach ($character in $invalidCharacters) {
        $safeValue = $safeValue.Replace([string]$character, '_')
    }

    if ([string]::IsNullOrWhiteSpace($safeValue)) {
        return $Fallback
    }

    return $safeValue
}

function Convert-ByteSizeToGigabytes {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return 0
    }

    if ($Value -is [long] -or $Value -is [int] -or $Value -is [double] -or $Value -is [decimal]) {
        return [Math]::Round(([double]$Value / 1GB), 2)
    }

    $stringValue = [string]$Value

    if ($stringValue -match '\((?<bytes>[\d,]+)\s+bytes\)') {
        $bytes = [double](($Matches['bytes']) -replace ',', '')
        return [Math]::Round(($bytes / 1GB), 2)
    }

    if ($stringValue -match '(?<value>[\d.,]+)\s*(?<unit>KB|MB|GB|TB)') {
        $number = [double](($Matches['value']) -replace ',', '')
        switch ($Matches['unit'].ToUpperInvariant()) {
            'KB' { return [Math]::Round(($number / 1MB), 2) }
            'MB' { return [Math]::Round(($number / 1KB), 2) }
            'GB' { return [Math]::Round($number, 2) }
            'TB' { return [Math]::Round(($number * 1024), 2) }
        }
    }

    return $null
}

function Convert-QuotaToGigabytes {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Quota
    )

    if ($null -eq $Quota) {
        return $null
    }

    $quotaText = [string]$Quota
    if ($quotaText -match '(?i)unlimited') {
        return 'Unlimited'
    }

    return Convert-ByteSizeToGigabytes -Value $Quota
}

function Get-ProxyAddressRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$ProxyAddresses
    )

    foreach ($proxyAddress in $ProxyAddresses) {
        if ($null -eq $proxyAddress) {
            continue
        }

        $value = [string]$proxyAddress
        $parts = $value -split ':', 2

        [PSCustomObject]@{
            AddressType  = if ($parts.Count -gt 1) { $parts[0] } else { $null }
            EmailAddress = if ($parts.Count -gt 1) { $parts[1] } else { $value }
            IsPrimarySmtp = $value -cmatch '^SMTP:'
        }
    }
}

function Get-TenantReportPropertyValue {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

function Get-DirectoryObjectSummary {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$DirectoryObject
    )

    $result = [ordered]@{
        Id                = $null
        DisplayName       = $null
        UserPrincipalName = $null
        UserType          = $null
        ObjectType        = $null
    }

    if ($null -eq $DirectoryObject) {
        return [PSCustomObject]$result
    }

    $result.Id = Get-TenantReportPropertyValue -InputObject $DirectoryObject -Name 'Id'
    $result.DisplayName = Get-TenantReportPropertyValue -InputObject $DirectoryObject -Name 'DisplayName'
    $result.UserPrincipalName = Get-TenantReportPropertyValue -InputObject $DirectoryObject -Name 'UserPrincipalName'
    $result.UserType = Get-TenantReportPropertyValue -InputObject $DirectoryObject -Name 'UserType'

    $additionalProperties = Get-TenantReportPropertyValue -InputObject $DirectoryObject -Name 'AdditionalProperties'
    if ($additionalProperties -and $additionalProperties.ContainsKey('@odata.type')) {
        $result.ObjectType = $additionalProperties['@odata.type']
    }

    if ($result.Id -and $script:TenantReportState.Cache.ContainsKey('UsersById') -and
        $script:TenantReportState.Cache.UsersById.ContainsKey($result.Id)) {
        $user = $script:TenantReportState.Cache.UsersById[$result.Id]
        $result.DisplayName = $user.DisplayName
        $result.UserPrincipalName = $user.UserPrincipalName
        $result.UserType = $user.UserType
        $result.ObjectType = '#microsoft.graph.user'
    }

    if ($result.Id -and $script:TenantReportState.Cache.ContainsKey('GraphGroupsById') -and
        $script:TenantReportState.Cache.GraphGroupsById.ContainsKey($result.Id)) {
        $group = $script:TenantReportState.Cache.GraphGroupsById[$result.Id]
        $result.DisplayName = $group.DisplayName
        $result.ObjectType = '#microsoft.graph.group'
    }

    if ([string]::IsNullOrWhiteSpace($result.DisplayName) -and $additionalProperties) {
        $result.DisplayName = $additionalProperties['displayName']
        $result.UserPrincipalName = $additionalProperties['userPrincipalName']
    }

    return [PSCustomObject]$result
}

function Get-UserManagerSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )

    if ($script:TenantReportState.Cache.ManagerByUserId.ContainsKey($UserId)) {
        return $script:TenantReportState.Cache.ManagerByUserId[$UserId]
    }

    $managerSummary = [PSCustomObject]@{
        DisplayName       = $null
        UserPrincipalName = $null
    }

    try {
        # A 404 here usually means that the user has no manager; it is not a report failure.
        $manager = Get-MgUserManager -UserId $UserId -ErrorAction Stop
        $resolvedManager = Get-DirectoryObjectSummary -DirectoryObject $manager

        $managerSummary = [PSCustomObject]@{
            DisplayName       = $resolvedManager.DisplayName
            UserPrincipalName = $resolvedManager.UserPrincipalName
        }
    }
    catch {
        Write-Verbose "Manager is unavailable for user ID '$UserId'."
    }

    $script:TenantReportState.Cache.ManagerByUserId[$UserId] = $managerSummary
    return $managerSummary
}

function Disconnect-TenantReportServices {
    [CmdletBinding()]
    param()

    if ($script:TenantReportState.ConnectedSharePoint -and (Get-Command -Name Disconnect-SPOService -ErrorAction SilentlyContinue)) {
        try {
            Disconnect-SPOService -ErrorAction Stop
        }
        catch {
            Write-TenantReportLog -Level 'WARN' -Message "Unable to disconnect from SharePoint Online cleanly. $($_.Exception.Message)"
        }
    }

    if ($script:TenantReportState.ConnectedExchange) {
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        }
        catch {
            Write-TenantReportLog -Level 'WARN' -Message "Unable to disconnect from Exchange Online cleanly. $($_.Exception.Message)"
        }
    }

    if ($script:TenantReportState.ConnectedGraph) {
        try {
            Disconnect-MgGraph -ErrorAction Stop
        }
        catch {
            Write-TenantReportLog -Level 'WARN' -Message "Unable to disconnect from Microsoft Graph cleanly. $($_.Exception.Message)"
        }
    }
}

function Initialize-TenantReportCache {
    [CmdletBinding()]
    param()

    Write-TenantReportLog -Level 'INFO' -Message 'Retrieving reusable tenant inventory data.'

    $userProperties = @(
        'id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'employeeId',
        'jobTitle', 'mobilePhone', 'department', 'companyName', 'officeLocation',
        'streetAddress', 'city', 'postalCode', 'state', 'country', 'accountEnabled',
        'createdDateTime', 'userType', 'mail', 'onPremisesSyncEnabled', 'assignedLicenses'
    )

    $groupProperties = @(
        'id', 'displayName', 'description', 'createdDateTime', 'visibility',
        'securityEnabled', 'mailEnabled', 'groupTypes', 'membershipRule',
        'mail', 'mailNickname', 'resourceProvisioningOptions'
    )

    $script:TenantReportState.Cache.Users = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Entra users' -Operation {
            Get-MgUser -All -Property $userProperties -ErrorAction Stop
        }
    )

    $script:TenantReportState.Cache.UsersById = @{}
    foreach ($user in $script:TenantReportState.Cache.Users) {
        $script:TenantReportState.Cache.UsersById[$user.Id] = $user
    }

    $script:TenantReportState.Cache.ManagerByUserId = @{}

    $script:TenantReportState.Cache.GraphGroups = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Entra groups' -Operation {
            Get-MgGroup -All -Property $groupProperties -ErrorAction Stop
        }
    )

    $script:TenantReportState.Cache.GraphGroupsById = @{}
    foreach ($group in $script:TenantReportState.Cache.GraphGroups) {
        $script:TenantReportState.Cache.GraphGroupsById[$group.Id] = $group
    }

    $script:TenantReportState.Cache.Mailboxes = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Exchange mailboxes' -Operation {
            Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        }
    )

    $script:TenantReportState.Cache.DistributionGroups = @(
        Invoke-TenantReportReadOperation -Description 'retrieve distribution groups' -Operation {
            Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
        }
    )

    $script:TenantReportState.Cache.UnifiedGroups = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Microsoft 365 groups' -Operation {
            Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop
        }
    )
}

function Export-TenantReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^[A-Za-z0-9_&-]+$')]
        [string]$ReportName,

        [Parameter(Mandatory)]
        [scriptblock]$DataScript,

        [Parameter()]
        [switch]$StopOnError
    )

    $destination = Join-Path -Path $script:TenantReportState.ReportDirectory -ChildPath "$ReportName.csv"
    $startTime = Get-Date

    try {
        Write-TenantReportLog -Level 'INFO' -Message "Generating report '$ReportName'."

        # Export-Csv receives the stream directly, avoiding expensive array concatenation.
        & $DataScript | Export-Csv -LiteralPath $destination -NoTypeInformation -Encoding utf8 -Force

        if (-not (Test-Path -LiteralPath $destination)) {
            $null = New-Item -Path $destination -ItemType File -Force
        }

        $script:TenantReportState.Manifest.Add([PSCustomObject]@{
            ReportName   = $ReportName
            Status       = 'Completed'
            Path         = $destination
            DurationSec  = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
            ErrorMessage = $null
        })

        Write-TenantReportLog -Level 'INFO' -Message "Completed report '$ReportName'."
    }
    catch {
        $message = $_.Exception.Message

        $script:TenantReportState.Manifest.Add([PSCustomObject]@{
            ReportName   = $ReportName
            Status       = 'Failed'
            Path         = $destination
            DurationSec  = [Math]::Round(((Get-Date) - $startTime).TotalSeconds, 2)
            ErrorMessage = $message
        })

        Write-TenantReportLog -Level 'ERROR' -Message "Report '$ReportName' failed. $message"

        if ($StopOnError) {
            throw
        }
    }
}

function Get-UserInventoryData {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$SkipManagerLookup
    )

    $users = $script:TenantReportState.Cache.Users
    $total = $users.Count
    $index = 0

    foreach ($user in $users) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting user inventory' -Current $index -Total $total -Status $user.UserPrincipalName

        $manager = if ($SkipManagerLookup) {
            [PSCustomObject]@{ DisplayName = $null; UserPrincipalName = $null }
        }
        else {
            Get-UserManagerSummary -UserId $user.Id
        }

        [PSCustomObject]@{
            ObjectId                 = $user.Id
            UserPrincipalName        = $user.UserPrincipalName
            DisplayName              = $user.DisplayName
            FirstName                = $user.GivenName
            LastName                 = $user.Surname
            EmployeeId               = $user.EmployeeId
            ManagerDisplayName       = $manager.DisplayName
            ManagerUserPrincipalName = $manager.UserPrincipalName
            EmailAddress             = $user.Mail
            AccountStatus            = if ($user.AccountEnabled) { 'Enabled' } else { 'Disabled' }
            SyncStatus               = if ($user.OnPremisesSyncEnabled -eq $true) { 'Synced' } else { 'Cloud Only' }
            UserType                 = $user.UserType
            JobTitle                 = $user.JobTitle
            Mobile                   = $user.MobilePhone
            Licensed                 = if (@($user.AssignedLicenses).Count -gt 0) { 'Yes' } else { 'No' }
            Department               = $user.Department
            Company                  = $user.CompanyName
            Office                   = $user.OfficeLocation
            Street                   = $user.StreetAddress
            City                     = $user.City
            PostalCode               = $user.PostalCode
            State                    = $user.State
            Country                  = $user.Country
            AccountCreatedOn         = $user.CreatedDateTime
        }
    }

    Write-Progress -Activity 'Exporting user inventory' -Completed
}

function Get-EntraDeviceData {
    [CmdletBinding()]
    param()

    $deviceProperties = @(
        'id', 'deviceId', 'displayName', 'accountEnabled', 'isCompliant', 'trustType',
        'approximateLastSignInDateTime', 'manufacturer', 'model', 'operatingSystem',
        'operatingSystemVersion', 'deviceOwnership', 'managementType', 'profileType',
        'registrationDateTime'
    )

    $devices = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Entra devices' -Operation {
            Get-MgDevice -All -Property $deviceProperties -ErrorAction Stop
        }
    )

    $total = $devices.Count
    $index = 0

    foreach ($device in $devices) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting Entra device inventory' -Current $index -Total $total -Status $device.DisplayName

        $registeredUsers = @(
            Invoke-TenantReportReadOperation -Description "retrieve registered users for device '$($device.Id)'" -Operation {
                Get-MgDeviceRegisteredUser -DeviceId $device.Id -All -ErrorAction Stop
            }
        )

        $registeredOwners = @(
            Invoke-TenantReportReadOperation -Description "retrieve registered owners for device '$($device.Id)'" -Operation {
                Get-MgDeviceRegisteredOwner -DeviceId $device.Id -All -ErrorAction Stop
            }
        )

        $directMemberships = @(
            Invoke-TenantReportReadOperation -Description "retrieve direct group memberships for device '$($device.Id)'" -Operation {
                Get-MgDeviceMemberOf -DeviceId $device.Id -All -ErrorAction Stop
            }
        )

        $userInfo = @($registeredUsers | ForEach-Object { Get-DirectoryObjectSummary -DirectoryObject $_ })
        $ownerInfo = @($registeredOwners | ForEach-Object { Get-DirectoryObjectSummary -DirectoryObject $_ })

        # memberOf can include groups and administrative units. Export only groups as intended by the original report.
        $groups = foreach ($membership in $directMemberships) {
            if ($script:TenantReportState.Cache.GraphGroupsById.ContainsKey($membership.Id)) {
                $script:TenantReportState.Cache.GraphGroupsById[$membership.Id]
            }
        }

        if (-not $groups) {
            $groups = @($null)
        }

        foreach ($group in $groups) {
            [PSCustomObject]@{
                DeviceName             = $device.DisplayName
                DeviceId               = $device.DeviceId
                DeviceObjectId         = $device.Id
                DeviceStatus           = if ($device.AccountEnabled) { 'Enabled' } else { 'Disabled' }
                UserId                 = ($userInfo.Id -join ', ')
                UserDisplayName        = ($userInfo.DisplayName -join ', ')
                UserPrincipalName      = ($userInfo.UserPrincipalName -join ', ')
                OwnerId                = ($ownerInfo.Id -join ', ')
                OwnerDisplayName       = ($ownerInfo.DisplayName -join ', ')
                OwnerPrincipalName     = ($ownerInfo.UserPrincipalName -join ', ')
                GroupId                = if ($group) { $group.Id } else { 'No Groups' }
                GroupName              = if ($group) { $group.DisplayName } else { 'No Groups' }
                ComplianceStatus       = $device.IsCompliant
                JoinType               = $device.TrustType
                LastSignInDateTime     = $device.ApproximateLastSignInDateTime
                Manufacturer           = $device.Manufacturer
                Model                  = $device.Model
                OperatingSystem        = $device.OperatingSystem
                OperatingSystemVersion = $device.OperatingSystemVersion
                DeviceOwnership        = $device.DeviceOwnership
                ManagementType         = $device.ManagementType
                ProfileType            = $device.ProfileType
                RegistrationDateTime   = $device.RegistrationDateTime
            }
        }
    }

    Write-Progress -Activity 'Exporting Entra device inventory' -Completed
}

function Get-ProxyAddressReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Recipients,

        [Parameter(Mandatory)]
        [ValidateSet('DistributionGroup', 'Microsoft365Group', 'Mailbox')]
        [string]$RecipientCategory
    )

    $total = $Recipients.Count
    $index = 0

    foreach ($recipient in $Recipients) {
        $index++
        Write-TenantReportProgress -Activity "Exporting $RecipientCategory proxy addresses" -Current $index -Total $total -Status $recipient.DisplayName

        foreach ($proxy in Get-ProxyAddressRecord -ProxyAddresses @($recipient.EmailAddresses)) {
            [PSCustomObject]@{
                Name                 = $recipient.Name
                DisplayName          = $recipient.DisplayName
                RecipientTypeDetails = $recipient.RecipientTypeDetails
                UserPrincipalName    = if ($RecipientCategory -eq 'Mailbox') { $recipient.UserPrincipalName } else { $null }
                PrimarySmtpAddress   = $recipient.PrimarySmtpAddress
                AddressType          = $proxy.AddressType
                EmailAddress         = $proxy.EmailAddress
                IsPrimarySmtp        = $proxy.IsPrimarySmtp
            }
        }
    }

    Write-Progress -Activity "Exporting $RecipientCategory proxy addresses" -Completed
}

function Get-MailboxPermissionData {
    [CmdletBinding()]
    param()

    $mailboxes = $script:TenantReportState.Cache.Mailboxes
    $total = $mailboxes.Count
    $index = 0

    foreach ($mailbox in $mailboxes) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting mailbox permissions' -Current $index -Total $total -Status $mailbox.DisplayName

        $fullAccessPermissions = @(
            Invoke-TenantReportReadOperation -Description "retrieve Full Access permissions for '$($mailbox.UserPrincipalName)'" -Operation {
                Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop |
                    Where-Object {
                        $_.AccessRights -contains 'FullAccess' -and
                        -not $_.IsInherited -and
                        ([string]$_.User) -ne 'NT AUTHORITY\SELF'
                    }
            }
        )

        foreach ($permission in $fullAccessPermissions) {
            [PSCustomObject]@{
                DisplayName          = $mailbox.DisplayName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                UserPrincipalName    = $mailbox.UserPrincipalName
                AccessRights         = 'Full Access'
                User                 = $permission.User
                IsInherited          = $permission.IsInherited
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
            }
        }

        $sendAsPermissions = @(
            Invoke-TenantReportReadOperation -Description "retrieve Send As permissions for '$($mailbox.UserPrincipalName)'" -Operation {
                Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction Stop |
                    Where-Object {
                        $_.AccessRights -contains 'SendAs' -and
                        ([string]$_.Trustee) -ne 'NT AUTHORITY\SELF'
                    }
            }
        )

        foreach ($permission in $sendAsPermissions) {
            [PSCustomObject]@{
                DisplayName          = $mailbox.DisplayName
                PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                UserPrincipalName    = $mailbox.UserPrincipalName
                AccessRights         = 'Send As'
                User                 = $permission.Trustee
                IsInherited          = $permission.IsInherited
                RecipientTypeDetails = $mailbox.RecipientTypeDetails
            }
        }

        foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
            try {
                $recipient = Get-Recipient -Identity $delegate -ErrorAction Stop

                [PSCustomObject]@{
                    DisplayName          = $mailbox.DisplayName
                    PrimarySmtpAddress   = $mailbox.PrimarySmtpAddress
                    UserPrincipalName    = $mailbox.UserPrincipalName
                    AccessRights         = 'Send on Behalf'
                    User                 = $recipient.PrimarySmtpAddress
                    IsInherited          = $false
                    RecipientTypeDetails = $mailbox.RecipientTypeDetails
                }
            }
            catch {
                Write-TenantReportLog -Level 'WARN' -Message (
                    "Unable to resolve Send on Behalf delegate '$delegate' for mailbox '$($mailbox.UserPrincipalName)'. $($_.Exception.Message)"
                )
            }
        }
    }

    Write-Progress -Activity 'Exporting mailbox permissions' -Completed
}

function Get-LegacyExchangeDnData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Recipients,

        [Parameter(Mandatory)]
        [ValidateSet('Mailbox', 'DistributionGroup', 'Microsoft365Group')]
        [string]$RecipientCategory
    )

    foreach ($recipient in $Recipients) {
        [PSCustomObject]@{
            DisplayName        = $recipient.DisplayName
            PrimarySmtpAddress = $recipient.PrimarySmtpAddress
            UserPrincipalName  = if ($RecipientCategory -eq 'Mailbox') { $recipient.UserPrincipalName } else { $null }
            LegacyExchangeDn   = $recipient.LegacyExchangeDn
        }
    }
}

function Get-IntuneManagedDeviceData {
    [CmdletBinding()]
    param()

    $properties = @(
        'id', 'deviceName', 'userDisplayName', 'userPrincipalName', 'emailAddress', 'userId',
        'operatingSystem', 'osVersion', 'model', 'manufacturer', 'imei',
        'totalStorageSpaceInBytes', 'freeStorageSpaceInBytes', 'enrolledDateTime',
        'complianceState', 'managedDeviceOwnerType', 'managementAgent', 'azureADRegistered',
        'azureADDeviceId', 'isEncrypted', 'serialNumber', 'deviceRegistrationState'
    )

    $devices = @(
        Invoke-TenantReportReadOperation -Description 'retrieve Intune managed devices' -Operation {
            Get-MgDeviceManagementManagedDevice -All -Property $properties -ErrorAction Stop
        }
    )

    $total = $devices.Count
    $index = 0

    foreach ($device in $devices) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting Intune managed devices' -Current $index -Total $total -Status $device.DeviceName

        [PSCustomObject]@{
            DeviceName              = $device.DeviceName
            DeviceId                = $device.Id
            UserDisplayName         = $device.UserDisplayName
            UserPrincipalName       = $device.UserPrincipalName
            UserEmailAddress        = $device.EmailAddress
            UserId                  = $device.UserId
            OperatingSystem         = $device.OperatingSystem
            OsVersion               = $device.OsVersion
            DeviceModel             = $device.Model
            DeviceManufacturer      = $device.Manufacturer
            DeviceImeiNumber        = $device.Imei
            TotalSpaceInGb          = if ($null -ne $device.TotalStorageSpaceInBytes) { [Math]::Round(($device.TotalStorageSpaceInBytes / 1GB), 2) } else { $null }
            FreeSpaceInGb           = if ($null -ne $device.FreeStorageSpaceInBytes) { [Math]::Round(($device.FreeStorageSpaceInBytes / 1GB), 2) } else { $null }
            EnrolledDateTime        = $device.EnrolledDateTime
            ComplianceStatus        = $device.ComplianceState
            DeviceOwnerType         = $device.ManagedDeviceOwnerType
            DeviceManagementAgent   = $device.ManagementAgent
            AzureAdDeviceRegistered = $device.AzureAdRegistered
            AzureAdDeviceId         = $device.AzureAdDeviceId
            IsEncrypted             = $device.IsEncrypted
            DeviceSerialNumber      = $device.SerialNumber
            DeviceState             = $device.DeviceRegistrationState
        }
    }

    Write-Progress -Activity 'Exporting Intune managed devices' -Completed
}

function Get-UserLicenseData {
    [CmdletBinding()]
    param()

    $subscribedSkus = @(
        Invoke-TenantReportReadOperation -Description 'retrieve subscribed SKUs' -Operation {
            Get-MgSubscribedSku -All -ErrorAction Stop
        }
    )

    $skuMap = @{}
    foreach ($sku in $subscribedSkus) {
        $skuMap[[string]$sku.SkuId] = $sku.SkuPartNumber
    }

    foreach ($user in $script:TenantReportState.Cache.Users | Where-Object { @($_.AssignedLicenses).Count -gt 0 }) {
        $licenses = foreach ($assignedLicense in @($user.AssignedLicenses)) {
            $skuId = [string]$assignedLicense.SkuId
            if ($skuMap.ContainsKey($skuId)) {
                $skuMap[$skuId]
            }
            else {
                "UnknownSku:$skuId"
            }
        }

        [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            StreetAddress     = $user.StreetAddress
            Licenses          = ($licenses -join ', ')
        }
    }
}

function Get-SelfServicePasswordResetData {
    [CmdletBinding()]
    param()

    $users = $script:TenantReportState.Cache.Users
    $total = $users.Count
    $index = 0

    foreach ($user in $users) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting SSPR methods' -Current $index -Total $total -Status $user.UserPrincipalName

        $emailMethods = @(
            Invoke-TenantReportReadOperation -Description "retrieve email authentication method for '$($user.UserPrincipalName)'" -Operation {
                Get-MgUserAuthenticationEmailMethod -UserId $user.Id -ErrorAction Stop
            }
        )

        $phoneMethods = @(
            Invoke-TenantReportReadOperation -Description "retrieve phone authentication method for '$($user.UserPrincipalName)'" -Operation {
                Get-MgUserAuthenticationPhoneMethod -UserId $user.Id -ErrorAction Stop
            }
        )

        $emails = @($emailMethods | ForEach-Object { $_.EmailAddress } | Where-Object { $_ })
        $phones = @($phoneMethods | ForEach-Object { $_.PhoneNumber } | Where-Object { $_ })

        if ($emails.Count -eq 0) { $emails = @('No Email') }
        if ($phones.Count -eq 0) { $phones = @('No Phone') }

        foreach ($email in $emails) {
            foreach ($phone in $phones) {
                [PSCustomObject]@{
                    DisplayName       = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    Email             = $email
                    Phone             = $phone
                }
            }
        }
    }

    Write-Progress -Activity 'Exporting SSPR methods' -Completed
}

function Test-ArchiveEnabled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    if ($Mailbox.ArchiveStatus -eq 'Active') {
        return $true
    }

    if ($Mailbox.ArchiveGuid -and ([string]$Mailbox.ArchiveGuid -ne '00000000-0000-0000-0000-000000000000')) {
        return $true
    }

    return $false
}

function Get-MailboxSizeData {
    [CmdletBinding()]
    param()

    $mailboxes = $script:TenantReportState.Cache.Mailboxes
    $total = $mailboxes.Count
    $index = 0

    foreach ($mailbox in $mailboxes) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting mailbox and archive sizes' -Current $index -Total $total -Status $mailbox.UserPrincipalName

        $mailboxStatistics = Invoke-TenantReportReadOperation -Description "retrieve statistics for '$($mailbox.UserPrincipalName)'" -Operation {
            Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
        }

        $archiveEnabled = Test-ArchiveEnabled -Mailbox $mailbox
        $archiveStatistics = $null

        if ($archiveEnabled) {
            try {
                $archiveStatistics = Invoke-TenantReportReadOperation -Description "retrieve archive statistics for '$($mailbox.UserPrincipalName)'" -Operation {
                    Get-MailboxStatistics -Identity $mailbox.Identity -Archive -ErrorAction Stop
                }
            }
            catch {
                Write-TenantReportLog -Level 'WARN' -Message "Archive statistics are unavailable for '$($mailbox.UserPrincipalName)'. $($_.Exception.Message)"
            }
        }

        [PSCustomObject]@{
            UserName                    = $mailbox.DisplayName
            UserPrincipalName           = $mailbox.UserPrincipalName
            PrimarySmtpAddress          = $mailbox.PrimarySmtpAddress
            ObjectId                    = $mailbox.ExchangeGuid
            MailboxType                 = $mailbox.RecipientTypeDetails
            MailboxAllocationSizeInGb   = Convert-QuotaToGigabytes -Quota $mailbox.ProhibitSendReceiveQuota
            MailboxUsageSizeInGb        = Convert-ByteSizeToGigabytes -Value $mailboxStatistics.TotalItemSize
            MailboxItemCount            = $mailboxStatistics.ItemCount
            LastUserActionTime          = $mailboxStatistics.LastUserActionTime
            LastInteractionTime         = $mailboxStatistics.LastInteractionTime
            LastLogonTime               = $mailboxStatistics.LastLogonTime
            ArchiveEnabled              = if ($archiveEnabled) { 'Enabled' } else { 'Disabled' }
            ArchiveName                 = if ($archiveEnabled) { ($mailbox.ArchiveName -join '; ') } else { $null }
            ArchiveAllocationSizeInGb   = if ($archiveEnabled) { Convert-QuotaToGigabytes -Quota $mailbox.ArchiveQuota } else { $null }
            ArchiveUsageSizeInGb        = if ($archiveStatistics) { Convert-ByteSizeToGigabytes -Value $archiveStatistics.TotalItemSize } else { $null }
            ArchiveItemCount            = if ($archiveStatistics) { $archiveStatistics.ItemCount } else { $null }
            AutoExpandingArchiveEnabled = $mailbox.AutoExpandingArchiveEnabled
        }
    }

    Write-Progress -Activity 'Exporting mailbox and archive sizes' -Completed
}

function Get-DistributionGroupDetailData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.DistributionGroups) {
        [PSCustomObject]@{
            GroupDisplayName  = $group.DisplayName
            GroupEmailAddress = $group.PrimarySmtpAddress
            GroupType         = $group.RecipientTypeDetails
            ObjectId          = $group.Guid
        }
    }
}

function Get-DistributionGroupMemberData {
    [CmdletBinding()]
    param()

    $groups = $script:TenantReportState.Cache.DistributionGroups
    $total = $groups.Count
    $index = 0

    foreach ($group in $groups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting distribution group members' -Current $index -Total $total -Status $group.DisplayName

        $members = @(
            Invoke-TenantReportReadOperation -Description "retrieve members for distribution group '$($group.DisplayName)'" -Operation {
                Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop
            }
        )

        foreach ($member in $members) {
            [PSCustomObject]@{
                GroupName          = $group.DisplayName
                GroupEmailAddress  = $group.PrimarySmtpAddress
                GroupType          = $group.RecipientTypeDetails
                Member             = $member.DisplayName
                EmailAddress       = $member.PrimarySmtpAddress
                RecipientType      = $member.RecipientTypeDetails
            }
        }
    }

    Write-Progress -Activity 'Exporting distribution group members' -Completed
}

function Get-Microsoft365GroupDetailData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.UnifiedGroups) {
        [PSCustomObject]@{
            GroupDisplayName  = $group.DisplayName
            GroupEmailAddress = $group.PrimarySmtpAddress
            GroupType         = $group.GroupType
            ObjectId          = $group.ExchangeGuid
            AccessType        = $group.AccessType
            HasTeams          = if (@($group.ResourceProvisioningOptions) -contains 'Team') { 'WithTeams' } else { 'WithoutTeams' }
            CreationDate      = $group.WhenCreated
        }
    }
}

function Get-Microsoft365GroupCalendarData {
    [CmdletBinding()]
    param()

    $groups = $script:TenantReportState.Cache.UnifiedGroups
    $total = $groups.Count
    $index = 0

    foreach ($group in $groups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting Microsoft 365 group calendar counts' -Current $index -Total $total -Status $group.DisplayName

        $folderStatistics = @(
            Invoke-TenantReportReadOperation -Description "retrieve calendar statistics for Microsoft 365 group '$($group.DisplayName)'" -Operation {
                Get-MailboxFolderStatistics -Identity $group.Identity -FolderScope Calendar -ErrorAction Stop
            }
        )

        $calendarItemCount = ($folderStatistics | Measure-Object -Property ItemsInFolder -Sum).Sum

        [PSCustomObject]@{
            Name                  = $group.DisplayName
            PrimarySmtpAddress    = $group.PrimarySmtpAddress
            GroupType             = $group.GroupType
            AccessType            = $group.AccessType
            HasTeams              = if (@($group.ResourceProvisioningOptions) -contains 'Team') { 'WithTeams' } else { 'WithoutTeams' }
            CalendarItemsInFolder = $calendarItemCount
        }
    }

    Write-Progress -Activity 'Exporting Microsoft 365 group calendar counts' -Completed
}

function Get-OneDriveSizeData {
    [CmdletBinding()]
    param()

    $oneDrives = @(
        Invoke-TenantReportReadOperation -Description 'retrieve OneDrive sites' -Operation {
            Get-SPOSite -IncludePersonalSite $true -Limit All -Detailed -ErrorAction Stop |
                Where-Object { $_.Url -match '-my\.sharepoint\.(com|us|de|cn)/personal/' }
        }
    )

    $total = $oneDrives.Count
    $index = 0

    foreach ($oneDrive in $oneDrives) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting OneDrive size report' -Current $index -Total $total -Status $oneDrive.Title

        [PSCustomObject]@{
            Url                = $oneDrive.Url
            DisplayName        = $oneDrive.Title
            UserPrincipalName  = $oneDrive.Owner
            Status             = $oneDrive.Status
            UsedSpaceInGb      = [Math]::Round(($oneDrive.StorageUsageCurrent / 1024), 2)
            AllocatedSpaceInGb = [Math]::Round(($oneDrive.StorageQuota / 1024), 2)
            LastModifiedDate   = $oneDrive.LastContentModifiedDate
        }
    }

    Write-Progress -Activity 'Exporting OneDrive size report' -Completed
}

function Get-MailboxInventoryData {
    [CmdletBinding()]
    param()

    $mailboxes = $script:TenantReportState.Cache.Mailboxes
    $total = $mailboxes.Count
    $index = 0

    foreach ($mailbox in $mailboxes) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting mailbox inventory' -Current $index -Total $total -Status $mailbox.UserPrincipalName

        $statistics = Invoke-TenantReportReadOperation -Description "retrieve mailbox inventory statistics for '$($mailbox.UserPrincipalName)'" -Operation {
            Get-MailboxStatistics -Identity $mailbox.Identity -ErrorAction Stop
        }

        [PSCustomObject]@{
            DisplayName         = $mailbox.DisplayName
            PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
            UserPrincipalName   = $mailbox.UserPrincipalName
            ObjectId            = $mailbox.Guid
            MailboxCreationDate = $mailbox.WhenMailboxCreated
            LastUserActionTime  = $statistics.LastUserActionTime
            LastInteractionTime = $statistics.LastInteractionTime
            LastLogonTime       = $statistics.LastLogonTime
            LitigationStatus    = if ($mailbox.LitigationHoldEnabled) { 'Enabled' } else { 'Disabled' }
            MailboxTypeDetails  = $mailbox.RecipientTypeDetails
        }
    }

    Write-Progress -Activity 'Exporting mailbox inventory' -Completed
}

function Get-DeletedUserData {
    [CmdletBinding()]
    param()

    $deletedUsers = @(
        Invoke-TenantReportReadOperation -Description 'retrieve deleted users' -Operation {
            Get-MgDirectoryDeletedItemAsUser -All -Property 'id,displayName,userPrincipalName,deletedDateTime' -ErrorAction Stop
        }
    )

    foreach ($user in $deletedUsers) {
        [PSCustomObject]@{
            DisplayName       = $user.DisplayName
            ObjectId          = $user.Id
            UserPrincipalName = $user.UserPrincipalName
            DeletedDateTime   = $user.DeletedDateTime
        }
    }
}

function Get-SecurityGroupDetailData {
    [CmdletBinding()]
    param()

    $securityGroups = @(
        $script:TenantReportState.Cache.GraphGroups |
            Where-Object { $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false }
    )

    foreach ($group in $securityGroups) {
        [PSCustomObject]@{
            DisplayName     = $group.DisplayName
            ObjectId        = $group.Id
            Description     = $group.Description
            CreatedDateTime = $group.CreatedDateTime
            Visibility      = $group.Visibility
        }
    }
}

function Get-SecurityGroupMemberData {
    [CmdletBinding()]
    param()

    $securityGroups = @(
        $script:TenantReportState.Cache.GraphGroups |
            Where-Object { $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false }
    )

    $total = $securityGroups.Count
    $index = 0

    foreach ($group in $securityGroups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting security group members' -Current $index -Total $total -Status $group.DisplayName

        $members = @(
            Invoke-TenantReportReadOperation -Description "retrieve members for security group '$($group.DisplayName)'" -Operation {
                Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop
            }
        )

        foreach ($member in $members) {
            $summary = Get-DirectoryObjectSummary -DirectoryObject $member

            [PSCustomObject]@{
                GroupName          = $group.DisplayName
                GroupId            = $group.Id
                MemberName         = $summary.DisplayName
                MemberEmailAddress = $summary.UserPrincipalName
                MemberType         = $summary.ObjectType
                MemberUserType     = $summary.UserType
            }
        }
    }

    Write-Progress -Activity 'Exporting security group members' -Completed
}

function Get-Microsoft365GroupMemberData {
    [CmdletBinding()]
    param()

    $groups = $script:TenantReportState.Cache.UnifiedGroups
    $total = $groups.Count
    $index = 0

    foreach ($group in $groups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting Microsoft 365 group members' -Current $index -Total $total -Status $group.DisplayName

        $members = @(
            Invoke-TenantReportReadOperation -Description "retrieve members for Microsoft 365 group '$($group.DisplayName)'" -Operation {
                Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Member -ResultSize Unlimited -ErrorAction Stop
            }
        )

        foreach ($member in $members) {
            [PSCustomObject]@{
                GroupName          = $group.DisplayName
                GroupMail          = $group.PrimarySmtpAddress
                GroupType          = $group.GroupType
                MemberName         = $member.DisplayName
                MemberEmailAddress = $member.PrimarySmtpAddress
                RecipientType      = $member.RecipientTypeDetails
            }
        }
    }

    Write-Progress -Activity 'Exporting Microsoft 365 group members' -Completed
}

function Get-DistributionGroupOwnerData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.DistributionGroups) {
        $ownerAddresses = foreach ($owner in @($group.ManagedBy)) {
            try {
                (Get-Recipient -Identity $owner -ErrorAction Stop).PrimarySmtpAddress
            }
            catch {
                Write-TenantReportLog -Level 'WARN' -Message "Unable to resolve owner '$owner' for distribution group '$($group.DisplayName)'. $($_.Exception.Message)"
            }
        }

        [PSCustomObject]@{
            DisplayName           = $group.DisplayName
            PrimarySmtpAddress    = $group.PrimarySmtpAddress
            RecipientTypeDetails  = $group.RecipientTypeDetails
            OwnerEmailAddress     = ($ownerAddresses -join ', ')
        }
    }
}

function Get-Microsoft365GroupOwnerData {
    [CmdletBinding()]
    param()

    $groups = $script:TenantReportState.Cache.UnifiedGroups
    $total = $groups.Count
    $index = 0

    foreach ($group in $groups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting Microsoft 365 group owners' -Current $index -Total $total -Status $group.DisplayName

        $owners = @(
            Invoke-TenantReportReadOperation -Description "retrieve owners for Microsoft 365 group '$($group.DisplayName)'" -Operation {
                Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Owner -ResultSize Unlimited -ErrorAction Stop
            }
        )

        foreach ($owner in $owners) {
            [PSCustomObject]@{
                GroupName          = $group.DisplayName
                GroupMail          = $group.PrimarySmtpAddress
                GroupType          = $group.GroupType
                OwnerName          = $owner.DisplayName
                OwnerEmailAddress  = $owner.PrimarySmtpAddress
                RecipientType      = $owner.RecipientTypeDetails
            }
        }
    }

    Write-Progress -Activity 'Exporting Microsoft 365 group owners' -Completed
}

function Get-DistributionGroupSettingsData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.DistributionGroups) {
        [PSCustomObject]@{
            DisplayName                      = $group.DisplayName
            PrimarySmtpAddress               = $group.PrimarySmtpAddress
            RecipientTypeDetails             = $group.RecipientTypeDetails
            RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
            HiddenFromAddressListsEnabled    = $group.HiddenFromAddressListsEnabled
            MemberJoinRestriction            = $group.MemberJoinRestriction
            MemberDepartRestriction          = $group.MemberDepartRestriction
            ModerationEnabled                = $group.ModerationEnabled
        }
    }
}

function Get-Microsoft365GroupSettingsData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.UnifiedGroups) {
        [PSCustomObject]@{
            DisplayName                      = $group.DisplayName
            PrimarySmtpAddress               = $group.PrimarySmtpAddress
            AccessType                       = $group.AccessType
            RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
            ModerationEnabled                = $group.ModerationEnabled
            HiddenGroupMembershipEnabled     = $group.HiddenGroupMembershipEnabled
            ResourceProvisioningOptions      = ($group.ResourceProvisioningOptions -join ', ')
        }
    }
}

function Get-SecurityGroupOwnerData {
    [CmdletBinding()]
    param()

    $securityGroups = @(
        $script:TenantReportState.Cache.GraphGroups |
            Where-Object { $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false }
    )

    $total = $securityGroups.Count
    $index = 0

    foreach ($group in $securityGroups) {
        $index++
        Write-TenantReportProgress -Activity 'Exporting security group owners' -Current $index -Total $total -Status $group.DisplayName

        $owners = @(
            Invoke-TenantReportReadOperation -Description "retrieve owners for security group '$($group.DisplayName)'" -Operation {
                Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop
            }
        )

        foreach ($owner in $owners) {
            $summary = Get-DirectoryObjectSummary -DirectoryObject $owner

            [PSCustomObject]@{
                GroupName          = $group.DisplayName
                GroupObjectId      = $group.Id
                OwnerName          = $summary.DisplayName
                OwnerEmailAddress  = $summary.UserPrincipalName
                OwnerType          = $summary.ObjectType
                OwnerUserType      = $summary.UserType
            }
        }
    }

    Write-Progress -Activity 'Exporting security group owners' -Completed
}

function Get-DynamicGroupData {
    [CmdletBinding()]
    param()

    foreach ($group in $script:TenantReportState.Cache.GraphGroups |
        Where-Object { @($_.GroupTypes) -contains 'DynamicMembership' }) {
        [PSCustomObject]@{
            DisplayName    = $group.DisplayName
            ObjectId       = $group.Id
            MembershipRule = $group.MembershipRule
        }
    }
}

function Get-MailboxForwardingData {
    [CmdletBinding()]
    param()

    foreach ($mailbox in $script:TenantReportState.Cache.Mailboxes |
        Where-Object { $_.ForwardingAddress -or $_.ForwardingSmtpAddress }) {
        [PSCustomObject]@{
            DisplayName                = $mailbox.DisplayName
            PrimarySmtpAddress         = $mailbox.PrimarySmtpAddress
            UserPrincipalName          = $mailbox.UserPrincipalName
            ForwardTo                  = $mailbox.ForwardingAddress
            ForwardingSmtpAddress      = $mailbox.ForwardingSmtpAddress
            DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
        }
    }
}

function Get-DomainDetailData {
    [CmdletBinding()]
    param()

    $domains = @(
        Invoke-TenantReportReadOperation -Description 'retrieve tenant domains' -Operation {
            Get-MgDomain -All -ErrorAction Stop
        }
    )

    foreach ($domain in $domains) {
        [PSCustomObject]@{
            Name           = $domain.Id
            IsVerified     = $domain.IsVerified
            Authentication = $domain.AuthenticationType
            IsInitial      = $domain.IsInitial
        }
    }
}

function Export-UserProfilePhotos {
    [CmdletBinding()]
    param()

    $users = $script:TenantReportState.Cache.Users
    $total = $users.Count
    $index = 0

    foreach ($user in $users) {
        $index++
        Write-TenantReportProgress -Activity 'Downloading user profile photos' -Current $index -Total $total -Status $user.UserPrincipalName

        $fileName = '{0}.jpg' -f (Get-SafeTenantReportFileName -Value $user.UserPrincipalName -Fallback $user.Id)
        $destination = Join-Path -Path $script:TenantReportState.PhotoDirectory -ChildPath $fileName

        try {
            Invoke-TenantReportReadOperation -Description "download profile photo for '$($user.UserPrincipalName)'" -Operation {
                Get-MgUserPhotoContent -UserId $user.Id -OutFile $destination -ErrorAction Stop
            }
        }
        catch {
            # A missing profile photo is expected for many users. Keep it out of the error stream.
            Write-TenantReportLog -Level 'DEBUG' -Message "Photo not downloaded for '$($user.UserPrincipalName)'. $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity 'Downloading user profile photos' -Completed
}

function Get-UniqueExcelWorksheetName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProposedName,

        [Parameter(Mandatory)]
        [System.Collections.Generic.HashSet[string]]$UsedNames
    )

    $name = $ProposedName -replace '[:\\/?*\[\]]', '_'
    $name = $name.Substring(0, [Math]::Min(31, $name.Length))

    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = 'Report'
    }

    $candidate = $name
    $counter = 1

    while ($UsedNames.Contains($candidate)) {
        $suffix = "_$counter"
        $baseLength = [Math]::Min((31 - $suffix.Length), $name.Length)
        $candidate = $name.Substring(0, $baseLength) + $suffix
        $counter++
    }

    $null = $UsedNames.Add($candidate)
    return $candidate
}

function New-ExcelWorkbookFromCsvReports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantShortName
    )

    if (-not $IsWindows) {
        Write-TenantReportLog -Level 'WARN' -Message 'Skipping master workbook creation because Excel COM automation is available only on Windows.'
        return
    }

    $csvFiles = @(
        Get-ChildItem -LiteralPath $script:TenantReportState.ReportDirectory -Filter '*.csv' -File |
            Where-Object { $_.Name -ne 'Run_Manifest.csv' } |
            Sort-Object -Property Name
    )

    if ($csvFiles.Count -eq 0) {
        Write-TenantReportLog -Level 'WARN' -Message 'Skipping master workbook creation because no CSV reports were created.'
        return
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $usedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($csvFile in $csvFiles) {
            $workbookForCsv = $null

            try {
                $workbookForCsv = $excel.Workbooks.Open($csvFile.FullName)
                $worksheet = $workbookForCsv.Worksheets.Item(1)
                $worksheet.Copy($workbook.Worksheets.Item($workbook.Worksheets.Count))
                $copiedWorksheet = $workbook.Worksheets.Item($workbook.Worksheets.Count)
                $copiedWorksheet.Name = Get-UniqueExcelWorksheetName -ProposedName $csvFile.BaseName -UsedNames $usedNames
            }
            finally {
                if ($workbookForCsv) {
                    $workbookForCsv.Close($false)
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbookForCsv)
                }
            }
        }

        # Workbook.Add supplies a default worksheet. Remove it after at least one CSV has been copied.
        if ($workbook.Worksheets.Count -gt 1) {
            $workbook.Worksheets.Item(1).Delete()
        }

        $destination = Join-Path -Path $script:TenantReportState.ReportDirectory -ChildPath (
            'Master_TenantReport_{0}.xlsx' -f $TenantShortName
        )

        $workbook.SaveAs($destination)
        Write-TenantReportLog -Level 'INFO' -Message "Created master workbook '$destination'."
    }
    catch {
        Write-TenantReportLog -Level 'WARN' -Message "Unable to create the master workbook. $($_.Exception.Message)"
    }
    finally {
        if ($workbook) {
            $workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-RunManifest {
    [CmdletBinding()]
    param()

    $destination = Join-Path -Path $script:TenantReportState.ReportDirectory -ChildPath 'Run_Manifest.csv'
    $script:TenantReportState.Manifest | Export-Csv -LiteralPath $destination -NoTypeInformation -Encoding utf8 -Force
}

function Invoke-TenantInventoryReport {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseOutputPath,

        [Parameter()]
        [string]$TargetTenantDomain,

        [Parameter()]
        [string]$TargetSharePointAdminUrl,

        [Parameter()]
        [string]$RequestedLogPath,

        [Parameter()]
        [switch]$SkipPhotos,

        [Parameter()]
        [switch]$SkipWorkbook,

        [Parameter()]
        [switch]$SkipManagers,

        [Parameter()]
        [switch]$HideProgress,

        [Parameter()]
        [switch]$UseTranscript,

        [Parameter()]
        [switch]$FailOnReportError,

        [Parameter()]
        [int]$RetryAttempts,

        [Parameter()]
        [int]$RetryDelaySeconds
    )

    if (-not $PSCmdlet.ShouldProcess($BaseOutputPath, 'Generate local tenant inventory reports')) {
        return
    }

    $script:TenantReportState.NoProgress = $HideProgress.IsPresent
    $script:TenantReportState.MaxRetryAttempts = $RetryAttempts
    $script:TenantReportState.InitialRetryDelaySeconds = $RetryDelaySeconds
    $script:TenantReportState.Cache = @{}
    $script:TenantReportState.Manifest = [System.Collections.Generic.List[object]]::new()
    $script:TenantReportState.ConnectedGraph = $false
    $script:TenantReportState.ConnectedExchange = $false
    $script:TenantReportState.ConnectedSharePoint = $false
    $script:TenantReportState.TranscriptStarted = $false

    try {
        Import-TenantReportModules

        # Graph must be connected before an initial onmicrosoft.com domain can be discovered.
        $graphScopes = @(
            'User.Read.All',
            'UserAuthenticationMethod.Read.All',
            'Device.Read.All',
            'DeviceManagementManagedDevices.Read.All',
            'Group.Read.All',
            'GroupMember.Read.All',
            'Domain.Read.All',
            'LicenseAssignment.Read.All'
        )

        $context = Get-MgContext
        if (-not $context) {
            $graphConnectParameters = @{
                Scopes = $graphScopes
                NoWelcome = $true
            }

            if ($TargetTenantDomain) {
                $graphConnectParameters.TenantId = $TargetTenantDomain
            }

            Connect-MgGraph @graphConnectParameters
            $script:TenantReportState.ConnectedGraph = $true
        }

        $initialDomain = if ($TargetTenantDomain) {
            $TargetTenantDomain
        }
        else {
            $domain = Get-MgDomain -All -ErrorAction Stop | Where-Object { $_.IsInitial } | Select-Object -First 1

            if (-not $domain) {
                throw 'Unable to identify an initial .onmicrosoft.com tenant domain. Supply -TenantDomain explicitly.'
            }

            $domain.Id
        }

        $tenantShortName = $initialDomain -replace '\.onmicrosoft\.com$', ''

        $adminSiteUrl = if ($TargetSharePointAdminUrl) {
            $TargetSharePointAdminUrl
        }
        else {
            "https://$tenantShortName-admin.sharepoint.com"
        }

        Initialize-TenantReportDirectories -BaseOutputPath $BaseOutputPath -TenantShortName $tenantShortName -RequestedLogPath $RequestedLogPath
        Write-TenantReportLog -Level 'INFO' -Message "Starting tenant inventory export for '$initialDomain'."

        if ($UseTranscript) {
            $transcriptPath = Join-Path -Path $script:TenantReportState.ReportDirectory -ChildPath 'TenantInventoryReport.transcript.txt'
            Start-Transcript -LiteralPath $transcriptPath -Append -ErrorAction Stop
            $script:TenantReportState.TranscriptStarted = $true
        }

        # Graph is already connected at this point. Connect the two remaining service modules.
        $existingExchangeConnection = $false
        if (Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue) {
            try {
                $existingExchangeConnection = @(
                    Get-ConnectionInformation -ErrorAction Stop |
                        Where-Object { $_.State -eq 'Connected' -or -not $_.State }
                ).Count -gt 0
            }
            catch {
                # If detection is unavailable, establish and track a new connection.
                $existingExchangeConnection = $false
            }
        }

        if ($existingExchangeConnection) {
            Write-TenantReportLog -Level 'INFO' -Message 'Reusing the existing Exchange Online connection.'
        }
        else {
            Write-TenantReportLog -Level 'INFO' -Message 'Connecting to Exchange Online.'
            Connect-ExchangeOnline -ShowBanner:$false
            $script:TenantReportState.ConnectedExchange = $true
        }

        Write-TenantReportLog -Level 'INFO' -Message "Connecting to SharePoint Online admin center '$adminSiteUrl'."
        Connect-SPOService -Url $adminSiteUrl
        $script:TenantReportState.ConnectedSharePoint = $true

        Initialize-TenantReportCache

        $reportDefinitions = @(
            @{ Name = 'All_AzureAD_Users'; Data = { Get-UserInventoryData -SkipManagerLookup:$SkipManagers } },
            @{ Name = 'Azure_Device_Report'; Data = { Get-EntraDeviceData } },
            @{ Name = 'DL_Alias'; Data = { Get-ProxyAddressReportData -Recipients $script:TenantReportState.Cache.DistributionGroups -RecipientCategory DistributionGroup } },
            @{ Name = 'M365_Groups_Alias'; Data = { Get-ProxyAddressReportData -Recipients $script:TenantReportState.Cache.UnifiedGroups -RecipientCategory Microsoft365Group } },
            @{ Name = 'Mailbox_Alias'; Data = { Get-ProxyAddressReportData -Recipients $script:TenantReportState.Cache.Mailboxes -RecipientCategory Mailbox } },
            @{ Name = 'Mailbox_Permission'; Data = { Get-MailboxPermissionData } },
            @{ Name = 'X500_M365_Groups'; Data = { Get-LegacyExchangeDnData -Recipients $script:TenantReportState.Cache.UnifiedGroups -RecipientCategory Microsoft365Group } },
            @{ Name = 'X500_Mailbox'; Data = { Get-LegacyExchangeDnData -Recipients $script:TenantReportState.Cache.Mailboxes -RecipientCategory Mailbox } },
            @{ Name = 'X500_DL'; Data = { Get-LegacyExchangeDnData -Recipients $script:TenantReportState.Cache.DistributionGroups -RecipientCategory DistributionGroup } },
            @{ Name = 'Intune_Devices'; Data = { Get-IntuneManagedDeviceData } },
            @{ Name = 'Azure_License_Report'; Data = { Get-UserLicenseData } },
            @{ Name = 'SSPR_Details'; Data = { Get-SelfServicePasswordResetData } },
            @{ Name = 'Mailbox_&_Archive_Size_Report'; Data = { Get-MailboxSizeData } },
            @{ Name = 'DL_Groups_Details'; Data = { Get-DistributionGroupDetailData } },
            @{ Name = 'DL_Groups_Members'; Data = { Get-DistributionGroupMemberData } },
            @{ Name = 'M365_Groups_Details'; Data = { Get-Microsoft365GroupDetailData } },
            @{ Name = 'M365_Groups_Calendar_Items'; Data = { Get-Microsoft365GroupCalendarData } },
            @{ Name = 'Onedrive_Size'; Data = { Get-OneDriveSizeData } },
            @{ Name = 'All_Mailbox_Report'; Data = { Get-MailboxInventoryData } },
            @{ Name = 'Deleted_Users'; Data = { Get-DeletedUserData } },
            @{ Name = 'Security_Groups_Details'; Data = { Get-SecurityGroupDetailData } },
            @{ Name = 'Security_Groups_Members'; Data = { Get-SecurityGroupMemberData } },
            @{ Name = 'M365_Groups_Members'; Data = { Get-Microsoft365GroupMemberData } },
            @{ Name = 'DL_Owners'; Data = { Get-DistributionGroupOwnerData } },
            @{ Name = 'M365_Groups_Owners'; Data = { Get-Microsoft365GroupOwnerData } },
            @{ Name = 'DL_Settings'; Data = { Get-DistributionGroupSettingsData } },
            @{ Name = 'M365_Groups_Settings'; Data = { Get-Microsoft365GroupSettingsData } },
            @{ Name = 'Security_Groups_Owners'; Data = { Get-SecurityGroupOwnerData } },
            @{ Name = 'Dynamic_Groups'; Data = { Get-DynamicGroupData } },
            @{ Name = 'Mailbox_Forwarding'; Data = { Get-MailboxForwardingData } },
            @{ Name = 'All_Domains_Details'; Data = { Get-DomainDetailData } }
        )

        foreach ($definition in $reportDefinitions) {
            Export-TenantReportCsv -ReportName $definition.Name -DataScript $definition.Data -StopOnError:$FailOnReportError
        }

        if (-not $SkipPhotos) {
            try {
                Export-UserProfilePhotos
                $script:TenantReportState.Manifest.Add([PSCustomObject]@{
                    ReportName   = 'User_Profile_Photos'
                    Status       = 'Completed'
                    Path         = $script:TenantReportState.PhotoDirectory
                    DurationSec  = $null
                    ErrorMessage = $null
                })
            }
            catch {
                $message = $_.Exception.Message
                $script:TenantReportState.Manifest.Add([PSCustomObject]@{
                    ReportName   = 'User_Profile_Photos'
                    Status       = 'Failed'
                    Path         = $script:TenantReportState.PhotoDirectory
                    DurationSec  = $null
                    ErrorMessage = $message
                })
                Write-TenantReportLog -Level 'ERROR' -Message "User profile-photo export failed. $message"
                if ($FailOnReportError) {
                    throw
                }
            }
        }

        Export-RunManifest

        if (-not $SkipWorkbook) {
            New-ExcelWorkbookFromCsvReports -TenantShortName $tenantShortName
        }

        $completed = @($script:TenantReportState.Manifest | Where-Object Status -eq 'Completed').Count
        $failed = @($script:TenantReportState.Manifest | Where-Object Status -eq 'Failed').Count

        Write-TenantReportLog -Level 'INFO' -Message (
            "Tenant inventory export completed. Completed: $completed; Failed: $failed; Output: $($script:TenantReportState.ReportDirectory)"
        )

        return [PSCustomObject]@{
            TenantDomain     = $initialDomain
            ReportDirectory  = $script:TenantReportState.ReportDirectory
            PhotoDirectory   = if ($SkipPhotos) { $null } else { $script:TenantReportState.PhotoDirectory }
            LogPath          = $script:TenantReportState.LogPath
            CompletedReports = $completed
            FailedReports    = $failed
        }
    }
    catch {
        if ($script:TenantReportState.LogPath) {
            Write-TenantReportLog -Level 'ERROR' -Message "Tenant inventory export failed. $($_.Exception.Message)"
        }

        throw
    }
    finally {
        if ($script:TenantReportState.TranscriptStarted) {
            try {
                Stop-Transcript | Out-Null
            }
            catch {
                # The original failure is more important than transcript shutdown.
            }
        }

        Disconnect-TenantReportServices
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    Invoke-TenantInventoryReport `
        -BaseOutputPath $OutputPath `
        -TargetTenantDomain $TenantDomain `
        -TargetSharePointAdminUrl $SharePointAdminUrl `
        -RequestedLogPath $LogPath `
        -SkipPhotos:$SkipUserPhotos `
        -SkipWorkbook:$SkipMasterWorkbook `
        -SkipManagers:$SkipManagerLookup `
        -HideProgress:$NoProgress `
        -UseTranscript:$EnableTranscript `
        -FailOnReportError:$StopOnReportError `
        -RetryAttempts $MaxRetryAttempts `
        -RetryDelaySeconds $InitialRetryDelaySeconds
}
