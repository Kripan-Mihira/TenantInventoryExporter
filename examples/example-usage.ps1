<#
.SYNOPSIS
    Example execution patterns for Export-TenantInventoryReport.ps1.

.NOTES
    Replace all placeholder values before running.
    Do not store credentials, client secrets, access tokens, or tenant exports in this file.
#>

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..\src\Export-TenantInventoryReport.ps1'

# Recommended first run: CSV reports and logs only.
$parameters = @{
    TenantDomain          = '<TenantDomain>' # Example: contoso.onmicrosoft.com
    SharePointAdminUrl    = 'https://<TenantPrefix>-admin.sharepoint.com'
    OutputPath            = '<OutputPath>'
    SkipUserPhotos        = $true
    SkipMasterWorkbook    = $true
    EnableTranscript      = $true
    NoProgress            = $false
    MaxRetryAttempts      = 4
    InitialRetryDelaySeconds = 2
    Verbose               = $true
}

& $scriptPath @parameters

# Full report collection with photos and an Excel workbook:
# & $scriptPath `
#     -TenantDomain '<TenantDomain>' `
#     -SharePointAdminUrl 'https://<TenantPrefix>-admin.sharepoint.com' `
#     -OutputPath '<OutputPath>'

# Fast large-tenant run (skips user manager lookups, photos, and Excel):
# & $scriptPath `
#     -OutputPath '<OutputPath>' `
#     -SkipManagerLookup `
#     -SkipUserPhotos `
#     -SkipMasterWorkbook `
#     -NoProgress

# Dry run of local file-operation intent:
# & $scriptPath -OutputPath '<OutputPath>' -WhatIf
