# Tenant Inventory Reporter

A production-oriented PowerShell reporting tool that exports a read-only inventory of a Microsoft 365 tenant to CSV files. It consolidates the reporting intent of the original script into a modular, documented, testable, and safer implementation.

> **Important:** The generated reports can contain personal and operationally sensitive information, including user profile data, authentication-method details, mailbox delegates, device details, group membership, email addresses, forwarding, and profile photos. Store, transfer, and retain them according to your organization’s security and privacy requirements.

## Purpose

`Export-TenantInventoryReport.ps1` collects tenant inventory data from:

- Microsoft Entra ID / Microsoft Graph
- Microsoft Intune
- Exchange Online
- SharePoint Online and OneDrive
- Microsoft 365 Groups and distribution groups

The script is **read-only against Microsoft 365**. It does not modify tenant objects, memberships, mailboxes, licenses, policies, or devices. It creates local folders, CSV files, a log, an optional transcript, downloaded profile photos, and an optional Excel workbook.

## Features

- Uses Microsoft Graph PowerShell instead of the retired AzureAD module.
- Uses explicit parameters instead of fixed tenant paths and identifiers.
- Separates data collection into named functions.
- Adds validation, logging, retry handling for transient read failures, progress reporting, cleanup, and a run manifest.
- Eliminates `Format-Table` from data pipelines and avoids repeated `+=` array construction.
- Exports CSV reports in UTF-8.
- Handles no-manager, no-photo, no-alias, no-group, and no-archive scenarios more safely.
- Attempts to close only service connections opened by the script.
- Optionally creates a master Excel workbook from the CSV reports.

## Prerequisites

| Requirement | Details |
|---|---|
| Operating system | Windows. SharePoint Online Management Shell support and optional Excel COM workbook generation require Windows. |
| PowerShell | PowerShell **7.2 or later**. |
| Microsoft Graph PowerShell SDK | `Microsoft.Graph` |
| Exchange module | `ExchangeOnlineManagement` |
| SharePoint module | `Microsoft.Online.SharePoint.PowerShell` |
| Microsoft Excel | Required only when generating the optional master `.xlsx` workbook. |
| Network | Access to Microsoft Graph, Exchange Online, and SharePoint Online endpoints. |

The Graph PowerShell SDK is recommended with PowerShell 7 or later. The SharePoint Online module is imported using Windows PowerShell compatibility when the script runs in PowerShell 7.

## Installation

Clone or download this repository, then install the required modules in a PowerShell session:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

Confirm the script is available:

```powershell
Get-ChildItem .\src\Export-TenantInventoryReport.ps1
```

For a signed or restricted environment, follow your organization’s approved script-signing and execution-policy process. Do not lower execution policy or bypass security controls merely to run this tool.

## Authentication and permissions

This release uses **interactive delegated authentication**. It does not accept passwords, client secrets, access tokens, certificates, or tenant credentials as command-line parameters.

### Microsoft Graph delegated scopes requested

The script requests only the Graph delegated scopes needed by its report set:

```text
User.Read.All
UserAuthenticationMethod.Read.All
Device.Read.All
DeviceManagementManagedDevices.Read.All
Group.Read.All
GroupMember.Read.All
Domain.Read.All
LicenseAssignment.Read.All
```

Your tenant may require administrator consent. Authentication-method reporting also requires an appropriately privileged Microsoft Entra role. In many environments, a Global Reader, Authentication Administrator, or Privileged Authentication Administrator role is required; note that phone numbers can be masked for an Authentication Administrator.

### Exchange Online

Use an account with Exchange Online RBAC permissions sufficient to read recipients, mailboxes, mailbox statistics, mailbox permissions, group membership, and group settings. The exact least-privilege role assignment depends on your tenant’s custom RBAC model.

### SharePoint Online

Use an account with SharePoint Administrator access to retrieve OneDrive site inventory.

## Usage

### Standard interactive run

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -SkipMasterWorkbook
```

The script determines the initial `<tenant>.onmicrosoft.com` domain from Microsoft Graph and derives the SharePoint admin URL.

### Explicit tenant and SharePoint admin URL

Use this pattern for multi-geo, sovereign cloud, non-standard, or otherwise non-default SharePoint admin endpoints:

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -TenantDomain '<TenantDomain>' `
    -SharePointAdminUrl 'https://<TenantPrefix>-admin.sharepoint.com' `
    -OutputPath '<OutputPath>' `
    -SkipUserPhotos `
    -SkipMasterWorkbook `
    -Verbose
```

### Dry-run file-operation check

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -WhatIf
```

### Stop on the first report failure and create a transcript

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -EnableTranscript `
    -StopOnReportError
```

## Parameters

| Parameter | Type | Default | Description |
|---|---:|---|---|
| `OutputPath` | String | `./TenantReports` | Base folder for `<tenant>/Reports` and `<tenant>/Photos`. |
| `TenantDomain` | String | Discovered | Initial `<tenant>.onmicrosoft.com` domain. Use when automatic discovery is unsuitable. |
| `SharePointAdminUrl` | String | Derived | SharePoint admin center URL. Supply explicitly for non-standard environments. |
| `LogPath` | String | `<ReportDirectory>/TenantInventoryReport.log` | Local log-file path. |
| `SkipUserPhotos` | Switch | Off | Does not download user profile photos. |
| `SkipMasterWorkbook` | Switch | Off | Does not create the Excel workbook. Recommended if Excel is unavailable. |
| `SkipManagerLookup` | Switch | Off | Skips per-user Graph manager lookups and leaves manager columns blank. Faster for very large tenants. |
| `NoProgress` | Switch | Off | Suppresses progress bars. Useful for pipelines and logs. |
| `EnableTranscript` | Switch | Off | Creates or appends a local PowerShell transcript. |
| `StopOnReportError` | Switch | Off | Stops the run when a report fails. Default behavior logs the failure and continues to independent reports. |
| `MaxRetryAttempts` | Integer | `4` | Maximum attempts for transient read failures. |
| `InitialRetryDelaySeconds` | Integer | `2` | Initial exponential-backoff delay. |

## Output

The tool writes its results to:

```text
<OutputPath>/
└── <TenantPrefix>/
    ├── Reports/
    │   ├── All_AzureAD_Users.csv
    │   ├── Azure_Device_Report.csv
    │   ├── Intune_Devices.csv
    │   ├── Mailbox_Permission.csv
    │   ├── Mailbox_&_Archive_Size_Report.csv
    │   ├── Onedrive_Size.csv
    │   ├── Security_Groups_Members.csv
    │   ├── M365_Groups_Members.csv
    │   ├── ...additional report CSVs...
    │   ├── Run_Manifest.csv
    │   ├── TenantInventoryReport.log
    │   └── Master_TenantReport_<TenantPrefix>.xlsx  # unless skipped
    └── Photos/
        └── <UserPrincipalName>.jpg                  # unless skipped
```

The complete report set includes Entra users, Entra devices, Intune devices, licenses, SSPR methods, domains, deleted users, OneDrive usage, mailboxes, mailbox sizes, mailbox permissions, mailbox forwarding, distribution groups, Microsoft 365 groups, group settings, owners, members, aliases, and Legacy Exchange DN values.

`Run_Manifest.csv` identifies each report, completion status, local path, execution duration, and error message where applicable.

## Logging

The log uses ISO 8601 timestamps and `DEBUG`, `INFO`, `WARN`, and `ERROR` levels. By default it is written to:

```text
<OutputPath>\<TenantPrefix>\Reports\TenantInventoryReport.log
```

Use `-LogPath '<LogPath>'` to select another local file location. `-EnableTranscript` creates a separate transcript at:

```text
<ReportDirectory>\TenantInventoryReport.transcript.txt
```

## Security considerations

- Do not commit generated CSV, XLSX, log, transcript, or photo files to source control.
- Restrict file-system permissions on the output folder.
- Treat SSPR data, forwarding configuration, group membership, mailbox permissions, and profile photos as sensitive.
- Review retention requirements before storing generated reports.
- Use a dedicated administrative account where your organization requires separation of duties.
- Grant the minimum Graph delegated scopes and Exchange/SharePoint roles needed for this task.
- The script deliberately avoids hard-coded passwords, tenant IDs, client secrets, tokens, certificate thumbprints, internal paths, and email addresses.
- Do not place secrets in `examples/`, source files, transcripts, issue descriptions, or pull requests.

## Limitations and operational notes

- This is an interactive delegated-authentication script; unattended app-only authentication is intentionally not included because it was not part of the original script’s functionality.
- Large tenants can take substantial time because some data, such as mailbox permissions, SSPR methods, group memberships, device registrations, calendar statistics, and profile photos, requires per-object calls.
- `SkipManagerLookup` can materially reduce calls for large user populations.
- Microsoft Graph, Exchange Online, and SharePoint Online can throttle requests. The script retries only errors that appear transient; non-transient authorization and configuration errors fail the affected report.
- The master workbook depends on desktop Microsoft Excel for Windows. Use `-SkipMasterWorkbook` on automation hosts, systems without Excel, or when Excel COM is not allowed.
- OneDrive personal-site discovery depends on the SharePoint Online Management Shell and the tenant’s SharePoint host pattern. Supply `-SharePointAdminUrl` explicitly where automatic derivation is not valid.
- Report schema can vary as Microsoft service APIs evolve. Validate output in a non-production tenant before organization-wide use.

## Troubleshooting

### `Required command ... is unavailable`

Install the missing module in the same PowerShell environment, restart PowerShell, and run again.

```powershell
Get-Module -ListAvailable Microsoft.Graph, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell
```

### `Insufficient privileges` or a Graph authorization error

Verify that the account has accepted or received consent for the requested Graph scopes and has the required Microsoft Entra role for authentication methods. Also confirm sufficient Exchange Online RBAC and SharePoint Administrator access.

### SharePoint connection fails

Pass the SharePoint admin URL explicitly:

```powershell
-SharePointAdminUrl 'https://<TenantPrefix>-admin.sharepoint.com'
```

Confirm that the account is a SharePoint Administrator and that the module can be imported through Windows PowerShell compatibility.

### Master workbook is not created

Install desktop Microsoft Excel for Windows or run the script with:

```powershell
-SkipMasterWorkbook
```

CSV files and the run manifest remain available even if workbook creation is skipped or fails.

### A single report fails while others succeed

Review `Run_Manifest.csv` and `TenantInventoryReport.log`. Re-run with:

```powershell
-StopOnReportError -Verbose
```

to stop at the first failure and surface its context.

## Repository structure

```text
/
├── README.md
├── LICENSE
├── .gitignore
├── src/
│   └── Export-TenantInventoryReport.ps1
├── docs/
│   └── usage-guide.md
├── examples/
│   └── example-usage.ps1
├── tests/
│   └── Export-TenantInventoryReport.Tests.ps1
└── CHANGELOG.md
```

## Testing

Install Pester and run the unit tests:

```powershell
Install-Module Pester -Scope CurrentUser -Force
Invoke-Pester -Path .\tests\Export-TenantInventoryReport.Tests.ps1 -Output Detailed
```

The included tests cover deterministic helper functions and local CSV export behavior. They do not connect to Microsoft 365.

## Contributing

1. Fork the repository and create a focused branch.
2. Do not add tenant data, secrets, screenshots containing personal data, exported CSV files, or profile photos.
3. Keep changes compatible with PowerShell 7.2+ on Windows.
4. Add or update Pester tests for changed helper behavior.
5. Run Pester locally before opening a pull request.
6. Explain tenant-impact, permission, output-schema, and backward-compatibility considerations in the pull request.

## Disclaimer

This project is provided as a community automation example without warranty. Validate it in a non-production or representative tenant before use, review the code and requested permissions, and ensure that its data collection and storage comply with your organization’s policies and applicable law.
