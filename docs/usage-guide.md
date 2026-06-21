# Usage guide

## Scope

`Export-TenantInventoryReport.ps1` is a read-only tenant-reporting script. It creates local output files and does not modify Microsoft 365 objects.

## Before you run it

1. Use a Windows device with PowerShell 7.2 or later.
2. Install the required modules:

   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
   ```

3. Confirm that your account can:
   - Consent to or use the Graph read scopes requested by the script.
   - Read Exchange recipients, mailboxes, mailbox statistics, permissions, and group data.
   - Connect as a SharePoint Administrator to the SharePoint admin center.
4. Select a protected output folder with enough free disk space.

## Recommended first run

Avoid downloading photos and creating the workbook during your first validation run:

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -TenantDomain '<TenantDomain>' `
    -SharePointAdminUrl 'https://<TenantPrefix>-admin.sharepoint.com' `
    -OutputPath '<OutputPath>' `
    -SkipUserPhotos `
    -SkipMasterWorkbook `
    -Verbose
```

Review these files when the run completes:

```text
<OutputPath>\<TenantPrefix>\Reports\Run_Manifest.csv
<OutputPath>\<TenantPrefix>\Reports\TenantInventoryReport.log
```

## Common execution patterns

### Full local inventory

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>'
```

### Large tenant, faster user inventory

The user report normally resolves a manager for each user. Turn this off when manager fields are not required:

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -SkipManagerLookup `
    -SkipUserPhotos `
    -SkipMasterWorkbook
```

### Automation-host friendly local output

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -NoProgress `
    -SkipMasterWorkbook `
    -EnableTranscript
```

### Fail immediately for validation or troubleshooting

```powershell
.\src\Export-TenantInventoryReport.ps1 `
    -OutputPath '<OutputPath>' `
    -SkipUserPhotos `
    -SkipMasterWorkbook `
    -StopOnReportError `
    -Verbose
```

## Output folders

The script creates the following structure:

```text
<OutputPath>\
└── <TenantPrefix>\
    ├── Reports\
    │   ├── *.csv
    │   ├── Run_Manifest.csv
    │   ├── TenantInventoryReport.log
    │   └── TenantInventoryReport.transcript.txt  # when enabled
    └── Photos\
        └── *.jpg                                # unless skipped
```

## Operational guidance

- Test with a non-production tenant or a controlled administrative account first.
- Do not use an unprotected shared folder for output.
- Use `-SkipMasterWorkbook` on systems without Microsoft Excel.
- Use `-SkipUserPhotos` where photos are not required.
- Use `-SkipManagerLookup` to reduce Graph requests in very large tenants.
- If an individual report fails, inspect `Run_Manifest.csv`; independent reports continue unless `-StopOnReportError` is supplied.
- Do not edit generated report files while the script is still running.

## Safe publication practices

Before pushing to GitHub, confirm that the repository does not contain:

- Tenant reports, CSV exports, Excel workbooks, logs, transcripts, or profile photos
- Personal email addresses or user principal names
- Tenant IDs, domain names, client IDs, secrets, certificate thumbprints, access tokens, or internal URLs
- Screenshots that expose tenant data
