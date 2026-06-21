# Requires -Modules Pester
# Unit tests only. No Microsoft 365 tenant connection is made.

BeforeAll {
    $scriptUnderTest = Join-Path -Path $PSScriptRoot -ChildPath '..\src\Export-TenantInventoryReport.ps1'
    . $scriptUnderTest
}

Describe 'Convert-ByteSizeToGigabytes' {
    It 'returns zero for a null value' {
        Convert-ByteSizeToGigabytes -Value $null | Should -Be 0
    }

    It 'converts bytes to gigabytes' {
        Convert-ByteSizeToGigabytes -Value 1073741824 | Should -Be 1
    }

    It 'parses Exchange style total-item-size text' {
        $result = Convert-ByteSizeToGigabytes -Value '1 GB (1,073,741,824 bytes)'
        $result | Should -Be 1
    }

    It 'parses a TB value' {
        Convert-ByteSizeToGigabytes -Value '1.5 TB' | Should -Be 1536
    }
}

Describe 'Convert-QuotaToGigabytes' {
    It 'returns Unlimited unchanged' {
        Convert-QuotaToGigabytes -Quota 'Unlimited' | Should -Be 'Unlimited'
    }

    It 'converts a numeric quota' {
        Convert-QuotaToGigabytes -Quota '50 GB (53,687,091,200 bytes)' | Should -Be 50
    }
}

Describe 'Get-ProxyAddressRecord' {
    It 'separates proxy type, address, and primary SMTP state' {
        $records = @(
            Get-ProxyAddressRecord -ProxyAddresses @(
                'SMTP:primary@contoso.example',
                'smtp:alias@contoso.example',
                'X500:/o=Contoso/ou=Exchange Administrative Group'
            )
        )

        $records.Count | Should -Be 3
        $records[0].AddressType | Should -Be 'SMTP'
        $records[0].EmailAddress | Should -Be 'primary@contoso.example'
        $records[0].IsPrimarySmtp | Should -BeTrue
        $records[1].IsPrimarySmtp | Should -BeFalse
        $records[2].AddressType | Should -Be 'X500'
    }
}

Describe 'Get-SafeTenantReportFileName' {
    It 'uses the fallback when value is empty' {
        Get-SafeTenantReportFileName -Value '' -Fallback 'UserId' | Should -Be 'UserId'
    }

    It 'removes invalid file-name characters' {
        $safeName = Get-SafeTenantReportFileName -Value 'user/name:invalid'
        $safeName | Should -Not -Match '[\\/:*?"<>|]'
    }
}

Describe 'Get-UniqueExcelWorksheetName' {
    It 'limits worksheet names to 31 characters and keeps them unique' {
        $usedNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $longName = 'This is a deliberately long report file name that exceeds Excel limits'

        $first = Get-UniqueExcelWorksheetName -ProposedName $longName -UsedNames $usedNames
        $second = Get-UniqueExcelWorksheetName -ProposedName $longName -UsedNames $usedNames

        $first.Length | Should -BeLessOrEqual 31
        $second.Length | Should -BeLessOrEqual 31
        $first | Should -Not -Be $second
    }
}

Describe 'Export-TenantReportCsv' {
    BeforeEach {
        $script:TenantReportState.ReportDirectory = $TestDrive
        $script:TenantReportState.LogPath = Join-Path -Path $TestDrive -ChildPath 'test.log'
        $script:TenantReportState.Manifest = [System.Collections.Generic.List[object]]::new()
        New-Item -Path $script:TenantReportState.LogPath -ItemType File -Force | Out-Null
    }

    It 'writes local CSV output and a completed manifest entry' {
        Export-TenantReportCsv -ReportName 'Sample_Report' -DataScript {
            [PSCustomObject]@{
                Name  = 'Example'
                Value = 42
            }
        }

        $csvPath = Join-Path -Path $TestDrive -ChildPath 'Sample_Report.csv'
        Test-Path -LiteralPath $csvPath | Should -BeTrue
        (Import-Csv -LiteralPath $csvPath).Name | Should -Be 'Example'
        $script:TenantReportState.Manifest.Count | Should -Be 1
        $script:TenantReportState.Manifest[0].Status | Should -Be 'Completed'
    }
}
