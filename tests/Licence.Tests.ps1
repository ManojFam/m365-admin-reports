BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\M365AdminReports\M365AdminReports.psd1') -Force
    . (Join-Path $PSScriptRoot 'GraphStubs.ps1')
}

Describe 'Get-M365LicenceReport (mock tenant)' {
    BeforeAll {
        $report = @(Get-M365LicenceReport -Mock)
        $bySku = @{}
        foreach ($row in $report) { $bySku[$row.SkuPartNumber] = $row }
    }

    It 'returns one row per SKU with a friendly name' {
        $report.Count | Should -Be 5
        $bySku['M365EDU_A3_FACULTY'].Licence | Should -Be 'Microsoft 365 A3 for faculty'
        $bySku['AAD_PREMIUM'].Licence | Should -Be 'Microsoft Entra ID P1'
    }

    It 'works out total, assigned, available and percent used' {
        $row = $bySku['STANDARDWOFFPACK_FACULTY']
        $row.Total | Should -Be 25
        $row.Assigned | Should -Be 8
        $row.Available | Should -Be 17
        $row.PercentUsed | Should -Be 32
    }

    It 'lists licences assigned to blocked accounts' {
        $row = $bySku['M365EDU_A3_FACULTY']
        $row.AssignedToBlocked | Should -Be 1
        $row.BlockedUsers | Should -Be 'georgie.lam@example.com'
    }

    It 'warns when no licences are left' {
        $bySku['M365EDU_A3_FACULTY'].Available | Should -Be 0
        $bySku['M365EDU_A3_FACULTY'].Status | Should -Be 'Warning'
        $bySku['M365EDU_A3_FACULTY'].Finding | Should -BeLike '*No licences left*'
    }

    It 'warns when a subscription is in its grace period' {
        $bySku['POWER_BI_PRO'].SubscriptionStatus | Should -Be 'Warning'
        $bySku['POWER_BI_PRO'].Status | Should -Be 'Warning'
    }

    It 'reports a healthy SKU as OK' {
        $bySku['AAD_PREMIUM'].Status | Should -Be 'OK'
        $bySku['AAD_PREMIUM'].Finding | Should -Be ''
    }
}

Describe 'Get-M365LicenceReport (live mode)' {
    BeforeAll { Add-LiveStub }
    AfterAll { Remove-LiveStub }

    BeforeEach {
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { New-GraphContextMock }
        Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith { }
    }

    It 'marks an over-assigned SKU as Critical' {
        Mock -CommandName Get-MgSubscribedSku -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ SkuId = 'sku-1'; SkuPartNumber = 'SPE_E3'; CapabilityStatus = 'Enabled'; ConsumedUnits = 12; PrepaidUnits = [pscustomobject]@{ Enabled = 10 } }
        }

        $row = Get-M365LicenceReport
        $row.Available | Should -Be -2
        $row.Status | Should -Be 'Critical'
        $row.Finding | Should -Be 'Over-assigned by 2'
    }

    It 'marks a suspended subscription as Critical and keeps unknown SKU part numbers' {
        Mock -CommandName Get-MgSubscribedSku -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ SkuId = 'sku-2'; SkuPartNumber = 'SOME_NEW_SKU'; CapabilityStatus = 'Suspended'; ConsumedUnits = 0; PrepaidUnits = [pscustomobject]@{ Enabled = 0 } }
        }

        $row = Get-M365LicenceReport
        $row.Licence | Should -Be 'SOME_NEW_SKU'
        $row.Status | Should -Be 'Critical'
        $row.PercentUsed | Should -Be 0
    }
}
