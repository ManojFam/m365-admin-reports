BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\M365AdminReports\M365AdminReports.psd1') -Force
    . (Join-Path $PSScriptRoot 'GraphStubs.ps1')
}

Describe 'Get-M365StaleAccountReport (mock tenant)' {
    BeforeAll {
        $report = @(Get-M365StaleAccountReport -Mock)
        $byUpn = @{}
        foreach ($row in $report) { $byUpn[$row.UserPrincipalName] = $row }
    }

    It 'finds the expected 7 accounts at the default 90 days' {
        $report.Count | Should -Be 7
        @($report | Where-Object Finding -eq 'Inactive').Count | Should -Be 2
        @($report | Where-Object Finding -eq 'NeverSignedIn').Count | Should -Be 2
        @($report | Where-Object Finding -eq 'DisabledLicensed').Count | Should -Be 3
    }

    It 'measures days since sign-in from the snapshot date' {
        $row = $byUpn['eli.foster@example.com']
        $row.Finding | Should -Be 'Inactive'
        $row.LastSignIn | Should -Be '2026-06-02'
        $row.DaysSinceSignIn | Should -Be 166
    }

    It 'includes inactive guest accounts' {
        $byUpn['oakley.grant_partner.example#EXT#@example.com'].Finding | Should -Be 'Inactive'
    }

    It 'flags an old account that has never been used' {
        $byUpn['jules.ortega@example.com'].Finding | Should -Be 'NeverSignedIn'
        $byUpn['jules.ortega@example.com'].LastSignIn | Should -Be 'Never'
    }

    It 'does not flag a new starter who has not signed in yet' {
        $byUpn.ContainsKey('kai.bennett@example.com') | Should -BeFalse
    }

    It 'uses the non-interactive sign-in when it is newer' {
        # Parker last signed in interactively in July, but Outlook on the phone is still syncing.
        $byUpn.ContainsKey('parker.young@example.com') | Should -BeFalse
    }

    It 'flags disabled accounts that still hold a licence, with the licence name' {
        $row = $byUpn['georgie.lam@example.com']
        $row.Finding | Should -Be 'DisabledLicensed'
        $row.AccountEnabled | Should -BeFalse
        $row.Licences | Should -Be 'Microsoft 365 A3 for faculty'
    }

    It 'does not flag disabled accounts with no licence' {
        $byUpn.ContainsKey('indy.rowe@example.com') | Should -BeFalse
        $byUpn.ContainsKey('reception@example.com') | Should -BeFalse
    }

    It 'gives every row a recommendation' {
        $report.Recommendation | Should -Not -Contain $null
    }

    It 'honours -InactiveDays' {
        $strict = @(Get-M365StaleAccountReport -Mock -InactiveDays 30)
        # The break-glass account last signed in 45 days before the snapshot.
        $strict.UserPrincipalName | Should -Contain 'breakglass.01@example.com'
        $strict.Count | Should -BeGreaterThan $report.Count
    }

    It 'honours -ReferenceDate' {
        $later = @(Get-M365StaleAccountReport -Mock -ReferenceDate '2027-06-01')
        @($later | Where-Object Finding -eq 'Inactive').Count | Should -BeGreaterThan 2
    }

    It 'rejects an out-of-range -InactiveDays' {
        { Get-M365StaleAccountReport -Mock -InactiveDays 0 } | Should -Throw
    }
}

Describe 'Get-M365StaleAccountReport (live mode)' {
    BeforeAll { Add-LiveStub }
    AfterAll { Remove-LiveStub }

    It 'measures from today and maps licence names from Graph' {
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { New-GraphContextMock }
        Mock -CommandName Get-MgOrganization -ModuleName M365AdminReports -MockWith { [pscustomobject]@{ DisplayName = 'Test tenant' } }
        Mock -CommandName Get-MgSubscribedSku -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ SkuId = 'sku-1'; SkuPartNumber = 'SPE_E3'; CapabilityStatus = 'Enabled'; ConsumedUnits = 1; PrepaidUnits = [pscustomobject]@{ Enabled = 5 } }
        }
        Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{
                Id = 'id-1'; DisplayName = 'Old Account'; UserPrincipalName = 'old.account@example.com'
                UserType = 'Member'; AccountEnabled = $true; CreatedDateTime = (Get-Date).AddYears(-3)
                AssignedLicenses = @([pscustomobject]@{ SkuId = 'sku-1' })
                SignInActivity = [pscustomobject]@{ LastSignInDateTime = (Get-Date).AddDays(-120); LastNonInteractiveSignInDateTime = $null }
            }
        }

        $row = Get-M365StaleAccountReport
        $row.Finding | Should -Be 'Inactive'
        $row.DaysSinceSignIn | Should -BeIn 119, 120
        $row.Licences | Should -Be 'Microsoft 365 E3'
    }
}
