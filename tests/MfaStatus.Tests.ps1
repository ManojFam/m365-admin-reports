BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\M365AdminReports\M365AdminReports.psd1') -Force
    . (Join-Path $PSScriptRoot 'GraphStubs.ps1')
}

Describe 'Get-M365MfaStatusReport (mock tenant)' {
    BeforeAll {
        $report = @(Get-M365MfaStatusReport -Mock)
        $byUpn = @{}
        foreach ($row in $report) { $byUpn[$row.UserPrincipalName] = $row }
    }

    It 'flags an admin with no MFA as Critical' {
        $row = $byUpn['blake.harper@example.com']
        $row.IsAdmin | Should -BeTrue
        $row.MfaRegistered | Should -BeFalse
        $row.Status | Should -Be 'Critical'
    }

    It 'flags a normal user with no MFA as Warning' {
        $byUpn['dana.whitfield@example.com'].Status | Should -Be 'Warning'
        $byUpn['dana.whitfield@example.com'].Finding | Should -Be 'No MFA method registered'
    }

    It 'warns when an admin only has phone (SMS/voice) MFA' {
        $row = $byUpn['morgan.ellis@example.com']
        $row.MfaRegistered | Should -BeTrue
        $row.Status | Should -Be 'Warning'
        $row.Finding | Should -BeLike '*phone*'
    }

    It 'reports admins with strong methods as OK' {
        $byUpn['avery.quinn@example.com'].Status | Should -Be 'OK'
        $byUpn['breakglass.01@example.com'].Status | Should -Be 'OK'
    }

    It 'lists Critical rows first' {
        $report[0].Status | Should -Be 'Critical'
        $report[-1].Status | Should -Be 'OK'
    }

    It 'leaves out disabled accounts and guests by default' {
        $report.AccountEnabled | Should -Not -Contain $false
        $report.UserType | Should -Not -Contain 'Guest'
        $report.Count | Should -Be 14
    }

    It 'includes them with -IncludeDisabled and -IncludeGuests' {
        $all = @(Get-M365MfaStatusReport -Mock -IncludeDisabled -IncludeGuests)
        $all.Count | Should -Be 22
        $all.UserType | Should -Contain 'Guest'
    }

    It 'joins the registered methods into one string' {
        $byUpn['noor.haddad@example.com'].MethodsRegistered | Should -Be 'windowsHelloForBusiness; microsoftAuthenticatorPush'
    }

    It 'accepts -MockPath for a custom data folder' {
        $folder = Join-Path $TestDrive 'custom'
        Copy-Item -Path (Join-Path $PSScriptRoot '..\M365AdminReports\mock') -Destination $folder -Recurse
        @(Get-M365MfaStatusReport -MockPath $folder).Count | Should -Be 14
    }

    It 'fails clearly when -MockPath does not exist' {
        { Get-M365MfaStatusReport -MockPath (Join-Path $TestDrive 'missing') } | Should -Throw '*Mock data folder not found*'
    }
}

Describe 'Get-M365MfaStatusReport (live mode)' {
    BeforeAll { Add-LiveStub }
    AfterAll { Remove-LiveStub }

    It 'fails with a clear message when Graph is not connected' {
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { $null }

        { Get-M365MfaStatusReport } | Should -Throw '*Connect-MgGraph*'
    }

    It 'warns when the sign-in is missing a scope' {
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ AuthType = 'Delegated'; Scopes = @('User.Read.All') }
        }
        Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith { }
        Mock -CommandName Get-MgReportAuthenticationMethodUserRegistrationDetail -ModuleName M365AdminReports -MockWith { }

        Get-M365MfaStatusReport -WarningVariable warnings -WarningAction SilentlyContinue | Out-Null
        "$warnings" | Should -BeLike '*Reports.Read.All*'
    }

    It 'reads users and registration details from Graph' {
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { New-GraphContextMock }
        # Graph SDK objects use PascalCase property names and real DateTime values.
        Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{
                Id = 'id-1'; DisplayName = 'Test Admin'; UserPrincipalName = 'test.admin@example.com'
                UserType = 'Member'; AccountEnabled = $true; CreatedDateTime = (Get-Date).AddYears(-1)
                AssignedLicenses = @(); SignInActivity = [pscustomobject]@{ LastSignInDateTime = (Get-Date).AddDays(-1) }
            }
        }
        Mock -CommandName Get-MgReportAuthenticationMethodUserRegistrationDetail -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ Id = 'id-1'; IsAdmin = $true; IsMfaRegistered = $false; DefaultMfaMethod = 'none'; MethodsRegistered = @() }
        }

        $row = Get-M365MfaStatusReport
        $row.DisplayName | Should -Be 'Test Admin'
        $row.Status | Should -Be 'Critical'
        Should -Invoke -CommandName Get-MgUser -ModuleName M365AdminReports -Times 1 -Exactly -ParameterFilter {
            $All -and $Property -contains 'signInActivity'
        }
    }
}
