BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\M365AdminReports\M365AdminReports.psd1') -Force
    . (Join-Path $PSScriptRoot 'GraphStubs.ps1')
}

Describe 'Get-M365SharedMailboxAudit (mock tenant)' {
    BeforeAll {
        $report = @(Get-M365SharedMailboxAudit -Mock)
        $byName = @{}
        foreach ($row in $report) { $byName[$row.Mailbox] = $row }
    }

    It 'returns every shared mailbox with its Full Access and Send As holders' {
        $report.Count | Should -Be 5
        $byName['Reception'].FullAccess | Should -Be 'avery.quinn@example.com; casey.doyle@example.com'
        $byName['Reception'].SendAs | Should -Be 'casey.doyle@example.com'
        $byName['Reception'].FullAccessCount | Should -Be 2
        $byName['Reception'].Status | Should -Be 'OK'
    }

    It 'warns when sign-in is enabled on the mailbox account' {
        $row = $byName['Student Enquiries']
        $row.SignInBlocked | Should -BeFalse
        $row.Status | Should -Be 'Warning'
        $row.Finding | Should -BeLike '*Sign-in is enabled*'
    }

    It 'warns when a disabled account still has access' {
        $byName['Finance Invoices'].Finding | Should -BeLike '*georgie.lam@example.com*'
        $byName['Student Enquiries'].Finding | Should -BeLike '*harper.singh@example.com*'
    }

    It 'warns when nobody has access' {
        $byName['IT Alerts'].FullAccessCount | Should -Be 0
        $byName['IT Alerts'].Finding | Should -Be 'Nobody has Full Access or Send As'
    }
}

Describe 'Get-M365SharedMailboxAudit (live mode)' {
    BeforeAll { Add-LiveStub }
    AfterAll { Remove-LiveStub }

    It 'fails with a clear message when Exchange Online is not connected' {
        Mock -CommandName Get-ConnectionInformation -ModuleName M365AdminReports -MockWith { }

        { Get-M365SharedMailboxAudit } | Should -Throw '*Connect-ExchangeOnline*'
    }

    It 'reads permissions with the EXO cmdlets and skips system entries' {
        Mock -CommandName Get-ConnectionInformation -ModuleName M365AdminReports -MockWith { [pscustomobject]@{ State = 'Connected' } }
        Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { New-GraphContextMock }
        Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ Id = '1'; DisplayName = 'Help Desk'; UserPrincipalName = 'helpdesk@example.com'; AccountEnabled = $false }
            [pscustomobject]@{ Id = '2'; DisplayName = 'Sam Lee'; UserPrincipalName = 'sam.lee@example.com'; AccountEnabled = $true }
        }
        Mock -CommandName Get-EXOMailbox -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ DisplayName = 'Help Desk'; PrimarySmtpAddress = 'helpdesk@example.com'; UserPrincipalName = 'helpdesk@example.com' }
        }
        Mock -CommandName Get-EXOMailboxPermission -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess', 'ReadPermission'); IsInherited = $false }
            [pscustomobject]@{ User = 'sam.lee@example.com'; AccessRights = @('FullAccess'); IsInherited = $false }
            [pscustomobject]@{ User = 'admin@example.com'; AccessRights = @('FullAccess'); IsInherited = $true }
        }
        Mock -CommandName Get-EXORecipientPermission -ModuleName M365AdminReports -MockWith {
            [pscustomobject]@{ Trustee = 'NT AUTHORITY\SELF'; AccessRights = @('SendAs') }
            [pscustomobject]@{ Trustee = 'sam.lee@example.com'; AccessRights = @('SendAs') }
        }

        $row = Get-M365SharedMailboxAudit
        $row.FullAccess | Should -Be 'sam.lee@example.com'
        $row.SendAs | Should -Be 'sam.lee@example.com'
        $row.SignInBlocked | Should -BeTrue
        $row.Status | Should -Be 'OK'
        Should -Invoke -CommandName Get-EXOMailbox -ModuleName M365AdminReports -Times 1 -Exactly -ParameterFilter {
            $RecipientTypeDetails -eq 'SharedMailbox'
        }
    }
}
