BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\M365AdminReports\M365AdminReports.psd1') -Force
    . (Join-Path $PSScriptRoot 'GraphStubs.ps1')
}

Describe 'Export-M365AdminReport (mock tenant)' {
    BeforeAll {
        $out = Join-Path $TestDrive 'report'
        $files = @(Export-M365AdminReport -Mock -OutputPath $out)
        $html = Get-Content -Path (Join-Path $out 'm365-admin-report.html') -Raw
    }

    It 'creates the folder and writes one HTML page and four CSVs' {
        $files.Name | Sort-Object | Should -Be @(
            'licences.csv', 'm365-admin-report.html', 'mfa-status.csv', 'shared-mailboxes.csv', 'stale-accounts.csv'
        )
    }

    It 'writes every row to the CSVs' {
        @(Import-Csv -Path (Join-Path $out 'mfa-status.csv')).Count | Should -Be 14
        @(Import-Csv -Path (Join-Path $out 'stale-accounts.csv')).Count | Should -Be 7
        @(Import-Csv -Path (Join-Path $out 'licences.csv')).Count | Should -Be 5
        @(Import-Csv -Path (Join-Path $out 'shared-mailboxes.csv')).Count | Should -Be 5
    }

    It 'has a section for each report and the tenant name' {
        foreach ($title in 'MFA registration', 'Stale accounts', 'Licences', 'Shared mailboxes') {
            $html | Should -Match "<h2>$title</h2>"
        }
        $html | Should -Match 'Example Training Institute \(fictional\)'
        $html | Should -Match 'Data as of 2026-11-16'
    }

    It 'only lists MFA rows that need attention in the HTML' {
        $html | Should -Match 'blake\.harper@example\.com'
        $html | Should -Not -Match 'rory\.chen@example\.com</td><td>False</td><td>microsoftAuthenticatorPush'
        $html | Should -Match '8 row\(s\) with status OK are in the CSV only'
    }

    It 'is self-contained (no external scripts or stylesheets)' {
        $html | Should -Not -Match '<script'
        $html | Should -Not -Match '<link'
    }

    It 'passes -InactiveDays through to the stale account report' {
        $strictOut = Join-Path $TestDrive 'strict'
        Export-M365AdminReport -Mock -OutputPath $strictOut -InactiveDays 30 | Out-Null
        @(Import-Csv -Path (Join-Path $strictOut 'stale-accounts.csv')).Count | Should -BeGreaterThan 7
    }

    It 'skips the shared mailbox audit with -SkipSharedMailbox' {
        $skipOut = Join-Path $TestDrive 'skip'
        $skipped = @(Export-M365AdminReport -Mock -OutputPath $skipOut -SkipSharedMailbox)
        $skipped.Name | Should -Not -Contain 'shared-mailboxes.csv'
        Get-Content -Path (Join-Path $skipOut 'm365-admin-report.html') -Raw | Should -Match 'Skipped: turned off'
    }

    It 'HTML-encodes values from the directory' {
        $folder = Join-Path $TestDrive 'evil'
        Copy-Item -Path (Join-Path $PSScriptRoot '..\M365AdminReports\mock') -Destination $folder -Recurse
        $usersFile = Join-Path $folder 'users.json'
        (Get-Content -Path $usersFile -Raw) -replace '"Blake Harper"', '"<script>alert(1)</script> & Co"' |
            Set-Content -Path $usersFile -Encoding UTF8

        Export-M365AdminReport -MockPath $folder -OutputPath (Join-Path $folder 'out') | Out-Null
        $evil = Get-Content -Path (Join-Path $folder 'out\m365-admin-report.html') -Raw
        $evil | Should -Not -Match '<script>'
        $evil | Should -Match '&lt;script&gt;alert\(1\)&lt;/script&gt; &amp; Co'
    }
}

Describe 'Export-M365AdminReport (live mode)' {
    It 'refuses to run without a Graph connection' {
        Add-LiveStub
        try {
            Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { $null }
            { Export-M365AdminReport -OutputPath (Join-Path $TestDrive 'live') } | Should -Throw '*Connect-MgGraph*'
        }
        finally {
            Remove-LiveStub
        }
    }

    It 'skips the shared mailbox audit with a warning when Exchange Online is not connected' {
        Add-LiveStub
        try {
            Mock -CommandName Get-MgContext -ModuleName M365AdminReports -MockWith { New-GraphContextMock }
            Mock -CommandName Get-MgOrganization -ModuleName M365AdminReports -MockWith { [pscustomobject]@{ DisplayName = 'Test tenant' } }
            Mock -CommandName Get-MgUser -ModuleName M365AdminReports -MockWith { }
            Mock -CommandName Get-MgSubscribedSku -ModuleName M365AdminReports -MockWith { }
            Mock -CommandName Get-MgReportAuthenticationMethodUserRegistrationDetail -ModuleName M365AdminReports -MockWith { }
            Mock -CommandName Get-ConnectionInformation -ModuleName M365AdminReports -MockWith { }
            Mock -CommandName Get-EXOMailbox -ModuleName M365AdminReports -MockWith { }

            $liveOut = Join-Path $TestDrive 'live'
            $files = @(Export-M365AdminReport -OutputPath $liveOut -WarningVariable warnings -WarningAction SilentlyContinue)
            "$warnings" | Should -BeLike '*Shared mailbox audit skipped*'
            $files.Name | Should -Contain 'm365-admin-report.html'
            $files.Name | Should -Not -Contain 'shared-mailboxes.csv'
            Should -Invoke -CommandName Get-EXOMailbox -ModuleName M365AdminReports -Times 0 -Exactly
        }
        finally {
            Remove-LiveStub
        }
    }
}
