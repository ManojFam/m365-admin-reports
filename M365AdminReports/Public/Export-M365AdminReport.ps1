function Export-M365AdminReport {
    <#
    .SYNOPSIS
        Runs every report and saves one HTML page plus a CSV for each report.

    .DESCRIPTION
        Runs Get-M365MfaStatusReport, Get-M365StaleAccountReport, Get-M365LicenceReport and
        Get-M365SharedMailboxAudit, then writes these files to -OutputPath:

          m365-admin-report.html   summary cards and one table per report (self-contained)
          mfa-status.csv           every enabled member account and its MFA status
          stale-accounts.csv       inactive, never-used and disabled-but-licensed accounts
          licences.csv             licence counts per SKU
          shared-mailboxes.csv     shared mailbox access

        To keep the page short, the MFA table in the HTML only lists accounts that need
        attention; the CSV has everyone. Returns the files that were written.

        In live mode the shared mailbox audit needs Connect-ExchangeOnline. If Exchange Online
        is not connected it is skipped with a warning and the other reports still run.

        Live reports contain real names and addresses: keep them out of source control
        (output/ is in .gitignore) and treat them as internal.

    .PARAMETER OutputPath
        Folder to write the files to. It is created if it does not exist.

    .PARAMETER InactiveDays
        Passed to Get-M365StaleAccountReport. Default 90.

    .PARAMETER SkipSharedMailbox
        Do not run the shared mailbox audit.

    .EXAMPLE
        Export-M365AdminReport -Mock -OutputPath .\output
        Invoke-Item .\output\m365-admin-report.html

    .EXAMPLE
        Connect-MgGraph -Scopes User.Read.All, AuditLog.Read.All, Reports.Read.All, Organization.Read.All
        Connect-ExchangeOnline -ShowBanner:$false
        Export-M365AdminReport -OutputPath ".\output\$(Get-Date -Format yyyy-MM-dd)"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$OutputPath,

        [switch]$Mock,
        [string]$MockPath,

        [ValidateRange(1, 3650)]
        [int]$InactiveDays = 90,

        [switch]$SkipSharedMailbox
    )

    $mockFolder = Resolve-MockPath -Mock:$Mock -MockPath $MockPath
    if (-not $mockFolder) { Assert-GraphConnection }

    # Every report reads from the same place: the resolved mock folder, or the live tenant.
    $source = @{}
    if ($mockFolder) { $source.MockPath = $mockFolder }

    $tenant = Get-TenantInfo -MockPath $mockFolder

    $mfa = @(Get-M365MfaStatusReport @source)
    $stale = @(Get-M365StaleAccountReport @source -InactiveDays $InactiveDays)
    $licences = @(Get-M365LicenceReport @source)

    $mailboxSkipped = $null
    $mailboxes = @()
    if ($SkipSharedMailbox) {
        $mailboxSkipped = 'turned off with -SkipSharedMailbox.'
    }
    elseif (-not $mockFolder -and -not (Test-ExchangeConnection)) {
        $mailboxSkipped = 'not connected to Exchange Online (run Connect-ExchangeOnline).'
        Write-Warning "Shared mailbox audit skipped: $mailboxSkipped"
    }
    else {
        $mailboxes = @(Get-M365SharedMailboxAudit @source)
    }

    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory | Out-Null
    }
    $folder = (Resolve-Path -Path $OutputPath).ProviderPath

    $sections = @(
        @{
            Title       = 'MFA registration'
            Description = 'Enabled member accounts and whether they have an MFA method registered. Admins without MFA are Critical.'
            Columns     = 'DisplayName', 'UserPrincipalName', 'IsAdmin', 'MethodsRegistered', 'Status', 'Finding'
            Rows        = $mfa
            HideOk      = $true
            File        = 'mfa-status.csv'
        }
        @{
            Title       = 'Stale accounts'
            Description = "Enabled accounts with no sign-in for $InactiveDays days or more, accounts that have never been used, and disabled accounts that still hold a licence."
            Columns     = 'DisplayName', 'UserPrincipalName', 'LastSignIn', 'DaysSinceSignIn', 'Licences', 'Finding', 'Status', 'Recommendation'
            Rows        = $stale
            File        = 'stale-accounts.csv'
        }
        @{
            Title       = 'Licences'
            Description = 'Licences bought vs assigned for each SKU, and licences sitting on accounts with sign-in blocked.'
            Columns     = 'Licence', 'Total', 'Assigned', 'Available', 'PercentUsed', 'AssignedToBlocked', 'BlockedUsers', 'Status', 'Finding'
            Rows        = $licences
            File        = 'licences.csv'
        }
        @{
            Title       = 'Shared mailboxes'
            Description = 'Who has Full Access and Send As on each shared mailbox, and whether sign-in is blocked on the mailbox account.'
            Columns     = 'Mailbox', 'Address', 'FullAccess', 'SendAs', 'SignInBlocked', 'Status', 'Finding'
            Rows        = $mailboxes
            Skipped     = $mailboxSkipped
            File        = 'shared-mailboxes.csv'
        }
    )

    foreach ($section in $sections) {
        if ($section.Skipped) { continue }
        $csvPath = Join-Path -Path $folder -ChildPath $section.File
        if ($section.Rows.Count -gt 0) {
            $section.Rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }
        else {
            Set-Content -Path $csvPath -Value '' -Encoding UTF8
        }
        Get-Item -Path $csvPath
    }

    $htmlPath = Join-Path -Path $folder -ChildPath 'm365-admin-report.html'
    ConvertTo-AdminReportHtml -Tenant $tenant -Section $sections -GeneratedAt (Get-Date) |
        Set-Content -Path $htmlPath -Encoding UTF8
    Get-Item -Path $htmlPath
}
