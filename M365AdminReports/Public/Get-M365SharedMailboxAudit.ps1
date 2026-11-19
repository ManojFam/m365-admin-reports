function Get-M365SharedMailboxAudit {
    <#
    .SYNOPSIS
        Lists shared mailboxes and who has Full Access and Send As on each one.

    .DESCRIPTION
        One row per shared mailbox. Each one is checked for three common problems:
          - sign-in is enabled on the mailbox's own account (it should be blocked, so nobody
            can sign in to the shared mailbox directly with a password)
          - a disabled account (usually a leaver) still has Full Access or Send As
          - nobody has access at all (the mailbox may be unused, or mail is going unread)

        Live mode reads the mailboxes and permissions with the Exchange Online cmdlets
        (Get-EXOMailbox, Get-EXOMailboxPermission, Get-EXORecipientPermission) and the account
        status from Microsoft Graph, so it needs Connect-ExchangeOnline and Connect-MgGraph.

    .EXAMPLE
        Get-M365SharedMailboxAudit -Mock | Format-Table Mailbox, FullAccess, SendAs, Status

    .EXAMPLE
        Connect-ExchangeOnline -ShowBanner:$false
        Connect-MgGraph -Scopes User.Read.All
        Get-M365SharedMailboxAudit | Export-Csv .\output\shared-mailboxes.csv -NoTypeInformation
    #>
    [CmdletBinding()]
    param(
        [switch]$Mock,
        [string]$MockPath
    )

    $mockFolder = Resolve-MockPath -Mock:$Mock -MockPath $MockPath
    if (-not $mockFolder) {
        Assert-ExchangeConnection
        Assert-GraphConnection -RequiredScope 'User.Read.All'
    }

    # Index accounts by UPN (lower case) to look up whether each one is enabled.
    $accounts = @{}
    foreach ($user in Get-DirectoryUser -MockPath $mockFolder) {
        $accounts[$user.UserPrincipalName.ToLowerInvariant()] = $user
    }

    $rows = foreach ($mailbox in Get-SharedMailboxData -MockPath $mockFolder) {
        $statuses = @()
        $findings = @()

        $account = $accounts[([string]$mailbox.UserPrincipalName).ToLowerInvariant()]
        $signInBlocked = if ($account) { -not $account.AccountEnabled } else { $null }
        if ($signInBlocked -eq $false) {
            $statuses += 'Warning'; $findings += 'Sign-in is enabled on the shared mailbox account - block it'
        }

        $delegates = @(@($mailbox.FullAccess) + @($mailbox.SendAs) | Where-Object { $_ } | Sort-Object -Unique)
        $disabledDelegates = @($delegates | Where-Object {
                $delegate = $accounts[$_.ToLowerInvariant()]
                $delegate -and -not $delegate.AccountEnabled
            })
        if ($disabledDelegates.Count -gt 0) {
            $statuses += 'Warning'; $findings += "Disabled account(s) still have access: $(Join-List $disabledDelegates)"
        }

        if ($delegates.Count -eq 0) {
            $statuses += 'Warning'; $findings += 'Nobody has Full Access or Send As'
        }

        [pscustomobject]@{
            Mailbox         = $mailbox.DisplayName
            Address         = $mailbox.PrimarySmtpAddress
            FullAccess      = Join-List $mailbox.FullAccess
            SendAs          = Join-List $mailbox.SendAs
            FullAccessCount = @($mailbox.FullAccess).Count
            SendAsCount     = @($mailbox.SendAs).Count
            SignInBlocked   = $signInBlocked
            Status          = Get-WorstStatus $statuses
            Finding         = Join-List $findings
        }
    }

    $rows | Sort-Object -Property (Get-StatusSortKey), Mailbox
}
