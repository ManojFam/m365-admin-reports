function Get-M365StaleAccountReport {
    <#
    .SYNOPSIS
        Finds enabled accounts nobody has signed in to for a while, and disabled accounts
        that still hold a licence.

    .DESCRIPTION
        Returns one row per account that needs a look. Each row has a Finding:
          Inactive          enabled, last sign-in at least -InactiveDays ago
          NeverSignedIn     enabled, never signed in, and created at least -InactiveDays ago
                            (brand-new starters who have not signed in yet are not flagged)
          DisabledLicensed  disabled, but still has a licence assigned (money being wasted)

        "Last sign-in" is the later of the interactive and non-interactive sign-in dates, so a
        user whose phone keeps Outlook connected is not reported as inactive.

        In mock mode ages are measured from the date the sample data was taken (in tenant.json),
        so the results do not change as time passes. Use -ReferenceDate to pick another date.

        Live mode needs Connect-MgGraph with User.Read.All, AuditLog.Read.All and
        Organization.Read.All. Sign-in dates need an Entra ID P1 or P2 licence in the tenant.

    .PARAMETER InactiveDays
        How many days without a sign-in counts as inactive. Default 90.

    .PARAMETER ReferenceDate
        Measure ages from this date instead of today (live) or the snapshot date (mock).

    .EXAMPLE
        Get-M365StaleAccountReport -Mock | Format-Table DisplayName, LastSignIn, DaysSinceSignIn, Finding

    .EXAMPLE
        Get-M365StaleAccountReport -InactiveDays 60 | Export-Csv .\output\stale.csv -NoTypeInformation
    #>
    [CmdletBinding()]
    param(
        [switch]$Mock,
        [string]$MockPath,

        [ValidateRange(1, 3650)]
        [int]$InactiveDays = 90,

        [datetime]$ReferenceDate
    )

    $mockFolder = Resolve-MockPath -Mock:$Mock -MockPath $MockPath
    if (-not $mockFolder) { Assert-GraphConnection }

    $asOf = if ($PSBoundParameters.ContainsKey('ReferenceDate')) {
        $ReferenceDate.ToUniversalTime()
    }
    else {
        (Get-TenantInfo -MockPath $mockFolder).ReferenceDate
    }

    $skuNames = @{}
    foreach ($sku in Get-SubscribedSkuData -MockPath $mockFolder) {
        $skuNames[$sku.SkuId] = Get-SkuFriendlyName $sku.SkuPartNumber
    }

    $recommendation = @{
        Inactive         = 'Check with the manager; disable the account if the person has left.'
        NeverSignedIn    = 'Check the account is still needed; disable it if not.'
        DisabledLicensed = 'Remove the licence, or note why it is kept (e.g. a mailbox over 50 GB).'
    }

    $rows = foreach ($user in Get-DirectoryUser -MockPath $mockFolder) {
        $lastSignIn = @($user.LastInteractiveSignIn, $user.LastNonInteractiveSignIn) |
            Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1
        $daysSince = if ($lastSignIn) { [int][math]::Floor(($asOf - $lastSignIn).TotalDays) } else { $null }

        $finding = $null
        if ($user.AccountEnabled) {
            if ($lastSignIn) {
                if ($daysSince -ge $InactiveDays) { $finding = 'Inactive' }
            }
            elseif (-not $user.CreatedDateTime -or ($asOf - $user.CreatedDateTime).TotalDays -ge $InactiveDays) {
                $finding = 'NeverSignedIn'
            }
        }
        elseif ($user.SkuIds.Count -gt 0) {
            $finding = 'DisabledLicensed'
        }

        if (-not $finding) { continue }

        $licences = foreach ($skuId in $user.SkuIds) {
            if ($skuNames.ContainsKey($skuId)) { $skuNames[$skuId] } else { $skuId }
        }

        [pscustomobject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            UserType          = $user.UserType
            AccountEnabled    = $user.AccountEnabled
            Department        = $user.Department
            LastSignIn        = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { 'Never' }
            DaysSinceSignIn   = $daysSince
            Licences          = Join-List $licences
            Finding           = $finding
            Status            = 'Warning'
            Recommendation    = $recommendation[$finding]
        }
    }

    $rows | Sort-Object -Property Finding, DisplayName
}
