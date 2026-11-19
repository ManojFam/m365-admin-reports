function Get-M365LicenceReport {
    <#
    .SYNOPSIS
        Shows assigned vs available licences for each SKU, and which licences sit on accounts
        with sign-in blocked.

    .DESCRIPTION
        One row per licence SKU in the tenant:
          Total              licences bought (prepaid units that are enabled)
          Assigned           licences given to users (consumed units)
          Available          Total - Assigned
          AssignedToBlocked  how many are on accounts whose sign-in is blocked (disabled)
          BlockedUsers       those accounts, so the licences can be reclaimed

        Status rules:
          Critical  more licences assigned than bought, or the subscription is suspended / locked out
          Warning   none left to give out, subscription in its grace period ("Warning" state),
                    or licences assigned to blocked accounts
          OK        otherwise

        Live mode needs Connect-MgGraph with User.Read.All and Organization.Read.All.

    .EXAMPLE
        Get-M365LicenceReport -Mock | Format-Table Licence, Total, Assigned, Available, AssignedToBlocked, Status

    .EXAMPLE
        Get-M365LicenceReport | Where-Object AssignedToBlocked -gt 0 | Select-Object Licence, BlockedUsers
    #>
    [CmdletBinding()]
    param(
        [switch]$Mock,
        [string]$MockPath
    )

    $mockFolder = Resolve-MockPath -Mock:$Mock -MockPath $MockPath
    if (-not $mockFolder) { Assert-GraphConnection }

    $blockedUsers = @(Get-DirectoryUser -MockPath $mockFolder | Where-Object { -not $_.AccountEnabled })

    $rows = foreach ($sku in Get-SubscribedSkuData -MockPath $mockFolder) {
        $available = $sku.Total - $sku.Assigned
        $percentUsed = if ($sku.Total -gt 0) { [math]::Round($sku.Assigned / $sku.Total * 100) } else { 0 }
        $blocked = @($blockedUsers | Where-Object { $_.SkuIds -contains $sku.SkuId } |
            Sort-Object -Property UserPrincipalName | Select-Object -ExpandProperty UserPrincipalName)

        $statuses = @()
        $findings = @()
        switch ($sku.CapabilityStatus) {
            'Enabled' { }
            'Warning' { $statuses += 'Warning'; $findings += 'Subscription is in its grace period - renew or it will be suspended' }
            default { $statuses += 'Critical'; $findings += "Subscription status is $($sku.CapabilityStatus)" }
        }
        if ($available -lt 0) {
            $statuses += 'Critical'; $findings += "Over-assigned by $(-$available)"
        }
        elseif ($available -eq 0) {
            $statuses += 'Warning'; $findings += 'No licences left to assign'
        }
        if ($blocked.Count -gt 0) {
            $statuses += 'Warning'; $findings += "$($blocked.Count) assigned to blocked account(s) - reclaim"
        }

        [pscustomobject]@{
            Licence            = Get-SkuFriendlyName $sku.SkuPartNumber
            SkuPartNumber      = $sku.SkuPartNumber
            SubscriptionStatus = $sku.CapabilityStatus
            Total              = $sku.Total
            Assigned           = $sku.Assigned
            Available          = $available
            PercentUsed        = $percentUsed
            AssignedToBlocked  = $blocked.Count
            BlockedUsers       = Join-List $blocked
            Status             = Get-WorstStatus $statuses
            Finding            = Join-List $findings
        }
    }

    $rows | Sort-Object -Property (Get-StatusSortKey), Licence
}
