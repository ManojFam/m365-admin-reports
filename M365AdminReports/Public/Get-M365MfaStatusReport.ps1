function Get-M365MfaStatusReport {
    <#
    .SYNOPSIS
        Lists users and whether they have an MFA method registered. Flags admins without MFA.

    .DESCRIPTION
        Joins the user list with the Entra ID "User registration details" report.
        Each user gets a status:
          Critical  an admin with no MFA method registered
          Warning   a user with no MFA method registered, or an admin whose only MFA method
                    is a phone (SMS / voice), which is the easiest method to phish or SIM-swap
          OK        MFA is registered

        Disabled accounts and guests are left out by default: disabled accounts cannot sign in,
        and guests register MFA in their own tenant.

        Live mode needs Connect-MgGraph with User.Read.All, AuditLog.Read.All and Reports.Read.All.

    .PARAMETER Mock
        Use the fictional sample tenant that ships with the module instead of a live tenant.

    .PARAMETER MockPath
        Use your own folder of mock JSON files (implies -Mock).

    .PARAMETER IncludeGuests
        Also list guest (B2B) accounts.

    .PARAMETER IncludeDisabled
        Also list disabled accounts.

    .EXAMPLE
        Get-M365MfaStatusReport -Mock | Format-Table DisplayName, IsAdmin, MfaRegistered, Status, Finding

    .EXAMPLE
        Connect-MgGraph -Scopes User.Read.All, AuditLog.Read.All, Reports.Read.All
        Get-M365MfaStatusReport | Where-Object Status -ne 'OK' | Export-Csv .\output\mfa-gaps.csv -NoTypeInformation
    #>
    [CmdletBinding()]
    param(
        [switch]$Mock,
        [string]$MockPath,
        [switch]$IncludeGuests,
        [switch]$IncludeDisabled
    )

    $mockFolder = Resolve-MockPath -Mock:$Mock -MockPath $MockPath
    if (-not $mockFolder) { Assert-GraphConnection }

    # Look up each user's registration record by object ID.
    $registration = @{}
    foreach ($record in Get-MfaRegistration -MockPath $mockFolder) {
        $registration[$record.Id] = $record
    }

    # Phone methods count as MFA but are the weakest; email only counts for password reset.
    $phoneMethods = 'mobilePhone', 'alternateMobilePhone', 'officePhone'
    $notMfaMethods = 'email', 'securityQuestion'

    $rows = foreach ($user in Get-DirectoryUser -MockPath $mockFolder) {
        if (-not $IncludeDisabled -and -not $user.AccountEnabled) { continue }
        if (-not $IncludeGuests -and $user.UserType -eq 'Guest') { continue }

        $record = $registration[$user.Id]
        $isAdmin = [bool]($record -and $record.IsAdmin)
        $mfaRegistered = [bool]($record -and $record.IsMfaRegistered)
        $methods = if ($record) { $record.MethodsRegistered } else { @() }
        $strongMethods = @($methods | Where-Object { $phoneMethods -notcontains $_ -and $notMfaMethods -notcontains $_ })

        if (-not $mfaRegistered -and $isAdmin) {
            $status = 'Critical'; $finding = 'Admin account has no MFA method registered'
        }
        elseif (-not $mfaRegistered) {
            $status = 'Warning'; $finding = 'No MFA method registered'
        }
        elseif ($isAdmin -and $strongMethods.Count -eq 0) {
            $status = 'Warning'; $finding = 'Admin relies on phone (SMS/voice) MFA only'
        }
        else {
            $status = 'OK'; $finding = ''
        }

        [pscustomobject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            UserType          = $user.UserType
            AccountEnabled    = $user.AccountEnabled
            IsAdmin           = $isAdmin
            MfaRegistered     = $mfaRegistered
            DefaultMethod     = if ($record) { $record.DefaultMfaMethod } else { '' }
            MethodsRegistered = Join-List $methods
            Status            = $status
            Finding           = $finding
        }
    }

    $rows | Sort-Object -Property (Get-StatusSortKey), DisplayName
}
