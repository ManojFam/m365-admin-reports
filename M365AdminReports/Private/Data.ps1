# The data layer. Each function has two branches:
#   mock  read a JSON file from the mock folder (same shape as the Graph / Exchange response)
#   live  call the Microsoft Graph or Exchange Online cmdlet
# Both branches return the same simple objects, so the report logic is written once and the
# mock data runs exactly the same code path as a real tenant.

function Read-MockFile {
    param(
        [Parameter(Mandatory)][string]$MockPath,
        [Parameter(Mandatory)][string]$Name
    )

    $file = Join-Path -Path $MockPath -ChildPath "$Name.json"
    if (-not (Test-Path -Path $file -PathType Leaf)) {
        throw "Mock data file not found: $file"
    }
    $data = ConvertFrom-Json -InputObject (Get-Content -Path $file -Raw -Encoding UTF8)
    # Unroll arrays so callers get one object per item on 5.1 and 7 alike.
    foreach ($item in $data) { $item }
}

function ConvertTo-UtcDate {
    # Graph returns DateTime objects; JSON gives ISO 8601 strings (PowerShell 7 converts those
    # to DateTime itself, 5.1 does not). Everything is compared in UTC.
    param($Value)

    if ($null -eq $Value -or "$Value" -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $styles = [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    return [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture, $styles)
}

function Get-TenantInfo {
    param([string]$MockPath)

    if ($MockPath) {
        $tenant = Read-MockFile -MockPath $MockPath -Name 'tenant'
        return [pscustomobject]@{
            Name          = [string]$tenant.displayName
            Source        = 'Mock data'
            # Mock data is a snapshot, so ages are measured from the day it was taken.
            ReferenceDate = ConvertTo-UtcDate $tenant.snapshotDate
        }
    }

    $org = Get-MgOrganization | Select-Object -First 1
    [pscustomobject]@{
        Name          = [string]$org.DisplayName
        Source        = 'Microsoft Graph'
        ReferenceDate = (Get-Date).ToUniversalTime()
    }
}

function Get-DirectoryUser {
    param([string]$MockPath)

    if ($MockPath) {
        $raw = Read-MockFile -MockPath $MockPath -Name 'users'
    }
    else {
        # signInActivity needs AuditLog.Read.All and an Entra ID P1/P2 licence in the tenant.
        $properties = 'id', 'displayName', 'userPrincipalName', 'userType', 'accountEnabled',
            'department', 'createdDateTime', 'assignedLicenses', 'signInActivity'
        $raw = Get-MgUser -All -Property $properties
    }

    foreach ($user in $raw) {
        $userType = if ($user.userType) { [string]$user.userType } else { 'Member' }
        [pscustomobject]@{
            Id                       = [string]$user.id
            DisplayName              = [string]$user.displayName
            UserPrincipalName        = [string]$user.userPrincipalName
            UserType                 = $userType
            AccountEnabled           = [bool]$user.accountEnabled
            Department               = [string]$user.department
            CreatedDateTime          = ConvertTo-UtcDate $user.createdDateTime
            LastInteractiveSignIn    = ConvertTo-UtcDate $user.signInActivity.lastSignInDateTime
            LastNonInteractiveSignIn = ConvertTo-UtcDate $user.signInActivity.lastNonInteractiveSignInDateTime
            SkuIds                   = @($user.assignedLicenses | ForEach-Object { [string]$_.skuId } | Where-Object { $_ })
        }
    }
}

function Get-MfaRegistration {
    param([string]$MockPath)

    if ($MockPath) {
        $raw = Read-MockFile -MockPath $MockPath -Name 'mfa-registration'
    }
    else {
        # The "User registration details" report (Reports.Read.All + AuditLog.Read.All).
        $raw = Get-MgReportAuthenticationMethodUserRegistrationDetail -All
    }

    foreach ($row in $raw) {
        [pscustomobject]@{
            Id                = [string]$row.id
            UserPrincipalName = [string]$row.userPrincipalName
            IsAdmin           = [bool]$row.isAdmin
            IsMfaRegistered   = [bool]$row.isMfaRegistered
            DefaultMfaMethod  = [string]$row.defaultMfaMethod
            MethodsRegistered = @($row.methodsRegistered | ForEach-Object { [string]$_ } | Where-Object { $_ })
        }
    }
}

function Get-SubscribedSkuData {
    param([string]$MockPath)

    if ($MockPath) {
        $raw = Read-MockFile -MockPath $MockPath -Name 'subscribed-skus'
    }
    else {
        $raw = Get-MgSubscribedSku -All
    }

    foreach ($sku in $raw) {
        [pscustomobject]@{
            SkuId            = [string]$sku.skuId
            SkuPartNumber    = [string]$sku.skuPartNumber
            CapabilityStatus = [string]$sku.capabilityStatus
            Total            = [int]$sku.prepaidUnits.enabled
            Assigned         = [int]$sku.consumedUnits
        }
    }
}

function Get-SharedMailboxData {
    param([string]$MockPath)

    if ($MockPath) {
        foreach ($mailbox in Read-MockFile -MockPath $MockPath -Name 'shared-mailboxes') {
            [pscustomobject]@{
                DisplayName        = [string]$mailbox.displayName
                PrimarySmtpAddress = [string]$mailbox.primarySmtpAddress
                UserPrincipalName  = [string]$mailbox.userPrincipalName
                FullAccess         = @($mailbox.fullAccess | ForEach-Object { [string]$_ })
                SendAs             = @($mailbox.sendAs | ForEach-Object { [string]$_ })
            }
        }
        return
    }

    # Built-in system entries (NT AUTHORITY\SELF and orphaned SIDs) are not people, so skip them.
    $isSystem = { param($Name) $Name -like 'NT AUTHORITY\*' -or $Name -like 'S-1-5-*' }

    foreach ($mailbox in Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited) {
        $identity = [string]$mailbox.PrimarySmtpAddress

        $fullAccess = @(Get-EXOMailboxPermission -Identity $identity |
            Where-Object { ($_.AccessRights -join ',') -match 'FullAccess' -and -not $_.IsInherited -and -not (& $isSystem $_.User) } |
            ForEach-Object { [string]$_.User })

        $sendAs = @(Get-EXORecipientPermission -Identity $identity -AccessRights SendAs |
            Where-Object { -not (& $isSystem $_.Trustee) } |
            ForEach-Object { [string]$_.Trustee })

        [pscustomobject]@{
            DisplayName        = [string]$mailbox.DisplayName
            PrimarySmtpAddress = $identity
            UserPrincipalName  = [string]$mailbox.UserPrincipalName
            FullAccess         = $fullAccess
            SendAs             = $sendAs
        }
    }
}
