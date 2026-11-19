# Decides where the data comes from (mock JSON files or a live tenant) and checks that the
# live connections exist before any report runs.

# Read-only Microsoft Graph permissions. Nothing in this module changes the tenant.
#   User.Read.All          users, their licences and account status
#   AuditLog.Read.All      signInActivity (last sign-in dates) on each user
#   Reports.Read.All       the MFA / authentication methods registration report
#   Organization.Read.All  subscribed licences (SKUs) and the tenant name
$script:GraphScopes = @('User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All', 'Organization.Read.All')

function Resolve-MockPath {
    # Returns the folder of mock JSON files to read, or $null for live mode.
    # -MockPath on its own also turns mock mode on.
    param(
        [switch]$Mock,
        [string]$MockPath
    )

    if ($MockPath) {
        if (-not (Test-Path -Path $MockPath -PathType Container)) {
            throw "Mock data folder not found: $MockPath"
        }
        return (Resolve-Path -Path $MockPath).ProviderPath
    }
    if ($Mock) {
        return (Join-Path -Path $script:ModuleRoot -ChildPath 'mock')
    }
    return $null
}

function Assert-GraphConnection {
    # Stops with a clear message when the Graph SDK is missing or Connect-MgGraph has not been
    # run. Warns (but carries on) when a delegated sign-in is missing one of the scopes.
    param([string[]]$RequiredScope = $script:GraphScopes)

    if (-not (Get-Command -Name 'Get-MgContext' -ErrorAction SilentlyContinue)) {
        throw ('The Microsoft Graph PowerShell SDK is not installed. Install it with: ' +
            'Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Reports, ' +
            'Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser. ' +
            'Or add -Mock to run the report against the sample data.')
    }

    $context = Get-MgContext
    if (-not $context) {
        throw ("Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes $($RequiredScope -join ', ') " +
            '(or add -Mock to run the report against the sample data).')
    }

    if ($context.AuthType -eq 'Delegated') {
        $missing = @($RequiredScope | Where-Object { @($context.Scopes) -notcontains $_ })
        if ($missing.Count -gt 0) {
            Write-Warning ("The Graph sign-in is missing these scopes: $($missing -join ', '). " +
                'Some data may be empty or the request may fail. Reconnect with Connect-MgGraph -Scopes ' +
                ($RequiredScope -join ', '))
        }
    }
}

function Test-ExchangeConnection {
    # True when the ExchangeOnlineManagement module is loaded and Connect-ExchangeOnline has run.
    if (-not (Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)) {
        return $false
    }
    if (Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue) {
        return [bool](Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })
    }
    return $true
}

function Assert-ExchangeConnection {
    if (-not (Test-ExchangeConnection)) {
        throw ('Not connected to Exchange Online. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser; ' +
            'Connect-ExchangeOnline (a read-only role such as View-Only Recipients or Global Reader is enough). ' +
            'Or add -Mock to run the audit against the sample data.')
    }
}
