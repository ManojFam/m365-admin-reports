# The Microsoft Graph and Exchange Online modules are not installed on the CI runners, and
# Pester can only mock a command that exists. These empty global functions stand in for the
# real cmdlets (with the parameters this module uses) so the live-mode branches can be tested
# with Mock.

$script:Stubs = @{
    'Get-MgContext'                                          = 'param()'
    'Get-MgOrganization'                                     = 'param()'
    'Get-MgUser'                                             = 'param([switch]$All, [string[]]$Property)'
    'Get-MgSubscribedSku'                                    = 'param([switch]$All)'
    'Get-MgReportAuthenticationMethodUserRegistrationDetail' = 'param([switch]$All)'
    'Get-EXOMailbox'                                         = 'param($RecipientTypeDetails, $ResultSize)'
    'Get-EXOMailboxPermission'                               = 'param($Identity)'
    'Get-EXORecipientPermission'                             = 'param($Identity, $AccessRights)'
    'Get-ConnectionInformation'                              = 'param()'
}

function Add-LiveStub {
    foreach ($name in $script:Stubs.Keys) {
        Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create($script:Stubs[$name]))
    }
}

function Remove-LiveStub {
    foreach ($name in $script:Stubs.Keys) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
}

function New-GraphContextMock {
    # A delegated Graph sign-in that has all four read-only scopes.
    [pscustomobject]@{
        AuthType = 'Delegated'
        Scopes   = @('User.Read.All', 'AuditLog.Read.All', 'Reports.Read.All', 'Organization.Read.All')
    }
}
