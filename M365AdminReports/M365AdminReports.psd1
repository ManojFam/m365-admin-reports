@{
    RootModule           = 'M365AdminReports.psm1'
    ModuleVersion        = '1.0.0'
    GUID                 = '7c2f4b9e-3d1a-4e8f-9b6c-5a0d2e7f1c43'
    Author               = 'Manoj Kandel'
    Copyright            = '(c) 2026 Manoj Kandel. MIT License.'
    Description          = 'Read-only Microsoft 365 / Entra ID admin reports (MFA registration, stale accounts, licences, shared mailbox access) using Microsoft Graph, with a mock mode that runs without a tenant.'
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    FunctionsToExport    = @(
        'Get-M365MfaStatusReport'
        'Get-M365StaleAccountReport'
        'Get-M365LicenceReport'
        'Get-M365SharedMailboxAudit'
        'Export-M365AdminReport'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()
    PrivateData          = @{
        PSData = @{
            Tags       = @('Microsoft365', 'EntraID', 'MicrosoftGraph', 'MFA', 'Licensing', 'ExchangeOnline', 'Reporting')
            ProjectUri = 'https://github.com/ManojFam/m365-admin-reports'
            LicenseUri = 'https://github.com/ManojFam/m365-admin-reports/blob/main/LICENSE'
        }
    }
}
