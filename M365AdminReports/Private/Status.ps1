# Shared rules for the OK / Warning / Critical status every report row gets.

$script:StatusRank = @{ OK = 0; Warning = 1; Critical = 2 }

function Get-WorstStatus {
    # The highest of the given statuses, e.g. OK + Warning = Warning. No input = OK.
    param([string[]]$Status)

    $worst = 'OK'
    foreach ($item in $Status) {
        if ($item -and $script:StatusRank[$item] -gt $script:StatusRank[$worst]) { $worst = $item }
    }
    return $worst
}

function Get-StatusSortKey {
    # Sorts Critical first, then Warning, then OK: Sort-Object -Property (Get-StatusSortKey), Name
    @{ Expression = { $script:StatusRank[$_.Status] }; Descending = $true }
}

function Join-List {
    # Arrays are joined into one string so they read well in a table and a CSV cell.
    param([object[]]$Item)

    (@($Item) | Where-Object { $_ }) -join '; '
}

# Friendly names for common licence SKUs. Anything not listed shows its part number.
# The full list is in Microsoft's "Product names and service plan identifiers for licensing".
$script:SkuFriendlyName = @{
    'STANDARDWOFFPACK_FACULTY' = 'Office 365 A1 for faculty'
    'STANDARDWOFFPACK_STUDENT' = 'Office 365 A1 for students'
    'M365EDU_A3_FACULTY'       = 'Microsoft 365 A3 for faculty'
    'M365EDU_A3_STUDENT'       = 'Microsoft 365 A3 for students'
    'M365EDU_A5_FACULTY'       = 'Microsoft 365 A5 for faculty'
    'SPE_E3'                   = 'Microsoft 365 E3'
    'SPE_E5'                   = 'Microsoft 365 E5'
    'SPB'                      = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'    = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
    'ENTERPRISEPACK'           = 'Office 365 E3'
    'EXCHANGESTANDARD'         = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'       = 'Exchange Online (Plan 2)'
    'POWER_BI_PRO'             = 'Power BI Pro'
    'POWER_BI_STANDARD'        = 'Power BI (free)'
    'FLOW_FREE'                = 'Power Automate Free'
    'AAD_PREMIUM'              = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'           = 'Microsoft Entra ID P2'
    'EMS'                      = 'Enterprise Mobility + Security E3'
    'INTUNE_A'                 = 'Microsoft Intune Plan 1'
}

function Get-SkuFriendlyName {
    param([string]$SkuPartNumber)

    if ($script:SkuFriendlyName.ContainsKey($SkuPartNumber)) { return $script:SkuFriendlyName[$SkuPartNumber] }
    return $SkuPartNumber
}
