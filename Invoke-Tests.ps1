<#
.SYNOPSIS
    Runs PSScriptAnalyzer (lint) and the Pester tests. Used locally and by GitHub Actions.

.EXAMPLE
    .\Invoke-Tests.ps1

.EXAMPLE
    .\Invoke-Tests.ps1 -SkipAnalyzer
#>
[CmdletBinding()]
param([switch]$SkipAnalyzer)

$ErrorActionPreference = 'Stop'

if (-not $SkipAnalyzer) {
    Import-Module PSScriptAnalyzer
    $issues = @(Invoke-ScriptAnalyzer -Path (Join-Path $PSScriptRoot 'M365AdminReports') -Recurse -Settings (Join-Path $PSScriptRoot 'PSScriptAnalyzerSettings.psd1'))
    if ($issues.Count -gt 0) {
        $issues | Format-Table -Property Severity, RuleName, ScriptName, Line, Message -AutoSize | Out-String -Width 200 | Write-Output
        throw "PSScriptAnalyzer found $($issues.Count) issue(s)."
    }
    Write-Output 'PSScriptAnalyzer: no issues.'
}

Import-Module Pester -MinimumVersion 5.5.0 -MaximumVersion 5.99.99

$config = New-PesterConfiguration
$config.Run.Path = Join-Path $PSScriptRoot 'tests'
$config.Run.Exit = $true
$config.Output.Verbosity = 'Detailed'
Invoke-Pester -Configuration $config
