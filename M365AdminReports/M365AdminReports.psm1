# Module loader: every function lives in its own file.
#   Private\  helpers used inside the module (not exported)
#   Public\   the commands people run (one file per command, file name = command name)
#   mock\     fictional sample tenant used by -Mock

$script:ModuleRoot = $PSScriptRoot

foreach ($folder in 'Private', 'Public') {
    Get-ChildItem -Path (Join-Path $PSScriptRoot $folder) -Filter '*.ps1' | ForEach-Object {
        . $_.FullName
    }
}

$publicFunctions = Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter '*.ps1' |
    Select-Object -ExpandProperty BaseName
Export-ModuleMember -Function $publicFunctions
