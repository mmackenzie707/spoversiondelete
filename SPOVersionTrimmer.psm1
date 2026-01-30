# SPOVersionTrimmer.psm1

Set-StrictMode -Version Latest

# Load private functions
Get-ChildItem -Path (Join-Path $PSScriptRoot 'private') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    ForEach-Object { . $_.FullName }

# Load public functions
Get-ChildItem -Path (Join-Path $PSScriptRoot 'public') -Filter '*.ps1' -ErrorAction SilentlyContinue |
    ForEach-Object { . $_.FullName }

Export-ModuleMember -Function @(
    'Connect-SPOVTDeviceLogin',
    'Get-SPOVTLibraryFilesPaged',
    'Trim-SPOVTFileVersions',
    'Invoke-SPOVTLibraryTrim'
)