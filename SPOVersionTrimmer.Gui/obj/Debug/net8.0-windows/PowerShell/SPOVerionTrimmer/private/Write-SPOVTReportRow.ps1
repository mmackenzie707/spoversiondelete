function Write-SPOVTReportRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [pscustomobject]$Row
    )

    if (-not (Test-Path $Path)) {
        $Row | Export-Csv -Path $Path -NoTypeInformation
    } else {
        $Row | Export-Csv -Path $Path -NoTypeInformation -Append
    }
}