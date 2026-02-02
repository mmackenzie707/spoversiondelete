function Trim-SPOVTFileVersions {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$FileRef,
        [Parameter(Mandatory)][string]$FileName,

        [ValidateRange(1, 5000)]
        [int]$KeepVersions = 10,

        [ValidateRange(1, 20)]
        [int]$MaxRetries = 8,

        [ValidateRange(2, 120)]
        [int]$MaxBackoffSeconds = 60
    )

    # Get-PnPFileVersion returns PREVIOUS versions only (not current). [3](https://github.com/pnp/powershell/issues/4608)[4](https://learn.microsoft.com/en-us/answers/questions/1373403/using-client-id-and-secret-to-pull-sharepoint-onli)
    # So: "KeepVersions (total)" = keep current (always) + keep (KeepVersions-1) previous versions.
    $keepPrevious = [Math]::Max(0, $KeepVersions - 1)

    $versions = Invoke-SPOVTRetry -MaxRetries $MaxRetries -MaxBackoffSeconds $MaxBackoffSeconds -Action {
        Get-PnPFileVersion -Url $FileRef
    }

    $previousCount = @($versions).Count
    $totalVersions = $previousCount + 1  # + current version [3](https://github.com/pnp/powershell/issues/4608)[4](https://learn.microsoft.com/en-us/answers/questions/1373403/using-client-id-and-secret-to-pull-sharepoint-onli)

    if ($totalVersions -le $KeepVersions -or $previousCount -le $keepPrevious) {
        return [pscustomobject]@{
            Timestamp        = (Get-Date).ToString('s')
            FileName         = $FileName
            FileRef          = $FileRef
            TotalVersions    = $totalVersions
            PreviousVersions = $previousCount
            KeepVersions     = $KeepVersions
            DeletedVersions  = 0
            Status           = 'Skipped'
            Error            = $null
        }
    }

    $sortedPrev = $versions | Sort-Object Created -Descending
    $toDelete   = $sortedPrev | Select-Object -Skip $keepPrevious
    $deleteCount = @($toDelete).Count

    $target = "$FileRef (delete $deleteCount previous; keep total $KeepVersions)"

    if ($PSCmdlet.ShouldProcess($target, 'Trim versions')) {
        foreach ($v in $toDelete) { $v.DeleteObject() }

        Invoke-SPOVTRetry -MaxRetries $MaxRetries -MaxBackoffSeconds $MaxBackoffSeconds -Action {
            Invoke-PnPQuery
        }
    }

    return [pscustomobject]@{
        Timestamp        = (Get-Date).ToString('s')
        FileName         = $FileName
        FileRef          = $FileRef
        TotalVersions    = $totalVersions
        PreviousVersions = $previousCount
        KeepVersions     = $KeepVersions
        DeletedVersions  = $deleteCount
        Status           = 'OK'
        Error            = $null
    }
}