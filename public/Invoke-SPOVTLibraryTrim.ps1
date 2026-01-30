function Invoke-SPOVTLibraryTrim {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$LibraryName,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId,

        [ValidateRange(1, 5000)]
        [int]$KeepVersions = 10,

        [ValidateRange(100, 5000)]
        [int]$PageSize = 2000,

        [ValidateRange(1, 20)]
        [int]$MaxRetries = 8,

        [ValidateRange(2, 120)]
        [int]$MaxBackoffSeconds = 60,

        [Parameter()]
        [string]$ReportPath = (Join-Path $PWD ("VersionTrim_{0}_{1:yyyyMMdd_HHmmss}.csv" -f ($LibraryName -replace '[\\/:*?"<>|]','_'), (Get-Date))),

        [switch]$NoProgress
    )

    Connect-SPOVTDeviceLogin -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId

    $library = Get-PnPList -Identity $LibraryName
    if ($library.BaseTemplate -ne 101) { throw "The specified list '$LibraryName' is not a document library." }

    $processed = 0
    $skipped   = 0
    $failed    = 0
    $deletedTotal = 0

    Get-SPOVTLibraryFilesPaged -Library $library -PageSize $PageSize -NoProgress:$NoProgress | ForEach-Object {
        $processed++
        $fileRef  = $_['FileRef']
        $fileName = $_['FileLeafRef']

        try {
            $row = Trim-SPOVTFileVersions -FileRef $fileRef -FileName $fileName -KeepVersions $KeepVersions `
                                          -MaxRetries $MaxRetries -MaxBackoffSeconds $MaxBackoffSeconds `
                                          -WhatIf:$WhatIfPreference

            if ($row.Status -eq 'Skipped') { $skipped++ }
            if ($row.Status -eq 'OK') { $deletedTotal += [int]$row.DeletedVersions }

            Write-SPOVTReportRow -Path $ReportPath -Row $row
        }
        catch {
            $failed++
            Write-SPOVTReportRow -Path $ReportPath -Row ([pscustomobject]@{
                Timestamp        = (Get-Date).ToString('s')
                FileName         = $fileName
                FileRef          = $fileRef
                TotalVersions    = $null
                PreviousVersions = $null
                KeepVersions     = $KeepVersions
                DeletedVersions  = 0
                Status           = 'FAILED'
                Error            = $_.Exception.Message
            })
        }
    }

    [pscustomobject]@{
        SiteUrl        = $SiteUrl
        LibraryName    = $LibraryName
        KeepVersions   = $KeepVersions
        ProcessedFiles = $processed
        SkippedFiles   = $skipped
        FailedFiles    = $failed
        DeletedVersionsTotal = $deletedTotal
        ReportPath     = $ReportPath
    }
}