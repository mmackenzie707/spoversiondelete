<#
.SYNOPSIS
  Delete all but the latest 4 versions of every file in a SharePoint library.
  Uses app-only authentication (client secret).

.DEPENDENCIES
  Install-Module -Name PnP.PowerShell -Force -Scope CurrentUser
#>

#region ---------------- APP-ONLY AUTH ----------------------------------
function Connect-SharePoint {
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,
        [Parameter(Mandatory=$true)]
        [string]$ClientId,
        [Parameter(Mandatory=$true)]
        [string]$TenantId
    )
    # 2.12 device-code (delegated) login – secret no longer needed
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -DeviceLogin
}
#endregion -------------- END AUTH --------------------------------------

#region ---------------- BUSINESS LOGIC ---------------------------------
function Remove-Versions {
    param (
        [array]$VersionsToDelete,
        [Microsoft.SharePoint.Client.File]$FileItem
    )
    foreach ($version in $VersionsToDelete) {
        Write-Output "Removing version created on: $($version.Created) with ID: $($version.ID)"
        $version.DeleteObject()
    }
    $FileItem.Context.ExecuteQuery()
}

function Show-CustomProgress {
    param ([int]$current, [int]$total, [string]$activity, [string]$status)
    $percent = [math]::Round(($current / $total) * 100)
    $bar = ('=' * ($percent / 2)).PadRight(50)
    Write-Host "$activity`n[$bar] $percent% - $status"
}
#endregion -------------- END BUSINESS ---------------------------------

#-----------------------------------------------------------------------
# MAIN
#-----------------------------------------------------------------------
$ErrorActionPreference = 'Stop'

# --- inputs ---
$SiteUrl     = Read-Host 'Enter the SharePoint Site URL'
$LibraryName = Read-Host 'Enter the Document Library Name'
$ClientId    = Read-Host 'Entra-ID Application (client) ID'
$TenantId    = Read-Host 'Entra-ID Tenant ID (GUID)'

# --- connect (app-only) ---
Connect-SharePoint -SiteUrl $SiteUrl -ClientId $ClientId -TenantId $TenantId
Write-Output "Connected to site: $SiteUrl"

# --- get library ---
$library = Get-PnPList -Identity $LibraryName -ErrorAction Stop
if ($library.BaseTemplate -ne 101) { throw 'The specified list is not a document library.' }
Write-Output "Retrieved library: $LibraryName"

# --- enumerate files (paginated) ---
$pageSize  = 2000
$listItems = @()
$batch     = 0
do {
    $query = "<View Scope='RecursiveAll'><RowLimit>$pageSize</RowLimit><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query></View>"
    $items = Get-PnPListItem -List $LibraryName -Query $query -PageSize $pageSize
    $listItems += $items
    $batch++
    Write-Output "Batch $batch`: retrieved $($items.Count) items"
} while ($items.Count -eq $pageSize -and $items.Count -gt 0)

# --- process files ---
$results     = @()
$totalFiles   = ($listItems | Where-Object { $_.FileSystemObjectType -eq 'File' }).Count
$currentFile  = 0

foreach ($file in $listItems) {
    if ($file.FileSystemObjectType -ne 'File') { continue }

    $currentFile++
    try {
        $fileUrl  = $file.FieldValues["FileRef"]
        $fileName = $file.FieldValues["FileLeafRef"]
        $fileItem = Get-PnPFile -Url $fileUrl -AsListItem

        $fileItem.Context.Load($fileItem)
        $fileItem.Context.Load($fileItem.Versions)
        $fileItem.Context.ExecuteQuery()

        $versions = $fileItem.Versions
        if (-not $versions) { Write-Warning "No versions for $fileName"; continue }

        $sorted = $versions | Sort-Object Created -Descending
        $start  = $sorted.Count
        if ($start -gt 4) {
            $toDelete = $sorted[4..($start - 1)]
            Remove-Versions -VersionsToDelete $toDelete -FileItem $fileItem
        }

        $results += [PSCustomObject]@{
            LibraryName          = $LibraryName
            FileName             = $fileName
            FileUrl              = $fileUrl
            StartingVersionCount = $start
            TotalVersions        = $sorted.Count
            DeletedVersionsCount = $start - $sorted.Count
        }
        Write-Output "Processed file: $fileName – Total Versions: $start"
    } catch {
        Write-Error "Failed to process file $fileName : $_"
        continue
    }
}

# --- final progress & CSV ---
Show-CustomProgress -current $currentFile -total $totalFiles -activity "Processing Files in $LibraryName" -status "Completed $currentFile of $totalFiles files"

$tenantName = ([uri]$SiteUrl).Host.Split('.')[0] -replace '^https://',''
$outputDir  = "C:\$tenantName"
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
$outputCsv  = Join-Path $outputDir "$LibraryName`_results.csv"
$results | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Output "Results exported to $outputCsv."
Write-Output "Deleted all but the latest 4 versions of files in the specified library."