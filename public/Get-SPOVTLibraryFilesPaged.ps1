function Get-SPOVTLibraryFilesPaged {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Microsoft.SharePoint.Client.List]$Library,

        [ValidateRange(100, 5000)]
        [int]$PageSize = 2000,

        [switch]$NoProgress
    )

    $script:seen = 0
    $total = $Library.ItemCount

    # Get-PnPListItem supports -PageSize and -ScriptBlock paging [2](https://pnp.github.io/powershell/cmdlets/Get-PnPFileVersion.html)
    Get-PnPListItem -List $Library -Fields "FileRef","FileLeafRef" -PageSize $PageSize -ScriptBlock {
        param($items)

        $script:seen += $items.Count

        if (-not $NoProgress -and $total -gt 0) {
            $pct = [Math]::Min(100, ($script:seen / $total) * 100)
            Write-Progress -Activity "Enumerating '$($Library.Title)'" -Status "$script:seen / $total" -PercentComplete $pct
        }

        # Common paging pattern: execute CSOM for each page [2](https://pnp.github.io/powershell/cmdlets/Get-PnPFileVersion.html)
        $items.Context.ExecuteQuery()
    } | Where-Object { $_.FileSystemObjectType -eq 'File' }
}