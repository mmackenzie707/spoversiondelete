# Function to delete versions and show details
function Delete-Versions {
    param (
        [array]$versionsToDelete,
        [object]$fileItem
    )

    foreach ($version in $versionsToDelete) {
        Write-Output "Deleting version created on: $($version.Created) with ID: $($version.ID)"
        $version.DeleteObject()
    }
    $fileItem.Context.ExecuteQuery()
}

# Function to show custom progress
function Show-CustomProgress {
    param (
        [int]$current,
        [int]$total,
        [string]$activity,
        [string]$status
    )

    $percentComplete = [math]::Round(($current / $total) * 100)
    $progressBar = "=" * ($percentComplete / 2)
    $progressBar = $progressBar.PadRight(50)

    Write-Host "$activity"
    Write-Host "[$progressBar] $percentComplete% - $status"
}

# Prompt for manual input of the Site URL
$SiteUrl = Read-Host "Enter the SharePoint Site URL"

# Extract the tenant name from the Site URL
$tenantName = ($SiteUrl -split '\.')[0] -replace 'https://', ''

# Prompt for manual input of the Document Library Name
$LibraryName = Read-Host "Enter the Document Library Name"

# App-only authentication parameters
$clientId = "7c313b94-e800-4058-a362-0c3f1d5473fb"
$tenantId = "5d784adb-69b9-4013-b010-cc2ec806aa6c"
$certificatePath = "C:\test\BMLScripts\bravomediacert.pfx"
$certificatePassword = "bravomedia1234!"

# Connect to SharePoint Online using app-only authentication
try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $clientId -Tenant $tenantId -CertificatePath $certificatePath -CertificatePassword (ConvertTo-SecureString $certificatePassword -AsPlainText -Force)
    Write-Output "Connected to site: $SiteUrl"
} catch {
    Write-Error "Failed to connect to site ${SiteUrl}: $_"
    exit 1
}

# Get the specified document library
try {
    $library = Get-PnPList -Identity $LibraryName -ErrorAction Stop
    if ($library -eq $null) {
        Write-Error "Library not found: $LibraryName"
        exit 1
    }
    if ($library.BaseTemplate -ne 101) {
        Write-Error "The specified list is not a document library."
        exit 1
    }
    Write-Output "Retrieved library: $LibraryName"
} catch {
    Write-Error "Failed to retrieve library ${LibraryName}: $_"
    exit 1
}

# Initialize variables
$results = @()
$pageSize = 2000
$listItems = @()
$batch = 0

# Get all files in the library using pagination
do {
    $query = "<View Scope='RecursiveAll'><RowLimit>${pageSize}</RowLimit><Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query></View>"
    try {
        $items = Get-PnPListItem -List $LibraryName -Query $query -PageSize $pageSize
        $listItems += $items
        $batch++
        Write-Output ("Batch ${batch}: Retrieved ${($items.Count)} items")
    } catch {
        Write-Error "Failed to retrieve files for library ${LibraryName} in site ${SiteUrl}: $_"
        exit 1
    }
} while ($items.Count -eq $pageSize -and $items.Count -gt 0)

$totalFiles = $listItems.Count
$currentFile = 0

foreach ($file in $listItems) {
    # Skip folders and only process files
    if ($file.FileSystemObjectType -ne "File") {
        continue
    }

    $currentFile++

    try {
        $fileUrl = $file.FieldValues["FileRef"]
        $fileName = $file.FieldValues["FileLeafRef"]
        $fileItem = Get-PnPFile -Url $fileUrl -AsListItem

        # Ensure the file item and its versions are loaded
        if ($fileItem -ne $null) {
            $fileItem.Context.Load($fileItem)
            $fileItem.Context.Load($fileItem.Versions)
            $fileItem.Context.ExecuteQuery()

            # Get all versions of the file
            $versions = $fileItem.Versions

            if ($versions -ne $null) {
                # Sort versions by creation date in descending order
                $sortedVersions = $versions | Sort-Object -Property Created -Descending

                # Keep the latest 4 versions and delete the rest
                $startingVersionCount = $sortedVersions.Count
                if ($startingVersionCount -gt 4) {
                    $versionsToDelete = $sortedVersions[4..($startingVersionCount - 1)]
                    Delete-Versions -versionsToDelete $versionsToDelete -fileItem $fileItem
                }

                # Collect results for CSV export
                $results += [PSCustomObject]@{
                    LibraryName = $LibraryName
                    FileName = $fileName
                    FileUrl = $fileUrl
                    StartingVersionCount = $startingVersionCount
                    TotalVersions = $sortedVersions.Count
                    DeletedVersionsCount = $startingVersionCount - $sortedVersions.Count
                }
                Write-Output "Processed file: ${fileName} - Total Versions: ${startingVersionCount}"
            } else {
                Write-Error "No versions found for file ${fileName} in library ${LibraryName} on site ${SiteUrl}"
            }
        } else {
            Write-Error "Failed to load file item for ${fileUrl} in library ${LibraryName} on site ${SiteUrl}"
        }
    } catch {
        Write-Error "Failed to process file ${fileName} in library ${LibraryName} on site ${SiteUrl}: $_"
        continue
    }
}

# Show final progress
Show-CustomProgress -current $currentFile -total $totalFiles -activity "Processing Files in ${LibraryName}" -status "Completed processing ${currentFile} of ${totalFiles} files"

# Export results to CSV
# Define the directory path dynamically based on the tenant name
$directoryPath = "C:\${tenantName}"

# Check if the directory exists, if not, create it
if (-not (Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath -Force
}

# Create the CSV file name dynamically
$csvFileName = "$LibraryName`_results.csv"
$csvPath = Join-Path -Path $directoryPath -ChildPath $csvFileName

# Export the results to the CSV file
$results | Export-Csv -Path $csvPath -NoTypeInformation

# Results Output
Write-Output "Results exported to $csvPath"
Write-Output "Deleted all but the latest 4 versions of files in the specified library."

# Show final progress bar at the bottom
Show-CustomProgress -current $currentFile -total $totalFiles -activity "Processing Files in ${LibraryName}" -status "Completed processing ${currentFile} of ${totalFiles} files"