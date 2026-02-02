# At the end of your script, output JSON:
$output = @{
    ok = $true
    title = "Site Title"
    url = $siteUrl
} | ConvertTo-Json -Compress

Write-Output $output