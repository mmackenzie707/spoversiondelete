function Connect-SPOVTDeviceLogin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SiteUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$TenantId
    )

    # DeviceLogin is a supported Connect-PnPOnline auth mode [1](https://www.sharepointdiary.com/2018/05/sharepoint-online-delete-version-history-using-pnp-powershell.html)
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $TenantId -DeviceLogin
}