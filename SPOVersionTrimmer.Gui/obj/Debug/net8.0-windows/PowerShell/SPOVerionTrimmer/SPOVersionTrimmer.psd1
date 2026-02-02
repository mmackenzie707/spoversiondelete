@{
    RootModule        = 'SPOVersionTrimmer.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b6c33f49-4fb2-4a5d-a5b7-9b5a9e6e1c41'
    Author            = 'Your Team'
    CompanyName       = 'Your Org'
    Copyright         = '(c) 2026 Your Org'
    Description       = 'Trim SharePoint Online file version history using PnP.PowerShell Device Login with paging and reporting.'

    PowerShellVersion = '7.2'

    RequiredModules   = @(
        @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }
    )

    FunctionsToExport = @(
        'Connect-SPOVTDeviceLogin',
        'Get-SPOVTLibraryFilesPaged',
        'Trim-SPOVTFileVersions',
        'Invoke-SPOVTLibraryTrim'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags       = @('SharePoint','PnP','Versioning','Cleanup','DeviceLogin')
            ProjectUri = 'https://pnp.github.io/powershell/'
        }
    }
}
``