@{
    RootModule = 'xMSGraph.PowerShell.psm1'
    ModuleVersion = '0.2'
    Author = 'Dustin Dortch'
    Description = 'Microsoft Graph API module for PowerShell'
    PowerShellVersion = '5.0'
    PowerShellHostVersion = '1.0'
    DotNetFrameworkVersion = '4.5.0.0'
    RequiredModules = @("Microsoft.ADAL.PowerShell")
    FunctionsToExport = "*"
    CmdletsToExport = "*"
    VariablesToExport = "*"
    AliasesToExport = "*"
    ModuleList = @("xMSGraph.PowerShell")
    DefaultCommandPrefix = ''
    FileList = @("xMSGraph.PowerShell.psd1","xMSGraph.PowerShell.psm1")
    PrivateData = @{
        PSData = @{
            Tags = @('Graph')
        }
    }
}