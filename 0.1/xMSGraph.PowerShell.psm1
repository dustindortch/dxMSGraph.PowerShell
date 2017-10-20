# Written and tested with PowerShell 5.0 and 5.1
# Requires Microsoft.ADAL.PowerShell module
#   > Install-Module Microsoft.ADAL.PowerShell

$Script:GraphApiVersion = 'v1.0'
$Script:PowerShellClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
$Script:ResourceId = "https://graph.microsoft.com"
$Script:RedirectUri = "urn:ietf:wg:oauth:2.0:oob"

#Import-Module Microsoft.ADAL.PowerShell # Install-Module Microsoft.ADAL.PowerShell

function _temp{[CmdletBinding(SupportsShouldProcess)] param() Write-Verbose "Temporary function to build list of parameters established for Advanced Functions."}
$Script:DefaultParams = (Get-Command _temp | Select-Object -ExpandProperty parameters).Keys
Remove-Item function:\_temp

# Connect-Graph
function Connect-Graph
{
    <#
    .SYNOPSIS
        Authenticate with Modern Authentication to the Microsoft Graph
    .DESCRIPTION
        Using the Microsoft.ADAL.PowerShell Module, authenticate against Azure Active Directory with Modern Authentication to the Microsoft Graph and retrieve a Bearer Token
    .EXAMPLE
        Connect-Graph -TenantDomain <TenantName>.onmicrosoft.com
    .INPUTS
        TenantDomain
    #>

    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory)]
        [Alias("Tenant")] 
        [String]$TenantDomain
    )

    Begin
    {
        $Script:TenantDomain = $TenantDomain
    }
    Process
    {
        $Token = Get-ADALAccessToken `
            -AuthorityName $Script:TenantDomain `
            -ClientId $Script:PowerShellClientId `
            -ResourceId $Script:ResourceId `
            -RedirectUri $Script:RedirectUri `
            -ForcePromptSignIn
        $Script:AuthHeader = "Bearer {0}" -f $Token
        $Applications = Get-GraphQuery -Query "applications" -GraphVersion beta -Filter "displayName eq 'Tenant Schema Extension App'"
        if ($Applications)
        {
            $Script:ExtensionGuid = $Applications.id
            $Script:CustomAttributes = Get-GraphQuery -Query "applications/$(Get-GraphAppId)/extensionProperties" -GraphVersion beta
        }
    }
    End
    {
    }
}

function Get-GraphQuery
{
    <#
    .SYNOPSIS
        Performs a simple query against the Microsoft Graph
    .DESCRIPTION
        This allows for queries to be made against the Microsoft Graph
    .EXAMPLE
        Get-GraphQuery -Query "me"
    .INPUTS
        Inputs to this cmdlet (if any)
    .OUTPUTS
        Output from this cmdlet (if any)
    .NOTES
        General notes
    .COMPONENT
        The component this cmdlet belongs to
    .FUNCTIONALITY
        The functionality that best describes this cmdlet
    #>

    [CmdletBinding()]
    Param
    (
        # Query help description
        [Parameter(Mandatory, 
            ValueFromPipeline,
            ValueFromPipelineByPropertyName, 
            ValueFromRemainingArguments, 
            Position=0)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("Q")] 
        [String]$Query,

        # Filter help description
        [Alias("F")]
        [String]$Filter,

        # Filter help description
        [Alias("S")]
        [String]$Select,

        # Filter help description
        [Alias("E")]
        [String]$Expand,

        # GraphVersion help description
        [ValidateSet('v1.0','beta')]
        [Alias("V","Version")]
        [String]$GraphVersion = $Script:GraphApiVersion,

        # Raw help description
        [Switch]$Raw
    )

    Begin
    {
        $Uri = "{0}/{1}/{2}"
        if($Filter -or $Select -or $Expand)
        {
            if ($Filter) {$Filter = "`$filter={0}" -f [System.Net.WebUtility]::UrlEncode($Filter)}
            if ($Select) {$Select = "`$select={0}" -f [System.Net.WebUtility]::UrlEncode($Select)}
            if ($Expand) {$Expand = "`$expand={0}" -f [System.Net.WebUtility]::UrlEncode($Expand)}
            $QueryString = (($Filter,$Select,$Expand -ne $null) -join "&")
            Write-Verbose "QUERY STRING: ${QueryString}"
            $Uri += "?{3}"
        }
    }
    Process
    {
        $Result = Invoke-RestMethod -Method Get -Header @{
            Authorization = $Script:AuthHeader
            'Content-Type' = "application/json"
        } -Uri ($Uri -f $Script:ResourceId, $GraphVersion, $Query, $QueryString)
    }
    End
    {
        if($Raw)
        {
            Return $Result | Select-Object * -ExcludeProperty "@odata.context"
        } else {
            Return $Result.value
        }
    }
}

function Get-GraphUser
{

    <#
    .SYNOPSIS
        Get Microsoft Graph users
    .DESCRIPTION
        Get Microsoft Graph users and attributes

        Works with custom attribute synchronization configured with Azure AD Connect
    .EXAMPLE
        Get-GraphUser [-UserPrincipal <UserPrincipalName>]
    .INPUTS
        UserPrincipalName
    #>

    [CmdletBinding()]
    Param
    (
        # UserPrincipalName help description
        [Parameter(Position=0)]
        [Alias("User")]
        [String]$UserPrincipalName
    )
    DynamicParam {
        $UserParameterDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
        if ($Script:CustomAttributes)
        {
            Write-Verbose "CUSTOM ATTRIBUTES SYNCHRONIZED"
            $Script:CustomAttributes | Where-Object {$_.targetObjects -Contains "User"} | ForEach-Object {
                $Type = [System.Type]"$($_.dataType)"
                $FullName = $_.Name
                $FriendlyName = $FullName.Replace("extension_$(Get-GraphAppId -Trim)_","")
                $UserAttribute = New-Object System.Management.Automation.ParameterAttribute
                $UserAttribute.Mandatory = $false
                $UserAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
                $UserAttributeCollection.Add($UserAttribute)
                $UserParameter = New-Object System.Management.Automation.RuntimeDefinedParameter($FriendlyName,$Type,$UserAttributeCollection)
                $UserParameterDictionary.Add($FriendlyName,$UserParameter)
            }
        }
        Return $UserParameterDictionary
    }

    Begin
    {
        $UnfilteredParams = New-Object System.Collections.ArrayList
        $Query = "users"
        if($UserPrincipalName)
        {
            $Query += "/${UserPrincipalName}"
            $UnfilteredParams.Add("UserPrincipalName") | Out-Null
        }
        Write-Verbose "DEFAULT PARAMETERS: ${Script:DefaultParams}"
        $Script:DefaultParams | ForEach-Object {
            $UnfilteredParams.Add($_) | Out-Null
        }
        Write-Verbose "UNFILTERED PARAMETERS: ${UnfilteredParams}"
        $Filter = $null
        ForEach($Key in $PSBoundParameters.Keys) {
            if ($UnfilteredParams -notcontains $Key)
            {
                Write-Verbose "ADDING ATTRIBUTE KEY: ${Key}"
                if ($Filter) {$Filter += " and "}
                $Filter += "{0} eq '{1}'" -f "extension_$(Get-GraphAppId -Trim)_${Key}",$PSBoundParameters.($Key)
                Write-Verbose "ADDING ATTRIBUTE VALUE: ${PSBoundParameters.($Key)}"
            }
        }
    }
    Process
    {
        $Users = Get-GraphQuery -Query $Query -GraphVersion beta -Filter $Filter -Raw
        if (!$UserPrincipalName)
        {
            $Users = $Users.Value
        }
        if ($Script:ExtensionGuid)
        {
            ForEach ($User in $Users) {
                $User | ForEach-Object {
                    $Properties = $_
                    ForEach ($Property in $_.PSObject.Properties) {
                        if ($Property.Name -like "extension_$(Get-GraphAppId -Trim)_*")
                        {
                            $FriendlyName = ($Property.Name).Replace("extension_$(Get-GraphAppId -Trim)_","")
                            $User | Add-Member `
                                -NotePropertyName $FriendlyName `
                                -NotePropertyValue $Property.Value
                            Write-Verbose "RENAMING PROPERTY: $($Property.Name) to $($FriendlyName)"
                        }
                    }
                }
            }
        $Users = $Users | Select-Object * -ExcludeProperty "extension_$(Get-GraphAppId -Trim)_*"
        }
    }
    End
    {
        Return $Users
    }
}

function Get-GraphAppId
{
    Param(
        [switch]$Trim
    )

    Begin {}
    Process {}
    End
    {
        if ($Trim -and $Script:ExtensionGuid)
        {
            Return $Script:ExtensionGuid.Replace("-","")
        } else {
            Return $Script:ExtensionGuid
        }
    }
}