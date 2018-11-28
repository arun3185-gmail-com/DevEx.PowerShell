
################################################################################################################################################################
# Quest Job Log Analyser
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString

[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10543"
[string] $ListTitle = "Doku IM-FS-AT/A&S"
#[string] $Global:QstLogFilePath = "C:\ProgramData\Dell\Migrator for Notes to SharePoint\Log Files\NMSP_08_09_2018_10_53_56.log"

[string] $Global:Tab = [char]9

[string] $Global:TimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $Global:ThisScriptRoot = @("J:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[string] $Global:ThisScriptName = $null

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}
else
{
    $Global:ThisScriptName = "PS_LN_Migration_ReconciliationReport"
}

[string] $Global:LogFilePath  = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"

################################################################################################################################################################
# Functions
################################################################################################################################################################

function Write-LogInfo()
{
    Param ([string] $Message)
    
    "$(Get-Date -Format $Global:TimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}


################################################################################################################################################################
# Main Program
################################################################################################################################################################

try
{
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
    $context.Credentials = $credentials
    
    ################################################################################
    # Files and Folders count
    ################################################################################
    
    $camlQry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $camlQry.ViewXml = "<View Scope='RecursiveAll'> <Query><Where><Eq> <FieldRef Name='FSObjType' /> <Value Type='Integer'>1</Value> </Eq></Where></Query> </View>"
    $web = $context.Web
    $list = $context.Web.Lists.GetByTitle($ListTitle)
    $listItems = $list.GetItems($camlQry)  
    $context.Load($web)
    $context.Load($list)
    $context.Load($listItems)

    $context.ExecuteQuery()
    $listFoldersCount = $listItems.Count
    
    Write-Host "Items Count :-"
    Write-Host "   Total Items - $($list.ItemCount)"
    Write-Host "   Folders Count - $($listFoldersCount)"
    Write-Host ""
    Write-Host ""

    ################################################################################
    # ACL groups to SharePoint Groups
    # AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup
    ################################################################################

    $webRoleAssignments = $web.RoleAssignments
    $context.Load($webRoleAssignments)
    $siteGroups = $web.SiteGroups
    $context.Load($siteGroups)
    $siteUsers = $web.SiteUsers
    $context.Load($siteUsers)
    $userInfoList = $web.SiteUserInfoList
    $context.Load($userInfoList)
    $userInfoListItemCollection = $userInfoList.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $context.Load($userInfoListItemCollection)
    $context.ExecuteQuery()
    

    foreach ($roleAssgn in $webRoleAssignments)
    {
        $prncpl = $roleAssgn.Member

        $context.Load($prncpl)
        $context.ExecuteQuery()

        Write-Host "$($roleAssgn.PrincipalId) - $($roleAssgn.Member.PrincipalType) - $($roleAssgn.Member.LoginName)"
    }

    Write-Host ""
    Write-Host ""
    
    foreach ($grp in $siteGroups)
    {
        Write-Host "$($grp.Id) - $($grp.PrincipalType) - $($grp.LoginName)"
    }

    Write-Host ""
    Write-Host ""
    
    foreach ($grp in $siteUsers)
    {
        Write-Host "$($grp.Id) - $($grp.PrincipalType) - $($grp.LoginName)"
    }

    foreach ($itm in $userInfoListItemCollection)
    {
        Write-Host "$($itm.Id) - $($itm["Title"])"
    }


    ################################################################################
    # Left Navigation
    ################################################################################
    $webNavigation = $web.Navigation
    $quickLaunchColl = $webNavigation.QuickLaunch
    $context.Load($web)
    $context.Load($webNavigation)
    $context.Load($quickLaunchColl)

    $context.ExecuteQuery()
    
    Write-Host "Quick Launch :-"
    foreach ($itm in $quickLaunchColl)
    {
        Write-Host ""
        Write-Host "$($itm.Title) [$($itm.Url)]"
    }

    

    
    ################################################################################
    # Views, Sorting
    # UI improvement webpart
    ################################################################################

    ################################################################################
    # Feat Disable - Publishing, Limited-access user permission
    ################################################################################
    
    $limitedUserFeatID = [guid]::new("7c637b23-06c4-472d-9a9a-7c175762c5c4")
    $spPubServerFeatID = [guid]::new("f6924d36-2fa8-4f0b-b16d-06b7250180fa")
    $siteFeatCollection = $context.Site.Features
    $context.Load($siteFeatCollection)

    $context.ExecuteQuery()
    $siteFeat1 = $siteFeatCollection | Where { $_.DefinitionID -eq $limitedUserFeatID }
    if ($siteFeat1)
    {
        Write-Host "$($limitedUserFeatID) - Activated"
    }
    else
    {
        Write-Host "$($limitedUserFeatID) - Not Activated"
    }

    $siteFeat2 = $siteFeatCollection | Where { $_.DefinitionID -eq $spPubServerFeatID }
    if ($siteFeat2)
    {
        Write-Host "$($spPubServerFeatID) - Activated"
    }
    else
    {
        Write-Host "$($spPubServerFeatID) - Not Activated"
    }

    ################################################################################
    # Ordering of columns - New/Edit/View form
    # SCA remove all except Migration team
    # Change owner of notes group
    # Remove ADM groups
    ################################################################################

}
catch
{
    Write-LogInfo $_.Exception.ToString()
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
finally
{
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
