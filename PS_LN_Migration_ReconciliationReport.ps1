
################################################################################################################################################################
# Quest Job Log Analyser
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString

[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10577"
[string] $ListTitle = "Changchun Teamdoc"
#[string] $Global:QstLogFilePath = "C:\ProgramData\Dell\Migrator for Notes to SharePoint\Log Files\NMSP_08_09_2018_10_53_56.log"

[string[]] $GroupsToBeRemoved = @("#ADM_Database_DellQuest_Migrator_Team","#ADM_Database_Developer","#ADM_Database_IS_Admin","#ADM_Database_IS_Support")
[string[]] $AllowedSiteCollAdminTitles = @("(GA4250 evonik) SharePoint Online - Site collection administrator","GMS_MDL_BS-BP-SiteCollection","SiteCollectionAdmin_O365_EMEA")
[string[]] $AllowedSiteCollAdminLogins = @("i:0#.f|membership|ga4250@evonik.onmicrosoft.com","c:0t.c|tenant|9c535337-1405-4bc3-9214-494cd2357bcf","c:0t.c|tenant|765575ea-6693-4cd2-a7cf-127878a3f536")


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
    
    $site = $context.Site
    $web = $context.Web
    $list = $context.Web.Lists.GetByTitle($ListTitle)
    $camlQryFolders = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $camlQryFolders.ViewXml = "<View Scope='RecursiveAll'> <Query><Where><Eq> <FieldRef Name='FSObjType' /> <Value Type='Integer'>1</Value> </Eq></Where></Query> </View>"
    $listItemsFolders = $list.GetItems($camlQryFolders)  
    
    $context.Load($site)
    $context.Load($web)
    $context.Load($list)
    $context.Load($listItemsFolders)
    $context.ExecuteQuery()
    
    Write-Host "Items Count :-"
    Write-Host "   Total Items - $($list.ItemCount)"
    Write-Host "   Folders Count - $($listItemsFolders.Count)"
    Write-Host ""
    Write-Host ""

    ################################################################################
    # ACL groups to SharePoint Groups
    # AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup
    ################################################################################
    
    <#
    $userInfoList = $web.SiteUserInfoList
    $context.Load($userInfoList)
    $userInfoListFields = $userInfoList.Fields
    $context.Load($userInfoListFields)
    $userInfoListItemCollection = $userInfoList.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $context.Load($userInfoListItemCollection)

    $webRoleAssignments = $web.RoleAssignments
    $context.Load($webRoleAssignments)
    #>

    $webAssociatedOwnerGroup = $web.AssociatedOwnerGroup
    $webAssociatedMemberGroup = $web.AssociatedMemberGroup
    $webAssociatedVisitorGroup = $web.AssociatedVisitorGroup
    $siteGroups = $web.SiteGroups
    $siteUsers = $web.SiteUsers
    
    $context.Load($webAssociatedOwnerGroup)
    $context.Load($webAssociatedMemberGroup)
    $context.Load($webAssociatedVisitorGroup)
    $context.Load($siteGroups)
    $context.Load($siteUsers)

    $context.ExecuteQuery()

    <#
    foreach ($roleAssgn in $webRoleAssignments)
    {
        $prncpl = $roleAssgn.Member
        $context.Load($prncpl)
        $context.ExecuteQuery()
        Write-Host "$($roleAssgn.PrincipalId) - $($roleAssgn.Member.PrincipalType) - $($roleAssgn.Member.LoginName)"
    }
    #>

    
    ################################################################################
    # Remove #ADM Groups, Change owner of other groups to default Owner
    ################################################################################

    foreach ($grp in $siteGroups)
    {
        if ($grp.LoginName -in $GroupsToBeRemoved)
        {
            $siteGroups.RemoveByLoginName($grp.LoginName)
        }
    }
    foreach ($grp in $siteGroups)
    {
        if ($grp.OwnerTitle -notin @("System Account",$webAssociatedOwnerGroup.LoginName))
        {
            Write-Host $grp.LoginName
            $grp.Owner = $webAssociatedOwnerGroup
            $grp.Update()
            $context.Load($grp)
        }
    }
    $context.Load($siteGroups)
    $context.ExecuteQuery()

    ################################################################################
    # Remove other site collection admins
    ################################################################################
    
    foreach ($usr in $siteUsers)
    {
        if ($usr.IsSiteAdmin)
        {
            if ($usr.LoginName -notin $AllowedSiteCollAdminLogins -or $usr.Title -notin $AllowedSiteCollAdminTitles)
            {
                $usr.IsSiteAdmin = $false
                $context.Load($usr)
            }
        }
    }

    $context.ExecuteQuery()
    
    
    ################################################################################
    # Views, Sorting
    ################################################################################
    
    $listViews = $list.Views
    $context.Load($listViews)
    $context.ExecuteQuery()

    ################################################################################
    # Left Navigation
    ################################################################################
    
    [Microsoft.SharePoint.Client.NavigationNodeCreationInformation] $newQLNodeInfo = $null
    [Microsoft.SharePoint.Client.NavigationNodeCreationInformation] $newSubNodeInfo = $null

    $webNavigation = $web.Navigation
    $quickLaunchColl = $webNavigation.QuickLaunch
    $context.Load($webNavigation)
    $context.Load($quickLaunchColl)

    $context.ExecuteQuery()
    
    Write-Host "Working on Quick Launch..."
    for ($i = ($quickLaunchColl.Count-1); $i -ge 0; $i--)
    {
        Write-Host "Removing - $($quickLaunchColl[$i].Title) [$($quickLaunchColl[$i].Url)]"
        #Write-LogInfo "Removing - $($quickLaunchColl[$i].Title) [$($quickLaunchColl[$i].Url)]"
        $quickLaunchColl[$i].DeleteObject()
    }
    $context.Load($quickLaunchColl)
    $context.ExecuteQuery()

    ####################
    
    $newQLNodeInfo = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
    $newQLNodeInfo.Title = $list.Title
    $newQLNodeInfo.Url = $listViews.Where({ $PSItem.DefaultView })[0].ServerRelativeUrl
    
    $context.Load($quickLaunchColl.Add($newQLNodeInfo))
    $context.ExecuteQuery()
    
    $newSubNodeInfo = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
    $newSubNodeInfo.Title = $listViews.Where({ $PSItem.DefaultView })[0].Title
    $newSubNodeInfo.Url = $listViews.Where({ $PSItem.DefaultView })[0].ServerRelativeUrl
    
    $context.Load($quickLaunchColl[0].Children)
    $context.ExecuteQuery()
    $context.Load($quickLaunchColl[0].Children.Add($newSubNodeInfo))
    
    foreach ($view in $listViews.Where({ -not($PSItem.DefaultView) }))
    {
        $newSubNodeInfo = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
        $newSubNodeInfo.Title = $view.Title
        $newSubNodeInfo.Url = $view.ServerRelativeUrl
        $newSubNodeInfo.AsLastNode = $true
        $context.Load($quickLaunchColl[0].Children.Add($newSubNodeInfo))
    }
    
    $context.Load($quickLaunchColl)
    $context.ExecuteQuery()
    
    ################################################################################
    # UI improvement webpart
    $list
    
    $listContTypes = $list.ContentTypes
    $context.Load($listContTypes)
    $context.ExecuteQuery()

    $pageDispItem = $context.Web.GetFileByServerRelativeUrl("/sites/10577/Lists/Changchun%20Teamdoc/DispForm.aspx")
    $context.Load($pageDispItem)
    $context.ExecuteQuery()

    $webPartManager = $pageDispItem.GetLimitedWebPartManager([System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
    $context.Load($webPartManager)
    $context.ExecuteQuery()

    $webPartDefnCollection = $webPartManager.WebParts
    $context.Load($webPartDefnCollection)
    $context.ExecuteQuery()

    

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
