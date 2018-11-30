
################################################################################################################################################################
# CSOM Workshop
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################
[string] $Username = "Arunkumar.Balasubramani@DxIt181130.onmicrosoft.com"
if ($Password -eq $null) { [SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString }
[string] $SiteUrl   = "https://dxit181130.sharepoint.com"
[string] $ListTitle = "DemoList1"

################################################################################################################################################################
$context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
$context.Credentials = $credentials    

$web = $context.Web
$context.Load($web)


$list = $web.Lists.GetByTitle($ListTitle)
$context.Load($list)

$fieldCollection = $list.Fields
$context.Load($fieldCollection)
$context.ExecuteQuery()

$demoColumnField = $fieldCollection.Where({$_.Title -eq "DemoColumn"})[0]
$demoColumnField.SchemaXml




$webNavigation = $web.Navigation
$quickLaunchColl = $webNavigation.QuickLaunch
$context.Load($webNavigation)
$context.Load($quickLaunchColl)

$context.ExecuteQuery()

foreach ($itm in $quickLaunchColl)
{
    Write-Host $itm.Title
    $itmChildren = $itm.Children
    $context.Load($itmChildren)
    $context.ExecuteQuery()
    if ($itmChildren.Count -gt 0)
    {
        foreach ($subItm in $itmChildren)
        {
            Write-Host "  $($subItm.Title)"
        }
    }
}



$headerLink = $quickLaunchColl[6]
$headerLink.DeleteObject()
$context.Load($quickLaunchColl)
$context.ExecuteQuery()





