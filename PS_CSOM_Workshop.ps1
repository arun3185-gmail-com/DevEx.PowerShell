################################################################################################################################################################
# CSOM Workshop
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################
[string] $Username = "Arunkumar.Balasubramani@dxit181001.onmicrosoft.com"
if ($Password -eq $null) { [SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString }
[string] $SiteUrl   = "https://dxit181001.sharepoint.com"
[string] $ListTitle = "DemoList1"

################################################################################################################################################################
$context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
$context.Credentials = $credentials    
    

$web = $context.Web
$context.Load($web)
$list = $web.Lists.GetByTitle($ListTitle)

$context.ExecuteQuery()











