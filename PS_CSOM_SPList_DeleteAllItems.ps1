
################################################################################################################################################################
# SharePoint delete all items
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString
[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10426"
[string] $ListTitle = "Info"

[string] $TimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $ThisScriptRoot = @("J:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[string] $ThisScriptName = $null

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}
else
{
    $ThisScriptName = "PS_CSOM_SPList_DeleteAllItems"
}

[string] $LogFilePath  = "$($ThisScriptRoot)\$($ThisScriptName).log"

################################################################################################################################################################
# Main Program
################################################################################################################################################################

try
{
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
    $context.Credentials = $credentials

    Write-Host "Querying web/list data..."

    $web = $context.Web
    $context.Load($web)

    $list = $context.Web.Lists.GetByTitle($ListTitle)
    
    $context.Load($list)
    $context.ExecuteQuery()

    $camlQry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $camlQry.ViewXml = "<View Scope='RecursiveAll'><RowLimit>100</RowLimit></View>"
    

    Write-Host "Querying SP list..."
        
    $listItems = $list.GetItems($camlQry)
    $context.Load($listItems)
    $context.ExecuteQuery()
        
    while ($listItems.Count -gt 0)
    {
        Write-Host "Number of recs : $($listItems.Count)"
        for ($i = 0; $i -lt $listItems.Count; $i++)
        {
            $listItem = $listItems[$i]
            try
            {
                $context.Load($listItem)
                $listItem.DeleteObject()
            }
            catch
            {
                Write-Host $_.Exception.Message
            }
        }

        $context.ExecuteQuery()
        $listItems = $list.GetItems($camlQry)
        $context.Load($listItems)
        $context.ExecuteQuery()
    }

}
catch
{
    #Write-Host    $_.Exception.ToString() -ForegroundColor Red
    throw
}
finally
{
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
