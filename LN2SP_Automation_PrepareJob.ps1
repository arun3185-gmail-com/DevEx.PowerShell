﻿
################################################################################################################################################################
# Lotus Notes to SharePoint Migration Automation
################################################################################################################################################################
#Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString
[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10426"
[string] $ListTitle = "Info"

[string] $DataFolderPath = "J:\Arun\Git\DevEx.Data"

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
    $ThisScriptName = "LN2SP_Automation_PrepareJob"
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
    
    [string] $dtTimeSuffix = (Get-Date -Format "yyyyMMdd_HHmmss")

    $camlQry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $camlQry.ViewXml = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>"
        
    do
    {
        Write-Host "Querying SP list..."
        
        $listItems = $list.GetItems($camlQry)
        $context.Load($listItems)
        $context.ExecuteQuery()
        
        Write-Host "Number of recs : $($listItems.Count)"
        
        foreach ($listItem in $listItems)
        {
            try
            {
                #
                [string] $migrationStatus = $listItem["MigrationStatus"].ToString()
                [string] $serverFilePath = $listItem["MigrationStatus"].ToString()
                [string[]] $serverFilePathParts = $serverFilePath.Split(@("!!"), [System.StringSplitOptions]::RemoveEmptyEntries)
                [string] $serverPath = $serverFilePathParts[0]
                [string] $lnFilePath = $serverFilePathParts[1]
                [string] $destinationType = $listItem["MigrationStatus"].ToString()
            }
            catch
            {
                Write-Host $_.Exception.Message
            }
            finally
            {
            }
        }

        $camlQry.ListItemCollectionPosition = $listItems.ListItemCollectionPosition;
        Write-Host $camlQry.ListItemCollectionPosition.PagingInfo

    }
    while ($camlQry.ListItemCollectionPosition -ne $null)


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
