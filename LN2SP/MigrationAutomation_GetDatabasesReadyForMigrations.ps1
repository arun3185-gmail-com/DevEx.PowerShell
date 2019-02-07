
################################################################################################################################################################
# Lotus Notes to SharePoint Migration Automation
#     GetDatabasesReadyForMigrations
################################################################################################################################################################
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.dll"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString
[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/7554"
[string] $ListTitle = "Master Source"

[string] $Global:Tab = [char]9
[string] $Global:LogTimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $ThisScriptRoot = "J:\LN2SP"
[string] $ThisScriptName = "MigrationAutomation_GetDatabasesReadyForMigrations"

[string] $Global:LogFilePath = "$($ThisScriptRoot)\MigrationAutomationLogs.log"
[string] $JobsMetadataFilePath = "$($ThisScriptRoot)\JobsMetadata.json"

################################################################################################################################################################
# Functions
################################################################################################################################################################

function Write-LogInfo()
{
    Param ([string] $Message)
    
    "$(Get-Date -Format $Global:LogTimeFormat):$($Global:Tab)GetDBReadyForMig$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[PSCustomObject[]] $Global:JobsMetadata = @()

try
{
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
    $context.Credentials = $credentials

    Write-Host "Querying web... $($SiteUrl)"

    $web = $context.Web
    $context.Load($web)

    $list = $context.Web.Lists.GetByTitle($ListTitle)
    $context.Load($list)
    
    $context.ExecuteQuery()
    
    #[string] $dtTimeSuffix = (Get-Date -Format "yyyyMMdd_HHmmss")

    $camlQry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $camlQry.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='TargetDecision' /><Value Type='Choice'>Migration</Value></Eq></Where></Query><RowLimit>5000</RowLimit></View>"

    do
    {
        Write-Host "Querying list... '$($ListTitle)'"
        
        $listItems = $list.GetItems($camlQry)
        $context.Load($listItems)
        $context.ExecuteQuery()
        
        Write-Host "Number of databases : $($listItems.Count)"
        
        foreach ($listItem in $listItems)
        {
            $newJobMetadata = New-Object PSObject -Property @{
                ListItemID = $listItem.ID
                Title = [string]::Empty
                ReplicaID = [string]::Empty
                FilepathReference = [string]::Empty
                ServerPath = [string]::Empty
                LNFilePath = [string]::Empty
                OwnerEmail = [string]::Empty
                TargetDecision = [string]::Empty
                Status = [string]::Empty
                LinkToSharepointSite = [string]::Empty
                Type = [string]::Empty
                DestinationType = [string]::Empty
                DocumentLevelSecurity = [string]::Empty
            }

            try
            {
                Write-Host $listItem["Title"].ToString()

                if ($listItem["Title"] -ne $null)                            { $newJobMetadata.Title = $listItem["Title"].ToString() }
                if ($listItem["ReplicaID"] -ne $null)                        { $newJobMetadata.ReplicaID = $listItem["ReplicaID"].ToString() }
                if ($listItem["TargetDecision"] -ne $null)                   { $newJobMetadata.TargetDecision = $listItem["TargetDecision"].ToString() }
                if ($listItem["Status"] -ne $null)                           { $newJobMetadata.Status = $listItem["Status"].ToString() }
                if ($listItem["LinkToSharepointSite"] -ne $null)             { $newJobMetadata.LinkToSharepointSite = $listItem["LinkToSharepointSite"].ToString() }
                if ($listItem["Type"] -ne $null)                             { $newJobMetadata.Type = $listItem["Type"].ToString() }
                if ($listItem["Destination_x0020_Type"] -ne $null)           { $newJobMetadata.DestinationType = $listItem["Destination_x0020_Type"].ToString() }
                if ($listItem["Document_x0020_Level_x0020_Secur"] -ne $null) { $newJobMetadata.DocumentLevelSecurity = $listItem["Document_x0020_Level_x0020_Secur"].ToString() }
                if ($listItem["FilepathReference"] -ne $null)
                {
                    $newJobMetadata.FilepathReference = $listItem["FilepathReference"].ToString()
                    [string] $strFilepathReference = $newJobMetadata.FilepathReference
                    [string[]] $partsOfFilepathReference = $strFilepathReference.Split(@("!!"), [System.StringSplitOptions]::RemoveEmptyEntries)

                    $newJobMetadata.ServerPath = $partsOfFilepathReference[0]
                    $newJobMetadata.LNFilePath = $partsOfFilepathReference[1]
                }
                if ($listItem["Owner"] -ne $null)
                {
                    [Microsoft.SharePoint.Client.FieldUserValue] $fuvOwner = [Microsoft.SharePoint.Client.FieldUserValue] $listItem["Owner"]
                    $newJobMetadata.OwnerEmail = $fuvOwner.Email
                }
                
            }
            catch
            {
                Write-LogInfo $_.Exception.Message
                Write-Host $_.Exception.Message -ForegroundColor Red
            }
            finally
            {
                $Global:JobsMetadata += $newJobMetadata
            }
        }

        $camlQry.ListItemCollectionPosition = $listItems.ListItemCollectionPosition;
        Write-Host $camlQry.ListItemCollectionPosition.PagingInfo

    }
    while ($camlQry.ListItemCollectionPosition -ne $null)

    ################################################################################
    Write-Host "Writing Metadata for job"

    $Global:JobsMetadata | ConvertTo-Json | Out-File -FilePath $JobsMetadataFilePath

    ################################################################################
}
catch
{
    Write-LogInfo $_.Exception.ToString()
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
