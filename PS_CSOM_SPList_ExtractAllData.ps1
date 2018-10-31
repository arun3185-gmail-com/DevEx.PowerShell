
################################################################################################################################################################
# SharePoint List Extract all data
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString
[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10426"
[string] $ListTitle = "Info"

[string] $DataFolderPath = "J:\Arun\Git\DevEx.Data"
[string] $XlFileNamePrefix = "$($SiteUrl.Substring($SiteUrl.IndexOf("sites/") + 6).Replace("/","_"))_$($ListTitle)"

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
    $ThisScriptName = "PS_CSOM_SPList_ExtractAllData"
}

[string] $LogFilePath  = "$($ThisScriptRoot)\$($ThisScriptName).log"

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet = $null

[string[]] $ignoreListColumnInternalNames = @("ID","LinkTitleNoMenu","LinkTitle","ComplianceAssetId","ContentType","ContentTypeId","_ModerationComments","LinkTitleNoMenu","_UIVersionString","Attachments","Edit","DocIcon","ItemChildCount","FolderChildCount","_ComplianceFlags","_ComplianceTag","_ComplianceTagWrittenTime","_ComplianceTagUserId","_IsRecord","AppAuthor","AppEditor")
[string[]] $selectedListColumnsInternalNames = @("ID")
[string[]] $ignoreDocLibColumnInternalNames = @("ID","ContentTypeId","ContentType","Created","Author","Modified","Editor","_HasCopyDestinations","_CopySource","_ModerationStatus","_ModerationComments","FileRef","FileDirRef","SortBehavior","PermMask","FileLeafRef","UniqueId","SyncClientId","ProgId","ScopeId","VirusStatus","_CheckinComment","LinkCheckedOutTitle","HTML_x0020_File_x0020_Type","_SourceUrl","_SharedFileIndex","_EditMenuTableStart","_EditMenuTableStart2","_EditMenuTableEnd","LinkFilenameNoMenu","LinkFilename","LinkFilename2","DocIcon","BaseName","FileSizeDisplay","MetaInfo","_Level","_IsCurrentVersion","Restricted","OriginatorId","NoExecute","ContentVersion","_ComplianceFlags","_ComplianceTag","_ComplianceTagWrittenTime","_ComplianceTagUserId","_IsRecord","BSN","_ListSchemaVersion","_Dirty","_Parsable","AccessPolicy","_VirusStatus","_VirusVendorID","_VirusInfo","_CommentFlags","_CommentCount","_LikeCount","_RmsTemplateId","_IpLabelId","_DisplayName","AppAuthor","AppEditor","SMTotalSize","SMLastModifiedDate","SMTotalFileStreamSize","SMTotalFileCount","SelectTitle","SelectFilename","Edit","owshiddenversion","_UIVersion","_UIVersionString","InstanceID","Order","WorkflowVersion","WorkflowInstanceID","ParentVersionString","ParentLeafName","DocConcurrencyNumber","ParentUniqueId","StreamHash","ComplianceAssetId","Title","TemplateUrl","xd_ProgID","xd_Signature","Combine","RepairDocument","_ShortcutUrl","_ShortcutSiteId","_ShortcutWebId","_ShortcutUniqueId")
[string[]] $selectedDocLibColumnsInternalNames = @("ID","Title","FSObjType","FileDirRef","FileLeafRef","FileRef","ServerUrl","EncodedAbsUrl","Created","Author","Modified","Editor")
[string[]] $selectedColumnsInternalNames = $null
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
    $fieldCollection = $list.Fields
    $context.Load($fieldCollection)
    $context.ExecuteQuery()

    $excelPkg = New-Object OfficeOpenXml.ExcelPackage
    $excelSheet = $excelPkg.Workbook.Worksheets.Add("Sheet1")
    $listBaseType = $list.BaseType

    if ($listBaseType -eq [Microsoft.SharePoint.Client.BaseType]::GenericList)
    {
        $selectedColumnsInternalNames = $selectedListColumnsInternalNames
        foreach ($field in $fieldCollection)
        {
            if ($field.InternalName -notin $ignoreListColumnInternalNames -and -not($field.Hidden))
            {            
                $selectedColumnsInternalNames += $field.InternalName
            }
        }
    }
    elseif ($listBaseType -eq [Microsoft.SharePoint.Client.BaseType]::DocumentLibrary)
    {
        $selectedColumnsInternalNames = $selectedDocLibColumnsInternalNames
        foreach ($field in $fieldCollection)
        {
            if ($field.InternalName -notin $ignoreDocLibColumnInternalNames -and -not($field.Hidden))
            {            
                $selectedColumnsInternalNames += $field.InternalName
            }
        }
    }

    Write-Host "Adding field headers to excel..."
    
    [int] $rowCounter = 1
    for ($i = 0; $i -lt $selectedColumnsInternalNames.Count; $i++)
    {
        $excelSheet.SetValue($rowCounter, $i + 1, $selectedColumnsInternalNames[$i])
    }
    $rowCounter++

    
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
                for ($i = 0; $i -lt $selectedColumnsInternalNames.Count; $i++)
                {
                    $colName = $selectedColumnsInternalNames[$i]
                    if ($listItem[$colName] -ne $null)
                    {
                        if ($colName -eq "ID")
                        {
                            $excelSheet.SetValue($rowCounter, $i + 1, [int]::Parse($listItem[$colName].ToString()))
                        }
                        elseif ($listItem[$colName].GetType().ToString() -eq "Microsoft.SharePoint.Client.FieldUserValue")
                        {
                            [Microsoft.SharePoint.Client.FieldUserValue] $fuv = [Microsoft.SharePoint.Client.FieldUserValue] $listItem[$colName]
                            $excelSheet.SetValue($rowCounter, $i + 1, $fuv.LookupValue)
                        }
                        elseif ($listItem[$colName].GetType().ToString() -eq "Microsoft.SharePoint.Client.FieldUrlValue")
                        {
                            [Microsoft.SharePoint.Client.FieldUrlValue] $fuv = [Microsoft.SharePoint.Client.FieldUrlValue] $listItem[$colName]
                            $excelSheet.SetValue($rowCounter, $i + 1, $fuv.Url)
                        }
                        elseif ($listItem[$colName].GetType().ToString() -eq "Microsoft.SharePoint.Client.FieldLookupValue")
                        {
                            [Microsoft.SharePoint.Client.FieldLookupValue] $flv = [Microsoft.SharePoint.Client.FieldLookupValue] $listItem[$colName]
                            $excelSheet.SetValue($rowCounter, $i + 1, $flv.LookupValue)
                        }
                        else
                        {
                            $excelSheet.SetValue($rowCounter, $i + 1, $listItem[$colName].ToString())
                        }
                    }
                }
            }
            catch
            {
                Write-Host $_.Exception.Message
            }
            finally
            {
                $rowCounter++
            }
        }

        $camlQry.ListItemCollectionPosition = $listItems.ListItemCollectionPosition;
        Write-Host $camlQry.ListItemCollectionPosition.PagingInfo

    } while ($camlQry.ListItemCollectionPosition -ne $null)


    
    [string] $xlFilePath = "$($DataFolderPath)\$($XlFileNamePrefix)_$($dtTimeSuffix).xlsx"

    if (!(Test-Path -Path $DataFolderPath)) { New-Item -Path $DataFolderPath -ItemType "directory" }

    $excelPkg.SaveAs((New-Object System.IO.FileInfo($xlFilePath)))
}
catch
{
    #Write-Host    $_.Exception.ToString() -ForegroundColor Red
    throw
}
finally
{
    if ($excelSheet -ne $null) { $excelSheet.Dispose(); $excelSheet = $null }
    if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
