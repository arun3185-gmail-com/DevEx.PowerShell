
################################################################################################################################################################
# Lotus Notes to SharePoint Migration Automation
#     Job Preparation
################################################################################################################################################################
Add-Type -Path "J:\LN2SP\References\LotusNotesWrapper\LN.vb"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.dll"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.Runtime.dll"
Import-Module "J:\LN2SP\References\EPPlus\EPPlus.dll"
################################################################################################################################################################

[string] $Global:Tab = [char]9
[string] $Global:LogTimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $ThisScriptRoot = "J:\LN2SP"
[string] $ThisScriptName = "MigrationAutomation_JobPreparation"

[string] $Global:LogFilePath = "$($ThisScriptRoot)\MigrationAutomationLogs.log"
[string] $JobsMetadataFilePath = "$($ThisScriptRoot)\JobsMetadataTest2.json"
[string] $JobTemplateFullPath = "$($ThisScriptRoot)\References\Files\Template0.pmjob"
[string] $JobInfosFullPath = "$($ThisScriptRoot)\JobInfo"
[string] $QuestJobsFullPath = "$($ThisScriptRoot)\QuestJobs"

[string] $SPUsername = "B13501@evonik.com"
[SecureString] $SPPassword = Read-Host "Enter Password for $($SPUsername)" -AsSecureString
[string] $SPSiteUrl   = $null
[string] $SPListTitle = $null

################################################################################################################################################################
# Functions
################################################################################################################################################################

function Write-LogInfo()
{
    Param ([string] $Message)
    
    "$(Get-Date -Format $Global:LogTimeFormat):$($Global:Tab)JobPreparation$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

function New-FormInfo()
{
    Param ([string] $FormName)
    
    $Local:newLNFormInfo = New-Object PSObject
    $Local:newLNFormInfo | Add-Member NoteProperty -TypeName "System.String" -Name "FormName" -Value $FormName
    $Local:newLNFormInfo | Add-Member NoteProperty -TypeName "System.Int32"  -Name "Count"    -Value 0

    Return $Local:newLNFormInfo
}

function New-DocInfo()
{
    Param
    (
        [string] $NoteID,
        [string] $UniversalID,
        [string] $Form,
        [string] $NotesURL
    )

    $Local:newDocInfo = New-Object PSObject
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "NoteID"       -Value $NoteID
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "UniversalID"  -Value $UniversalID
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "Form"         -Value $Form
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "NotesURL"     -Value $NotesURL
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "Created"      -Value $NotesURL
    $Local:newDocInfo | Add-Member NoteProperty -TypeName "System.String" -Name "LastModified" -Value $NotesURL

    Return $Local:newDocInfo
}

function New-SourceDefinitionColumn()
{
    Param
    (
        [string] $FieldName,
        [int]    $FieldTypeNumber,
        [string] $FieldTypeName = "",
        [string] $ReturnType = ""
    )

    $Local:SrcDefCol = New-Object PSObject
    $Local:SrcDefCol | Add-Member NoteProperty -TypeName "System.String" -Name "FieldName"       -Value $FieldName
    $Local:SrcDefCol | Add-Member NoteProperty -TypeName "System.Int32"  -Name "FieldTypeNumber" -Value $FieldTypeNumber
    $Local:SrcDefCol | Add-Member NoteProperty -TypeName "System.String" -Name "FieldTypeName"   -Value $FieldTypeName
    $Local:SrcDefCol | Add-Member NoteProperty -TypeName "System.String" -Name "ReturnType"      -Value $ReturnType

    Return $Local:SrcDefCol
}

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[LN.NotesSession] $nSession = $null
[LN.NotesDatabase] $nDatabase = $null
[LN.NotesDocumentCollection] $docCollection = $null

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Info = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Forms = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_SourceDef = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Documents = $null

[string] $jsonStrJobsMetadata = $null
[PSCustomObject[]] $JobsMetadata = $null
[xml] $pmJobXmlDoc = $null

[int] $rowCounter = 1


try
{
    $jsonStrJobsMetadata = [System.IO.File]::ReadAllText($JobsMetadataFilePath)
    $JobsMetadata = ConvertFrom-Json $jsonStrJobsMetadata

    $nSession = New-Object LN.NotesSession

    foreach ($jobMetadata in $JobsMetadata)
    {
        <#
        if (-not([string]::IsNullOrWhiteSpace($jobMetadata.Title)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.DestinationType)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.OwnerEmail)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.ServerPath)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.LNFilePath)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.LinkToSharepointSite)))
        { }
        #>

        if (-not([string]::IsNullOrWhiteSpace($jobMetadata.ServerPath)) -and
            -not([string]::IsNullOrWhiteSpace($jobMetadata.LNFilePath)))
        {
            [System.Collections.Generic.List[PSCustomObject]] $FormInfosList = $null
            [System.Collections.Generic.IEnumerable[PSCustomObject]] $MigFormInfosList = $null
            [System.Collections.Generic.List[PSCustomObject]] $SourceDefnColumnsList = $null
            [System.Collections.Generic.List[PSCustomObject]] $DocInfosList = $null
            [System.Collections.Generic.List[string]] $FormNamesList = $null

            try
            {
                $FormInfosList = New-Object System.Collections.Generic.List[PSCustomObject]
                $MigFormInfosList = New-Object System.Collections.Generic.List[PSCustomObject]
                $SourceDefnColumnsList = New-Object System.Collections.Generic.List[PSCustomObject]
                $DocInfosList = New-Object System.Collections.Generic.List[PSCustomObject]
                $FormNamesList = New-Object System.Collections.Generic.List[string]

                ################################################################################

                $nDatabase = $nSession.GetDatabase($jobMetadata.ServerPath, $jobMetadata.LNFilePath)
                [string] $ServerName = $nDatabase.Server.Split('/')[0]
            
                $excelPkg = New-Object OfficeOpenXml.ExcelPackage
                $excelSheet_Info = $excelPkg.Workbook.Worksheets.Add("Info")

                $excelSheet_Info.SetValue(1, 1, "Title")
                $excelSheet_Info.SetValue(1, 2, $nDatabase.Title)
                $excelSheet_Info.SetValue(2, 1, "ReplicaID")
                $excelSheet_Info.SetValue(2, 2, $nDatabase.ReplicaID)
                $excelSheet_Info.SetValue(3, 1, "TemplateName")
                $excelSheet_Info.SetValue(3, 2, $nDatabase.TemplateName)
                $excelSheet_Info.SetValue(4, 1, "Server")
                $excelSheet_Info.SetValue(4, 2, $nDatabase.Server)
                $excelSheet_Info.SetValue(5, 1, "FilePath")
                $excelSheet_Info.SetValue(5, 2, $nDatabase.FilePath)
                $excelSheet_Info.SetValue(6, 1, "FileName")
                $excelSheet_Info.SetValue(6, 2, $nDatabase.FileName)
                $excelSheet_Info.SetValue(7, 1, "NotesURL")
                $excelSheet_Info.SetValue(7, 2, $nDatabase.NotesURL)

                $nForms = $nDatabase.Forms

                for ($i = 0; $i -lt $nForms.Length; $i++)
                {
                    if ($FormInfosList.Where({ $PSItem.FormName -eq $nForms[$i].Name }).Count -eq 0)
                    {
                        $FormInfosList.Add((New-FormInfo -FormName $nForms[$i].Name))
                    }
                }

                $excelSheet_Info.SetValue(9, 1, "Forms Count")
                $excelSheet_Info.SetValue(9, 2, $nForms.Length)

                ################################################################################
                
                $docCollection = $nDatabase.AllDocuments
                [LN.NotesDocument] $doc = $docCollection.GetFirstDocument()
                while ($doc -ne $null)
                {
                    $lnDocInfo = New-DocInfo -NoteID $doc.NoteID -UniversalID $doc.UniversalID -Form $doc.GetFirstItem("Form").Text -NotesURL $doc.NotesURL
                    $DocInfosList.Add($lnDocInfo)

                    if ($FormInfosList.Where({ $PSItem.FormName -eq $doc.GetFirstItem("Form").Text }).Count -eq 0)
                    {
                        $FormInfosList.Add((New-FormInfo -FormName $doc.GetFirstItem("Form").Text))
                    }
                    $FormInfosList.Where({ $PSItem.FormName -eq $doc.GetFirstItem("Form").Text })[0].Count++

                    $doc = $docCollection.GetNextDocument($doc)
                }
                
                $excelSheet_Info.SetValue(10, 1, "Documents Count")
                $excelSheet_Info.SetValue(10, 2, $docCollection.Count)

                ################################################################################

                $excelSheet_Forms = $excelPkg.Workbook.Worksheets.Add("Forms")
                $rowCounter = 1

                $excelSheet_Forms.SetValue($rowCounter, 1, "Form Name")
                $excelSheet_Forms.SetValue($rowCounter, 2, "Docs Count")
                $rowCounter++

                for ($i = 0; $i -lt $FormInfosList.Count; $i++)
                {
                    $excelSheet_Forms.SetValue($rowCounter, 1, $FormInfosList[$i].FormName)
                    $excelSheet_Forms.SetValue($rowCounter, 2, $FormInfosList[$i].Count)
                    $rowCounter++
                }
            
                ################################################################################

                $MigFormInfosList = $FormInfosList.Where({ $PSItem.Count -gt 10 -and -not($PSItem.FormName.StartsWith("$")) -and -not($PSItem.FormName.StartsWith("(")) -and -not($PSItem.FormName.StartsWith("_")) })
                foreach ($formInfo in $MigFormInfosList) { $FormNamesList.Add($formInfo.FormName) }

                for ($i = 0; $i -lt $nDatabase.Forms.Length; $i++)
                {
                    if ($MigFormInfosList.Where({ $PSItem.FormName -eq $nDatabase.Forms[$i].Name }).Count -eq 1)
                    {
                        foreach ($fieldName in $nDatabase.Forms[$i].Fields)
                        {
                            if ($SourceDefnColumnsList.Where({ $PSItem.FieldName -eq $fieldName }).Count -eq 0)
                            {
                                $srcDefColumn = New-SourceDefinitionColumn -FieldName $fieldName -FieldTypeNumber ($nDatabase.Forms[$i].GetFieldType($fieldName))
                                $SourceDefnColumnsList.Add($srcDefColumn)
                            }
                        }
                    }
                }

                [string[]] $checkedUNIDs = @()
                while ($SourceDefnColumnsList.Where({ $PSItem.FieldTypeNumber -eq 0 }).Count -gt 0)
                {
                    $qlnDocInfo = $DocInfosList.Where({ $PSItem.UniversalID -notin $checkedUNIDs -and $PSItem.Form -in $FormNamesList }, 'Default', 1)[0]
                
                    [LN.NotesDocument] $qDoc = $nDatabase.GetDocumentByUNID($qlnDocInfo.UniversalID)
                    for ($i = 0; $i -lt $qDoc.Items.Length; $i++)
                    {
                        [LN.NotesItem] $itm = $qDoc.Items[$i]
                        $SourceDefnColumnsList.Where({ $PSItem.FieldTypeNumber -eq 0 -and $PSItem.FieldName -eq $itm.Name }).ForEach({ $PSItem.FieldTypeNumber = $itm.Type; $PSItem.FieldTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $itm.Type) })
                    }

                    $checkedUNIDs += $qlnDocInfo.UniversalID
                }
                $SourceDefnColumnsList.Where({ [string]::IsNullOrWhiteSpace($PSItem.FieldTypeName) }).ForEach({ $PSItem.FieldTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $PSItem.FieldTypeNumber) })
                
                ################################################################################

                $excelSheet_SourceDef = $excelPkg.Workbook.Worksheets.Add("SourceDef")
                $rowCounter = 1

                $excelSheet_SourceDef.SetValue($rowCounter, 1, "FieldName")
                $excelSheet_SourceDef.SetValue($rowCounter, 2, "FieldTypeNumber")
                $excelSheet_SourceDef.SetValue($rowCounter, 3, "FieldTypeName")
                $rowCounter++

                for ($i = 0; $i -lt $SourceDefnColumnsList.Count; $i++)
                {
                    $excelSheet_SourceDef.SetValue($rowCounter, 1, $SourceDefnColumnsList[$i].FieldName)
                    $excelSheet_SourceDef.SetValue($rowCounter, 2, $SourceDefnColumnsList[$i].FieldTypeNumber)
                    $excelSheet_SourceDef.SetValue($rowCounter, 3, $SourceDefnColumnsList[$i].FieldTypeName)
                    $rowCounter++
                }

                ################################################################################

                $pmJobXmlDoc = New-Object xml
                $pmJobXmlDoc.Load($JobTemplateFullPath)
                
                ########################################
                # /TransferJob/QuerySource
                ########################################

                [string[]] $serverNameParts = $jobMetadata.ServerPath.Split('/')
                if ($serverNameParts.Count -gt 1)
                {
                    [System.Xml.XmlNode] $jobNode_QrySrc_ConnString = $pmJobXmlDoc.SelectSingleNode("/TransferJob/QuerySource/ConnectionString")
                
                    $jobNode_QrySrc_ConnString.InnerXml = "server='CN=$($serverNameParts[0])"
                    for ($i = 1; $i -lt ($serverNameParts.Count-1); $i++)
                    {
                        $jobNode_QrySrc_ConnString.InnerXml += "/OU=$($serverNameParts[$i])"
                    }
                    $jobNode_QrySrc_ConnString.InnerXml += "/O=$($serverNameParts[$serverNameParts.Count-1])'; database='$($LNFilePath)'; zone=utc"
                }

                ########################################
                # /TransferJob/SourceDefinition
                ########################################

                [System.Xml.XmlNode] $jobNode_SrcDef = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition")
                $jobNode_SrcDef.Attributes["Name"].InnerText = $nDatabase.Title
                $jobNode_SrcDef.Attributes["Templates"].InnerText = $nDatabase.DesignTemplateName

                [System.Xml.XmlNode] $jobNode_SrcDef_QrySpec = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec")

                [System.Xml.XmlNode] $jobNode_SrcDef_ReplicaID = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/ReplicaId")
                $jobNode_SrcDef_ReplicaID.InnerText = $nDatabase.ReplicaID

                [System.Xml.XmlNode] $jobNode_SrcDef_Unid = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/UNID")
                
                foreach ($sourceDefnColumn in $SourceDefnColumnsList)
                {
                    <#
                     # Insert columns
                    #>

                    [System.Xml.XmlElement] $colElmnt = $pmJobXmlDoc.CreateElement("Column")

                    if ($sourceDefnColumn.FieldTypeName -in @("NUMBERS"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "Number")
                        $colElmnt.SetAttribute("Option", "Multi")
                    }
                    if ($sourceDefnColumn.FieldTypeName -in @("DATETIMES","DATETIMES_1025","NAMES","READERS","AUTHORS","TEXT","TEXT_1281","RICHTEXT"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "String")
                        $colElmnt.SetAttribute("Option", "Multi")
                    }
                    <#
                    if ($sourceDefnColumn.FieldTypeName -in @("DATETIMES","DATETIMES_1025"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "Date")
                        $colElmnt.SetAttribute("Option", "Multi")
                    }
                    if ($sourceDefnColumn.FieldTypeName -in @("NAMES","READERS","AUTHORS"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "User")
                        $colElmnt.SetAttribute("Option", "Multi")
                    }
                    if ($sourceDefnColumn.FieldTypeName -in @("TEXT","TEXT_1281"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "String")
                        $colElmnt.SetAttribute("Option", "Multi")
                    }
                    if ($sourceDefnColumn.FieldTypeName -in @("RICHTEXT"))
                    {
                        $colElmnt.SetAttribute("ColumnType", "Item")
                        $colElmnt.SetAttribute("ReturnType", "HtmlString")
                        $colElmnt.SetAttribute("Option", "Html")
                    }
                    #>
                    
                    $colElmnt.SetAttribute("Value", $sourceDefnColumn.FieldName)
                    #$colElmnt.SetAttribute("Alias", "")

                    $jobNode_SrcDef_QrySpec.InsertAfter($colElmnt, $jobNode_SrcDef_Unid)
                }
                
                [System.Xml.XmlNode] $jobNode_SrcDef_Forms = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/Forms")
                $jobNode_SrcDef_Forms.InnerText = [string]::Join(";", $FormNamesList.ToArray())

                ########################################
                # SharePoint List/Library creations
                ########################################

                try
                {
                    Write-Host "Querying web/list data..."

                    $spoCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPUsername , $SPPassword)
                    $spContext = New-Object Microsoft.SharePoint.Client.ClientContext($SPSiteUrl)
                    $spContext.Credentials = $spoCredentials

                    $spWeb = $context.Web
                    $spLists = $context.Web.Lists

                    $spContext.Load($spWeb)
                    $spContext.Load($spLists)
                    $spContext.ExecuteQuery()

                    $spList = $spLists | Where { $_.Title -eq $SPListTitle }
                    $spListFields = $null

                    if ($spList)
                    {
                        Write-Host "$($SPListTitle) - SharePoint List/Library Exists"

                        $spList = $spContext.Web.Lists.GetByTitle($SPListTitle)
                        $spListFields = $spList.Fields
                    }
                    else
                    {
                        Write-Host "$($SPListTitle) - Creating SharePoint List/Library..."

                        $newSPListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
                        $newSPListInfo.QuickLaunchOption = [Microsoft.SharePoint.Client.QuickLaunchOptions]::On
                        $newSPListInfo.Title = $SPListTitle
                        
                        if ($jobMetadata.DestinationType -eq "List")
                        {
                            $newSPListInfo.TemplateType = [int] [Microsoft.SharePoint.Client.ListTemplateType]::GenericList
                        }
                        elseif ($jobMetadata.DestinationType.Contains("Doc") -or $jobMetadata.DestinationType.Contains("Lib"))
                        {
                            $newSPListInfo.TemplateType = [int] [Microsoft.SharePoint.Client.ListTemplateType]::DocumentLibrary
                        }
                        
                        $spList = $spContext.Web.Lists.Add($newSPListInfo)
                        $spListFields = $spList.Fields
                    }
                    
                    $spContext.Load($spList)
                    $spContext.Load($spListFields)
                    $spContext.ExecuteQuery()

                    foreach ($sourceDefnColumn in $SourceDefnColumnsList)
                    {
                    }
                }
                catch
                {
                    Write-Host $_.Exception.ToString() -ForegroundColor Red
                }

                ########################################
                # /TransferJob/SharePointConnection
                ########################################

                [System.Xml.XmlNode] $jobNode_SP_Web = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointConnection/Web")
                [System.Xml.XmlNode] $jobNode_SP_List = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointConnection/List")
                
                $jobNode_SP_Web.InnerText = ""
                $jobNode_SP_List.InnerText = ""

                ########################################
                # /TransferJob/JobOptions
                ########################################
                
                [System.Xml.XmlNode] $jobNode_JobOpts_DupDocHandling = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/DuplicateDocumentHandling")
                [System.Xml.XmlNode] $jobNode_JobOpts_PreserveIdentities = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/PreserveIdentities")
                [System.Xml.XmlNode] $jobNode_JobOpts_PreserveDates = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/PreserveDates")
                
                [System.Xml.XmlNode] $jobNode_JobOpts_QryOptns_DelMigDocs = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/QueryOptions/DeleteMigratedDocuments")
                [System.Xml.XmlNode] $jobNode_JobOpts_QryOptns_ExtRecPat = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/QueryOptions/ExtractRecurrencePatterns")
                [System.Xml.XmlNode] $jobNode_JobOpts_QryOptns_ExtDocSec = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/QueryOptions/ExtractDocSecurity")

                [System.Xml.XmlNode] $jobNode_JobOpts_UsrMapOptns_MapFailSubs = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/UserMappingOptions/MappingFailureSubstitution")
                [System.Xml.XmlNode] $jobNode_JobOpts_UsrMapOptns_MapForUsr = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/UserMappingOptions/UseMappingFailureSubstitutionForUserFields")
                [System.Xml.XmlNode] $jobNode_JobOpts_UsrMapOptns_DefUsr = $pmJobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/UserMappingOptions/DefaultUserName")

                $jobNode_JobOpts_DupDocHandling.InnerText = "Replace"
                $jobNode_JobOpts_PreserveIdentities.InnerText = "true"
                $jobNode_JobOpts_PreserveDates.InnerText = "true"

                $jobNode_JobOpts_QryOptns_DelMigDocs.InnerText = "false"
                $jobNode_JobOpts_QryOptns_ExtRecPat.InnerText = "false"
                $jobNode_JobOpts_QryOptns_ExtDocSec.InnerText = "true"
                
                $jobNode_JobOpts_UsrMapOptns_MapFailSubs.InnerText = "DefaultIdentity"
                $jobNode_JobOpts_UsrMapOptns_MapForUsr.InnerText = "true"
                $jobNode_JobOpts_UsrMapOptns_DefUsr.InnerText = $jobMetadata.OwnerEmail

                ########################################
                # /TransferJob/SecurityMapping
                ########################################

                [System.Xml.XmlNode] $jobNode_SecMap = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SecurityMapping")
                $jobNode_SecMap.Attributes["Enabled"].InnerText = "true"

                ########################################
                # /TransferJob/SharePointTargetDefinition
                ########################################

                [System.Xml.XmlNode] $jobNode_SPDef = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition")
                [System.Xml.XmlNode] $jobNode_SPDef_ExtIcons = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/ExtractIcons")
                [System.Xml.XmlNode] $jobNode_SPDef_MigCustProps = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/MigrateCustomProperties")
                [System.Xml.XmlNode] $jobNode_SPDef_IsDocLib = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/IsDocLib")
                [System.Xml.XmlNode] $jobNode_SPDef_IsDiscs = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/IsDiscussion")
                [System.Xml.XmlNode] $jobNode_SPDef_IsEvnt = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/IsEvents")
                [System.Xml.XmlNode] $jobNode_SPDef_Attchmts = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/AllowAttachments")

                $jobNode_SPDef.Attributes["Name"].InnerText = "Custom List"
                $jobNode_SPDef.Attributes["Templates"].InnerText = "Custom List"
                $jobNode_SPDef.Attributes["SharePointTemplates"].InnerText = "Custom List"

                $jobNode_SPDef_ExtIcons.InnerText = "true"
                $jobNode_SPDef_MigCustProps.InnerText = "false"
                $jobNode_SPDef_IsDocLib.InnerText = "false"
                $jobNode_SPDef_IsDiscs.InnerText = "false"
                $jobNode_SPDef_IsEvnt.InnerText = "false"
                $jobNode_SPDef_Attchmts.InnerText = "true"

                <#
                 # Fields
                #>
                
                [System.Xml.XmlNode] $jobNode_SPDef_VwsOwrt = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/ViewsOverwriteExisting")
                [System.Xml.XmlNode] $jobNode_SPDef_EnbVers = $pmJobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition/EnableVersioning")

                $jobNode_SPDef_VwsOwrt.InnerText = "false"
                $jobNode_SPDef_EnbVers.InnerText = "false"

                ########################################
                # /TransferJob/Mapping
                ########################################

                <#
                 # Mapping
                #>
                

                ################################################################################
                
                [string] $genericFileName = "$($nDatabase.Title) on $($ServerName)"
                foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $genericFileName = $genericFileName.Replace($c, '_') }
                
                [string] $jobFilePath = "$($QuestJobsFullPath)\$($genericFileName).pmjob"
                [string] $xlFilePath = "$($JobInfosFullPath)\$($genericFileName).xlsx"

                ################################################################################
                
                Write-Host "Saving - $($genericFileName)"
                $pmJobXmlDoc.Save($jobFilePath)
                $excelPkg.SaveAs((New-Object System.IO.FileInfo($xlFilePath)))

                ################################################################################
            }
            catch
            {
                #Write-LogInfo $_.Exception.ToString()
                throw
            }
            finally
            {
                if ($excelSheet_Documents -ne $null) { $excelSheet_Documents.Dispose(); $excelSheet_Documents = $null }
                if ($excelSheet_SourceDef -ne $null) { $excelSheet_SourceDef.Dispose(); $excelSheet_SourceDef = $null }
                if ($excelSheet_Forms -ne $null) { $excelSheet_Forms.Dispose(); $excelSheet_Forms = $null }
                if ($excelSheet_Info -ne $null) { $excelSheet_Info.Dispose(); $excelSheet_Info = $null }
                if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }
            }

        }
    }

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
