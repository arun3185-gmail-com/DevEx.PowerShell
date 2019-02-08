
################################################################################################################################################################
# Lotus Notes to SharePoint Migration Automation
#     Job Sheet Preparation
################################################################################################################################################################
Add-Type -Path "J:\LN2SP\References\LotusNotesWrapper\LN.vb"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.dll"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.Runtime.dll"
Import-Module "J:\LN2SP\References\EPPlus\EPPlus.dll"
################################################################################################################################################################

[string[]] $MigrationStatusOptions = @("Preparation","Ready","Migrate","Completed")
[string[]] $MigrationColumnOptions = @("No", "Yes", "TitleField", "SubFoldersField")

[string] $Global:Tab = [char]9
[string] $Global:LogTimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $ThisScriptRoot = "J:\LN2SP"
[string] $ThisScriptName = "MigrationAutomation_JobPreparation"

[string] $Global:LogFilePath = "$($ThisScriptRoot)\MigrationAutomationLogs.log"
[string] $JobsMetadataFilePath = "$($ThisScriptRoot)\JobsMetadataTest2.json"
#[string] $JobTemplateFullPath = "$($ThisScriptRoot)\References\Files\Template0.pmjob"
[string] $JobInfosFullPath = "$($ThisScriptRoot)\JobInfo"
#[string] $QuestJobsFullPath = "$($ThisScriptRoot)\QuestJobs"

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
#[xml] $pmJobXmlDoc = $null

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

                $excelSheet_Info      = $excelPkg.Workbook.Worksheets.Add("Info")
                $excelSheet_Forms     = $excelPkg.Workbook.Worksheets.Add("Forms")
                $excelSheet_SourceDef = $excelPkg.Workbook.Worksheets.Add("SourceDef")
                $excelSheet_DocInfos  = $excelPkg.Workbook.Worksheets.Add("DocInfos")


                ################################################################################
                # Sheet 1 - Info
                ################################################################################

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

                $excelSheet_Info.SetValue(8, 1, "Forms Count")
                $excelSheet_Info.SetValue(8, 2, $nForms.Length)

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
                
                $excelSheet_Info.SetValue(9, 1, "Documents Count")
                $excelSheet_Info.SetValue(9, 2, $docCollection.Count)


                ################################################################################
                # Sheet 2 - Forms
                ################################################################################

                $rowCounter = 0
                
                $rowCounter++
                $excelSheet_Forms.SetValue($rowCounter, 1, "Form Name")
                $excelSheet_Forms.SetValue($rowCounter, 2, "Docs Count")

                for ($i = 0; $i -lt $FormInfosList.Count; $i++)
                {
                    $rowCounter++
                    $excelSheet_Forms.SetValue($rowCounter, 1, $FormInfosList[$i].FormName)
                    $excelSheet_Forms.SetValue($rowCounter, 2, $FormInfosList[$i].Count)
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
                # Sheet 3 - SourceDef
                ################################################################################

                $rowCounter = 0
                
                $rowCounter++
                $excelSheet_SourceDef.SetValue($rowCounter, 1, "FieldName")
                $excelSheet_SourceDef.SetValue($rowCounter, 2, "FieldTypeNumber")
                $excelSheet_SourceDef.SetValue($rowCounter, 3, "FieldTypeName")

                for ($i = 0; $i -lt $SourceDefnColumnsList.Count; $i++)
                {
                    $rowCounter++
                    $excelSheet_SourceDef.SetValue($rowCounter, 1, $SourceDefnColumnsList[$i].FieldName)
                    $excelSheet_SourceDef.SetValue($rowCounter, 2, $SourceDefnColumnsList[$i].FieldTypeNumber)
                    $excelSheet_SourceDef.SetValue($rowCounter, 3, $SourceDefnColumnsList[$i].FieldTypeName)
                    $excelSheet_SourceDef.SetValue($rowCounter, 4, $MigrationColumnOptions[0])
                }

                [OfficeOpenXml.ExcelAddress] $xlAddrColumnOptsCell = New-Object OfficeOpenXml.ExcelAddress(2, 4, $rowCounter, 4)
                [OfficeOpenXml.DataValidation.ExcelDataValidationList] $xlColumnOptsDataValidationList = $excelSheet_SourceDef.DataValidations.AddListValidation($xlAddrColumnOptsCell.Address)
                
                foreach ($MigOpt in $MigrationColumnOptions)
                {
                    $xlColumnOptsDataValidationList.Formula.Values.Add($MigOpt)
                }
                $xlColumnOptsDataValidationList.AllowBlank = $false

                ################################################################################

                $excelSheet_Info.SetValue(10, 1, "Status")
                $excelSheet_Info.SetValue(10, 2, $MigrationStatusOptions[0])
                
                [OfficeOpenXml.ExcelAddress] $xlAddrStatusCell = New-Object OfficeOpenXml.ExcelAddress(10, 2, 10, 2)
                [OfficeOpenXml.DataValidation.ExcelDataValidationList] $xlStatusDataValidationList = $excelSheet_Info.DataValidations.AddListValidation($xlAddrStatusCell.Address)

                foreach ($MigOpt in $MigrationStatusOptions)
                {
                    $xlStatusDataValidationList.Formula.Values.Add($MigOpt)
                }
                $xlStatusDataValidationList.AllowBlank = $false
                

                ################################################################################
                # Sheet 4 - DocInfos
                ################################################################################

                $rowCounter = 0
                
                $rowCounter++
                $excelSheet_DocInfos.SetValue($rowCounter, 1, "NoteID")
                $excelSheet_DocInfos.SetValue($rowCounter, 2, "UniversalID")
                $excelSheet_DocInfos.SetValue($rowCounter, 3, "Form")
                $excelSheet_DocInfos.SetValue($rowCounter, 4, "NotesURL")
                $excelSheet_DocInfos.SetValue($rowCounter, 5, "Created")
                $excelSheet_DocInfos.SetValue($rowCounter, 6, "LastModified")

                for ($i = 0; $i -lt $DocInfosList.Count; $i++)
                {
                    $rowCounter++
                    $excelSheet_DocInfos.SetValue($rowCounter, 1, $DocInfosList[$i].NoteID)
                    $excelSheet_DocInfos.SetValue($rowCounter, 2, $DocInfosList[$i].UniversalID)
                    $excelSheet_DocInfos.SetValue($rowCounter, 3, $DocInfosList[$i].Form)
                    $excelSheet_DocInfos.SetValue($rowCounter, 4, $DocInfosList[$i].NotesURL)
                    $excelSheet_DocInfos.SetValue($rowCounter, 5, $DocInfosList[$i].Created)
                    $excelSheet_DocInfos.SetValue($rowCounter, 6, $DocInfosList[$i].LastModified)
                }


                ################################################################################

                <#
                Add data validation formula.
                #>

                ################################################################################

                $excelSheet_Info.Cells[1, 1, 10, 2].AutoFitColumns()

                ################################################################################
                
                [string] $genericFileName = "$($nDatabase.Title) on $($ServerName)"
                foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $genericFileName = $genericFileName.Replace($c, '_') }
                
                [string] $xlFilePath = "$($JobInfosFullPath)\$($genericFileName).xlsx"

                ################################################################################
                
                Write-Host "Saving - $($genericFileName)"
                
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
