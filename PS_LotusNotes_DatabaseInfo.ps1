
################################################################################################################################################################
# Lotus Notes All Documents List
################################################################################################################################################################
Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
################################################################################################################################################################

[string] $ServerName = "APACApp01/Server/Evonik"
[string] $LNFilePath = "teamdoc/td0000103.nsf"

[string] $Global:Tab = [char]9
[string] $Global:LogTimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $Global:ThisScriptRoot = @("J:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[string] $Global:ThisScriptName = $null

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}
else
{
    $Global:ThisScriptName = "PS_LotusNotes_DatabaseInfo"
}
[string] $Global:LogFilePath  = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"


[string[]] $AdditionalFields = @("theme","division1","division2","division3","division4","DocNr","docdescription","Keywords","status")

################################################################################################################################################################
# Functions
################################################################################################################################################################

function Write-LogInfo()
{
    Param ([string] $Message)
    
    "$(Get-Date -Format $Global:LogTimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

function CreateFormInfoObject()
{
    Param ([string] $FormName)
    
    $Local:newLNFormInfo = New-Object PSObject
    $Local:newLNFormInfo | Add-Member NoteProperty -TypeName "System.String" -Name "FormName" -Value $FormName
    $Local:newLNFormInfo | Add-Member NoteProperty -TypeName "System.Int32"  -Name "Count"    -Value 0

    Return $Local:newLNFormInfo
}

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[LN.NotesSession] $nSession = $null
[LN.NotesDatabase] $nDatabase = $null
[LN.NotesFormCollection] $nForms = $null
[LN.NotesDocumentCollection] $docCollection = $null
[PSCustomObject[]] $Global:LNFormInfos = @()

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Info = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Forms = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet_Documents = $null

[int] $rowCounter = 1
[int] $colCounter = 1

try
{
    Write-Host "Initializing NotesSession..."

    $nSession = New-Object LN.NotesSession
    $nDatabase = $nSession.GetDatabase($ServerName, $LNFilePath)
    $docCollection = $nDatabase.AllDocuments
    $nForms = $nDatabase.Forms

    ################################################################################

    Write-Host "Reading Database..."

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

    $excelSheet_Info.SetValue(9, 1, "Forms Count")
    $excelSheet_Info.SetValue(9, 2, $nForms.Length)
    $excelSheet_Info.SetValue(10, 1, "Documents Count")
    $excelSheet_Info.SetValue(10, 2, $docCollection.Count)

    $excelSheet_Info.Cells["A1:B10"].AutoFitColumns()

    
    ################################################################################

    Write-Host "Reading Forms..."

    $excelSheet_Forms = $excelPkg.Workbook.Worksheets.Add("Forms")
    $rowCounter = 1

    $excelSheet_Forms.SetValue($rowCounter, 1, "Form Name")
    $excelSheet_Forms.SetValue($rowCounter, 2, "Docs Count")
    
    for ($i = 0; $i -lt $nForms.Length; $i++)
    {
        if ($Global:LNFormInfos.Where({ $PSItem.FormName -eq $nForms[$i].Name }).Count -eq 0)
        {
            $rowCounter++
            $excelSheet_Forms.SetValue($rowCounter, 1, $nForms[$i].Name)
            $Global:LNFormInfos += CreateFormInfoObject -FormName $nForms[$i].Name
        }
    }
    

    ################################################################################
    
    Write-Host "Reading Documents..."

    $excelSheet_Documents = $excelPkg.Workbook.Worksheets.Add("Documents")
    $rowCounter = 1
    $colCounter = 1
    
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "NoteID")
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "UniversalID")
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "Form")
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "NotesURL")
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "Created")
    $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), "LastModified")

    
    foreach ($AddnField in $AdditionalFields)
    {
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $AddnField)
    }

    [LN.NotesDocument] $doc = $docCollection.GetFirstDocument()
    while ($doc -ne $null)
    {
        $rowCounter++
        $colCounter = 1

        if ($Global:LNFormInfos.Where({ $PSItem.FormName -eq $doc.GetFirstItem("Form").Text }).Count -eq 0)
        {
            $Global:LNFormInfos += CreateFormInfoObject -FormName $doc.GetFirstItem("Form").Text
        }
        $Global:LNFormInfos.Where({ $PSItem.FormName -eq $doc.GetFirstItem("Form").Text })[0].Count++

        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.NoteID)
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.UniversalID)
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.GetFirstItem("Form").Text)
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.NotesURL)
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.Created)
        $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.LastModified)

        foreach ($AddnField in $AdditionalFields)
        {
            $excelSheet_Documents.SetValue($rowCounter, ($colCounter++), $doc.GetFirstItem($AddnField).Text)
        }
        
        $doc = $docCollection.GetNextDocument($doc)
    }
    
    ################################################################################
    
    Write-Host "Updating Forms..."

    for ($i = 0; $i -lt $Global:LNFormInfos.Count; $i++)
    {
        $excelSheet_Forms.SetValue(($i + 2), 1, $Global:LNFormInfos[$i].FormName)
        $excelSheet_Forms.SetValue(($i + 2), 2, $Global:LNFormInfos[$i].Count)
    }

    ################################################################################

    Write-Host "Saving Excel..."

    [string] $xlFileName = "DatabaseInfo - $($nDatabase.Title).xlsx"
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars())
    {
        $xlFileName = $xlFileName.Replace($c, '_')
    }
    [string] $xlFilePath = "$($Global:ThisScriptRoot)\$($xlFileName)"

    $excelPkg.SaveAs((New-Object System.IO.FileInfo($xlFilePath)))

    ################################################################################
}
catch
{
    Write-LogInfo $_.Exception.ToString()
    throw
}
finally
{
    if ($excelSheet_Info -ne $null) { $excelSheet_Info.Dispose(); $excelSheet_Info = $null }
    if ($excelSheet_Forms -ne $null) { $excelSheet_Forms.Dispose(); $excelSheet_Forms = $null }
    if ($excelSheet_Documents -ne $null) { $excelSheet_Documents.Dispose(); $excelSheet_Documents = $null }
    if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
