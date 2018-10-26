
################################################################################################################################################################

Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"

################################################################################################################################################################

#[string] $XmlFilePath = "J:\Arun\DevEx\Logs\VB.Net_Ex2.xml"

[string] $BackupPath = "J:\Arun\DevEx\Data"
[string] $XlFileNamePrefix = "Birmingham_ESH_Observations"
[string] $SheetName = "Birmingham_ESH_Observations"

[string] $ServerName = "AmericasApp02/Server/Evonik"
[string] $LNFilePath = "HN/CIAOBhamESHOb.nsf"

$ArrayOfDefaultFields = 
@(
    @("NoteID"       , { param ($NotesDoc) $NotesDoc.NoteID }),
    @("UniversalID"  , { param ($NotesDoc) $NotesDoc.UniversalID }),
    @("Created"      , { param ($NotesDoc) $NotesDoc.Created }),
    @("LastModified" , { param ($NotesDoc) $NotesDoc.LastModified }),
    @("Form"         , { param ($NotesDoc) $NotesDoc.GetFirstItem("Form").Text })
)

$ArrayOfCurntDBFields = @("ObserverCanon","DateObservation","ObserverArea","ObservationArea","Observation","JobObserved","PeopleObservedInt","Comments","Attachments","ReviewerCanon","Status","ReviewerComments","DateSubmitted")

<#
doccreated
docmodified
ModifiedBy
unid
#>

################################################################################################################################################################

try
{
    [LN.NotesSession]  $nSession  = New-Object LN.NotesSession
    [LN.NotesDatabase] $nDatabase = $nSession.GetDatabase($ServerName, $LNFilePath)
    [LN.NotesDocumentCollection] $docCollection = $nDatabase.AllDocuments
    
    [int] $rowCounter = 1
    [int] $colCounter = 1

    ################################################################################

    Write-Host "NotesURL : $($nDatabase.NotesURL)"
    Write-Host "Document collection Count - $($docCollection.Count)"
    
    ################################################################################

    $excelPkg = New-Object OfficeOpenXml.ExcelPackage
    $excelSheet = $excelPkg.Workbook.Worksheets.Add($SheetName)    
    
    foreach($defaultField in $ArrayOfDefaultFields)
    {
        $excelSheet.SetValue($rowCounter, ($colCounter++), $defaultField[0])
    }
    foreach($curntDBField in $ArrayOfCurntDBFields)
    {
        $excelSheet.SetValue($rowCounter, ($colCounter++), $curntDBField)
    }
    $rowCounter++

    ################################################################################
    
    [LN.NotesDocument] $doc = $docCollection.GetFirstDocument()

    while ($doc -ne $null)
    {
        $colCounter = 1
        foreach($defaultField in $ArrayOfDefaultFields)
        {
            $excelSheet.SetValue($rowCounter, ($colCounter++), (& $defaultField[1] $doc))
        }
        foreach($curntDBField in $ArrayOfCurntDBFields)
        {
            $excelSheet.SetValue($rowCounter, ($colCounter++), $doc.GetFirstItem($curntDBField).Text)
        }
        
        Write-Host $doc.NoteID
        $rowCounter++
        $doc = $docCollection.GetNextDocument($doc)
    }    
    
    ################################################################################

    [string] $dtTimeSuffix = (Get-Date -Format "yyyyMMdd_HHmmss")
    [string] $xlFilePath = "$($BackupPath)\$($XlFileNamePrefix)_$($dtTimeSuffix).xlsx"

    if (!(Test-Path -Path $BackupPath)) { New-Item -Path $BackupPath -ItemType "directory" }
    $excelPkg.SaveAs((New-Object System.IO.FileInfo($xlFilePath)))

    ################################################################################
}
catch
{
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
