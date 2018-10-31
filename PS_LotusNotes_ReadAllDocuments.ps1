
################################################################################################################################################################

Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"

################################################################################################################################################################

[string] $DataFolderPath = "J:\Arun\Git\DevEx.Data"

[string] $ServerName = "EMEAARFN01/Server/Evonik"
[string] $LNFilePath = "Abteilungen/PKM-R/archiv-pkmrabre.nsf"

$ArrayOfDefaultFields = 
@(
    @("NoteID"       , { param ($NotesDoc) $NotesDoc.NoteID }),
    @("UniversalID"  , { param ($NotesDoc) $NotesDoc.UniversalID }),
    @("Created"      , { param ($NotesDoc) $NotesDoc.Created }),
    @("LastModified" , { param ($NotesDoc) $NotesDoc.LastModified }),
    @("Form"         , { param ($NotesDoc) $NotesDoc.GetFirstItem("Form").Text })
)

$ArrayOfCurntDBFields = @("beschreibung","thema","gliederung_1","datum","author","leser","info_mail","info","ablageort","ablageort2","aktualisiert_am")

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

    [string] $XlFileNamePrefix = $nDatabase.Title
    Write-Host "NotesURL : $($nDatabase.NotesURL)"
    Write-Host "Document collection Count - $($docCollection.Count)"
    
    ################################################################################

    $excelPkg = New-Object OfficeOpenXml.ExcelPackage
    $excelSheet = $excelPkg.Workbook.Worksheets.Add("Sheet1")
    
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
    [string] $xlFilePath = "$($DataFolderPath)\$($XlFileNamePrefix)_$($dtTimeSuffix).xlsx"

    if (!(Test-Path -Path $DataFolderPath)) { New-Item -Path $DataFolderPath -ItemType "directory" }
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
