
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"
################################################################################################################################################################

[string] $Script:XlFilePath = "F:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $Script:XlDoc = $null
[System.Reflection.MethodInfo] $Script:AddNewPartMethodInfo = $null


################################################################################################################################################################

Function Create-ExcelFile()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $FilePath
    )


    $Script:XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($FilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wbPart = $Script:XlDoc.AddWorkbookPart()
    $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    
    
    Return $Script:XlDoc
}


Function Open-ExcelFile()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [Parameter(Mandatory = $true)]
        [bool] $IsEditable
    )


    $Script:XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($FilePath, $IsEditable)
    if ($Script:XlDoc.WorkbookPart -eq $null)
    {
        $wbPart = $Script:XlDoc.AddWorkbookPart()
        $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    }

    
    Return $Script:XlDoc
}


Function Add-Sheet()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )


    $wbPart = $XlDoc.WorkbookPart

    [Type[]] $emptyTypeArray = @()
    $Script:AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $Script:AddNewPartMethodInfo.Invoke($wbPart, @())    
    
    $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $wsPart.Worksheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)

    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = [uint32]1
    $sheet.Name = $SheetName

    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)
    $sheets.Append($sheet)

    Return $wsPart
}


Function Get-Sheet()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )
    
    
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $XlDoc.WorkbookPart
    [DocumentFormat.OpenXml.Spreadsheet.Workbook]$wb = $wbPart.Workbook

    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null

    [System.Reflection.MethodInfo] $getFirstChildMethodInfo = [DocumentFormat.OpenXml.Spreadsheet.Workbook].GetMethod("GetFirstChild", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Sheets])
    [Type[]] $emptyTypeArray = @()


    $sheets = $getFirstChildMethodInfo.Invoke($wb, @())
    $qSheets = $sheets.ChildElements.Where({ $_.Name -eq $SheetName })
        
    if ($qSheets.Count() -ge 1)
    {
        $sheet = $qSheets.First()
    }
    else
    {
        Return $null
    }

    $wsPart = $wbPart.GetPartById($sheet.Id.Value)

    Return $wsPart
}


Function Save-ExcelFile()
{
    $wbPart.Workbook.Save()
    if ($Script:XlDoc -ne $null) { $Script:XlDoc.Close(); $Script:XlDoc.Dispose() }
}

################################################################################################################################################################

