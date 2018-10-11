
Import-Module "F:\Arun\DevEx\NuPkg\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"

[string] $xlFilePath = "F:\Arun\DevEx\Data\OpenXmlSheet.xlsx"
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $xlDoc = $null

try
{
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $null
    [DocumentFormat.OpenXml.Spreadsheet.Workbook] $wb = $null
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null
    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $ws = $null
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheets] $sheets = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
    [System.Reflection.MethodInfo] $addNewPartMethodInfo = $null
    
    $xlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($xlFilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wb = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook

    $wbPart = $xlDoc.AddWorkbookPart()
    $wbPart.Workbook = $wb

    # WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
    [Type[]] $emptyTypeArray = @()
    $addNewPartMethod = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $addNewPartMethod.Invoke($wbPart, @()) 

    $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)
    $wsPart.Worksheet = $ws

    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)

    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = [uint32]1
    $sheet.Name = "Sheet1"

    $sheets.Append($sheet)

    $wbPart.Workbook.Save()
}
catch
{
    Write-Host $_.Exception.ToString()
}
finally
{

    if ($xlDoc -ne $null) { $xlDoc.Dispose() }

}

<#

# [Content_Types].xml

<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
</Types>

    # _rels\.rels

    <?xml version="1.0" encoding="utf-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="R5b18cca2eefc4a88" />
    </Relationships>


    # xl\workbook.xml

    <?xml version="1.0" encoding="utf-8"?>
    <x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:sheets>
            <x:sheet name="Sheet1" sheetId="1" r:id="Re6ce4d2333fa47b6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
        </x:sheets>
    </x:workbook>


        # xl\_rels\workbook.xml.rels

        <?xml version="1.0" encoding="utf-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml" Id="Re6ce4d2333fa47b6" />
        </Relationships>


        # xl\worksheets\sheet1.xml

        <?xml version="1.0" encoding="utf-8"?>
        <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <x:sheetData />
        </x:worksheet>



#>
