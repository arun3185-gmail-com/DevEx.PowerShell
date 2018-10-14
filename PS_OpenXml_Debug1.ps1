
################################################################################################################################################################
# PowerShell Open Xml 3
################################################################################################################################################################
Import-Module "F:\Arun\Git\DevEx.References\NuPkg\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"
################################################################################################################################################################

[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc = $null
[System.Reflection.MethodInfo] $AddNewPartMethodInfo = $null

################################################################################################################################################################


Function Save-ExcelFile()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc
    )

    $XlDoc.WorkbookPart.Workbook.Save()
    if ($XlDoc -ne $null) { $XlDoc.Close(); $XlDoc.Dispose() }
}

<#
Function GetDataTable()
{
    [System.Data.DataTable] $dt = $null
    [System.Data.DataRow] $dr = $null
    [System.Data.DataColumn] $dc = $null

    $dt = New-Object System.Data.DataTable
    $dc = $dt.Columns.Add("EmployeeID")
    $dc = $dt.Columns.Add("EmpName")
    $dc = $dt.Columns.Add("Designation")

    $dr = $dt.NewRow()
    $dr["EmployeeID"] = 1
    $dr["EmpName"] = "Arun"
    $dr["Designation"] = "Developer"
    $dt.Rows.Add($dr)

    $dr = $dt.NewRow()
    $dr["EmployeeID"] = 2
    $dr["EmpName"] = "Sangeetha"
    $dr["Designation"] = "Developer"
    $dt.Rows.Add($dr)

    $dr = $dt.NewRow()
    $dr["EmployeeID"] = 3
    $dr["EmpName"] = "Thejaswini"
    $dr["Designation"] = "Developer"
    $dt.Rows.Add($dr)

    Return $dt
}
#>

################################################################################################################################################################

Try
{
    $FilePath = "F:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"

    $XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($FilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wbPart = $XlDoc.AddWorkbookPart()
    $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook

    [Type[]] $emptyTypeArray = @()
    $AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $AddNewPartMethodInfo.Invoke($wbPart, @())

    $wsPart.Worksheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet(New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData)

    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = [uint32]1
    $sheet.Name = "Sheet1"

    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)
    $sheets.Append($sheet)

    $txt = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $txt.Text = "Test"
    $ins = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $ins.AppendChild($txt)


    $Ws = $wsPart.Worksheet 
    [uint32] $RowIndex = 1
    [Type[]] $emptyTypeArray = @()
    [System.Reflection.MethodInfo] $getFirstChildMethodInfo = [DocumentFormat.OpenXml.Spreadsheet.Worksheet].GetMethod("GetFirstChild", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SheetData])

    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $getFirstChildMethodInfo.Invoke($Ws, @())

    [DocumentFormat.OpenXml.Spreadsheet.Row] $row = $sheetData.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Row] -and $_.RowIndex -eq $RowIndex} ).First

    if ($row -eq $null)
    {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = $RowIndex
        $sheetData.AppendChild($row)
    }

    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $cell = $row.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Cell] -and $_.CellReference.Value -eq  "$($ColumnName)$($Row.RowIndex)"} ).First

    if ($cell -eq $null)
    {
        $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
        $cell.CellReference = "$($ColumnName)$($row.RowIndex)"
        $row.AppendChild($cell)
    }
    $cell.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $cell.AppendChild($ins)

    Save-ExcelFile -XlDoc $XlDoc
}
Catch
{
    Write-Host $_.Exception.ToString()
}
Finally
{

}

################################################################################################################################################################
