
################################################################################################################################################################
# PowerShell Open Xml 3
################################################################################################################################################################
Import-Module "F:\Arun\Git\DevEx.References\NuPkg\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"
################################################################################################################################################################

[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc = $null
[System.Reflection.MethodInfo] $AddNewPartMethodInfo = $null

################################################################################################################################################################



Function Create-ExcelFile()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $FilePath
    )


    $XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($FilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wbPart = $XlDoc.AddWorkbookPart()
    $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    
    
    Return $XlDoc
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


    $XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($FilePath, $IsEditable)
    if ($XlDoc.WorkbookPart -eq $null)
    {
        $wbPart = $XlDoc.AddWorkbookPart()
        $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    }

    
    Return $XlDoc
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
    $AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $AddNewPartMethodInfo.Invoke($wbPart, @())    
    
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


    $sheets = $wb.FirstChild
    $qSheets = $sheets.ChildElements.Where({ $_.Name -eq $SheetName })
        
    if ($qSheets.Count -ge 1)
    {
        $sheet = $qSheets.FirstChild
    }
    else
    {
        Return $null
    }

    $wsPart = $wbPart.GetPartById($sheet.Id.Value)

    Return $wsPart
}

Function Get-Row()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $Ws,

        [Parameter(Mandatory = $true)]
        [System.UInt32] $RowIndex
    )


    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $null
    [DocumentFormat.OpenXml.Spreadsheet.Row] $row = $null

    [System.Reflection.MethodInfo] $getFirstChildMethodInfo = [DocumentFormat.OpenXml.Spreadsheet.Worksheet].GetMethod("GetFirstChild", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SheetData])
    [Type[]] $emptyTypeArray = @()

    $sheetData = $getFirstChildMethodInfo.Invoke($Ws, @())

    $row = $sheetData.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Row] -and $_.RowIndex -eq $RowIndex} ).First

    if ($row -eq $null)
    {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = $RowIndex
        $sheetData.AppendChild($row)
    }
    
    Return $row
}

Function Get-Cell()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.Row] $Row,

        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )


    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $null
    
    $cell = $Row.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Cell] -and $_.CellReference.Value -eq  "$($ColumnName)$($Row.RowIndex)"} ).First

    if ($cell -eq $null)
    {
        $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
        $cell.CellReference = "$($ColumnName)$($Row.RowIndex)"
        $Row.AppendChild($cell)
    }

    Return $cell
}


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
    $XlDoc = Create-ExcelFile -FilePath "F:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
    $wksPart = Add-Sheet -XlDoc $XlDoc -SheetName "Sheet0"

    $txt = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $txt.Text = "Test"
    $ins = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $ins.AppendChild($txt)


    $r = Get-Row -Ws $wksPart.Worksheet -RowIndex 1
    $c = Get-Cell -Row $r -ColumnName "A"
    $c.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $c.AppendChild($ins)

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
