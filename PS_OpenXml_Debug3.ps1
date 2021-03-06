
################################################################################################################################################################
# Open Xml debug 3
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"
################################################################################################################################################################


Function Init-Excel()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ref] $ExcelDocRef
    )
    
    
    [Type[]] $emptyTypeArray = @()
    [System.Reflection.MethodInfo] $AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $null

    if ($ExcelDocRef.Value.WorkbookPart -eq $null)
    {
        $wbPart = $ExcelDocRef.Value.AddWorkbookPart()
        $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    }
    
    # $ExcelDocRef.Value.WorkbookPart.Workbook.Save()
}


Function New-ExcelSheet()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ref] $ExcelDocRef,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )
    
    
    [Type[]] $emptyTypeArray = @()
    [System.Reflection.MethodInfo] $AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])

    [uint32] $maxSheetId = 1
    if ($ExcelDocRef.Value.WorkbookPart.Workbook.Sheets.ChildElements.Count -gt 0)
    {
        $maxSheetId = ($ExcelDocRef.Value.WorkbookPart.Workbook.Sheets | Sort-Object -Property SheetId | Select-Object -Last 1).SheetId.Value
        $maxSheetId++
    }
    
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $ExcelDocRef.Value.WorkbookPart
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $AddNewPartMethodInfo.Invoke($wbPart, @())

    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $wsPart.Worksheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)
    
    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = $maxSheetId
    $sheet.Name = $SheetName

    [DocumentFormat.OpenXml.Spreadsheet.Sheets] $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)
    $sheets.Append($sheet)

    # Return $sheetData
}


Function Delete-ExcelSheet()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ref] $ExcelDocRef,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )
    
    
    [DocumentFormat.OpenXml.Spreadsheet.Sheet[]] $qSheets = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $ExcelDocRef.Value.WorkbookPart

    $qSheets = $ExcelDocRef.Value.WorkbookPart.Workbook.Sheets.Where({ $_.Name.HasValue -and $_.Name.Value -eq $SheetName })
    if ($qSheets.Count -ge 1)
    {
        $sheet = $qSheets[0]
    }

    $wsPart = ([DocumentFormat.OpenXml.Packaging.WorksheetPart]$ExcelDocRef.Value.WorkbookPart.GetPartById($sheet.Id))
    $sheet.Remove()
    $wbPart.DeletePart($wsPart)
}


Function Get-ExcelWorkSheetPart()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ref] $ExcelDocRef,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )
    
    [DocumentFormat.OpenXml.Spreadsheet.Sheet[]] $qSheets = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null

    $qSheets = $ExcelDocRef.Value.WorkbookPart.Workbook.Sheets.Where({ $_.Name.HasValue -and $_.Name.Value -eq $SheetName })
    
    if ($qSheets.Count -ge 1)
    {
        $sheet = $qSheets[0]
        $wsPart = ([DocumentFormat.OpenXml.Packaging.WorksheetPart]$ExcelDocRef.Value.WorkbookPart.GetPartById($sheet.Id))
    }

    Return $wsPart
}


Function Get-ExcelRow()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.SheetData] $SheetData,

        [Parameter(Mandatory = $true)]
        [System.UInt32] $RowIndex
    )
    
    
    [DocumentFormat.OpenXml.Spreadsheet.Row] $row = $SheetData.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Row] -and $_.RowIndex -eq $RowIndex} ).First
    if ($row -eq $null)
    {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = $RowIndex
        $SheetData.AppendChild($row)
    }
    
    Return $row
}


Function Get-ExcelCell()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.Row] $Row,

        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )


    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $Row.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Cell] -and $_.CellReference.Value -eq  "$($ColumnName)$($Row.RowIndex)"} ).First

    if ($cell -eq $null)
    {
        $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
        $cell.CellReference = "$($ColumnName)$($Row.RowIndex)"
        $Row.AppendChild($cell)
    }

    Return $cell
}


Function Get-ExcelCell()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.SheetData] $SheetData,

        [Parameter(Mandatory = $true)]
        [System.UInt32] $RowIndex,

        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )


    [DocumentFormat.OpenXml.Spreadsheet.Row] $row = $SheetData.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Row] -and $_.RowIndex -eq $RowIndex} ).First
    if ($row -eq $null)
    {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = $RowIndex
        $SheetData.AppendChild($row)
    }
    
    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $row.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Cell] -and $_.CellReference.Value -eq  "$($ColumnName)$($row.RowIndex)"} ).First
    if ($cell -eq $null)
    {
        $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
        $cell.CellReference = "$($ColumnName)$($row.RowIndex)"
        $row.AppendChild($cell)
    }

    Return $cell
}



Function Set-ExcelCell()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [ref] $ExcelCellRef,

        [Parameter(Mandatory = $true)]
        [ref] $RowIndex
    )
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


[string] $filePath = "D:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
[string] $filePath1 = "D:\Arun\Git\DevEx.Data\TestExcel.xlsx"
[string] $sheetName = "Sheet2"
[System.UInt32] $rowIndex = 1
[bool] $isEditable = $true
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $excelDoc = $null

Try
{
    $excelDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($filePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $excelDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($filePath1, $isEditable)
    
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wkShtPart = $null
    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $wkSht = $null
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $shtDt = $null

    $wkShtPart = Get-ExcelWorkSheetPart -ExcelDocRef ([ref]$excelDoc) -SheetName "Sheet3"
    $wkSht = $wkShtPart.Worksheet
    $shtDt = $wkSht.Where({ $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.SheetData] })[0]


}
Catch
{
    Write-Host $_.Exception.ToString()
}
Finally
{
    $excelDoc.WorkbookPart.Workbook.Save()
    if ($excelDoc -ne $null) { $excelDoc.Close(); $excelDoc.Dispose() }
}

################################################################################################################################################################
