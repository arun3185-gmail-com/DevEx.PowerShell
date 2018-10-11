
Import-Module "D:\Arun\DevEx\NuPkg\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"

[string] $xlFilePath = "D:\Arun\DevEx\Data\OpenXmlSheet1.xlsx"

[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $xlDoc = $null


function InsertSharedStringItem()
{
    Param ([String] $text, [DocumentFormat.OpenXml.Packaging.SharedStringTablePart] $shareStringPart)

    if ($shareStringPart.SharedStringTable -eq $null)
    {
        $shareStringPart.SharedStringTable = New-Object DocumentFormat.OpenXml.Spreadsheet.SharedStringTable
    }

    [int] $i = 0

    [Type[]] $arr = @()
    #$elementsMethod = [DocumentFormat.OpenXml.Spreadsheet.SharedStringTable].GetMethod("Elements", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SharedStringItem])
    $elementsMethod = [DocumentFormat.OpenXml.Spreadsheet.SharedStringTable].GetMethods().Where( { $_.Name -eq "Elements" } )[0].MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SharedStringItem])

    $listSharedStringItem = $elementsMethod.Invoke($shareStringPart.SharedStringTable, @())
    foreach ($item in $listSharedStringItem)
    {
        if ($item.InnerText = $text)
        {
            return $i
        }
        $i++
    }
    
    
    $shareStringPart.SharedStringTable.AppendChild((New-Object DocumentFormat.OpenXml.Spreadsheet.SharedStringItem((New-Object DocumentFormat.OpenXml.Spreadsheet.Text($text)))))
    $shareStringPart.SharedStringTable.Save()

    return $i
}

function InsertCellInWorksheet()
{
    Param ([string] $columnName, [System.UInt32] $rowIndex, [DocumentFormat.OpenXml.Packaging.WorksheetPart] $worksheetPart)

    [Type[]] $arr = @()
    $worksheetGetFirstChildMethod = [DocumentFormat.OpenXml.Spreadsheet.Worksheet].GetMethod("GetFirstChild", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SheetData])
    #$sheetDataElementsMethod = [DocumentFormat.OpenXml.Spreadsheet.SheetData].GetMethod("Elements", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Row])
    $sheetDataElementsMethod = [DocumentFormat.OpenXml.Spreadsheet.SheetData].GetMethods().Where( { $_.Name -eq "Elements" } )[0].MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Row])
    #$rowElementsMethod = [DocumentFormat.OpenXml.Spreadsheet.Row].GetMethod("Elements", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Cell])
    $rowElementsMethod = [DocumentFormat.OpenXml.Spreadsheet.Row].GetMethods().Where( { $_.Name -eq "Elements" } )[0].MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Cell])
    
    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $worksheet = $worksheetPart.Worksheet
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $worksheetGetFirstChildMethod.Invoke($worksheet, @())
    [string] $cellReference = $columnName + $rowIndex

    [DocumentFormat.OpenXml.Spreadsheet.Row] $row = $null
    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $null
    
    $allRows = $sheetDataElementsMethod.Invoke($sheetData, @())
    $qRow = $allRows | Where-Object RowIndex -EQ $rowIndex
    
    if ($qRow -ne $null -and $qRow.Count() -gt 0)
    {
        $row = $qRow.First()
    }
    else
    {
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = $rowIndex
        $sheetData.Append($row)
    }

    $allCells = $rowElementsMethod.Invoke($row, @())
    $qCell = $allCells | Where-Object CellReference.Value -EQ $cellReference

    if ($qCell -ne $null -and $qCell.Count() -gt 0)
    {
        $cell = $qCell.First()
    }
    else
    {
        [DocumentFormat.OpenXml.Spreadsheet.Cell] $refCell = $null

        $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
        $cell.CellReference = $cellReference
        $row.InsertBefore($cell, $refCell)

        $worksheet.Save()
    }

    return $cell
}

try
{
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $null
    [DocumentFormat.OpenXml.Spreadsheet.Workbook] $wb = $null
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null
    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $ws = $null
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheets] $sheets = $null
    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
    [Type[]] $arr = @()
    [Type[]] $arrTypes = @([DocumentFormat.OpenXml.Spreadsheet.CellValues])
    [System.Reflection.MethodInfo] $workbookPartAddNewPartMethodInfo_WorksheetPart = $null
    [System.Reflection.MethodInfo] $workbookPartAddNewPartMethodInfo_SharedStringTablePart = $null
    [System.Reflection.MethodInfo] $workbookPartGetPartsOfTypeMethodInfo = $null
    [System.Reflection.ConstructorInfo] $enumValueConstructorInfo = $null

    [System.Data.DataTable] $dt = $null
    [System.Data.DataRow] $dr = $null
    [System.Data.DataColumn] $dc = $null
    

    $xlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($xlFilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wb = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    
    $wbPart = $xlDoc.AddWorkbookPart();
    $wbPart.Workbook = $wb

    # WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
    
    $workbookPartAddNewPartMethodInfo_WorksheetPart = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $workbookPartAddNewPartMethodInfo_WorksheetPart.Invoke($wbPart, @()) 

    $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)    
    $wsPart.Worksheet = $ws

    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)

    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = [System.UInt32]1
    $sheet.Name = "Sheet1"

    $sheets.Append($sheet)
    
    [string] $cl = ""
    [System.UInt32] $row = 2
    [int] $index
    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell
    
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

    foreach ($dr in $dt.Rows)
    {
        for ($idx = 0; $idx -lt $dt.Columns.Count; $idx++)
        {
            if ($idx -ge 26)
            {
                $cl = [string]::Concat("A", [System.Convert]::ToChar(65 + $idx - 26))                
            }
            else
            {
                $cl = [System.Convert]::ToChar(65 + $idx)
            }

            [DocumentFormat.OpenXml.Packaging.SharedStringTablePart] $shareStringPart = $null
            $workbookPartGetPartsOfTypeMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("GetPartsOfType", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.SharedStringTablePart])
            $allSharedStringTableParts = $workbookPartGetPartsOfTypeMethodInfo.Invoke($wbPart, @())

            if ($allSharedStringTableParts.Count -gt 0)
            {
                $shareStringPart = $allSharedStringTableParts.First()
            }
            else
            {
                $workbookPartAddNewPartMethodInfo_SharedStringTablePart = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $arr).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.SharedStringTablePart])
                $shareStringPart = $workbookPartAddNewPartMethodInfo_SharedStringTablePart.Invoke($wbPart, @())
            }
            
            $enumValueConstructorInfo = [DocumentFormat.OpenXml.EnumValue[DocumentFormat.OpenXml.Spreadsheet.CellValues]].GetConstructor($arrTypes)
            
            if ($row -eq 2)
            {
                $index = InsertSharedStringItem $dt.Columns[$idx].ColumnName $shareStringPart
                $cell = InsertCellInWorksheet $cl ($row - 1) $wsPart
                $cell.CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue $index.ToString()
                $cell.DataType = $enumValueConstructorInfo.Invoke([DocumentFormat.OpenXml.Spreadsheet.CellValues]::SharedString)
            }
            
            $index = InsertSharedStringItem ([System.Convert]::ToString($dr[$idx])) $shareStringPart
            $cell = InsertCellInWorksheet $cl $row $wsPart
            $cell.CellValue = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue $index.ToString()
            $cell.DataType = $enumValueConstructorInfo.Invoke([DocumentFormat.OpenXml.Spreadsheet.CellValues]::SharedString)
        }
        $row++
    }
    
    $wbPart.Workbook.Save()
    $xlDoc.Close()
}

finally
{
    if ($xlDoc -ne $null) { $xlDoc.Dispose() }
    if ($dt -ne $null) { $dt.Dispose() }
}
