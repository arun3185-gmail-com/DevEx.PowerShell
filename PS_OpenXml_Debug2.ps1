
Import-Module "F:\Arun\Git\DevEx\NuPkg\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"

[string] $xlFilePath = "F:\Arun\Git\DevEx\Data\OpenXmlSheet2.xlsx"

[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $xlDoc = $null

function AddCellWithText()
{
    Param ([string] $text)
    
    [Type[]] $arr1 = @([DocumentFormat.OpenXml.Spreadsheet.Text].GetType())
    [Type[]] $arr2 = @([DocumentFormat.OpenXml.Spreadsheet.InlineString])
    $InlineString_AppendChild_Text_MethodInfo = [DocumentFormat.OpenXml.Spreadsheet.InlineString].GetMethod("AppendChild").MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.Text])
    $Cell_AppendChild_InlineString_MethodInfo = [DocumentFormat.OpenXml.Spreadsheet.Cell].GetMethod("AppendChild").MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.InlineString])
    [DocumentFormat.OpenXml.Spreadsheet.Cell] $c1 = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
    $c1.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString

    [DocumentFormat.OpenXml.Spreadsheet.InlineString] $inlineStr = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    [DocumentFormat.OpenXml.Spreadsheet.Text] $t = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $t.Text = $text
    #$inlineStr.AppendChild($t)
    $InlineString_AppendChild_Text_MethodInfo.Invoke($inlineStr, $t)

    #$c1.AppendChild($inlineStr)
    $Cell_AppendChild_InlineString_MethodInfo.Invoke($c1, $inlineStr)

    $c1
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


    [int] $rowindex = 1
    foreach ($dr in $dt.Rows)
    {
        [DocumentFormat.OpenXml.Spreadsheet.Row] $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = [UInt32]$rowindex

        if ($rowindex -eq 1)
        {
            $cell01 = AddCellWithText "EmployeeID"
            $cell02 = AddCellWithText "EmpName"
            $cell03 = AddCellWithText "Designation"
            $row.AppendChild($cell01)
            $row.AppendChild($cell02)
            $row.AppendChild($cell03)
        }
        else
        {
            $cell01 = AddCellWithText $dr["EmployeeID"].ToString()
            $cell02 = AddCellWithText $dr["EmpName"].ToString()
            $cell03 = AddCellWithText $dr["Designation"].ToString()
            $row.AppendChild($cell01)
            $row.AppendChild($cell02)
            $row.AppendChild($cell03)
        }

        $sheetData.AppendChild($row)
        $rowindex
    }
    
    $wbPart.Workbook.Save()
    $xlDoc.Close()
}

finally
{
    if ($xlDoc -ne $null) { $xlDoc.Dispose() }
    if ($dt -ne $null) { $dt.Dispose() }
}
