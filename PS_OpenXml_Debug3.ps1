################################################################################################################################################################
# PowerShell Open Xml 3
################################################################################################################################################################

Import-Module "D:\Arun\Git\DevEx.References\NuGet\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"

################################################################################################################################################################

[string] $xlFilePath = "F:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $xlDoc = $null
[DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $xlDoc = $null
[DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $null
[DocumentFormat.OpenXml.Spreadsheet.Workbook] $wb = $null
[DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $null
[DocumentFormat.OpenXml.Spreadsheet.Worksheet] $ws = $null
[DocumentFormat.OpenXml.Spreadsheet.SheetData] $sheetData = $null
[DocumentFormat.OpenXml.Spreadsheet.Sheets] $sheets = $null
[DocumentFormat.OpenXml.Spreadsheet.Sheet] $sheet = $null
[System.Reflection.MethodInfo] $addNewPartMethodInfo = $null

################################################################################################################################################################

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


Function Create-Excel()
{
    Param ([string] $xlFilePath, [string] $SheetName = "Sheet1")

    $xlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($xlFilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wb = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    $wbPart = $xlDoc.AddWorkbookPart()
    $wbPart.Workbook = $wb

    [Type[]] $emptyTypeArray = @()
    $addNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $wsPart = $addNewPartMethodInfo.Invoke($wbPart, @())

    $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)
    $wsPart.Worksheet = $ws

    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $wbPart.GetIdOfPart($wsPart)
    $sheet.SheetId = [uint32]1
    $sheet.Name = $SheetName

    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)
    $sheets.Append($sheet)
}


Function Open-Excel()
{
    Param ([string] $xlFilePath, [bool] $IsEditable)

    $xlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($xlFilePath, $IsEditable)
    if ($xlDoc.WorkbookPart -eq $null)
    {
        $wbPart = $xlDoc.AddWorkbookPart()
        $wb = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
        $wbPart.Workbook = $wb
    }
    else
    {
        $xlDoc.WorkbookPart = $wbPart
    }


    if ($wbPart.WorksheetParts.Count() -lt 1)
    {
        # WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
        [Type[]] $emptyTypeArray = @()
        $addNewPartMethod = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])
        $wsPart = $addNewPartMethod.Invoke($wbPart, @()) 
    }
    else
    {
        $wsPart = $wbPart.WorksheetParts.First()
    }


    if ($wsPart.Worksheet -eq $null)
    {
        $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
        $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)
        $wsPart.Worksheet = $ws
    }
    else
    {
        $ws = $wsPart.Worksheet
        if ($ws.ChildElements.Count -gt 0 -and $ws.ChildElements.Where({$_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.SheetData]}).Count() -gt 0)
        {
            $sheetData = $ws.ChildElements.Where({$_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.SheetData]})[0]
        }
        else
        {
            $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
            $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($sheetData)
            $wsPart.Worksheet = $ws
        }
    }


    if ($wbPart.Workbook.Sheets -eq $null)
    {
        $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
        $sheets = $wbPart.Workbook.AppendChild($sheets)
    }
    else
    {
        $sheets = $wbPart.Workbook.Sheets
    }


    if ($sheets.ChildElements.Count -lt 1)
    {
        $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
        $sheet.Id = $wbPart.GetIdOfPart($wsPart)
        $sheet.SheetId = [uint32]1
        $sheet.Name = "Sheet1"
        $sheets.Append($sheet)
    }
    else
    {
        $sheet = $sheets.ChildElements.GetItem(0)
    }


}


Function Save-Excel()
{
    $wbPart.Workbook.Save()
    if ($xlDoc -ne $null) { $xlDoc.Close(); $xlDoc.Dispose() }
}


################################################################################################################################################################

Try
{
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
Catch
{
    Write-Host $_.Exception.ToString()
}
Finally
{

    if ($xlDoc -ne $null) { $xlDoc.Dispose() }

}

################################################################################################################################################################
