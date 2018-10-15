
################################################################################################################################################################
# Open Xml debug 3
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\documentformat.openxml.2.8.1\lib\net40\DocumentFormat.OpenXml.dll"
################################################################################################################################################################


Function Create-ExcelFile()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [Parameter(Mandatory = $false)]
        [string] $DefaultSheetName = "Sheet1"
    )

    [Type[]] $emptyTypeArray = @()
    [System.Reflection.MethodInfo] $AddNewPartMethodInfo = [DocumentFormat.OpenXml.Packaging.WorkbookPart].GetMethod("AddNewPart", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Packaging.WorksheetPart])


    [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($FilePath, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    [DocumentFormat.OpenXml.Packaging.WorkbookPart] $wbPart = $XlDoc.AddWorkbookPart()
    $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    
    [DocumentFormat.OpenXml.Packaging.WorksheetPart] $wsPart = $AddNewPartMethodInfo.Invoke($wbPart, @())
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $defaultSheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $defaultWorkSheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet($defaultSheetData)
    $wsPart.Worksheet = $defaultWorkSheet

    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $defaultSheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $defaultSheet.Id = $wbPart.GetIdOfPart($wsPart)
    $defaultSheet.SheetId = [uint32]1
    $defaultSheet.Name = $DefaultSheetName

    [DocumentFormat.OpenXml.Spreadsheet.Sheets] $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $sheets = $wbPart.Workbook.AppendChild($sheets)
    $sheets.Append($defaultSheet)
    
    $XlDoc.WorkbookPart.Workbook.Save()

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
        [bool] $IsEditable,

        [Parameter(Mandatory = $false)]
        [string] $DefaultSheetName = "Sheet1"
    )

    [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($FilePath, $IsEditable)    

    Return $XlDoc
}



Function Get-WorkSheet()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument] $XlDoc,

        [Parameter(Mandatory = $true)]
        [string] $SheetName
    )


    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $defaultSheet = $null
    
    for ([int] $i = 0; $i -lt $XlDoc.WorkbookPart.Workbook.ChildElements.Count; $i++)
    {
        if ($XlDoc.WorkbookPart.Workbook.ChildElements.GetItem($i).GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Sheets])
        {
            [DocumentFormat.OpenXml.OpenXmlCompositeElement] $oxce = $XlDoc.WorkbookPart.Workbook.ChildElements.GetItem($i)
            for ([int] $j = 0; $j -lt $oxce.ChildElements.Count; $j++)
            {
                if ($oxce.ChildElements.GetItem($j).GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Sheet])
                {
                    [DocumentFormat.OpenXml.OpenXmlLeafElement] $oxle = $oxce.ChildElements.GetItem($j)
                    $defaultSheet = $oxle
                    break
                }
            }

        }
        if ($defaultSheet -ne $null) { break; }
    }

    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $defaultWorkSheet = ([DocumentFormat.OpenXml.Packaging.WorksheetPart]$XlDoc.WorkbookPart.GetPartById($defaultSheet.Id)).Worksheet
    
    Return $defaultWorkSheet
}


Function Get-SheetData()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $Worksheet
    )
    
    [Type[]] $emptyTypeArray = @()
    [System.Reflection.MethodInfo] $getFirstChildMethodInfo = [DocumentFormat.OpenXml.Spreadsheet.Worksheet].GetMethod("GetFirstChild", $emptyTypeArray).MakeGenericMethod([DocumentFormat.OpenXml.Spreadsheet.SheetData])
    
    [DocumentFormat.OpenXml.Spreadsheet.SheetData] $defaultSheetData = $getFirstChildMethodInfo.Invoke($Worksheet, @())

    Return $defaultSheetData
}


Function Get-Row()
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


    [DocumentFormat.OpenXml.Spreadsheet.Cell] $cell = $Row.ChildElements.Where( { $_.GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Cell] -and $_.CellReference.Value -eq  "$($ColumnName)$($Row.RowIndex)"} ).First

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


[string] $filePath = "D:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
[string] $sheetName = "SheetOne"
[System.UInt32] $rowIndex = 1


Try
{
    
    $excelDoc = Create-ExcelFile -FilePath $filePath -DefaultSheetName $sheetName
    # $excelDoc = Open-ExcelFile -FilePath $filePath -IsEditable $true
    $wkSheet = Get-WorkSheet -XlDoc $excelDoc -SheetName $sheetName
    $sheetDt = Get-SheetData -Worksheet $wkSheet

    <##>

    [DocumentFormat.OpenXml.Spreadsheet.Sheet] $defaultSheet = $null

    for ([int] $i = 0; $i -lt $excelDoc.WorkbookPart.Workbook.ChildElements.Count; $i++)
    {
        if ($excelDoc.WorkbookPart.Workbook.ChildElements.GetItem($i).GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Sheets])
        {
            [DocumentFormat.OpenXml.OpenXmlCompositeElement] $oxce = $excelDoc.WorkbookPart.Workbook.ChildElements.GetItem($i)
            for ([int] $j = 0; $j -lt $oxce.ChildElements.Count; $j++)
            {
                if ($oxce.ChildElements.GetItem($j).GetType() -eq [DocumentFormat.OpenXml.Spreadsheet.Sheet])
                {
                    [DocumentFormat.OpenXml.OpenXmlLeafElement] $oxle = $oxce.ChildElements.GetItem($j)
                    $defaultSheet = $oxle
                    break
                }
            }

        }
    }

    [DocumentFormat.OpenXml.Spreadsheet.Worksheet] $defaultWorkSheet = ([DocumentFormat.OpenXml.Packaging.WorksheetPart]$excelDoc.WorkbookPart.GetPartById($defaultSheet.Id)).Worksheet

    


    $txt = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $txt.Text = "Test"
    $ins = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $ins.AppendChild($txt)

    $r = Get-Row -SheetData $sheetDt -RowIndex $rowIndex
    $c = Get-Cell -Row $r -ColumnName "A"
    $c.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $c.AppendChild($ins)

    Save-ExcelFile -XlDoc $excelDoc
}
Catch
{
    Write-Host $_.Exception.ToString()
}
Finally
{

}

################################################################################################################################################################
