
################################################################################################################################################################
# EPP Excel Test 2
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
################################################################################################################################################################

[string] $XlFilePath = "D:\Arun\TestEPP2.xlsx"

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet1 = $null

[OfficeOpenXml.DataValidation.ExcelDataValidationList] $valList = $null

try
{
    [System.IO.FileInfo] $XlFileInfo = New-Object System.IO.FileInfo($XlFilePath)
    $excelPkg = New-Object OfficeOpenXml.ExcelPackage($XlFileInfo)

    <################################################################################
    $excelSheet1 = $excelPkg.Workbook.Worksheets["Sheet1"]
    [OfficeOpenXml.ExcelAddress] $xlAddrStatusCell = New-Object OfficeOpenXml.ExcelAddress(10, 2, 10, 2)
    
    
    $valList = $excelSheet1.DataValidations.AddListValidation("A1:A10")
    $valList.Formula.Values.Add("Opt1")
    $valList.Formula.Values.Add("Opt2")
    $valList.Formula.Values.Add("Opt3")
    $valList.Formula.Values.Add("Opt4")
    $valList.AllowBlank = $true
    $valList.GetType()
    ################################################################################>


    ################################################################################
    $excelSheet1 = $excelPkg.Workbook.Worksheets["Sheet1"]
    
    $xlTable.Columns[0].Name = "ID"
    $xlTable.Columns[1].Name = "Name"
    $xlTable.Columns[2].Name = "Amount"
    $xlTable.ShowFilter = $true

    $excelSheet1.SetValue(2, 1, 1001)
    $excelSheet1.SetValue(2, 2, "Arun")
    $excelSheet1.SetValue(2, 3, 2000)

    $excelSheet1.SetValue(3, 1, 1002)
    $excelSheet1.SetValue(3, 2, "Sangeetha")
    $excelSheet1.SetValue(3, 3, 3000)

    $excelSheet1.SetValue(4, 1, 1003)
    $excelSheet1.SetValue(4, 2, "Thejaswini")
    $excelSheet1.SetValue(4, 3, 10000)

    [OfficeOpenXml.ExcelAddress] $xlAddrTable = New-Object OfficeOpenXml.ExcelAddress(1, 3, 4, 3)
    [OfficeOpenXml.Table.ExcelTable] $xlTable = $excelSheet1.Tables.Add($xlAddrTable, "TestTable");
    
    #$xlRange.AutoFitColumns()

    ################################################################################
}
catch
{
    #Write-Host    $_.Exception.ToString() -ForegroundColor Red
    throw
}
finally
{
    $excelPkg.Save()
    if ($excelSheet -ne $null) { $excelSheet.Dispose(); $excelSheet = $null }
    if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
