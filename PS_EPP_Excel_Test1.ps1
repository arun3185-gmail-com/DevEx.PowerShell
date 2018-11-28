
################################################################################################################################################################
# Lotus Notes All Documents List
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
################################################################################################################################################################

[string] $XlFilePath = "J:\Arun\EvonikPoCBackups\EWA_Tracker_Sheet_181122.xlsx"

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet = $null

try
{
    [System.IO.FileInfo] $XlFileInfo = New-Object System.IO.FileInfo($XlFilePath)

    $excelPkg = New-Object OfficeOpenXml.ExcelPackage($XlFileInfo)
    $excelSheet = $excelPkg.Workbook.Worksheets[1]
    
    [int] $colCount = $excelSheet.Dimension.End.Column
    [int] $rowCount = $excelSheet.Dimension.End.Row

    for ($i = 2; $i -le $rowCount; $i++)
    {
        "$($excelSheet.Cells[$i, 1].Value);$($excelSheet.Cells[$i, 1].Text)"
    }

    ################################################################################
}
catch
{
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
finally
{
    if ($excelSheet -ne $null) { $excelSheet.Dispose(); $excelSheet = $null }
    if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
