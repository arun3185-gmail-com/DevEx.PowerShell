



$excelDoc = Create-ExcelFile -FilePath "F:\Arun\Git\DevEx.Data\OpenXmlSheet.xlsx"
$wksPart = Add-Sheet -XlDoc $excelDoc -SheetName "Sheet0"

$txt = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
$txt.Text = "Test"
$ins = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
$ins.AppendChild($txt)


$r = Get-Row -Ws $wksPart.Worksheet -RowIndex 1
$c = Get-Cell -Row $r -ColumnName "A"
$c.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
$c.AppendChild($ins)

Save-ExcelFile -XlDoc $excelDoc