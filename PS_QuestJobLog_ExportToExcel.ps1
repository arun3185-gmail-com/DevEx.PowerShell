
################################################################################################################################################################
# Quest Job Log Export to Excel
################################################################################################################################################################

Import-Module "J:\Arun\Git\DevEx.References\NuGet\newtonsoft.json.11.0.2\lib\net40\Newtonsoft.Json.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"

################################################################################################################################################################

[string] $QuestJobLogPath = "J:\QuestJobLogs"
[string] $QuestJobLogFileName = "10426_KMZ_Abteilungsinfo_NMSP_10_31_2018_12_12_34"

[string] $Global:Tab = [char]9
[string] $Global:LogTimeFormat  = "[yyyy-MM-dd HH:mm:ss.fff]"
[string] $Global:ThisScriptRoot = @("J:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[string] $Global:ThisScriptName = $null

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}
else
{
    $Global:ThisScriptName = "PS_QuestJobLog_ExportToExcel"
}
[string] $Global:LogFilePath  = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"


if (-not($QuestJobLogFileName.EndsWith(".log")))
{
    $QuestJobLogFileName += ".log"
}
[string] $QstLogFilePath = "$($QuestJobLogPath)\$($QuestJobLogFileName)"   #"J:\QuestJobLogs\10426_PKMR_Abrechnung_NMSP_10_29_2018_18_46_45.log"
[string] $XlFileName = $QuestJobLogFileName.Replace(".log", ".xlsx")
[string] $SheetName = "Sheet1"

################################################################################################################################################################
# Functions
################################################################################################################################################################

function Write-LogInfo()
{
    Param ([string] $Message)
    
    "$(Get-Date -Format $Global:LogTimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}


################################################################################################################################################################
# Main Program
################################################################################################################################################################

[System.Xml.XmlDocument] $QstJobXmlLogDoc = $null
[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet = $null
[System.Collections.ArrayList] $columnNames = @()

try
{
    [string] $xlFilePath = "$($QuestJobLogPath)\$($XlFileName)"
    [System.IO.FileInfo] $xlFileInfo = New-Object System.IO.FileInfo($xlFilePath)


    $QstJobXmlLogDoc = New-Object System.Xml.XmlDocument
    $QstJobXmlLogDoc.Load($QstLogFilePath)
    [string] $strJsonQstJob = [Newtonsoft.Json.JsonConvert]::SerializeXmlNode($QstJobXmlLogDoc)
    $QstJobHash = ConvertFrom-Json $strJsonQstJob

    $newIdx = $columnNames.Add("LogType")
    $newIdx = $columnNames.Add("@date")
    $newIdx = $columnNames.Add("@severity")
    $newIdx = $columnNames.Add("@stage")
    $newIdx = $columnNames.Add("context")
    
    $excelPkg = New-Object OfficeOpenXml.ExcelPackage($xlFileInfo)        
    if ($excelPkg.Workbook.Worksheets[$SheetName] -ne $null)
    {
        $excelSheet = $excelPkg.Workbook.Worksheets.Delete($SheetName)
    }
    $excelSheet = $excelPkg.Workbook.Worksheets.Add($SheetName)

    [int] $rowCounter = 1
    for ($i = 0; $i -lt $columnNames.Count; $i++)
    {
        $excelSheet.SetValue(1, $i + 1, $columnNames[$i])
    }
    
    foreach ($hshEntry in $QstJobHash.log.entry)
    {
        $rowCounter++
        $excelSheet.SetValue($rowCounter, $columnNames.IndexOf("LogType") + 1, "Entry")

        $props = Get-Member -InputObject $hshEntry -MemberType NoteProperty
        foreach ($prop in $props)
        {
            $propIdx = $columnNames.IndexOf($prop.Name)
            if ($propIdx -eq -1)
            {
                $propIdx = $columnNames.Add($prop.Name)
                $excelSheet.SetValue(1, $propIdx + 1, $prop.Name)
            }
            $propValue = $hshEntry | Select-Object -ExpandProperty $prop.Name
            $excelSheet.SetValue($rowCounter, $propIdx + 1, $propValue)
        }
    }
    
    $props = Get-Member -InputObject $QstJobHash.log -MemberType NoteProperty    
    if ($props.Where({$_.Name -eq "summary"}).Count -ge 1)
    {
        foreach ($hshEntry in $QstJobHash.log.summary)
        {
            $rowCounter++
            $excelSheet.SetValue($rowCounter, $columnNames.IndexOf("LogType") + 1, "Summary")
            
            $props = Get-Member -InputObject $hshEntry -MemberType NoteProperty
            foreach ($prop in $props)
            {
                $propIdx = $columnNames.IndexOf($prop.Name)
                if ($propIdx -eq -1)
                {
                    $propIdx = $columnNames.Add($prop.Name)
                    $excelSheet.SetValue(1, $propIdx + 1, $prop.Name)
                }
                $propValue = $hshEntry | Select-Object -ExpandProperty $prop.Name
                $excelSheet.SetValue($rowCounter, $propIdx + 1, $propValue)
            }
        }
    }

    
    $excelPkg.Save()
    
}
catch
{
    Write-LogInfo $_.Exception.ToString()
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
finally
{
    if ($excelSheet -ne $null) { $excelSheet.Dispose(); $excelSheet = $null }
    if ($excelPkg -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

################################################################################################################################################################

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
