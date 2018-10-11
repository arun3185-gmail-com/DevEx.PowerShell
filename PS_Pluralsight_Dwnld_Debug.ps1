
################################################################################################################################################################
# Pluralsight Downloader   
################################################################################################################################################################

Add-Type -AssemblyName System.Web

Import-Module "F:\Arun\DevEx\NuPkg\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Import-Module "F:\Arun\DevEx\NuPkg\htmlagilitypack.1.8.5\lib\Net40\HtmlAgilityPack.dll"

################################################################################################################################################################

[String] $Global:CourseMetaDataXL = "F:\Arun\DevEx\Data\PluralsightCoursesMetadata.xlsx"
[String] $Global:UserName = "ben.andrews@I180618.onmicrosoft.com"
[String] $Global:Password = "P@180618"

[String] $Global:HostUri       = "https://app.pluralsight.com"
[String] $Global:LoginPageUrl  = "https://app.pluralsight.com/id?redirectTo=/library/"
[String] $Global:LogoutPageUrl = "https://app.pluralsight.com/id/signout"

[Boolean] $Global:Flag_DisableVideoDownload    = $false

[String] $Global:Tab        = [char]9
[String] $Global:TimeFormat = "[yyyy-MM-dd HH:mm:ss.fff]"

[String] $Global:ThisScriptRoot       = @("F:\Arun\DevEx\PS", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[String] $Global:ThisScriptName       = "PS_Pluralsight_Dwnld"
[String] $Global:ResourceListFileName = "ResourceCheckList.csv"

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}

[String] $Global:LogFilePath      = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"
[String] $Global:DownloadLocation = "$($Global:ThisScriptRoot)\Pluralsight"

[System.__ComObject] $Global:INetExp    = $null
[System.__ComObject] $Global:XmlHttpObj = $null
[String] $Global:CookieString           = $null

################################################################################################################################################################
# Functions
################################################################################################################################################################

Function Write-LogInfo()
{
    Param ([String] $Message)
    
    "$(Get-Date -Format $Global:TimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

Function Get-ResourceList()
{
    Param ([string] $CourseUrl)

    $Local:courseUrlName = $CourseUrl.Substring($CourseUrl.LastIndexOf('/') + 1)

    $Local:InfoPSObj = New-Object PSObject
    $Local:InfoPSObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value "CourseInfo"
    $Local:InfoPSObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value "CourseInfo"
    $Local:InfoPSObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value "$($CourseUrl)"
    $Local:InfoPSObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value "CourseInfo ($($Local:courseUrlName)).txt"
    $Local:InfoPSObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value $CourseUrl
    $Local:InfoPSObj | Add-Member NoteProperty -Name "StatusCode"       -Value 0

    $Local:MetadataPSObj = New-Object PSObject
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value "CourseMetadataJson"
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value "CourseMetadataJson"
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value "$($CourseUrl)"
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value "CourseMetadataJson.json"
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value "https://app.pluralsight.com/learner/content/courses/$($Local:courseUrlName)"
    $Local:MetadataPSObj | Add-Member NoteProperty -Name "StatusCode"       -Value 0

    $Local:TranscriptPSObj = New-Object PSObject
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value "CourseTranscriptJson"
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value "CourseTranscriptJson"
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value "$($CourseUrl)"
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value "CourseTranscriptJson.json"
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value "https://app.pluralsight.com/learner/courses/$($Local:courseUrlName)/transcript"
    $Local:TranscriptPSObj | Add-Member NoteProperty -Name "StatusCode"       -Value 0

    $Local:ExercisePSObj = New-Object PSObject
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value "ExerciseFile"
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value "ExerciseFile"
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value "$($CourseUrl)/exercise-files"
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value ""
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value ""
    $Local:ExercisePSObj | Add-Member NoteProperty -Name "StatusCode"       -Value 0

    Return @($Local:InfoPSObj, $Local:MetadataPSObj, $Local:TranscriptPSObj, $Local:ExercisePSObj)
}


################################################################################################################################################################
# Main Program
################################################################################################################################################################

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet = $null
[System.Net.WebClient] $wbClient = $null
[HtmlAgilityPack.HtmlDocument] $htmlDoc = $null

Try
{
    $excelPkg = New-Object OfficeOpenXml.ExcelPackage((New-Object System.IO.FileInfo($Global:CourseMetaDataXL)))
    $excelSheet = $excelPkg.Workbook.Worksheets[1]


}
Catch
{
    Write-LogInfo $_.Exception.ToString()
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
Finally
{
    if ($excelSheet        -ne $null) { $excelSheet.Dispose(); $excelSheet = $null }
    if ($excelPkg          -ne $null) { $excelPkg.Dispose(); $excelPkg = $null }
    if ($wbClient          -ne $null) { $wbClient.Dispose(); $wbClient = $null }    
    if ($Global:INetExp    -ne $null) { $Global:INetExp.Quit(); $refCnt = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Global:INetExp); $Global:INetExp = $null    }
    if ($Global:XmlHttpObj -ne $null) { $refCnt = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Global:XmlHttpObj); $Global:XmlHttpObj = $null }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

################################################################################################################################################################

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
