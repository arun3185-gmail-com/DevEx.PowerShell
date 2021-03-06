﻿
################################################################################################################################################################
# Pluralsight Downloader   
################################################################################################################################################################

Add-Type -AssemblyName System.Web

Import-Module "D:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\htmlagilitypack.1.8.5\lib\Net40\HtmlAgilityPack.dll"

################################################################################################################################################################

[String] $Global:CourseMetaDataXL = "D:\Arun\Git\DevEx.Data\PluralsightCoursesMetadata.xlsx"
[String] $Global:UserName = "ben.andrews@I180618.onmicrosoft.com"
[String] $Global:Password = "P@180618"

[String] $Global:HostUri       = "https://app.pluralsight.com"
[String] $Global:LoginPageUrl  = "https://app.pluralsight.com/id?redirectTo=/library/"
[String] $Global:LogoutPageUrl = "https://app.pluralsight.com/id/signout"

[Boolean] $Global:Flag_DisableExerciseDownload = $false
[Boolean] $Global:Flag_DisableVideoDownload    = $false

[String] $Global:Tab        = [char]9
[String] $Global:TimeFormat = "[yyyy-MM-dd HH:mm:ss.fff]"

[String] $Global:ThisScriptRoot       = @("D:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[String] $Global:ThisScriptName       = "PS_Pluralsight_Dwnld_Debug"
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

[char[]] $Script:Alphas = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()

Function Convert-ExcelColumnNumberToName()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [int] $Number
    )

    if ($Number -lt 1)
    {
        throw New-Object System.ApplicationException("number must be greater than or equal to 1")
    }

    [int] $mod = $Number % 26
    [int] $coefOf26 = ($Number - $mod) / 26
    [int] $coefOf676 = ($Number - (26 * $coefOf26) - $mod) / 676
    [System.Text.StringBuilder] $colNameBuilder = New-Object System.Text.StringBuilder(3)

    if ($coefOf676 -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($coefOf676 -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }
    
    if ($coefOf26 -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($coefOf26 -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }

    if ($mod -eq 0) { $colNameBuilder.Append($Script:Alphas[25]) }
    elseif ($mod -gt 0) { $colNameBuilder.Append($Script:Alphas[$mod - 1]) }


    Return $colNameBuilder.ToString()
}

Function Convert-ExcelColumnNameToNumber()
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string] $Name
    )

    [int] $colNameLength = $Name.Length
    [int] $number = 0

    if ($colNameLength -ge 1) { $number +=       ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 1]) + 1) }
    if ($colNameLength -ge 2) { $number +=  26 * ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 2]) + 1) }
    if ($colNameLength -ge 3) { $number += 676 * ([array]::IndexOf($Script:Alphas, $Name[$colNameLength - 3]) + 1) }


    Return $number
}

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

[System.Net.WebClient] $wbClient = $null
[HtmlAgilityPack.HtmlDocument] $htmlDoc = $null

[OfficeOpenXml.ExcelPackage] $excelPkg = $null
[OfficeOpenXml.ExcelWorksheet] $excelSheet = $null

Try
{
    ################################################################################
    # Login
    Write-LogInfo "Logging In"
    Write-Host    "Logging In..."
    ################################################################################
    
    
    $Global:INetExp = New-Object -ComObject InternetExplorer.Application
    $Global:INetExp.Visible = $true
    $Global:INetExp.Silent = $true

    $Global:INetExp.Navigate($Global:LoginPageUrl)

    while ($Global:INetExp.Busy) { Start-Sleep -Milliseconds 100 }
    $elmtUserName = $Global:INetExp.Document.getElementById("Username")
    $elmtPassword = $Global:INetExp.Document.getElementById("Password")
    if ($elmtUserName               -ne $null    -and 
        $elmtUserName.GetTypeCode() -ne "DBNull" -and 
        $elmtPassword               -ne $null    -and 
        $elmtPassword.GetTypeCode() -ne "DBNull" )
    {
        try
        {
            $elmtUserName.value = $Global:UserName
            $elmtPassword.value = $Global:Password
            $elmtLoginButton = $Global:INetExp.Document.getElementById("login")
            $elmtLoginButton.click()

            while ($Global:INetExp.Busy) { Start-Sleep -Milliseconds 100 }
        }
        Catch
        {
            Write-Host $_.Exception.Message
        }
    }
    $Global:CookieString = $Global:INetExp.Document.cookie


    $wbClient = New-Object System.Net.WebClient
    $wbClient.Headers.Add([System.Net.HttpRequestHeader]::Cookie, $Global:CookieString)
    
    ################################################################################
    # Download Courses
    Write-LogInfo "Downloading Courses"
    Write-Host    "Downloading Courses..."
    ################################################################################

    $excelPkg = New-Object OfficeOpenXml.ExcelPackage((New-Object System.IO.FileInfo($Global:CourseMetaDataXL)))
    $excelSheet = $excelPkg.Workbook.Worksheets[1]
    
    [string[]] $ArrayOfColumnHeaders = @("Category","Status","CourseName","Rating","Level","Updated","Duration","CourseUrl")
    [int] $excelColumnCount = $excelSheet.Dimension.Columns
    [int] $excelRowCount = $excelSheet.Dimension.Rows

    for ($i = 2; $i -le $excelRowCount; $i++)
    {
        if ($excelSheet.Cells[$i, 2].Text -eq "Completed") { continue; }

        ################################################################################
        # Opening Course Url and Extracting Title, Info, Description
        #    $courseUrl = $excelSheet.Cells[$i, 8].Text
        ################################################################################

        $courseUrl = $excelSheet.Cells[$i, 8].Text

        Write-LogInfo "   Opening - $($courseUrl)"
        Write-Host    "   Opening... $($courseUrl)"

        $mainPageResponseString = $wbClient.DownloadString($courseUrl)

        Write-LogInfo "      Read - Title, Info, Description"
        Write-Host    "      Read - Title, Info, Description"

        $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
        $htmlDoc.LoadHtml($mainPageResponseString)
        
        $courseDirectoryInfo = $null
        $courseUrlName       = $courseUrl.Substring($courseUrl.LastIndexOf('/') + 1)
        $courseTitle         = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim())
        $courseInfo          = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectSingleNode("//div[@id='course-page-description']").SelectNodes("//div[@class='text-component']")[0].InnerText)
        $courseDescription   = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectNodes("//div[@class='course-info-tile-right']")[0].SelectNodes("p").InnerText)
        $validCourseTitle    = [string]::Join("", $courseTitle.Split([System.IO.Path]::GetInvalidFileNameChars()))
        $courseInfoRespJson  = $null
        $ArrResourceUrls     = $null

        ################################################################################
        # Getting Resource List
        ################################################################################

        $courseDirectoryInfo = New-Item -Path "$($Global:DownloadLocation)\$($validCourseTitle)" -ItemType Directory -Force
        
        if (Test-Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)")
        {
            $ArrResourceUrls = Import-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"
        }
        if ($ArrResourceUrls -eq $null)
        {
            $ArrResourceUrls = Get-ResourceList $courseUrl
            $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"
        }

        ################################################################################
        # Saving Course Information
        ################################################################################

        if ($ArrResourceUrls[0].StatusCode -lt 2)
        {
            Write-LogInfo "      Saving Course Information"
            Write-Host    "      Saving Course Information"
            
            $Local:sbCourseInfo = New-Object System.Text.StringBuilder

            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("Course Title:")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("--------------")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine($courseTitle)
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine()
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("Course Info:")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("-------------")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine($courseInfo)
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine()
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("Course Description:")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine("--------------------")
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine($courseDescription)
            $Local:sbCourseInfo = $Local:sbCourseInfo.AppendLine()
            
            $Local:sbCourseInfo.ToString() | Out-File -FilePath "$($courseDirectoryInfo.FullName)\$($ArrResourceUrls[0].RelativeFilePath)" -Encoding utf8

            $ArrResourceUrls[0].StatusCode = 2
        }

        ################################################################################
        # Downloading and Saving Course Metadata Json
        ################################################################################
        
        if ($ArrResourceUrls[1].StatusCode -lt 2)
        {
            Write-LogInfo "      Downloading and Saving Course Metadata Json"
            Write-Host    "      Downloading and Saving Course Metadata Json"
            
            $courseInfoRespJson = $wbClient.DownloadString($ArrResourceUrls[1].ResourceUrl)
            $courseInfoRespJson | Out-File -FilePath "$($courseDirectoryInfo.FullName)\$($ArrResourceUrls[1].RelativeFilePath)"

            $ArrResourceUrls[1].StatusCode = 2
        }
        else
        {
            Write-LogInfo "      Reading Course Metadata Json"
            Write-Host    "      Reading Course Metadata Json"

            $courseInfoRespJson = Get-Content -Path "$($courseDirectoryInfo.FullName)\$($ArrResourceUrls[1].RelativeFilePath)"
        }


        ################################################################################
        # Downloading and Saving Transcript Json
        ################################################################################

        <#
        if ($ArrResourceUrls[2].StatusCode -lt 2)
        {
            Write-LogInfo "      Downloading and Saving Transcript Json"
            Write-Host    "      Downloading and Saving Transcript Json"

            if ($Global:XmlHttpObj -eq $null)
            {
                $Global:XmlHttpObj = New-Object -ComObject Msxml2.XMLHTTP
            }
            $Global:XmlHttpObj.open("GET", $ArrResourceUrls[2].ResourceUrl, $false)
            $Global:XmlHttpObj.setRequestHeader("Cookie", $Global:CookieString)
            $Global:XmlHttpObj.send()
            $Global:XmlHttpObj.responseText | Out-File -FilePath "$($courseDirectoryInfo.FullName)\$($ArrResourceUrls[2].RelativeFilePath)"
            if ($Global:XmlHttpObj.status -eq 200)
            {
                $ArrResourceUrls[2].StatusCode = 2
            }
        }

        ################################################################################
        # Downloading and Saving Exercise files
        ################################################################################

        if ($ArrResourceUrls[3].StatusCode -lt 2) { }
        #>

        ################################################################################
        # Appending Videos to Resource List
        ################################################################################

        if ($ArrResourceUrls.Count -le 4)
        {
            Write-LogInfo "      Appending Videos to Resource List"
            Write-Host    "      Appending Videos to Resource List"

            $jsonHash = ConvertFrom-Json -InputObject $courseInfoRespJson
            for ($i = 0; $i -lt $jsonHash.modules.Count; $i++)
            {
                for ($j = 0; $j -lt $jsonHash.modules[$i].clips.Count; $j++)
                {
                    $validModuleName = "{0}. {1}" -f ($i+1), [string]::Join("", $jsonHash.modules[$i].title.Split([System.IO.Path]::GetInvalidFileNameChars()))
                    $validClipName   = "{0}. {1}" -f ($j+1), [string]::Join("", $jsonHash.modules[$i].clips[$j].title.Split([System.IO.Path]::GetInvalidFileNameChars()))

                    $VidPSObj = New-Object PSObject
                    $VidPSObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value $validModuleName
                    $VidPSObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value $validClipName
                    $VidPSObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value $jsonHash.modules[$i].clips[$j].playerUrl
                    $VidPSObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value ""
                    $VidPSObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value ""
                    $VidPSObj | Add-Member NoteProperty -Name "StatusCode"       -Value 0
                    $ArrResourceUrls += $VidPSObj
                }
            }
        }
        
        $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"

        ################################################################################
        # Getting and Downloading Videos
        # $vidResrcItem = $ArrResourceUrls[0]
        ################################################################################

        if ($ArrResourceUrls -ne $null -and $ArrResourceUrls.Count -gt 4)
        {
            for ($i = 4; $i -lt $ArrResourceUrls.Count; $i++)
            {
                if ($ArrResourceUrls[$i].StatusCode -eq 2) { continue }
                $vidResrcItem = $ArrResourceUrls[$i]

                ################################################################################
                # Get Video
                ################################################################################
                if ($vidResrcItem.StatusCode -eq 0)
                {
                    Write-LogInfo "      Getting     Video url - $($vidResrcItem.ResourceTitle)\$($vidResrcItem.ResourceSubTitle)"
                    Write-Host    "      Getting     Video url... $($vidResrcItem.ResourceTitle)\$($vidResrcItem.ResourceSubTitle)"

                    $exitCounter  = 0
                    $elmtVideoUrl = $null

                    Try
                    {
                        $Global:INetExp.Navigate("$($Global:HostUri)$($vidResrcItem.ResourcePageUrl)")
                        while ($Global:INetExp.Busy) { Start-Sleep -Milliseconds 100 }

                        while ($exitCounter -le 20)
                        {
                            Start-Sleep -Milliseconds 1000

                            if ($Global:INetExp.ReadyState -eq 4)
                            {
                                $divMainElement = $Global:INetExp.Document.getElementById("main")
                                $sectionElement = $Global:INetExp.Document.getElementById("app")
                                $divVideoContainer = $Global:INetExp.Document.getElementById("video-container")
                                $elmtVideoUrl = $Global:INetExp.Document.getElementById("vjs_video_3_html5_api")

                                if ($divMainElement -ne $null -and $sectionElement -ne $null -and $divVideoContainer -ne $null -and -not([string]::IsNullOrWhiteSpace($elmtVideoUrl.src)))
                                {
                                    $exitCounter = 21
                                }
                            }

                            $exitCounter++
                        }

                        $elmtVideoUrl = $Global:INetExp.Document.getElementById("vjs_video_3_html5_api")
                        if (-not([string]::IsNullOrWhiteSpace($elmtVideoUrl.src)))
                        {
                            $videoUri = New-Object System.Uri($elmtVideoUrl.src)
                            $uriFileName = $videoUri.Segments[$videoUri.Segments.Count - 1]
                            $extName = $uriFileName.Substring($uriFileName.LastIndexOf('.'))

                            $vidResrcItem.RelativeFilePath = "$($vidResrcItem.ResourceTitle)\$($vidResrcItem.ResourceSubTitle)$($extName)"
                            $vidResrcItem.ResourceUrl = $elmtVideoUrl.src
                            $vidResrcItem.StatusCode = 1
                        }
                    }
                    Catch
                    {
                        Write-LogInfo $_.Exception.ToString()
                        Write-Host    $_.Exception.Message -ForegroundColor Red
                    }
                }

                ################################################################################
                # Download Video
                ################################################################################
                
                if ($vidResrcItem.StatusCode -eq 1 -and -not($Global:Flag_DisableVideoDownload))
                {
                    Write-LogInfo "      Downloading Video url - $($vidResrcItem.RelativeFilePath)"
                    Write-Host    "      Downloading Video url... $($vidResrcItem.RelativeFilePath)"

                    $Local:moduleDirectoryInfo = $null                    

                    Try
                    {
                        $Local:moduleDirectoryInfo = New-Item -Path "$($courseDirectoryInfo.FullName)\$($vidResrcItem.ResourceTitle)" -ItemType Directory -Force
                        $wbClient.DownloadFile($vidResrcItem.ResourceUrl, "$($courseDirectoryInfo.FullName)\$($vidResrcItem.RelativeFilePath)")
                        $vidResrcItem.StatusCode = 2
                        
                    }
                    Catch
                    {
                        Write-LogInfo $_.Exception.ToString()
                        Write-Host    $_.Exception.Message -ForegroundColor Red
                    }
                }

                ################################################################################
                $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"
            }

        }


        ################################################################################
        # Update Metadata file
        ################################################################################

        if ($ArrResourceUrls -ne $null)
        {
            if ($ArrResourceUrls.Count -gt 4 -and $ArrResourceUrls.Where({ $_.StatusCode -ne 2 }).Count -gt 0)
            {
                $excelSheet.Cells[$i, 2].Value = "Completed"
            }
            else
            {
                $excelSheet.Cells[$i, 2].Value = "InComplete"
            }
        }
        else
        {
            $excelSheet.Cells[$i, 2].Value = "NotStarted"
        }
        
        $excelPkg.Save()

        ################################################################################
    }
    
    
    ################################################################################
    # Logout
    Write-LogInfo "Logging Out"
    Write-Host    "Logging Out..."
    ################################################################################

    $Global:INetExp.Navigate($Global:LogoutPageUrl)
    while ($Global:INetExp.Busy) { Start-Sleep -Milliseconds 100 }
    ################################################################################

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
