
################################################################################################################################################################
# Pluralsight Downloader Selenium
################################################################################################################################################################

Add-Type -AssemblyName System.Web

Import-Module "D:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\htmlagilitypack.1.8.5\lib\Net40\HtmlAgilityPack.dll"

Import-Module "D:\Arun\Git\DevEx.References\NuGet\selenium.webdriver.3.14.0\lib\net40\WebDriver.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\selenium.webdriverbackedselenium.3.14.0\lib\net40\Selenium.WebDriverBackedSelenium.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\selenium.support.3.14.0\lib\net40\WebDriver.Support.dll"

################################################################################################################################################################
<#

https://app.pluralsight.com/library/courses/microsoft-azure-developers-what-to-use
https://app.pluralsight.com/library/courses/developing-dotnet-microsoft-azure-getting-started
https://app.pluralsight.com/library/courses/azure-logic-apps-fundamentals
https://app.pluralsight.com/library/courses/introduction-azure-app-services
https://app.pluralsight.com/library/courses/azure-paas-building-global-app

https://app.pluralsight.com/library/courses/graph-databases-neo4j-introduction
https://app.pluralsight.com/library/courses/understanding-machine-learning
https://app.pluralsight.com/library/courses/r-programming-fundamentals
https://app.pluralsight.com/library/courses/tree-based-models-classification
https://app.pluralsight.com/library/courses/understanding-applying-linear-regression

https://app.pluralsight.com/library/courses/python-understanding-machine-learning
https://app.pluralsight.com/library/courses/machine-learning-algorithms
https://app.pluralsight.com/library/courses/advanced-machine-learning-encog-pt2
https://app.pluralsight.com/library/courses/introduction-to-machine-learning-encog
https://app.pluralsight.com/library/courses/building-sentiment-analysis-systems-python

https://app.pluralsight.com/library/courses/introduction-microsoft-office-open-xml

#>

################################################################################################################################################################
#[String] $Global:UserName = "Alex.Grayson@DxIT180818.onmicrosoft.com"
#[String] $Global:Password = "Plur@1sight"
[String] $Global:CourseMetaDataXL = "D:\Arun\Git\DevEx.Data\PluralsightCoursesMetadata - Copy.xlsx"
[String] $Global:UserName = "john.travolta@I180618.onmicrosoft.com"
[String] $Global:Password = "P@180618"

[String] $Global:HostUri       = "https://app.pluralsight.com"
[String] $Global:LoginPageUrl  = "https://app.pluralsight.com/id?redirectTo=/library/"
[String] $Global:LogoutPageUrl = "https://app.pluralsight.com/id/signout"

[Boolean] $Global:Flag_DisableExerciseDownload = $false
[Boolean] $Global:Flag_DisableVideoDownload    = $false

[String] $Global:Tab        = [char]9
[String] $Global:TimeFormat = "[yyyy-MM-dd HH:mm:ss.fff]"

[string] $Global:ChromeDriverLocation = "D:\Arun\Git\DevEx.References\NuGet\chromedriver_win32"
[String] $Global:ThisScriptRoot       = @("D:\Arun\Git\DevEx.PowerShell", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[String] $Global:ThisScriptName       = "PS_Pluralsight_Dwnld_Selenium"
[String] $Global:ResourceListFileName = "ResourceCheckList.csv"

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}

[String] $Global:LogFilePath      = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"
[String] $Global:DownloadLocation = "$($Global:ThisScriptRoot)\Pluralsight"
[String] $Global:FileDownloadLocation = "$($Global:ThisScriptRoot)\Pluralsight\Dwnlds"


[String] $Global:CookieString           = $null

################################################################################################################################################################
# Functions
################################################################################################################################################################

Function Write-LogInfo()
{
    Param ([String] $Message)
    
    "$(Get-Date -Format $Global:TimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

Function Create-ResourceItem()
{
    Param
    (        
        [Parameter(Mandatory = $false)]
        [string] $ResourceTitle = "",

        [Parameter(Mandatory = $false)]
        [string] $ResourceSubTitle = "",

        [Parameter(Mandatory = $false)]
        [string] $ResourcePageUrl = "",

        [Parameter(Mandatory = $false)]
        [string] $RelativeFilePath = "",

        [Parameter(Mandatory = $false)]
        [string] $ResourceUrl = "",

        [Parameter(Mandatory = $false)]
        [int] $StatusCode = 0
    )

    $Local:ResObj = New-Object PSObject
    $Local:ResObj | Add-Member NoteProperty -Name "ResourceTitle"    -Value $ResourceTitle
    $Local:ResObj | Add-Member NoteProperty -Name "ResourceSubTitle" -Value $ResourceSubTitle
    $Local:ResObj | Add-Member NoteProperty -Name "ResourcePageUrl"  -Value $ResourcePageUrl
    $Local:ResObj | Add-Member NoteProperty -Name "RelativeFilePath" -Value $RelativeFilePath
    $Local:ResObj | Add-Member NoteProperty -Name "ResourceUrl"      -Value $ResourceUrl
    $Local:ResObj | Add-Member NoteProperty -Name "StatusCode"       -Value $StatusCode
    
    Return $Local:ResObj
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

[OpenQA.Selenium.Chrome.ChromeDriver] $webDriver = $null
[OpenQA.Selenium.Chrome.ChromeOptions] $chrmOpts = $null
[System.IO.FileSystemWatcher] $downloadsCatcher = $null

[string[]] $chromeTabs = $null

Try
{
    ################################################################################
    # Login
    Write-LogInfo "Logging In"
    Write-Host    "Logging In..."
    ################################################################################

    $chrmOpts = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $chrmOpts.AddUserProfilePreference("download.prompt_for_download", $false)
    $chrmOpts.AddUserProfilePreference("download.directory_upgrade", $true)
    $chrmOpts.AddUserProfilePreference("download.default_directory", $Global:FileDownloadLocation)
    
    $webDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($Global:ChromeDriverLocation, $chrmOpts)
    
    $webDriver.Navigate().GoToUrl($Global:LoginPageUrl)
    
    [System.Collections.Generic.IEnumerable[OpenQA.Selenium.IWebElement]] $loginElements = $webDriver.FindElementsById("login")

    if ($loginElements.Count -ge 1)
    {
        try
        {
            [OpenQA.Selenium.IWebElement] $loginBtnWebElmnt = $loginElements[0]
            [OpenQA.Selenium.IWebElement] $usernameWebElmnt = $webDriver.FindElementById("Username")
            [OpenQA.Selenium.IWebElement] $passwordWebElmnt = $webDriver.FindElementById("Password")

            $usernameWebElmnt.SendKeys($Global:UserName)
            $passwordWebElmnt.SendKeys($Global:Password)
            $loginBtnWebElmnt.Click()
        }
        catch
        {
            Write-Host $_.Exception.Message
        }
    }
    
    $Global:CookieString = ""
    foreach ($wbDrvCookie in $webDriver.Manage().Cookies.AllCookies)
    {
        $Global:CookieString += "$($wbDrvCookie.ToString());"
    }
    
    #[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Ssl3 -bor [System.Net.SecurityProtocolType]::Tls -bor [System.Net.SecurityProtocolType]::Tls11 -bor [System.Net.SecurityProtocolType]::Tls12
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

        $webDriver.Navigate().GoToUrl("$($courseUrl)/description")

        Write-LogInfo "      Read - Title, Info, Description"
        Write-Host    "      Read - Title, Info, Description"
        
        <#
        $htmlWebElmnt = $webDriver.FindElementByTagName("html");
        $htmlContents = ([string]([OpenQA.Selenium.IJavaScriptExecutor]$webDriver).ExecuteScript("return arguments[0].innerHTML;", $htmlWebElmnt))
        #$htmlContents = $webDriver.PageSource
        #>

        $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
        $htmlDoc.LoadHtml($webDriver.PageSource)
        
        <#
        $mainPageResponseString = $wbClient.DownloadString($courseUrl)
        $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
        $htmlDoc.LoadHtml($mainPageResponseString)
        #>

        $courseDirectoryInfo = $null
        $courseUrlName       = $courseUrl.Substring($courseUrl.LastIndexOf('/') + 1)
        $courseTitle         = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim())
        $courseInfo          = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectSingleNode("//div[@id='ps-main']").SelectNodes("//p[@class='course-hero__excerpt']")[0].InnerText)
        $courseDescription   = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectNodes("//div[@class='l-course-page__content']")[0].SelectNodes("p").InnerText)
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
            
            $webDriver.Navigate().GoToUrl($ArrResourceUrls[1].ResourceUrl)
            $htmlWebElmnt = $webDriver.FindElementByTagName("html");
            $courseInfoRespJson = ([string]([OpenQA.Selenium.IJavaScriptExecutor]$webDriver).ExecuteScript("return arguments[0].outerText;", $htmlWebElmnt))

            #$courseInfoRespJson = $wbClient.DownloadString($ArrResourceUrls[1].ResourceUrl)
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

        if ($ArrResourceUrls[2].StatusCode -lt 2)
        {
            Write-LogInfo "      Downloading and Saving Transcript Json"
            Write-Host    "      Downloading and Saving Transcript Json"

            $webDriver.Navigate().GoToUrl($ArrResourceUrls[2].ResourceUrl)
            $courseTranscriptJson = $webDriver.FindElementByTagName("html").GetAttribute("innerText")
            $courseTranscriptJson | Out-File -FilePath "$($courseDirectoryInfo.FullName)\$($ArrResourceUrls[2].RelativeFilePath)"
            
            # if status 200
            $ArrResourceUrls[2].StatusCode = 2
        }

        ################################################################################
        # Downloading and Saving Exercise file
        ################################################################################
        
        if ($ArrResourceUrls[3].StatusCode -lt 2)
        {
            Write-LogInfo "      Downloading and Saving Exercise file"
            Write-Host    "      Downloading and Saving Exercise file"        
        
            $webDriver.Navigate().GoToUrl("$($courseUrl)/exercise-files")
            
            Get-ChildItem -Path $Global:FileDownloadLocation | Remove-Item

            [System.IO.DirectoryInfo] $dirInfo = New-Object System.IO.DirectoryInfo($Global:FileDownloadLocation)
            [int] $initFilesCount = $dirInfo.GetFiles().Count
            [int] $counter = 0

            $dwnldBtnWebElmnt = $webDriver.FindElementsByTagName("button").Where({ $PSItem.Text.Contains("Download") -and $PSItem.Text.Contains("exercise") })[0]
            $dwnldBtnWebElmnt.Click()
            
            do { $counter++; Start-Sleep -Milliseconds 5000; }
            While ($dirInfo.GetFiles().Count -eq $initFilesCount -and $counter -le 3)


            if ($dirInfo.GetFiles().Count -gt 0 -and $dirInfo.GetFiles().Where({ $PSItem.Extension -in @(".tmp",".crdownload") }).Count -eq 1)
            {
                $counter = 0

                do { $counter++; Start-Sleep -Milliseconds 5000; }
                While ($dirInfo.GetFiles().Where({ $PSItem.Extension -in @(".tmp",".crdownload") }).Count -eq 1 -and $counter -le 10)

                if ($dirInfo.GetFiles().Where({ $PSItem.Extension -in @(".tmp",".crdownload") }).Count -eq 0 -and $dirInfo.GetFiles().Count -eq 1)
                {
                    Move-Item -Path "$($Global:FileDownloadLocation)\*.*" -Destination $courseDirectoryInfo.FullName
                    $ArrResourceUrls[3].StatusCode = 2

                    Write-LogInfo "      Downloaded and saved Exercise file"
                    Write-Host    "      Downloaded and saved Exercise file"
                }
                else
                {
                    Write-LogInfo "      Exercise file download timeout! Skipping"
                    Write-Host    "      Exercise file download timeout! Skipping..."
                }
            }
            else
            {
                Write-LogInfo "      Looks like nothing is downloading"
                Write-Host    "      Looks like nothing is downloading..."
            }
        }

        $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"

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

                    try
                    {
                        $webDriver.Navigate().GoToUrl("$($Global:HostUri)$($vidResrcItem.ResourcePageUrl)")
                        Start-Sleep -Milliseconds 2000

                        [OpenQA.Selenium.Support.UI.WebDriverWait] $videoLoadWait = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($webDriver, [timespan]::FromSeconds(100))
                        $videoElements = $videoLoadWait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::TagName("video")))
                        
                        #$videoElements = $webDriver.FindElementsByTagName("video")
                        
                        if ($videoElements.Count -gt 0 -and  (-not([string]::IsNullOrWhiteSpace($videoElements[0].GetAttribute("src")))))
                        {
                            $elmtVideoUrl = $videoElements[0]

                            $videoUri = New-Object System.Uri($elmtVideoUrl.GetAttribute("src"))
                            $uriFileName = $videoUri.Segments[$videoUri.Segments.Count - 1]
                            $extName = $uriFileName.Substring($uriFileName.LastIndexOf('.'))

                            $vidResrcItem.RelativeFilePath = "$($vidResrcItem.ResourceTitle)\$($vidResrcItem.ResourceSubTitle)$($extName)"
                            $vidResrcItem.ResourceUrl = $elmtVideoUrl.GetAttribute("src")
                            $vidResrcItem.StatusCode = 1

                            $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"
                        }
                    }
                    catch
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

                        $ArrResourceUrls | Export-Csv -Path "$($courseDirectoryInfo.FullName)\$($Global:ResourceListFileName)"
                        
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

    $webDriver.Navigate().GoToUrl($Global:LogoutPageUrl)

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
    if ($webDriver         -ne $null) { $webDriver.Close(); $webDriver.Quit(); $webDriver.Dispose(); $webDriver = $null }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

################################################################################################################################################################

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
