################################################################################################################################################################
# Pluralsight Course Metadata   
################################################################################################################################################################

Add-Type -AssemblyName System.Web

Import-Module "F:\Arun\DevEx\NuPkg\htmlagilitypack.1.8.5\lib\Net40\HtmlAgilityPack.dll"
Import-Module "F:\Arun\DevEx\NuPkg\selenium.webdriver.3.14.0\lib\net40\WebDriver.dll"
Import-Module "F:\Arun\DevEx\NuPkg\selenium.webdriverbackedselenium.3.14.0\lib\net40\Selenium.WebDriverBackedSelenium.dll"
Import-Module "F:\Arun\DevEx\NuPkg\selenium.support.3.14.0\lib\net40\WebDriver.Support.dll"

################################################################################################################################################################
# Declarations
################################################################################################################################################################

[String] $Global:CourseMetadataCsv    = "F:\Arun\Pluralsight\CoursesDetails.csv"
[string] $Global:ChromeDriverLocation = "F:\Arun\DevEx\NuPkg\chromedriver_win32"
[String] $Global:ThisScriptRoot       = @("D:\Arun\DevEx\PS", $PSScriptRoot)[($PSScriptRoot -ne $null -and $PSScriptRoot.Length -gt 0)]
[String] $Global:ThisScriptName       = "PS_Pluralsight_Metadata"

if ($PSCommandPath -ne $null -and $PSCommandPath.Length -gt 0)
{
    $idx = $PSCommandPath.LastIndexOf('\') + 1
    $Global:ThisScriptName = $PSCommandPath.Substring($idx, $PSCommandPath.LastIndexOf('.') - $idx)
}

[String] $Global:LogFilePath      = "$($Global:ThisScriptRoot)\$($Global:ThisScriptName).log"

################################################################################################################################################################
# Functions
################################################################################################################################################################

Function Write-LogInfo()
{
    Param ([String] $Message)
    
    "$(Get-Date -Format $Global:TimeFormat):$($Global:Tab)$($Message)" | Out-File -FilePath $Global:LogFilePath -Append
}

################################################################################################################################################################
# Main Program
################################################################################################################################################################


[HtmlAgilityPack.HtmlDocument] $htmlDoc = $null
[OpenQA.Selenium.Chrome.ChromeDriver] $webDriver = $null

Try
{
    $coursesData = Import-Csv $Global:CourseMetadataCsv

    $webDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($Global:ChromeDriverLocation)

    foreach ($courseData in $coursesData)
    {
        $courseData.Author = ""
        # $courseData = $coursesData[18]
        if (-not([string]::IsNullOrWhiteSpace($courseData.CourseUrl)))
        {
            $webDriver.Navigate().GoToUrl($courseData.CourseUrl)

            $htmlWebElmnt = $webDriver.FindElementByTagName("html");
            $htmlContents = ([string]([OpenQA.Selenium.IJavaScriptExecutor]$webDriver).ExecuteScript("return arguments[0].outerHTML;", $htmlWebElmnt))

            $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
            $htmlDoc.LoadHtml($htmlContents)
            
            $courseName = [System.Web.HttpUtility]::HtmlDecode($htmlDoc.DocumentNode.SelectSingleNode("//h1").InnerText.Trim())
            $courseAuthor = $htmlDoc.DocumentNode.SelectSingleNode("//h5[@class='title--alternate']").SelectSingleNode("a").InnerText
            $courseInfoElement = $htmlDoc.DocumentNode.SelectSingleNode("//div[@id='course-description-tile-info']")
            
            $courseStarNodes = $courseInfoElement.SelectSingleNode("//div[@class='course-info__row--right course-info__row--rating']")
            $courseRating = $courseStarNodes.SelectNodes("i[@class='fa fa-star']").Count.ToString() + @(" ", ".5 ")[$courseStarNodes.SelectNodes("i[@class='fa fa-star-half-o']").Count]
            $courseRating_Nos = $courseInfoElement.SelectSingleNode("//div[@class='course-info__row--right course-info__row--rating']").SelectSingleNode("span").InnerText.Trim()
            $courseLevel = $courseInfoElement.SelectSingleNode("//div[@class='course-info__row--right difficulty-level']").InnerText.Trim()
            $courseUpdated = $courseInfoElement.SelectNodes("//div[@class='course-info__row--right']")[0].InnerText.Trim()
            $courseDuration = $courseInfoElement.SelectNodes("//div[@class='course-info__row--right']")[1].InnerText.Trim()

            $courseData.Author = $courseAuthor
            $courseData.Rating = "$courseRating$courseRating_Nos"
            $courseData.Level = "$courseLevel"
            $courseData.Updated = $courseUpdated
            $courseData.Duration = $courseDuration

        }
        $courseData.Author
    }

    $coursesData | Export-Csv -Path $Global:CourseMetadataCsv -NoTypeInformation
}
Catch
{
    Write-LogInfo $_.Exception.ToString()
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
Finally
{
    if ($XlDoc -ne $null) { $XlDoc.Dispose(); }
    if ($webDriver -ne $null) { $webDriver.Close(); $webDriver.Quit(); $webDriver.Dispose(); $webDriver = $null }
    
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

################################################################################################################################################################

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
