

Import-Module "F:\Arun\DevEx\NuPkg\htmlagilitypack.1.8.5\lib\Net40\HtmlAgilityPack.dll"

[HtmlAgilityPack.HtmlDocument] $htmlDoc = $null

$ChairUrlsText = @"

http://www.chairwale.com/chairdetails.php?id=82&catid=8
http://www.chairwale.com/chairdetails.php?id=86&catid=8
http://www.chairwale.com/chairdetails.php?id=357&catid=8
http://www.chairwale.com/chairdetails.php?id=443&catid=8
http://www.chairwale.com/chairdetails.php?id=453&catid=8
http://www.chairwale.com/chairdetails.php?id=515&catid=8
http://www.chairwale.com/chairdetails.php?id=531&catid=8
http://www.chairwale.com/chairdetails.php?id=533&catid=8
http://www.chairwale.com/chairdetails.php?id=552&catid=8
http://www.chairwale.com/chairdetails.php?id=577&catid=8
http://www.chairwale.com/chairdetails.php?id=581&catid=8
http://www.chairwale.com/chairdetails.php?id=590&catid=8
http://www.chairwale.com/chairdetails.php?id=592&catid=8
http://www.chairwale.com/chairdetails.php?id=763&catid=8

"@

$ArrChairUrls = $ChairUrlsText.Split([System.Environment]::NewLine, [System.StringSplitOptions]::RemoveEmptyEntries)

$hashChairs = @()

foreach ($chairUrl in  $ArrChairUrls)
{
    $webRqst = Invoke-WebRequest -Uri $chairUrl -UseBasicParsing

    $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
    $htmlDoc.LoadHtml($webRqst.Content)
    
    
    $hsData = @{}
    $arrQS = $chairUrl.Substring($chairUrl.LastIndexOf("?")+1).Split("&", [System.StringSplitOptions]::RemoveEmptyEntries)

    foreach ($qs in $arrQS)
    {
        $pair = $qs.Split("=", [System.StringSplitOptions]::RemoveEmptyEntries)
        $hsData.Add($pair[0],$pair[1])
    }
    
    $hsData.Add("Name", $htmlDoc.GetElementbyId("content").SelectNodes("//*[contains(@class,'produttitle')]")[0].InnerText.Trim())
    $hsData.Add("Material", $htmlDoc.DocumentNode.SelectNodes("//*[contains(@class,'productspecification')]")[0].ChildNodes[7].ChildNodes[3].InnerText.Trim())
    $hsData.Add("Price", $htmlDoc.DocumentNode.SelectNodes("//*[contains(@class,'col-xs-6 col-sm-7 col-md-7 provalue price redtext boldtext increasetext')]")[0].InnerText.Trim())
    $imgRqst = Invoke-WebRequest -Uri "http://www.chairwale.com/$($htmlDoc.DocumentNode.SelectNodes("//*[contains(@class,'produtpic')]//img")[0].Attributes["src"].Value)" -UseBasicParsing

    $hsData.Add("PicData", [Convert]::ToBase64String($imgRqst.Content))

    $hashChairs += $hsData
}



ConvertTo-Json $hashChairs | Out-File -FilePath "F:\Arun\DevEx\PS\chairs.json"
