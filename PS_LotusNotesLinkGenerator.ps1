
################################################################################################################################################################
# Lotus Notes link generator
################################################################################################################################################################


[string] $UrlPrefix = "notes:///c12571d00062c633/0/"

[string] $UNIDTextList = @"


A2405606978E24E8C12577000051C107
893E5763AE05AC32C1257E73003A4214
B1943256E86779A6C1257E73003B1D8E
8ED5028F87E1DECEC1257E73003C04C4
D7102DC30AC28025C1257E73004162F5
9B4348C2234BD4F7C1257E730043603D
1E0C3F99E00CC58CC1257E7300437689
C041314727A3E78FC12580E4002086C9


"@


[System.Text.StringBuilder] $sbHtml = New-Object System.Text.StringBuilder

try
{
    [string[]] $UNIDs = $UNIDTextList.Split(@([System.Environment]::NewLine), [System.StringSplitOptions]::RemoveEmptyEntries)

    $sbHtml = $sbHtml.AppendLine("<html>")
    $sbHtml = $sbHtml.AppendLine("<head>")
    $sbHtml = $sbHtml.AppendLine("<title>Lotus Notes Link</title>")
    $sbHtml = $sbHtml.AppendLine("</head>")
    $sbHtml = $sbHtml.AppendLine("<body>")
    foreach ($unid in $UNIDs)
    {
        $sbHtml = $sbHtml.AppendLine("<a href='$($UrlPrefix)$($unid)'>$($unid)</a> <br />")
    }
    $sbHtml = $sbHtml.AppendLine("</body>")
    $sbHtml = $sbHtml.AppendLine("</html>")

    $sbHtml.ToString() | Out-File -FilePath "J:\Arun\Git\DevEx.PowerShell\LotusNotesLink.html"
}
catch { throw }
finally { }

################################################################################################################################################################
