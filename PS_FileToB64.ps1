
$srcFilePath = ""

[byte[]] $fileAsBytes = [System.IO.File]::ReadAllBytes($srcFilePath)
[string] $strBase64 = [System.Convert]::ToBase64String($fileAsBytes)

Set-Content -Encoding Ascii -Path "$($srcFilePath.Remove($srcFilePath.LastIndexOf("."))).txt" -Value $strBase64


$dstFilePath = ""

[string] $strBase64 = Get-Content -Encoding Ascii -Path $dstFilePath
[byte[]] $fileAsBytes = [System.Convert]::FromBase64String($strBase64)
[System.IO.File]::WriteAllBytes("$($dstFilePath.Remove($dstFilePath.LastIndexOf("."))).zip", $fileAsBytes)
