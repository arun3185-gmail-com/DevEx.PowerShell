
################################################################################################################################################################

$srcFolder = ""

$srcFilePath = "$($srcFolder).zip"
Compress-Archive -Path $srcFolder -DestinationPath $srcFilePath
[byte[]] $fileAsBytes = [System.IO.File]::ReadAllBytes($srcFilePath)
[string] $strBase64 = [System.Convert]::ToBase64String($fileAsBytes)
Set-Content -Encoding Ascii -Path "$($srcFilePath.Remove($srcFilePath.LastIndexOf("."))).txt" -Value $strBase64
Remove-Item -Path $srcFilePath
################################################################################################################################################################

$dstFilePath = ""

Add-Type -AssemblyName System.IO.Compression.FileSystem
[string] $strBase64 = Get-Content -Encoding Ascii -Path $dstFilePath
[byte[]] $fileAsBytes = [System.Convert]::FromBase64String($strBase64)
[System.IO.File]::WriteAllBytes("$($dstFilePath.Remove($dstFilePath.LastIndexOf("."))).zip", $fileAsBytes)
[System.IO.Compression.ZipFile]::ExtractToDirectory("$($dstFilePath.Remove($dstFilePath.LastIndexOf("."))).zip","$($dstFilePath.Remove($dstFilePath.LastIndexOf(".")).Remove($dstFilePath.LastIndexOf("\")))")
Remove-Item -Path "$($dstFilePath.Remove($dstFilePath.LastIndexOf("."))).zip"
################################################################################################################################################################

