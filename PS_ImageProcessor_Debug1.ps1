
################################################################################################################################################################
# ImageProcessor Debug 1
################################################################################################################################################################
Import-Module "D:\Arun\Git\DevEx.References\NuGet\imageprocessor.2.6.2.25\lib\net452\ImageProcessor.dll"
################################################################################################################################################################

[string] $srcFilePath = "C:\Users\arunkumar.b08\Pictures\VietnamStairs_EN-AU4320366505_1366x768.jpg"

[byte[]] $srcFileBytes = [System.IO.File]::ReadAllBytes($srcFilePath)

[ImageProcessor.Imaging.Formats.ISupportedImageFormat] $format = New-Object ImageProcessor.Imaging.Formats.JpegFormat
$format.Quality = 70
[System.Drawing.Size] $size = New-Object System.Drawing.Size(150, 0)

[System.IO.MemoryStream] $inStream = $null
[System.IO.MemoryStream] $outStream = $null
[ImageProcessor.ImageFactory] $imgFactory = $null

try
{
    $inStream = New-Object System.IO.MemoryStream($srcFileBytes)
    $outStream = New-Object System.IO.MemoryStream
    $imgFactory = New-Object ImageProcessor.ImageFactory($true)

    $imgFactory.Load($inStream).Resize($size).Format($format).Save($outStream)

    [bytes[]] $dstFileBytes = New-Object bytes[$($outStream.Length)]
    
    
}
catch
{}
finally
{
    if ($imgFactory -ne $null) { $imgFactory.Dispose(); $imgFactory = $null }
    if ($outStream -ne $null) { $outStream.Dispose(); $outStream = $null }
    if ($inStream -ne $null) { $inStream.Dispose(); $inStream = $null }
}





<#
try
{}
catch
{}
finally
{}
#>

################################################################################################################################################################
# Main Program
################################################################################################################################################################



Write-Host ""
Write-Host "Done!"

################################################################################################################################################################

