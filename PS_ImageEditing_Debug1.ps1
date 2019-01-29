
################################################################################################################################################################
# Image editing Debug 1
################################################################################################################################################################

[string] $srcFilePath = "C:\Users\arunkumar.b08\Pictures\initial-car.jpg"
[string] $dstFilePath = "C:\Users\arunkumar.b08\Pictures\initial-car_New.jpg"

[System.Drawing.Image] $srcImage = $null
[System.Drawing.Graphics] $grpSrcImage = $null

try
{
    $srcImage = [System.Drawing.Image]::FromFile($srcFilePath)
    
    Write-Host "PixelFormat : $($srcImage.PixelFormat)"

    Write-Host "VerticalResolution   : $($srcImage.VerticalResolution)"
    Write-Host "HorizontalResolution : $($srcImage.HorizontalResolution)"
    
    Write-Host "Height : $($srcImage.Height)"
    Write-Host "Width  : $($srcImage.Width)"

    Write-Host "Size.Height : $($srcImage.Size.Height)"
    Write-Host "Size.Width  : $($srcImage.Size.Width)"

    Write-Host "PhysicalDimension.Height : $($srcImage.PhysicalDimension.Height)"
    Write-Host "PhysicalDimension.Width  : $($srcImage.PhysicalDimension.Width)"


    $grpSrcImage = [System.Drawing.Graphics]::FromImage($srcImage)
    $grpSrcImage.DrawLine([System.Drawing.Pens]::Black, 10, 10, 20, 20)

    $srcImage.Save($dstFilePath)
}
catch
{}
finally
{
    if ($grpSrcImage -ne $null) { $grpSrcImage.Dispose(); $grpSrcImage = $null }
    if ($srcImage -ne $null) { $srcImage.Dispose(); $srcImage = $null }
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

