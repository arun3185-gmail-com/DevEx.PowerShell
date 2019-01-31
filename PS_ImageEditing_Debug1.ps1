
################################################################################################################################################################
# Image editing Debug 1
################################################################################################################################################################

<#
try{}
catch{}
finally{}
#>

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[string] $srcBlankFilePath = "C:\Users\arunkumar.b08\Pictures\Blank_24-bit.jpg"
[string] $srcFilePath = "C:\Users\arunkumar.b08\Pictures\initial-car.jpg"
[string] $dstFilePath = "C:\Users\arunkumar.b08\Pictures\initial-car_New.jpg"

[System.Drawing.Image] $srcBlankImage = $null
[System.Drawing.Image] $srcImage = $null
[System.Drawing.Graphics] $grpDstImage = $null
[System.Drawing.Bitmap] $destBMP = $null

try
{
    
    $srcImage = [System.Drawing.Image]::FromFile($srcFilePath)
    
    Write-Host "PixelFormat : $($srcImage.PixelFormat)"
    
    Write-Host "HorizontalResolution : $($srcImage.HorizontalResolution)"
    Write-Host "VerticalResolution   : $($srcImage.VerticalResolution)"
    
    Write-Host "Width  : $($srcImage.Width)"
    Write-Host "Height : $($srcImage.Height)"
    
    Write-Host "Size.Width  : $($srcImage.Size.Width)"
    Write-Host "Size.Height : $($srcImage.Size.Height)"
    
    Write-Host "PhysicalDimension.Width  : $($srcImage.PhysicalDimension.Width)"
    Write-Host "PhysicalDimension.Height : $($srcImage.PhysicalDimension.Height)"
    
    $srcBlankImage = [System.Drawing.Image]::FromFile($srcBlankFilePath)
    $grpDstImage = [System.Drawing.Graphics]::FromImage($srcBlankImage)
    #$destBMP = New-Object System.Drawing.Bitmap (1200, 900, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
    #$grpDstImage = [System.Drawing.Graphics]::FromImage($destBMP)

    [System.Drawing.Rectangle] $destRect1 = New-Object System.Drawing.Rectangle(0, 0, 600, 450)
    [System.Drawing.Rectangle] $destRect2 = New-Object System.Drawing.Rectangle(600, 0, 600, 450)
    [System.Drawing.Rectangle] $destRect3 = New-Object System.Drawing.Rectangle(0, 450, 600, 450)
    [System.Drawing.Rectangle] $destRect4 = New-Object System.Drawing.Rectangle(600, 450, 600, 450)

    $grpDstImage.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::Default

    $grpDstImage.DrawImage($srcImage, $destRect1, 0, 0, $srcImage.Width, $srcImage.Height, [System.Drawing.GraphicsUnit]::Pixel)
    $grpDstImage.DrawImage($srcImage, $destRect2, 0, 0, $srcImage.Width, $srcImage.Height, [System.Drawing.GraphicsUnit]::Pixel)
    $grpDstImage.DrawImage($srcImage, $destRect3, 0, 0, $srcImage.Width, $srcImage.Height, [System.Drawing.GraphicsUnit]::Pixel)
    $grpDstImage.DrawImage($srcImage, $destRect4, 0, 0, $srcImage.Width, $srcImage.Height, [System.Drawing.GraphicsUnit]::Pixel)


    #$grpDstImage.DrawLine([System.Drawing.Pens]::Black, 0, 0, 600, 450)
    #$grpDstImage.DrawLine([System.Drawing.Pens]::Black, 600, 0, 0, 450)


    $srcBlankImage.Save($dstFilePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    #$destBMP.Save($dstFilePath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
}
catch
{}
finally
{
    if ($destBMP -ne $null) { $destBMP.Dispose(); $destBMP = $null }
    if ($grpDstImage -ne $null) { $grpDstImage.Dispose(); $grpDstImage = $null }
    if ($srcImage -ne $null) { $srcImage.Dispose(); $srcImage = $null }
    if ($srcBlankImage -ne $null) { $srcBlankImage.Dispose(); $srcBlankImage = $null }
}


Write-Host ""
Write-Host "Done!"

################################################################################################################################################################

