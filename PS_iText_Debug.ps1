################################################################################################################################################################

Import-Module "D:\Arun\Git\DevEx.References\NuGet\itextsharp.5.5.13\lib\itextsharp.dll"

################################################################################################################################################################
# Main Program
################################################################################################################################################################

[string] $dstFilePath = "C:\Users\arunkumar.b08\Documents\Personal\Dest.pdf"

[string[]] $srcFilePaths = @("C:\Users\arunkumar.b08\Documents\Personal\Src1.pdf",
"C:\Users\arunkumar.b08\Documents\Personal\Src2.pdf",
"C:\Users\arunkumar.b08\Documents\Personal\Src3.pdf",
"C:\Users\arunkumar.b08\Documents\Personal\Src4.pdf",
"C:\Users\arunkumar.b08\Documents\Personal\Src5.pdf",
"C:\Users\arunkumar.b08\Documents\Personal\Src6.pdf")

[System.IO.FileStream] $opStream = $null

[string] $pwd = ""

try
{
    $opStream = New-Object System.IO.FileStream($dstFilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    
    [iTextSharp.text.Document] $document = New-Object iTextSharp.text.Document
    [iTextSharp.text.pdf.PdfCopy] $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $opStream)
    if (-not([string]::IsNullOrWhiteSpace($pwd)))
    {
        $pdfCopy.SetEncryption([System.Text.Encoding]::UTF8.GetBytes($pwd), [System.Text.Encoding]::UTF8.GetBytes($pwd), [iTextSharp.text.pdf.PdfCopy]::ALLOW_PRINTING, [iTextSharp.text.pdf.PdfCopy]::ENCRYPTION_AES_256)
    }
    
    [iTextSharp.text.pdf.PdfReader] $reader = $null
    
    try
    {
        $document.Open()
        
        foreach ($srcFilePath in $srcFilePaths)
        {
            $srcFilePath
            $reader = New-Object iTextSharp.text.pdf.PdfReader($srcFilePath)
            $pdfCopy.AddDocument($reader)
            $reader.Close()
        }
    }
    catch
    {
        if ($reader -ne $null) { $reader.Close() }
    }
    finally
    {
        if ($document -ne $null) { $document.Dispose() }
    }

    $opStream.Close()
}
catch
{
    Write-Host    $_.Exception.ToString() -ForegroundColor Red
}
finally
{
    if ($opStream -ne $null) { $opStream.Dispose() }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

################################################################################################################################################################

Write-Host ""
Write-Host "END!"
#$input = Read-Host "Hit 'Enter' key to close window!"

################################################################################################################################################################
