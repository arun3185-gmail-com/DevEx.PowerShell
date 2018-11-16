
################################################################################################################################################################

Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"

################################################################################################################################################################

[string] $ServerName = "EMEAAWES01/Server/Evonik"
[string] $LNFilePath = "betrieb/TeamDoku_S8_BR.nsf"

$ArrayOfDefaultFields = 
@(
    @("NoteID"       , { param ($NotesDoc) $NotesDoc.NoteID }),
    @("UniversalID"  , { param ($NotesDoc) $NotesDoc.UniversalID }),
    @("Created"      , { param ($NotesDoc) $NotesDoc.Created }),
    @("LastModified" , { param ($NotesDoc) $NotesDoc.LastModified }),
    @("Form"         , { param ($NotesDoc) $NotesDoc.GetFirstItem("Form").Text })
)

################################################################################################################################################################

try
{
    [LN.NotesSession]  $nSession  = New-Object LN.NotesSession
    [LN.NotesDatabase] $nDatabase = $nSession.GetDatabase($ServerName, $LNFilePath)
    [LN.NotesDocumentCollection] $docCollection = $nDatabase.AllDocuments
    
    [string] $XlFileNamePrefix = $nDatabase.Title
    Write-Host "NotesURL : $($nDatabase.NotesURL)"
    Write-Host "Document collection Count - $($docCollection.Count)"
    
    [LN.NotesDocument] $doc = $docCollection.GetFirstDocument()
    while ($doc -ne $null)
    {
        $doc.GetFirstItem("Form").Text
        $doc = $docCollection.GetNextDocument($doc)
    }

    [LN.NotesFormCollection] $nForms = $nDatabase.Forms
    $nForms.Length
    for ($i = 0; $i -lt $nForms.Length; $i++)
    {
        $nForms[$i].Name
    }
    


    for ($i = 0; $i -lt $doc.Items.Length; $i++)
    {
        [LN.NotesItem] $itm = $doc.Items[$i]
        [string] $dtTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $itm.Type)
        "$($itm.Name) - $($dtTypeName)"
    }


    ################################################################################
}
catch
{
    throw
}
finally
{
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
