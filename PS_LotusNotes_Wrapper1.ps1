
################################################################################################################################################################

Import-Module "J:\Arun\Git\DevEx.References\NuGet\epplus.4.5.2.1\lib\net40\EPPlus.dll"
Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"

################################################################################################################################################################

[string] $ServerName = "EMEAARFN01/Server/Evonik"
[string] $LNFilePath = "Abteilungen/PKM-R/archiv-pkmrabre.nsf"

$ArrayOfDefaultFields = 
@(
    @("NoteID"       , { param ($NotesDoc) $NotesDoc.NoteID }),
    @("UniversalID"  , { param ($NotesDoc) $NotesDoc.UniversalID }),
    @("Created"      , { param ($NotesDoc) $NotesDoc.Created }),
    @("LastModified" , { param ($NotesDoc) $NotesDoc.LastModified }),
    @("Form"         , { param ($NotesDoc) $NotesDoc.GetFirstItem("Form").Text })
)

$ArrayOfCurntDBFields = @("ObserverCanon","DateObservation","ObserverArea","ObservationArea","Observation","JobObserved","PeopleObservedInt","Comments","Attachments","ReviewerCanon","Status","ReviewerComments","DateSubmitted")

<#
doccreated
docmodified
ModifiedBy
unid
#>

################################################################################################################################################################

try
{
    [LN.NotesSession]  $nSession  = New-Object LN.NotesSession
    [LN.NotesDatabase] $nDatabase = $nSession.GetDatabase($ServerName, $LNFilePath)
    [LN.NotesDocumentCollection] $docCollection = $nDatabase.AllDocuments

    ################################################################################

    Write-Host "NotesURL : $($nDatabase.NotesURL)"
    Write-Host "Document collection Count - $($docCollection.Count)"
    Write-Host "Forms:"

    [LN.NotesFormCollection] $nForms = $nDatabase.Forms

    $nForms.Length
    for ($i = 0; $i -lt $nForms.Length; $i++)
    {
        "    $($nForms[$i].Name)"
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
