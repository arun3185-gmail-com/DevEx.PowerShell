﻿
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

    [string] $XlFileNamePrefix = $nDatabase.Title
    Write-Host "NotesURL : $($nDatabase.NotesURL)"
    Write-Host "Document collection Count - $($docCollection.Count)"
    Write-Host "Forms:"

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
    }
    


    for ($i = 0; $i -lt $doc.Items.Length; $i++)
    {
        [LN.NotesItem] $itm = $doc.Items[$i]
        [string] $dtTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $itm.Type)   #[Enum].GetName(GetType(EggSizeEnum))
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
