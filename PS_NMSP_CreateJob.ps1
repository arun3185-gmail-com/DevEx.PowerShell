
################################################################################################################################################################
# NMSP Create job
################################################################################################################################################################

Add-Type -Path "J:\Arun\Git\DevEx.VB.Net\LN.vb"

################################################################################################################################################################

[string] $ServerName = "EMEAARFN01/Server/Evonik"
[string] $LNFilePath = "Abteilungen/PKM-R/archiv-pkmrabre.nsf"

[string] $jobFileFullPath = "J:\QuestJobs\Template0.pmjob"
[string] $NewJobFileFullPath = "J:\QuestJobs\Template1.pmjob"

[xml] $jobXmlDoc = $null

[string] $formNamesDelimited = "doku;doku_1;doku_2;doku_3;doku_4"

[string] $webUrl = "https://evonik.sharepoint.com/sites/10426"
[string] $listTitle = "Archiv"
[string] $defaultUserName = "Nadine.Jost@evonik.com"

################################################################################################################################################################

function New-SourceDefinitionColumn()
{
    Param
    (
        [string] $FieldName,
        [int]    $FieldTypeNumber,
        [string] $FieldTypeName = ""
    )

    $Local:SrcDefCol = New-Object PSObject
    $Local:SrcDefCol | Add-Member NoteProperty -Name "FieldName" -Value $FieldName
    $Local:SrcDefCol | Add-Member NoteProperty -Name "FieldTypeNumber" -Value $FieldTypeNumber
    $Local:SrcDefCol | Add-Member NoteProperty -Name "FieldTypeName" -Value $FieldTypeName

    Return $Local:SrcDefCol
}

try
{
    [LN.NotesSession]  $nSession  = New-Object LN.NotesSession
    [LN.NotesDatabase] $nDatabase = $nSession.GetDatabase($ServerName, $LNFilePath)
    [LN.NotesDocumentCollection] $docCollection = $nDatabase.AllDocuments
    
    
    
    $jobXmlDoc = New-Object xml
    $jobXmlDoc.Load($jobFileFullPath)

    [string[]] $serverNameParts = $ServerName.Split('/')
    if ($serverNameParts.Count -eq 3)
    {
        [System.Xml.XmlNode] $connStrNode = $jobXmlDoc.SelectSingleNode("/TransferJob/QuerySource/ConnectionString")
        $connStrNode.InnerXml = "server='CN=$($serverNameParts[0])/OU=$($serverNameParts[1])/O=$($serverNameParts[2])'; database='$($LNFilePath)'; zone=utc"

        [System.Xml.XmlNode] $srcDefNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition")
        $srcDefNode.Attributes["Name"].InnerText = $nDatabase.Title
        $srcDefNode.Attributes["Templates"].InnerText = $nDatabase.DesignTemplateName

        [System.Xml.XmlNode] $replicaIDNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/ReplicaId")
        $replicaIDNode.InnerText = $nDatabase.ReplicaID

        [PSCustomObject[]] $srcDefColumns = @()
        [string[]] $arrFormNames = $formNamesDelimited.Split(';')
        [int[]] $arrFormIndices = @()

        for ($i = 0; $i -lt $nDatabase.Forms.Length; $i++)
        {
            if ($nDatabase.Forms[$i].Name -in $arrFormNames) { $arrFormIndices += $i }
        }
        
        foreach ($formIdx in $arrFormIndices)
        {
            foreach ($fieldName in $nDatabase.Forms[$formIdx].Fields)
            {
                if ($srcDefColumns.Where({ $PSItem.FieldName -eq $fieldName }).Count -eq 0)
                {
                    $srcDefColumn = New-SourceDefinitionColumn -FieldName $fieldName -FieldTypeNumber ($nDatabase.Forms[$formIdx].GetFieldType($fieldName))
                    $srcDefColumns += $srcDefColumn
                }
            }
        }

        [LN.NotesDocument] $doc = $docCollection.GetFirstDocument()
        while ($doc -ne $null)
        {
            if ($doc.GetFirstItem("Form").Text -in $arrFormNames)
            {
                for ($i = 0; $i -lt $doc.Items.Length; $i++)
                {
                    [LN.NotesItem] $itm = $doc.Items[$i]
                    $srcDefColumns.Where({ $PSItem.FieldTypeNumber -eq 0 -and $PSItem.FieldName -eq $itm.Name }).ForEach({ $PSItem.FieldTypeNumber = $itm.Type; $PSItem.FieldTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $itm.Type) })
                }
            }

            if ($srcDefColumns.Where({ $PSItem.FieldTypeNumber -eq 0 }).Count -eq 0) { exit }
            
            $doc = $docCollection.GetNextDocument($doc)
        }

        if ($srcDefColumns.Where({ [string]::IsNullOrWhiteSpace($PSItem.FieldTypeName) }).Count -gt 0)
        {
            $srcDefColumns.Where({ [string]::IsNullOrWhiteSpace($PSItem.FieldTypeName) }).ForEach({ $PSItem.FieldTypeName = [System.Enum]::GetName([LN.NotesItemDataType], $PSItem.FieldTypeNumber) })
        }

        [System.Xml.XmlNode] $querySpecNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec")
        [System.Xml.XmlNode] $unidNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/UNID")

        foreach ($srcDefColumn in $srcDefColumns)
        {
            [System.Xml.XmlElement] $xmlElmtColumn = $jobXmlDoc.CreateElement("Column")

            $xmlElmtColumn.SetAttribute("ColumnType", "Item")
            $xmlElmtColumn.SetAttribute("Value", $srcDefColumn.FieldName)
            
            if ($srcDefColumn.FieldTypeNumber -in @(1280,1281))
            {
                $xmlElmtColumn.SetAttribute("ReturnType", "String")
            }
            elseif ($srcDefColumn.FieldTypeNumber -in @(1024))
            {
                $xmlElmtColumn.SetAttribute("ReturnType", "Date")
            }
            elseif ($srcDefColumn.FieldTypeNumber -in @(1076))
            {
                $xmlElmtColumn.SetAttribute("ReturnType", "String")
                $xmlElmtColumn.SetAttribute("Option", "Multi")
            }
            else
            {
                $xmlElmtColumn.SetAttribute("ReturnType", "String")
            }

            $querySpecNode.InsertAfter($xmlElmtColumn, $unidNode)
            
            Write-Host "$($srcDefColumn.FieldTypeName) - $($srcDefColumn.FieldName)"
        }

        [System.Xml.XmlNode] $formsNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SourceDefinition/QuerySpec/Forms")
        $formsNode.InnerText = $formNamesDelimited

        [System.Xml.XmlNode] $sp_WebNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SharePointConnection/Web")
        $sp_WebNode.InnerText = $webUrl

        [System.Xml.XmlNode] $sp_ListNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SharePointConnection/List")
        $sp_ListNode.InnerText = $listTitle

        [System.Xml.XmlNode] $userMapping_DefaultUserNameNode = $jobXmlDoc.SelectSingleNode("/TransferJob/JobOptions/UserMappingOptions/DefaultUserName")
        $userMapping_DefaultUserNameNode.InnerText = $defaultUserName


        [System.Xml.XmlNode] $sharepointDefNode = $jobXmlDoc.SelectSingleNode("/TransferJob/SharePointTargetDefinition")
        
    }
    else
    {
        Write-Host "error in server name" -ForegroundColor Red
    }




    $jobXmlDoc.Save($NewJobFileFullPath)
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
