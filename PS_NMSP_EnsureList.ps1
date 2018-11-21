
################################################################################################################################################################
# NMSP Ensure SPList
################################################################################################################################################################
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "J:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
################################################################################################################################################################

[string] $Username = "B13501@evonik.com"
[SecureString] $Password = Read-Host "Enter Password for $($Username)" -AsSecureString
[string] $SiteUrl   = "https://evonik.sharepoint.com/sites/10426"
[string] $ListTitle = "Info"

[PSCustomObject[]] $SharePointFieldSchemaTemplates = New-Object PSCustomObject[] 10

################################################################################################################################################################
# Functions
################################################################################################################################################################

Function Init-FieldSchemaTemplates
{
    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Text)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Single line of text"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value ""
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Text' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' MaxLength='255' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Note)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Multiple lines of text"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Plain text"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='FALSE' AppendOnly='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Note)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Multiple lines of text"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Rich text (Bold, italics, text alignment, hyperlinks)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='TRUE' RichTextMode='Compatible' IsolateStyles='FALSE' AppendOnly='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Note)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Multiple lines of text"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Enhanced rich text (Rich text with pictures, tables, and hyperlinks)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' IsolateStyles='TRUE' RestrictedMode='TRUE' AppendOnly='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Choice)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Choice (menu to choose from)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Drop-Down Menu"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Choice' Format='Dropdown' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Choice)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Choice (menu to choose from)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Radio Buttons"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Choice' Format='RadioButtons' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Choice)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Choice (menu to choose from)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Checkboxes (allow multiple selections)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='MultiChoice' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Integer)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Number (1, 1.0, 100)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places (Automatic)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Integer)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Number (1, 1.0, 100)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Integer)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Number (1, 1.0, 100)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Min and Max"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Min='' Max='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Integer)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Number (1, 1.0, 100)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places | Min and Max"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' Min='' Max='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Currency)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Currency ($, ¥, €)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places (Automatic)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Currency)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Currency ($, ¥, €)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Currency)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Currency ($, ¥, €)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Min and Max"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Min='' Max='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Currency)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Currency ($, ¥, €)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Number of decimal places | Min and Max"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' Min='' Max='' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::DateTime)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Date and Time"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "DateOnly | Display Format (Standard)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateOnly' FriendlyDisplayFormat='Disabled' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::DateTime)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Date and Time"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "DateTime | Display Format (Standard)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateTime' FriendlyDisplayFormat='Disabled' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::DateTime)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Date and Time"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "DateOnly | Display Format (Friendly)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateOnly' FriendlyDisplayFormat='Relative' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::DateTime)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Date and Time"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "DateTime | Display Format (Friendly)"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateTime' FriendlyDisplayFormat='Relative' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Lookup)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Lookup (information already on this site)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value ""
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Lookup' Name='' StaticName='' DisplayName='' List='{876fa85c-ecd5-4654-80ec-e0412ce96f07}' ShowField='Title' RelationshipDeleteBehavior='None' Required='FALSE' EnforceUniqueValues='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Lookup)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Lookup (information already on this site)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value "Allow multiple values"
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='LookupMulti' Name='' StaticName='' DisplayName='' List='{876fa85c-ecd5-4654-80ec-e0412ce96f07}' ShowField='Title' Mult='TRUE' RelationshipDeleteBehavior='None' Required='FALSE' EnforceUniqueValues='FALSE' />"

    $fldSchm = New-Object PSObject
    $fldSchm | Add-Member NoteProperty -Name "FieldType" -TypeName "Microsoft.SharePoint.Client.FieldType" -Value ([Microsoft.SharePoint.Client.FieldType]::Boolean)
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeDescription" -Value "Yes/No (check box)"
    $fldSchm | Add-Member NoteProperty -Name "FieldTypeOptions" -Value ""
    $fldSchm | Add-Member NoteProperty -Name "SchemaXml" -Value "<Field Type='Boolean' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default>1</Default></Field>"

<#



<Field Type='User' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleOnly' UserSelectionScope='0' />
<Field Type='User' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Group='' Mult='FALSE' />
<Field Type='UserMulti' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleOnly' UserSelectionScope='3' Group='' Mult='TRUE' />

<Field Type='URL' Format='Hyperlink' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />
<Field Type='URL' Format='Image' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />




<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateOnly' LCID='1081' ResultType='Text' ReadOnly='TRUE'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
<FieldRefs><FieldRef Name='Title' /></FieldRefs>
</Field>


<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateOnly' LCID='1081' ResultType='Number' ReadOnly='TRUE' Required='FALSE' Decimals='2' Percentage='TRUE'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
</Field>


<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateOnly' LCID='1081' ResultType='Currency' Decimals='2'  ReadOnly='TRUE'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
<FieldRefs><FieldRef Name='Title' /></FieldRefs>
</Field>

<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateOnly' LCID='1081' ResultType='DateTime' ReadOnly='TRUE' CustomFormatter='' Required='FALSE' Version='1'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
</Field>

<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateTime' LCID='1081' ResultType='DateTime' ReadOnly='TRUE' CustomFormatter='' Required='FALSE' Version='2'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
</Field>

<Field Type='Calculated' Name='' StaticName='' DisplayName='' EnforceUniqueValues='FALSE' Indexed='FALSE' Format='DateOnly' LCID='1081' ResultType='Boolean' ReadOnly='TRUE'>
<Formula>=Title</Formula>
<FormulaDisplayNames>=Title</FormulaDisplayNames>
<FieldRefs><FieldRef Name='Title' /></FieldRefs>
</Field>





#>

}

################################################################################################################################################################
# Main Program
################################################################################################################################################################

try
{
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Username , $Password)
    $context.Credentials = $credentials

    Write-Host "Querying web/list data..."

    $web = $context.Web
    $context.Load($web)
    $lists = $context.Web.Lists
    $context.Load($lists)
    $context.ExecuteQuery()

    $list = $lists | Where { $_.Title -eq $ListTitle }

    $fieldCollection = $null

    if ($list)
    {
        Write-Host "$($ListTitle) - Exists"
        $list = $context.Web.Lists.GetByTitle($ListTitle)
        $fieldCollection = $list.Fields
        $context.Load($fieldCollection)
    }
    else
    {
        Write-Host "$($ListTitle) - Creating..."
        $listInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $listInfo.QuickLaunchOption = [Microsoft.SharePoint.Client.QuickLaunchOptions]::On
        $listInfo.Title = $ListTitle
        $listInfo.TemplateType = [int] [Microsoft.SharePoint.Client.ListTemplateType]::GenericList
        $list = $context.Web.Lists.Add($listInfo)
        $fieldCollection = $list.Fields
        $context.Load($fieldCollection)
    }

    $context.Load($list)
    $context.ExecuteQuery()





}
catch
{
    #Write-Host    $_.Exception.ToString() -ForegroundColor Red
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
