
################################################################################################################################################################
# Lotus Notes to SharePoint Migration Automation
#     SharePoint Field templates
################################################################################################################################################################

Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.dll"
Import-Module "J:\LN2SP\References\SharePoint\Microsoft.SharePoint.Client.Runtime.dll"

################################################################################################################################################################
# Functions
################################################################################################################################################################

Function Get-SPFieldsSchemaTemplate
{
    @(
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Text
            "FieldTypeDescription"  = "Single line of text"
            "FieldTypeOptions"      = ""
            "SchemaXml"             = "<Field Type='Text' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' MaxLength='255' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Note
            "FieldTypeDescription"  = "Multiple lines of text"
            "FieldTypeOptions"      = "Plain text"
            "SchemaXml"             = "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='FALSE' AppendOnly='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Note
            "FieldTypeDescription"  = "Multiple lines of text"
            "FieldTypeOptions"      = "Rich text (Bold, italics, text alignment, hyperlinks)"
            "SchemaXml"             = "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='TRUE' RichTextMode='Compatible' IsolateStyles='FALSE' AppendOnly='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Note
            "FieldTypeDescription"  = "Multiple lines of text"
            "FieldTypeOptions"      = "Enhanced rich text (Rich text with pictures, tables, and hyperlinks)"
            "SchemaXml"             = "<Field Type='Note' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' IsolateStyles='TRUE' RestrictedMode='TRUE' AppendOnly='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Choice
            "FieldTypeDescription"  = "Choice (menu to choose from)"
            "FieldTypeOptions"      = "Drop-Down Menu"
            "SchemaXml"             = "<Field Type='Choice' Format='Dropdown' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Choice
            "FieldTypeDescription"  = "Choice (menu to choose from)"
            "FieldTypeOptions"      = "Radio Buttons"
            "SchemaXml"             = "<Field Type='Choice' Format='RadioButtons' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Choice
            "FieldTypeDescription"  = "Choice (menu to choose from)"
            "FieldTypeOptions"      = "Checkboxes (allow multiple selections)"
            "SchemaXml"             = "<Field Type='MultiChoice' FillInChoice='FALSE' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default></Default><CHOICES><CHOICE></CHOICE></CHOICES></Field>"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Integer
            "FieldTypeDescription"  = "Number (1, 1.0, 100)"
            "FieldTypeOptions"      = "Number of decimal places (Automatic)"
            "SchemaXml"             = "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Integer
            "FieldTypeDescription"  = "Number (1, 1.0, 100)"
            "FieldTypeOptions"      = "Number of decimal places"
            "SchemaXml"             = "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Integer
            "FieldTypeDescription"  = "Number (1, 1.0, 100)"
            "FieldTypeOptions"      = "Min and Max"
            "SchemaXml"             = "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Min='' Max='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Integer
            "FieldTypeDescription"  = "Number (1, 1.0, 100)"
            "FieldTypeOptions"      = "Number of decimal places | Min and Max"
            "SchemaXml"             = "<Field Type='Number' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' Min='' Max='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Currency
            "FieldTypeDescription"  = "Currency ($, ¥, €)"
            "FieldTypeOptions"      = "Number of decimal places (Automatic)"
            "SchemaXml"             = "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Currency
            "FieldTypeDescription"  = "Currency ($, ¥, €)"
            "FieldTypeOptions"      = "Number of decimal places"
            "SchemaXml"             = "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Currency
            "FieldTypeDescription"  = "Currency ($, ¥, €)"
            "FieldTypeOptions"      = "Min and Max"
            "SchemaXml"             = "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Min='' Max='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Currency
            "FieldTypeDescription"  = "Currency ($, ¥, €)"
            "FieldTypeOptions"      = "Number of decimal places | Min and Max"
            "SchemaXml"             = "<Field Type='Currency' LCID='' Name='' DisplayName='' StaticName='' Percentage='FALSE' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' Decimals='' Min='' Max='' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::DateTime
            "FieldTypeDescription"  = "Date and Time"
            "FieldTypeOptions"      = "DateOnly | Display Format (Standard)"
            "SchemaXml"             = "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateOnly' FriendlyDisplayFormat='Disabled' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::DateTime
            "FieldTypeDescription"  = "Date and Time"
            "FieldTypeOptions"      = "DateTime | Display Format (Standard)"
            "SchemaXml"             = "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateTime' FriendlyDisplayFormat='Disabled' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::DateTime
            "FieldTypeDescription"  = "Date and Time"
            "FieldTypeOptions"      = "DateOnly | Display Format (Friendly)"
            "SchemaXml"             = "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateOnly' FriendlyDisplayFormat='Relative' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::DateTime
            "FieldTypeDescription"  = "Date and Time"
            "FieldTypeOptions"      = "DateTime | Display Format (Friendly)"
            "SchemaXml"             = "<Field Type='DateTime' Name='' StaticName='' DisplayName='' Format='DateTime' FriendlyDisplayFormat='Relative' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Lookup
            "FieldTypeDescription"  = "Lookup (information already on this site)"
            "FieldTypeOptions"      = ""
            "SchemaXml"             = "<Field Type='Lookup' Name='' StaticName='' DisplayName='' List='{876fa85c-ecd5-4654-80ec-e0412ce96f07}' ShowField='Title' RelationshipDeleteBehavior='None' Required='FALSE' EnforceUniqueValues='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Lookup
            "FieldTypeDescription"  = "Lookup (information already on this site)"
            "FieldTypeOptions"      = "Allow multiple values"
            "SchemaXml"             = "<Field Type='LookupMulti' Name='' StaticName='' DisplayName='' List='{876fa85c-ecd5-4654-80ec-e0412ce96f07}' ShowField='Title' Mult='TRUE' RelationshipDeleteBehavior='None' Required='FALSE' EnforceUniqueValues='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::Boolean
            "FieldTypeDescription"  = "Yes/No (check box)"
            "FieldTypeOptions"      = ""
            "SchemaXml"             = "<Field Type='Boolean' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' Indexed='FALSE'><Default>1</Default></Field>"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::User
            "FieldTypeDescription"  = "Person or Group"
            "FieldTypeOptions"      = "PeopleOnly"
            "SchemaXml"             = "<Field Type='User' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleOnly' UserSelectionScope='0' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::User
            "FieldTypeDescription"  = "Person or Group"
            "FieldTypeOptions"      = "PeopleAndGroups"
            "SchemaXml"             = "<Field Type='User' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleAndGroups' UserSelectionScope='0' Group='' Mult='FALSE' />"
        },
        @{
            "FieldType"             = [Microsoft.SharePoint.Client.FieldType]::User
            "FieldTypeDescription"  = "Person or Group"
            "FieldTypeOptions"      = "PeopleOnly | Multiple"
            "SchemaXml"             = "<Field Type='UserMulti' List='UserInfo' Name='' StaticName='' DisplayName='' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleOnly' UserSelectionScope='3' Group='' Mult='TRUE' />"
        }

    )
}


################################################################################################################################################################

<#
@{
    "FieldType"             = ""
    "FieldTypeDescription"  = ""
    "FieldTypeOptions"      = ""
    "SchemaXml"             = ""
}
#>

################################################################################################################################################################
