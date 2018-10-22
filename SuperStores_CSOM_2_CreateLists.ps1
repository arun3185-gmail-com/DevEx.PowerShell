
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Taxonomy.dll"

################################################################################
# SharePoint connection
################################################################################
$admin    = "arun.b180618@I180618.onmicrosoft.com"
$password = Read-Host "Enter Password for $($admin)" -AsSecureString
$siteUrl  = "https://i180618.sharepoint.com/sites/ProdCat1"

$context     = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
$context.Credentials = $credentials
################################################################################


################################################################################
# MMS Details
################################################################################
$mmsGroupName   = "Site Collection - i180618.sharepoint.com-sites-ProdCat1"
$mmsTermSetName = "Products"
$TermStoreID    = $null
$TermSetID      = $null
################################################################################


################################################################################################################################################################
# Function Get-MMS-Details
################################################################################################################################################################

Function Get-MMS-Details
{
    Param ($GroupName, $TermSetName)

    Write-Host "Getting MMS details..."
    #Bind to MMS, Term Store
    $mms = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
    $context.Load($mms)
    $termStore = $mms.GetDefaultSiteCollectionTermStore()
    $context.Load($termStore)
    $context.ExecuteQuery()

    #Bind to Group
    $group = $termStore.Groups.GetByName($GroupName)
    $context.Load($group)
    $context.ExecuteQuery()

    #Bind to Term Set
    $termSet = $group.TermSets.GetByName($TermSetName)
    $context.Load($termSet)
    $context.ExecuteQuery()

    Return New-Object PSObject -Property @{TermStoreID = $termStore.Id; TermSetID = $termSet.Id}
}

################################################################################################################################################################


################################################################################################################################################################
# Function Ensure-List
################################################################################################################################################################

Function Ensure-List()
{
    Param ($ListTitle, $FieldSchemaCollection)

    $lists = $context.Web.Lists
    $context.Load($lists)
    $context.ExecuteQuery()

    $list = $lists | Where { $_.Title -eq $ListTitle }
    $fieldColl = $null

    # Ensures list
    if($list)
    {
        Write-Host "$($ListTitle) - Exists"
        $list = $context.Web.Lists.GetByTitle($ListTitle)
        $fieldColl = $list.Fields
        $context.Load($fieldColl)
    }
    else
    {
        Write-Host "$($ListTitle) - Creating..."
        $listInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
        $listInfo.QuickLaunchOption = [Microsoft.SharePoint.Client.QuickLaunchOptions]::On
        $listInfo.Title = $ListTitle
        $listInfo.TemplateType = [int] [Microsoft.SharePoint.Client.ListTemplateType]::GenericList
        $list = $context.Web.Lists.Add($listInfo)
        $fieldColl = $list.Fields
        $context.Load($fieldColl)
    }
    
    $context.Load($list)
    $context.ExecuteQuery()

    # Ensures Fields
    foreach ($fldMetaData in $FieldSchemaCollection)
    {
        $fld = $fieldColl | Where { $_.InternalName -eq $fldMetaData.InternalName }

        if ($fld)
        {
            if ($fldMetaData.SchemaXml)
            {
                $fld.Title = (Select-Xml -Content $fldMetaData.SchemaXml -XPath "Field/@DisplayName")[0].Node.Value
                $fld.Update()
                $context.Load($fld)
            }
        }
        else
        {
            if ($fldMetaData.SchemaXml)
            {
                $fld = $list.Fields.AddFieldAsXml($fldMetaData.SchemaXml, $true, [Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldInternalNameHint)
                $list.Update()
                $context.Load($fld)
                
                if ((Select-Xml -Content $fldMetaData.SchemaXml -XPath "Field/@Type")[0].Node.Value -eq "TaxonomyFieldType")
                {
                    $txField = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($context, $fld)
                    $txField.SspId = $TermStoreID
                    $txField.TermSetId = $TermSetID
                    $txField.TargetTemplate = [System.String]::Empty
                    $txField.AnchorId = [System.Guid]::Empty
                    $txField.Update()
                }
                
            }
        }
    }

    if ($context.HasPendingRequest)
    {
        $context.ExecuteQuery()
    }

    Return $list.Id
}

################################################################################################################################################################



################################################################################################################################################################
# MMS Details
################################################################################################################################################################

$retValue = Get-MMS-Details -GroupName $mmsGroupName -TermSetName $mmsTermSetName

$TermStoreID = $retValue.TermStoreID
$TermSetID = $retValue.TermSetID

Write-Host "Term Store ID - $($TermStoreID)"
Write-Host "Term Set   ID - $($TermSetID)"
Write-Host ""

################################################################################################################################################################



################################################################################################################################################################
# List Details
################################################################################################################################################################

$list_Customer_Title = "Customers"
$list_Customer_SchemaCollection = @(
    New-Object PSObject -Property @{InternalName = "Title";        SchemaXml = "<Field Type='Text' Name='Title' StaticName='Title' DisplayName='CustomerID' />" }
    New-Object PSObject -Property @{InternalName = "CustomerName"; SchemaXml = "<Field Type='Text' Name='CustomerName' StaticName='CustomerName' DisplayName='Customer Name' />" }
    New-Object PSObject -Property @{InternalName = "Segment";      SchemaXml = "<Field Type='Choice' Name='Segment' StaticName='Segment' DisplayName='Segment' Format='RadioButtons'> <Default>Consumer</Default> <CHOICES><CHOICE>Consumer</CHOICE><CHOICE>Home Office</CHOICE><CHOICE>Corporate</CHOICE></CHOICES> </Field>" }
)

$list_Customer_ID = Ensure-List -ListTitle $list_Customer_Title -FieldSchemaCollection $list_Customer_SchemaCollection

Write-Host "List: $($list_Customer_Title), ID: $($list_Customer_ID)"
Write-Host ""

################################################################################################################################################################

$list_Product_Title = "Products"
$list_Product_SchemaCollection = @(
    New-Object PSObject -Property @{InternalName = "Title";        SchemaXml = "<Field Type='Text' Name='Title' StaticName='Title' DisplayName='ProductID' />" }
    New-Object PSObject -Property @{InternalName = "Category";     SchemaXml = "<Field Type='TaxonomyFieldType' Name='Category' StaticName='Category' DisplayName='Category' />"; GroupName = "Site Collection - i180618.sharepoint.com-sites-SuperStores"; TermSetName = "Products" }
    New-Object PSObject -Property @{InternalName = "ProductName";  SchemaXml = "<Field Type='Text' Name='ProductName' StaticName='ProductName' DisplayName='Product Name' />" }
)

$list_Product_ID = Ensure-List -ListTitle $list_Product_Title  -FieldSchemaCollection $list_Product_SchemaCollection

Write-Host "List: $($list_Product_Title), ID: $($list_Product_ID)"
Write-Host ""

################################################################################################################################################################

$list_Order_Title = "Orders"
$list_Order_SchemaCollection = @(
    New-Object PSObject -Property @{InternalName = "Title";      SchemaXml = "<Field Type='Text' Name='Title' StaticName='Title' DisplayName='OrderID' />" }
    New-Object PSObject -Property @{InternalName = "Customer";   SchemaXml = "<Field Type='Lookup' Name='Customer' StaticName='Customer' DisplayName='Customer' List='$($list_Customer_ID)' ShowField='Title' />" }
    New-Object PSObject -Property @{InternalName = "Product";    SchemaXml = "<Field Type='Lookup' Name='Product' StaticName='Product' DisplayName='Product' List='$($list_Product_ID)' ShowField='Title' />" }
    New-Object PSObject -Property @{InternalName = "OrderDate";  SchemaXml = "<Field Type='DateTime' Name='OrderDate' StaticName='OrderDate' DisplayName='Order Date' Format='DateOnly' />" }
    New-Object PSObject -Property @{InternalName = "ShipDate";   SchemaXml = "<Field Type='DateTime' Name='ShipDate' StaticName='ShipDate' DisplayName='Ship Date' Format='DateOnly' />" }
    New-Object PSObject -Property @{InternalName = "ShipMode";   SchemaXml = "<Field Type='Choice' Name='ShipMode' StaticName='ShipMode' DisplayName='Ship Mode' Format='RadioButtons'> <Default>Standard Class</Default> <CHOICES><CHOICE>Standard Class</CHOICE><CHOICE>First Class</CHOICE><CHOICE>Second Class</CHOICE><CHOICE>Same Day</CHOICE></CHOICES> </Field>" }
    New-Object PSObject -Property @{InternalName = "Country";    SchemaXml = "<Field Type='Text' Name='Country' StaticName='Country' DisplayName='Country' />" }
    New-Object PSObject -Property @{InternalName = "City";       SchemaXml = "<Field Type='Text' Name='City' StaticName='City' DisplayName='City' />" }
    New-Object PSObject -Property @{InternalName = "State";      SchemaXml = "<Field Type='Text' Name='State' StaticName='State' DisplayName='State' />" }
    New-Object PSObject -Property @{InternalName = "PostalCode"; SchemaXml = "<Field Type='Text' Name='PostalCode' StaticName='PostalCode' DisplayName='Postal Code' />" }
    New-Object PSObject -Property @{InternalName = "Region";     SchemaXml = "<Field Type='Text' Name='Region' StaticName='Region' DisplayName='Region' />" }
    New-Object PSObject -Property @{InternalName = "Sales";      SchemaXml = "<Field Type='Currency' Name='Sales' StaticName='Sales' DisplayName='Sales' LCID='1033' />" }
    New-Object PSObject -Property @{InternalName = "Quantity";   SchemaXml = "<Field Type='Number' Name='Quantity' StaticName='Quantity' DisplayName='Quantity' Decimals='0' />" }
    New-Object PSObject -Property @{InternalName = "Discount";   SchemaXml = "<Field Type='Number' Name='Discount' StaticName='Discount' DisplayName='Discount' />" }
    New-Object PSObject -Property @{InternalName = "Profit";     SchemaXml = "<Field Type='Currency' Name='Profit' StaticName='Profit' DisplayName='Profit' LCID='1033' />" }
)

$list_Order_ID = Ensure-List -ListTitle $list_Order_Title    -FieldSchemaCollection $list_Order_SchemaCollection

Write-Host "List: $($list_Order_Title), ID: $($list_Order_ID)"
Write-Host ""

################################################################################################################################################################

Write-Host ""
Write-Host "Done!"

################################################################################################################################################################
