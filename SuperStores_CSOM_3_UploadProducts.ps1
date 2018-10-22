
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Taxonomy.dll"

################################################################################
# SharePoint connection
################################################################################
$admin     = "arun.b180618@I180618.onmicrosoft.com"
$password  = Read-Host "Enter Password for $($admin)" -AsSecureString
$siteUrl   = "https://i180618.sharepoint.com/sites/ProdCat1"
$listTitle = "Products"

$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
$context.Credentials = $credentials
################################################################################


################################################################################
# Excel data source
################################################################################
$xlFilePath = "D:\Arun\DevEx\Data\Superstore_Products.xlsx"
$connString = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties='Excel 12.0 Xml;HDR=YES'; Data Source='" + $xlFilePath + "'"
$sqlQuery   = "SELECT ProductID, SubCategory, ProductName FROM [Superstore_Products$]"
################################################################################


################################################################################
# MMS Details
################################################################################
$groupName   = "Site Collection - i180618.sharepoint.com-sites-ProdCat1"
$termSetName = "Products"
################################################################################

$uploadBatchSize = 100
$batchCounter = 0

################################################################################


################################################################################

Write-Host "Getting MMS details..."

#Bind to MMS, Term Store
$mms = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($context)
$context.Load($mms)
$termStore = $mms.GetDefaultSiteCollectionTermStore()
$context.Load($termStore)
$context.Load($termStore.Groups)
$context.ExecuteQuery()

#Bind to Group
$group = $termStore.Groups.GetByName($groupName)
$context.Load($group)
$context.ExecuteQuery()

#Bind to Term Set
$termSet = $group.TermSets.GetByName($termSetName)
$context.Load($termSet)
$context.ExecuteQuery()

$terms = $termSet.GetAllTerms()
$context.Load($terms)
$context.ExecuteQuery()

$list = $context.Web.Lists.GetByTitle($listTitle)
$fields = $list.Fields
$field_Cat = $fields.GetByInternalNameOrTitle("Category")
$context.Load($fields)
$context.Load($field_Cat)
$context.Load($list)

$txField_Cat = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($context, $field_Cat)

################################################################################

Write-Host "Reading Excel..."

$sqlConnection = $null
$sqlCommand = $null
$sqlReader = $null

try
{
    $sqlConnection = New-Object System.Data.OleDb.OleDbConnection($connString)
    $sqlCommand = New-Object System.Data.OleDb.OleDbCommand($sqlQuery)
    $sqlCommand.Connection = $sqlConnection
    $sqlConnection.Open()
    $sqlReader = $sqlCommand.ExecuteReader()
    
    while ($sqlReader.Read())
    {
        $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $listItem = $list.AddItem($listItemInfo)
        
        $listItem["Title"] = $sqlReader[0].Tostring()

        $catTerm = $terms | ? { $_.Name -eq $sqlReader[1].Tostring() }
        $catTermV = New-Object Microsoft.SharePoint.Client.Taxonomy.TaxonomyFieldValue
        $catTermV.Label = $catTerm.Name
        $catTermV.TermGuid = $catTerm.Id
        $catTermV.WssId = -1
        $txField_Cat.SetFieldValueByValue($listItem, $catTermV)
    
        $listItem["ProductName"] = $sqlReader[2].Tostring()

        $listItem.Update()
        $context.Load($listItem)
        $batchCounter++
        Write-Host "Loaded for - $($sqlReader[0].Tostring())"
        
        if ($batchCounter -ge $uploadBatchSize)
        {
            Write-Host ""
            Write-Host "Executing query..."
            $batchCounter = 0
            $context.ExecuteQuery()
        }        
    }
    
    Write-Host ""
    Write-Host "Executing query..."
    $context.ExecuteQuery()
}
catch
{
    Write-Host "$($_.Exception.ToString())" -ForegroundColor Red
}
finally
{
    if ($sqlReader -ne $null) { $sqlReader.Close() }
    if ($sqlConnection -ne $null) { $sqlConnection.Close() }
}

Write-Host ""
Write-Host "Done!"

################################################################################
