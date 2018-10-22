
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"

################################################################################
# SharePoint connection
################################################################################
$admin     = "arun.b180618@I180618.onmicrosoft.com"
$password  = Read-Host "Enter Password for $($admin)" -AsSecureString
$siteUrl   = "https://i180618.sharepoint.com/sites/ProdCat1"
$listTitle = "Orders"
$listTitle_Customers = "Customers"
$listTitle_Products  = "Products"

$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
$context.Credentials = $credentials
################################################################################


################################################################################
# Excel data source
################################################################################
$xlFilePath = "D:\Arun\DevEx\Data\Superstore_Orders.xlsx"
$connString = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties='Excel 12.0 Xml;HDR=YES'; Data Source='" + $xlFilePath + "'"
$sqlQuery   = "SELECT OrderID, CustomerID, ProductID, OrderDate, ShipDate, ShipMode, Country, City, State, PostalCode, Region, Sales, Quantity, Discount, Profit FROM [Superstore_Orders$]"
################################################################################

$uploadBatchSize = 50
$batchCounter = 0

$arrCustomers = @()
$arrProducts = @()

################################################################################

Write-Host "Getting Lookup list data - $($listTitle_Customers)..."

$list_Customers = $context.Web.Lists.GetByTitle($listTitle_Customers)
$fields_Customers = $list_Customers.Fields
$context.Load($list_Customers)
$context.Load($fields_Customers)
$camlQry_Customers = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
$listItems_Customers = $list_Customers.GetItems($camlQry_Customers)
$context.Load($listItems_Customers)
$context.ExecuteQuery()
foreach ($listItem in $listItems_Customers)
{
    $newPSItem = New-Object PSObject
    $newPSItem | Add-Member NoteProperty -Name "ID" -Value $listItem["ID"].ToString()
    $newPSItem | Add-Member NoteProperty -Name "Title" -Value $listItem["Title"].ToString()
    $arrCustomers += $newPSItem
}

################################################################################

Write-Host "Getting Lookup list data - $($listTitle_Products)..."

$list_Products = $context.Web.Lists.GetByTitle($listTitle_Products)
$fields_Products = $list_Products.Fields
$context.Load($list_Products)
$context.Load($fields_Products)
$camlQry_Products = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
$listItems_Products = $list_Products.GetItems($camlQry_Products)
$context.Load($listItems_Products)
$context.ExecuteQuery()
foreach ($listItem in $listItems_Products)
{
    $newPSItem = New-Object PSObject
    $newPSItem | Add-Member NoteProperty -Name "ID" -Value $listItem["ID"].ToString()
    $newPSItem | Add-Member NoteProperty -Name "Title" -Value $listItem["Title"].ToString()
    $arrProducts += $newPSItem
}

################################################################################

$list = $context.Web.Lists.GetByTitle($listTitle)
$fields = $list.Fields
$context.Load($list)
$context.Load($fields)

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
        $flvCustomer = $null
        $flvProduct = $null
        $listItem = $list.AddItem($listItemInfo)
        $lkUpCustomer = ($arrCustomers | Where-Object -Property "Title" -eq $sqlReader[1].Tostring())[0]
        $lkUpProduct = ($arrProducts | Where-Object -Property "Title" -eq $sqlReader[2].Tostring())[0]
        
        $listItem["Title"] = $sqlReader[0].Tostring()
        $listItem["Customer"] = $lkUpCustomer.ID.ToString() + ";#" + $lkUpCustomer.Title
        $listItem["Product"] = $lkUpProduct.ID.ToString() + ";#" + $lkUpProduct.Title
        
        #$flvCustomer = [Microsoft.SharePoint.Client.FieldLookupValue]$listItem["Customer"]
        #if ($flvCustomer -eq $null) { $flvCustomer = New-Object Microsoft.SharePoint.Client.FieldLookupValue }
        #$flvCustomer.LookupId = [int]($arrCustomers | Where-Object -Property "Title" -eq $sqlReader[1].Tostring())[0].ID
        #$listItem["Customer"] = $flvCustomer
        
        #$flvProduct = [Microsoft.SharePoint.Client.FieldLookupValue]$listItem["Product"]
        #if ($flvProduct -eq $null) { $flvProduct = New-Object Microsoft.SharePoint.Client.FieldLookupValue }
        #$flvProduct.LookupId = [int]($arrProducts | Where-Object -Property "Title" -eq $sqlReader[2].Tostring())[0].ID
        #$listItem["Product"] = $flvProduct

        $listItem["OrderDate"] = $sqlReader[3].Tostring()
        $listItem["ShipDate"] = $sqlReader[4].Tostring()
        $listItem["ShipMode"] = $sqlReader[5].Tostring()
        $listItem["Country"] = $sqlReader[6].Tostring()
        $listItem["City"] = $sqlReader[7].Tostring()
        $listItem["State"] = $sqlReader[8].Tostring()
        $listItem["PostalCode"] = $sqlReader[9].Tostring()
        $listItem["Region"] = $sqlReader[10].Tostring()
        $listItem["Sales"] = $sqlReader[11].Tostring()
        $listItem["Quantity"] = $sqlReader[12].Tostring()
        $listItem["Discount"] = $sqlReader[13].Tostring()
        $listItem["Profit"] = $sqlReader[14].Tostring()
        
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
    $arrCustomers = $null
    $arrProducts = $null

    if ($sqlReader -ne $null)
    {
        $sqlReader.Close()
        $sqlReader.Dispose()
    }
    if ($sqlCommand -ne $null)
    {
        $sqlCommand.Dispose()
    }
    if ($sqlConnection -ne $null)
    {
        $sqlConnection.Close()
        $sqlConnection.Dispose()
    }
}

Write-Host ""
Write-Host "Done!"

################################################################################
