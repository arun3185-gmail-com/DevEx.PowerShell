
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.dll"
Import-Module "D:\Arun\Git\DevEx.References\NuGet\microsoft.sharepointonline.csom.16.1.7414.1200\lib\net45\Microsoft.SharePoint.Client.Runtime.dll"

################################################################################
# SharePoint connection
################################################################################
$admin     = "arun.b180618@I180618.onmicrosoft.com"
$password  = Read-Host "Enter Password for $($admin)" -AsSecureString
$siteUrl   = "https://i180618.sharepoint.com/sites/ProdCat1"
$listTitle = "Customers"

$context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($admin , $password)
$context.Credentials = $credentials
################################################################################


################################################################################
# Excel data source
################################################################################
$xlFilePath = "D:\Arun\DevEx\Data\Superstore_Customers.xlsx"
$connString = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties='Excel 12.0 Xml;HDR=YES'; Data Source='" + $xlFilePath + "'"
$sqlQuery   = "SELECT CustomerID, CustomerName, Segment FROM [Superstore_Customers$]"
################################################################################

$uploadBatchSize = 100
$batchCounter = 0

################################################################################

$list = $context.Web.Lists.GetByTitle($listTitle)
$fields = $list.Fields
$context.Load($fields)
$context.Load($list)

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
        $listItem["CustomerName"] = $sqlReader[1].Tostring()
        $listItem["Segment"] = $sqlReader[2].Tostring()
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
