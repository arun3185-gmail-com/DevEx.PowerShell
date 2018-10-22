
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
$mmsGroupName         = "Site Collection - i180618.sharepoint.com-sites-ProdCat1"
$mmsNewTermSetName    = "Products"
$mmsNewParentTermName = "Products"
$TermStoreID          = $null
$TermSetID            = $null
################################################################################


################################################################################
# Excel data source
################################################################################
$xlFilePath = "D:\Arun\DevEx\Data\Superstore_Terms.xlsx"
$connString = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties='Excel 12.0 Xml;HDR=YES'; Data Source='" + $xlFilePath + "'"
$sqlQuery   = "SELECT Category,SubCategory FROM [Superstore_Terms$]"
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
$group = $termStore.Groups.GetByName($mmsGroupName)
$context.Load($group)
$context.ExecuteQuery()

#Create Term Set
$termSet = $group.CreateTermSet($mmsNewTermSetName, [System.Guid]::NewGuid(), 1033)
$context.Load($termSet)
$context.ExecuteQuery()

#Create Parent Term
$parentTerm = $termSet.CreateTerm($mmsNewParentTermName, 1033, [System.Guid]::NewGuid().toString())
$context.Load($parentTerm)
$context.ExecuteQuery()


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

    $currCat = $null
    
    while ($sqlReader.Read())
    {
        $catTerm = $sqlReader[0].Tostring()
        $subCatTerm = $sqlReader[1].Tostring()
        
        if (($currCat -eq $null) -or ($currCat.Name -ne $catTerm))
        {
            $currCat = $parentTerm.CreateTerm($catTerm, 1033, [System.Guid]::NewGuid().toString())
            $context.Load($currCat)
            $context.ExecuteQuery()
        }

        $termNewSubCat = $currCat.CreateTerm($subCatTerm, 1033, [System.Guid]::NewGuid().toString())
        $context.Load($termNewSubCat)
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
    Write-Host "Closing Excel..."
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

################################################################################
