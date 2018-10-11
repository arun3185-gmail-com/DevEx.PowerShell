
[string] $script:xlFilePath = $null
[string] $script:connString = $null

function New-XlOleDb
{
    [CmdletBinding()]
    Param
    (        
        [Parameter(Mandatory = $true)]
        [string] $FilePath,

        [Parameter(Mandatory = $false)]
        [bool] $Overwrite = $false
    )

    # Begin
    Begin
    {
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # Process
    Process
    {
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        $script:xlFilePath = $FilePath
        $script:connString = "Provider=Microsoft.Jet.OLEDB.4.0; Extended Properties='Excel 8.0;IMEX=1;HDR=YES'; Data Source='" + $script:xlFilePath + "'"
        if ($Overwrite -and (Test-Path -Path $script:xlFilePath))
        {
            #Remove-Item -Path $script:xlFilePath
        }

        [string] $sqlQuery = $null
        [System.Data.OleDb.OleDbConnection] $sqlConnection = $null
        [System.Data.OleDb.OleDbCommand] $sqlCommand = $null

        try
        {
            $sqlQuery = "Create table [Sheet1$] (RecID INTEGER, Msg TEXT(10))"
            $sqlQuery = "Create table [Sheet1] ()"

            $sqlConnection = New-Object System.Data.OleDb.OleDbConnection($script:connString)
            $sqlCommand = New-Object System.Data.OleDb.OleDbCommand($sqlQuery)
            $sqlCommand.Connection = $sqlConnection
            $sqlConnection.Open()

            $sqlCommand.ExecuteNonQuery()
            $sqlConnection.Close()
        }
        catch
        {
            throw
        }
        finally
        {
            if ($sqlCommand -ne $null) { $sqlCommand.Dispose() }
            if ($sqlConnection -ne $null) { $sqlConnection.Dispose() }
        }


        Write-Verbose -Message "Leaving the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # End
    End
    {
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
}


function Get-XlOleDbAllRecords
{
    [CmdletBinding()]
    Param ($Command)

    # Begin
    Begin
    {
        Write-Verbose -Message "Entering the BEGIN block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # Process
    Process
    {
        Write-Verbose -Message "Entering the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        [string] $sqlQuery = $null
        [System.Data.OleDb.OleDbConnection] $sqlConnection = $null
        [System.Data.OleDb.OleDbCommand] $sqlCommand = $null

        try
        {
            $sqlQuery = "Select * from [Sheet1$]"
            $sqlConnection = New-Object System.Data.OleDb.OleDbConnection($script:connString)
            $sqlCommand = New-Object System.Data.OleDb.OleDbCommand($sqlQuery)
            $sqlCommand.Connection = $sqlConnection
            $sqlConnection.Open()

            $sqlCommand.ExecuteNonQuery()
            $sqlConnection.Close()
        }
        catch
        {
            throw
        }
        finally
        {
            if ($sqlCommand -ne $null) { $sqlCommand.Dispose() }
            if ($sqlConnection -ne $null) { $sqlConnection.Dispose() }
        }


        Write-Verbose -Message "Leaving the PROCESS block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }

    # End
    End
    {
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
    }
}

#New,Write
function Add-XlOleDbRecord
{
}

#Find,Search,Read
function Get-XlOleDbRecord
{
}

#Edit,Update,Write
function Set-XlOleDbRecord
{
}

#Clear
function Remove-XlOleDbRecord
{
}



#Save,Open,Close,Connect,Disconnect
function Use-XlOleDbRecord
{
}

<#
$xlFilePath = "D:\Arun\DevEx\Data\TestExcel.xlsx"
$connString = "Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties='Excel 12.0 Xml;HDR=YES'; Data Source='" + $xlFilePath + "'"
$sqlQuery = "CREATE TABLE tblCustomers (CustomerID INTEGER, FullName TEXT(50), Email TEXT(50))"

$sqlConnection = $null
$sqlCommand = $null

$sqlConnection = New-Object System.Data.OleDb.OleDbConnection($connString)
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($sqlQuery)
$sqlCommand.Connection = $sqlConnection
$sqlConnection.Open()

$sqlCommand.ExecuteNonQuery()
$sqlConnection.Close()

if ($sqlCommand -ne $null) { $sqlCommand.Dispose() }
if ($sqlConnection -ne $null) { $sqlConnection.Dispose() }
#>

<#
$strFileName = "D:\Arun\DevEx\PS\OleDb_WriteEx1.xlsx"
$strProvider = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSource = "Data Source ='" + $strfilename + "'"
$strExtend = "Extended Properties='Excel 12.0 Xml;HDR=YES'"
$strQuery = "Select Name,Type,Description,Path from [Sheet1$]"

$objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
$sqlCommand.Connection = $objConn
$objConn.Open()
$DataReader = $sqlCommand.ExecuteReader()

while ($DataReader.Read())
{
    Write-Host "$($DataReader[0].Tostring())|$($DataReader[1].Tostring())|$($DataReader[2].Tostring())|$($DataReader[3].Tostring())"
}

$DataReader.Close()
$objConn.Close()
#>

<#
$strFileName = "D:\Arun\DevEx\PS\OleDb_WriteEx1.xlsx"
$strProvider = "Provider=Microsoft.ACE.OLEDB.12.0"
$strDataSource = "Data Source ='" + $strfilename + "'"
$strExtend = "Extended Properties='Excel 12.0 Xml;HDR=YES'"
$strQuery = "Insert into [Sheet1$] (Name,Path,Description,Type) Values (?,?,?,?)"

$objConn = New-Object System.Data.OleDb.OleDbConnection("$strProvider;$strDataSource;$strExtend")
$sqlCommand = New-Object System.Data.OleDb.OleDbCommand($strQuery)
$sqlCommand.Connection = $objConn

$NameParam = $sqlCommand.Parameters.Add("Name","VarChar",80)
$PathParam = $sqlCommand.Parameters.Add("Path","VarChar",80)
$DescriptionParam = $sqlCommand.Parameters.Add("Description","VarChar",80)
$TypeParam = $sqlCommand.Parameters.Add("Type","UnsignedInt",16)


$objConn.Open()

$NameParam.Value = "Name1"
$PathParam.Value = "Path1"
$DescriptionParam.Value = "Desc1"
$TypeParam.Value = 101
$returnValue = $sqlCommand.ExecuteNonQuery()
Write-Host $returnValue

$NameParam.Value = "Name2"
$PathParam.Value = "Path2"
$DescriptionParam.Value = "Desc2"
$TypeParam.Value = 102
$returnValue = $sqlCommand.ExecuteNonQuery()
Write-Host $returnValue

$NameParam.Value = "Name3"
$PathParam.Value = "Path3"
$DescriptionParam.Value = "Desc3"
$TypeParam.Value = 103
$returnValue = $sqlCommand.ExecuteNonQuery()
Write-Host $returnValue

$objConn.Close()
#>
