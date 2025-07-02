<#
  Name: Liam Escusa
  Student ID: 011950351
#>

# Load required SQL Server module
if (Get-Module -Name sqlps) {
    Remove-Module -Name sqlps
}
Import-Module -Name SqlServer -ErrorAction Stop

# Define variables
$sqlInstance = ".\SQLEXPRESS"
$dbName = "ClientDB"
$tableName = "Client_A_Contacts"
$csvPath = ".\NewClientData.csv"

try {
    # D1: Check if the database exists and drop it if it does
    $db = Get-SqlDatabase -ServerInstance $sqlInstance -Name $dbName -ErrorAction SilentlyContinue
    if ($db) {
        Write-Host "Database '$dbName' exists. Dropping..." -ForegroundColor Yellow
        $dropQuery = @"
ALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
DROP DATABASE [$dbName];
"@
        Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $dropQuery -ErrorAction Stop
        Write-Host "Database '$dbName' has been dropped." -ForegroundColor Red
    } else {
        Write-Host "Database '$dbName' does not exist. Proceeding to create it." -ForegroundColor Green
    }

    # D2: Create the new database
    Invoke-Sqlcmd -ServerInstance $sqlInstance -Query "CREATE DATABASE [$dbName];" -ErrorAction Stop
    Write-Host "Database '$dbName' has been created." -ForegroundColor Cyan

    # D3: Create the Client_A_Contacts table
    $createTableQuery = @"
USE [$dbName];
CREATE TABLE dbo.$tableName (
    FirstName NVARCHAR(50),
    LastName NVARCHAR(50),
    City NVARCHAR(50),
    County NVARCHAR(50),
    Zip NVARCHAR(15),
    OfficePhone NVARCHAR(20),
    MobilePhone NVARCHAR(20)
);
"@
    Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $createTableQuery -ErrorAction Stop
    Write-Host "Table '$tableName' has been created." -ForegroundColor Cyan

    # D4: Import data from CSV into the table
    $csvData = Import-Csv -Path $csvPath
    foreach ($row in $csvData) {
        $insertQuery = @"
INSERT INTO [$dbName].dbo.$tableName (FirstName, LastName, City, County, Zip, OfficePhone, MobilePhone)
VALUES (N'$($row.first_name)', N'$($row.last_name)', N'$($row.city)', N'$($row.county)', N'$($row.zip)', N'$($row.officePhone)', N'$($row.mobilePhone)');
"@
        Invoke-Sqlcmd -ServerInstance $sqlInstance -Query $insertQuery -ErrorAction Stop
    }
    Write-Host "Data from 'NewClientData.csv' has been imported into '$tableName'." -ForegroundColor Cyan

    # D5: Output the results into SqlResults.txt
    Invoke-Sqlcmd -Database $dbName -ServerInstance $sqlInstance -Query "SELECT * FROM dbo.$tableName" > .\SqlResults.txt
    Write-Host "SQL results saved to SqlResults.txt." -ForegroundColor Green

} catch {
    # E: Catch and display any errors
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
