<#
    Liam Escusa
    Student ID: 011950351
#>

# Set variables
$dbName = "ClientDB"
$serverInstance = ".\SQLEXPRESS"
$tableName = "Client_A_Contacts"
$csvPath = ".\NewClientData.csv"

try {
    # Import required module
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlServer

    # Check if the database exists and delete if it does
    if (Get-SqlDatabase -Name $dbName -ServerInstance $serverInstance -ErrorAction SilentlyContinue) {
        $dropQuery = @"
ALTER DATABASE [$dbName] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
DROP DATABASE [$dbName];
"@
        Invoke-Sqlcmd -ServerInstance $serverInstance -Database "master" -Query $dropQuery
        Write-Host "Database '$dbName' existed and was deleted." -ForegroundColor Yellow
    } else {
        Write-Host "Database '$dbName' does not exist." -ForegroundColor Cyan
    }

    # Create the database
    Invoke-Sqlcmd -ServerInstance $serverInstance -Database "master" -Query "CREATE DATABASE [$dbName]"
    Write-Host "Database '$dbName' created successfully." -ForegroundColor Green

    # Create the table
    $createTableQuery = @"
CREATE TABLE [$tableName] (
    first_name NVARCHAR(50),
    last_name NVARCHAR(50),
    city NVARCHAR(50),
    county NVARCHAR(50),
    zip NVARCHAR(10),
    officePhone NVARCHAR(20),
    mobilePhone NVARCHAR(20)
)
"@
    Invoke-Sqlcmd -ServerInstance $serverInstance -Database $dbName -Query $createTableQuery
    Write-Host "Table '$tableName' created successfully." -ForegroundColor Green

    # Import data from CSV
    $rows = Import-Csv -Path $csvPath
    foreach ($row in $rows) {
        $insertQuery = @"
INSERT INTO [$tableName] (first_name, last_name, city, county, zip, officePhone, mobilePhone)
VALUES (
    N'$($row.first_name)',
    N'$($row.last_name)',
    N'$($row.city)',
    N'$($row.county)',
    N'$($row.zip)',
    N'$($row.officePhone)',
    N'$($row.mobilePhone)'
)
"@
        Invoke-Sqlcmd -ServerInstance $serverInstance -Database $dbName -Query $insertQuery
    }
    Write-Host "All records imported successfully from '$csvPath'." -ForegroundColor Green

    # Output the results to a text file
    Invoke-Sqlcmd -Database $dbName -ServerInstance $serverInstance -Query "SELECT * FROM dbo.$tableName" > .\SqlResults.txt
    Write-Host "SqlResults.txt has been created." -ForegroundColor Green

} catch {
    Write-Host "An error occurred!" -ForegroundColor Red
    Write-Host $_
}
