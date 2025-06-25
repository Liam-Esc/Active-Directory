<# 
    Liam Escusa 
    Student ID:  011950351
#>

# Check for existence of the Active Directory OU named "finance"

Import-Module ActiveDirectory

try {
    # Define OU distinguished name
    $ouDN = "OU=Finance,DC=consultingfirm,DC=com"

    # Check if the "Finance" OU exists
    $ou = Get-ADOrganizationalUnit -Filter { Name -eq "Finance" } -ErrorAction Stop

    if ($ou) {
        Write-Host "OU 'Finance' exists. Deleting it..."
        Remove-ADOrganizationalUnit -Identity $ou.DistinguishedName -Recursive -Confirm:$false
        Write-Host "OU 'Finance' deleted."
    }
} catch {
    Write-Host "OU 'Finance' does not exist. Creating it..."
}

try {
    # Create the OU
    New-ADOrganizationalUnit -Name "Finance" -Path "DC=consultingfirm,DC=com"
    Write-Host "OU 'Finance' created."
} catch {
    Write-Host "Error creating OU: $_"
}

try {
    # Import users from CSV and create user accounts
    $csvPath = ".\financePersonnel.csv"
    $users = Import-Csv -Path $csvPath

    foreach ($user in $users) {
        $firstName = $user.'First Name'
        $lastName = $user.'Last Name'
        $displayName = "$firstName $lastName"

        New-ADUser `
            -Name $displayName `
            -GivenName $firstName `
            -Surname $lastName `
            -DisplayName $displayName `
            -PostalCode $user.'Postal Code' `
            -OfficePhone $user.'Office Phone' `
            -MobilePhone $user.'Mobile Phone' `
            -Path "OU=Finance,DC=consultingfirm,DC=com" `
            -AccountPassword (ConvertTo-SecureString "P@ssw0rd123!" -AsPlainText -Force) `
            -Enabled $true
    }

    Write-Host "All users imported successfully into the 'Finance' OU."
} catch {
    Write-Host "Error importing users: $_"
}

# Output results to AdResults.txt
Get-ADUser -Filter * -SearchBase "OU=Finance,DC=consultingfirm,DC=com" -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt
Write-Host "Exported user data to AdResults.txt."
