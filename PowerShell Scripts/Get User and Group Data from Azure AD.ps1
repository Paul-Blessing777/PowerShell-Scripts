# Connect to Azure AD
#Connect-AzAccount
$clientId       = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"      
$clientSecret   = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"  
$tenantId       = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"       
$subscriptionId = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"   

$securePassword = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
$credential     = New-Object System.Management.Automation.PSCredential($clientId, $securePassword)

Connect-AzAccount -ServicePrincipal -Credential $credential -TenantId $tenantId -SubscriptionId $subscriptionId -WarningAction SilentlyContinue

# Set FilePath Variable
$csvFilePath = "\\Scripts\PowerShell\Azure AD Results\AllGroupsAndMembers.csv"
$excelFilePath = "\\Scripts\PowerShell\Azure AD Results\AllGroupsAndMembers.xlsx"

# Get all groups
$groups = Get-AzADGroup

# Create an array to store results 
$results = @()

# Process each group
foreach ($group in $groups) {
    try {
        # Get all members of the group
        $allMembers = Get-AzADGroupMember -GroupObjectId $group.Id

        # Filter members based on your conditions
        $filteredMembers = $allMembers | Where-Object {
            $_.AccountEnabled -eq $true
        }

        # Join the filtered members' UserPrincipalName with a semicolon
        $members = $filteredMembers.DisplayName -join ';'

      # Get all owners of the group
        $allOwners = Get-AzADGroupOwner -GroupId $group.Id

        # Filter owners based on AccountEnabled property
        $filteredOwners = $allOwners | Where-Object {
            $_.UserPrincipalName -ne $null
        }

        # Join the filtered owners' UserPrincipalName with a semicolon
        $owners = $filteredOwners.DisplayName -join ';'

        # Check for the groupType attribute to determine the group type
        if ($group.GroupType -contains "Unified") {
            $groupType = '365 Group'
        } elseif ($group.SecurityEnabled -eq $true) {
            $groupType = 'Security Group'
        } else {
            $groupType = 'Distribution List'
        }

        # Check if the group exists and meets the criteria
        $syncOnPrem = $group.OnPremisesSyncEnabled -eq $true

        # Create a custom PowerShell object to represent information about an Azure AD group.
        $result = [PSCustomObject]@{
            GroupName           = $group.DisplayName
            GroupType           = $groupType
            SyncedFromOnPremises = $syncOnPrem
            Members             = $members
            Owners              = $owners
            Email               = $group.Mail
            ObjectID            = $group.Id
        }

        # Add the result to the array
        $results += $result

    } catch {
        Write-Host "Error processing group $($group.DisplayName): $_"
    }
}

# Check if there are results before exporting to CSV
if ($results.Count -gt 0) {
    # Sort the results by GroupName
    $results = $results | Sort-Object GroupName

    # Export the CSV file
    $results | Export-Csv -Path $csvFilePath -NoTypeInformation -Force

    # Convert CSV to Excel using ImportExcel module and sort by GroupName
    Import-Csv $csvFilePath | Sort-Object GroupName | Export-Excel -Path $excelFilePath -AutoSize -AutoFilter -FreezeTopRow -WorksheetName "AllGroupsAndMembers"
} else {
    Write-Host "No data to export."
}

Disconnect-AzAccount