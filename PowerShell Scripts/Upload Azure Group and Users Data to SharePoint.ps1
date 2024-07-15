# Import the SharePoint PnP PowerShell module
# Import-Module AzureAD -Force
Import-Module -Name PnP.PowerShell
Import-Module -Name ImportExcel # This module requires PowerShell 7 to run our machines by default have 5


# Connect to SharePoint
$siteUrl = "https://companysharepoint.com/"

# Service account credentials
$serviceAccountUsername = "SharePointAuto@company.com"
$serviceAccountPassword = ConvertTo-SecureString "XXXXXXXX" -AsPlainText -Force

# Connect to SharePoint using service account
Connect-PnPOnline -Url $siteUrl -Credentials (New-Object System.Management.Automation.PSCredential($serviceAccountUsername, $serviceAccountPassword))

# Path to Excel file
$excelFilePath = "\Scripts\PowerShell\Azure AD Results\AllGroupsAndMembers.xlsx"

# SharePoint list name
$listName = "Groups"

# Load Excel data
$excelData = Import-Excel -Path $excelFilePath

# Iterate through each row in the Excel file
foreach ($row in $excelData) {
    # Check if "GroupName" column is not empty before processing
    if ($row."GroupName" -and $row."ObjectID") {
        Write-Host "Processing row: $($row.'GroupName')"

        # Extract Email and ObjectId values
        $email = if ($row."Email") { $row."Email" } else { "" }

        # Fetch all items in the SharePoint list
        $allItems = Get-PnPListItem -List $listName

        # Check if an item with the same "ObjectID" and "GroupName" exists
        $existingItem = $allItems | Where-Object { $_["ObjectID"] -eq $row."ObjectID" -and $_["Title"] -eq $row."GroupName" }

        # Create a SharePoint list item: Title, GroupType, ObjectID, Owners, and Members and SyncedFromOnPremises
        $listItemProperties = @{
            "Title"      = $row."GroupName"
            "GroupType"  = $row."GroupType"
            "Email"      = $email
            "ObjectID"   = $row."ObjectID"
            "GroupContents" = $row."SyncedFromOnPremises" 
        }

        # Add Owners to the "Owner" column, even if it's empty
        if ($row.Owners -and $row.Owners -ne "") {
            $listItemProperties.Add("Owner", @($row.Owners -split ';'))
        }

        # Add Members to the "Users" column, even if it's empty
        if ($row.Members -and $row.Members -ne "") {
            $listItemProperties.Add("Users", @($row.Members -split ';'))
        }

        if ($existingItem) {
            # Update the existing item using the Set-PnPListItem cmdlet
            Write-Host "Updating existing item in SharePoint..."
            $existingItem | forEach-Object { Set-PnpListItem -List $listName -Identity $_.Id -Values $listItemProperties }
        } else {
            # Add the item to the SharePoint list using the internal name
            Write-Host "Adding new item to SharePoint..."
            Add-PnpListItem -List $listName -Value $listItemProperties
        }
    }
}

# Disconnect from SharePoint Online
Disconnect-PnPOnline










#Legacy
#Connect-PnPOnline -Url $siteUrl -PnPManagementShell
# Azure AD App credentials
#$appId = "71d96e3d-0541-4d92-9fdd-5733d0cdb43d"
#$appSecret = "5~L8Q~wLF8Q1kkDK167mwO8srFpKDqFi6VKRfds2"
#$tenantId = "b9cb1efd-9c45-47ce-bca4-44d53248e00d"

#Connect to SharePoint using Azure AD App
#Connect-PnPOnline -ClientId $appId -ClientSecret $appSecret -Url $siteUrl -WarningAction Ignore

#Connect-PnPOnline -Url $siteUrl -UseWebLogin
