Import-Module ExchangeOnlineManagement
Import-Module PnP.PowerShell

$password = ConvertTo-SecureString "XXXXXXXXXXX" -AsPlainText -Force
$username = "SharePointAuto@company.com"
$credential = New-Object System.Management.Automation.PSCredential($username, $password)


Connect-ExchangeOnline -Credential $credential -ShowProgress $true

#{# Get User Information

$serviceAccountUsername = "SharePointAuto@company.com"
$serviceAccountPassword = ConvertTo-SecureString "XXXXXXXXXX" -AsPlainText -Force
$siteUrl  = "https://company.sharepoint.com"


# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -Credentials (New-Object System.Management.Automation.PSCredential($serviceAccountUsername, $serviceAccountPassword))

# Specify the list title
$listTitle = "Directory"

# Get items from the SharePoint list, including 'Department', 'Email', and 'Status' fields
$listItems = Get-PnPListItem -List $listTitle -Fields "Department", "Email", "Status"

# Create an array to store PowerShell objects
$users = @()

# Loop through each item and extract 'Department', 'Email', and 'Status' columns
foreach ($item in $listItems) {
    Write-Host "Processing item $($item.Id)"
    
    # Extract 'Department' field value
    $departmentField = $item.FieldValues["Department"]
    if ($departmentField) {
        $department = $departmentField.LookupValue
    } else {
        $department = ""
    }

    # Check if 'Email' field exists in the item
    if ($item.FieldValues.ContainsKey("Email")) {
        # Extract 'Email' field value
        $email = $item.FieldValues["Email"]
    } else {
        $email = ""
    }

    # Check if 'Status' field exists in the item and if it equals 'Active'
    if ($item.FieldValues.ContainsKey("Status") -and $item.FieldValues["Status"] -eq "Active") {
        # Create a PowerShell object for the row
        $user = [PSCustomObject]@{
            Department = $department
            Email = $email
        }

        # Add the object to the array
        $users += $user
    }
}

# Output the PowerShell objects
$users

$emails = foreach ($user in $users) {
    $user.Email
}

# Output the array of emails
$emails

$departments = foreach ($user in $users) {
    $user.Department
}
$departments = $departments | Sort-Object | Get-Unique
# Output the array of departments
$departments
#}
#---------------------------
#{# Get Shared Mail Box Information

$serviceAccountUsername = "SharePointAuto@company.com"
$serviceAccountPassword = ConvertTo-SecureString "XXXXXXXXX" -AsPlainText -Force
$siteUrl  = "https://company.sharepoint.com/IT"


# Connect to SharePoint Online
Connect-PnPOnline -Url $siteUrl -Credentials (New-Object System.Management.Automation.PSCredential($serviceAccountUsername, $serviceAccountPassword))

#list title
$listTitle = "Email"

# Get items from the SharePoint list, including 'Department' and 'Email' fields
$listItems = Get-PnPListItem -List $listTitle -Fields "Email", "Dept", 'AutoMapping'

# Initialize an empty array to store PowerShell objects
$sharedmailboxes = @()


# Loop through each item and extract 'Department' and 'Email' columns
foreach ($item in $listItems) {
    Write-Host "Processing item $($item.Id)"

    # Check if 'Email' field exists in the item
    if ($item.FieldValues.ContainsKey("Email")) {
        # Get 'Email' field value
        $email = $item.FieldValues["Email"]
    } else {
        $email = ""
    }

    # Check if 'Department' field exists in the item
    if ($item.FieldValues.ContainsKey("Dept")) {
        # Get 'Department' field value
        $dept = $item.FieldValues["Dept"]
    } else {
        $dept = ""
        Write-Host "Dept field does not exist in the item"
    }

        # Check if 'AutoMapping' field exists in the item
        if ($item.FieldValues.ContainsKey("AutoMapping")) {
        # Get 'Department' field value
        $automapping = $item.FieldValues["AutoMapping"]
    } else {
        $automapping = ""
        Write-Host "AutoMapping field does not exist in the item"
    }

    # Create a PowerShell object for the row
    $sharedmailbox = [PSCustomObject]@{
        Dept = $dept
        Email = $email
        AutoMapping = $automapping
    }

    # Add the object to the list
    $sharedmailboxes += $sharedmailbox
}

# Output the PowerShell objects
$sharedmailboxes
#}

#---------------------------
#{# Iterate through each shared mailbox and add users based on Dept to each mailbox if mailbox dept(s) = user's department

foreach ($sharedmailbox in $sharedmailboxes) {
    $SharedMailboxName = $sharedmailbox.Email
    $Mailbox = Get-EXOMailbox -Identity $SharedMailboxName

    # Initialize $AllPermissions as an empty array
    $AllPermissions = @()

    # Get users with SendAs permission
    $SendAsPermissions = Get-RecipientPermission -Identity $Mailbox.DistinguishedName

    # Get users with FullAccess permission
    $FullAccessPermissions = Get-MailboxPermission -Identity $SharedMailboxName | Where-Object { $_.AccessRights -eq "FullAccess" }

    # Get users with SendOnBehalf permission
    $SendOnBehalfPermissions = Get-MailboxPermission -Identity $SharedMailboxName | Where-Object { $_.AccessRights -eq "SendOnBehalf" }

    # Combine results by concatenating permissions arrays
    if ($SendAsPermissions) {
        $AllPermissions += $SendAsPermissions
    }
    if ($FullAccessPermissions) {
        $AllPermissions += $FullAccessPermissions
    }
    if ($SendOnBehalfPermissions) {
        $AllPermissions += $SendOnBehalfPermissions
    }

    # Check if $AllPermissions is not empty
    if ($AllPermissions) {
        Write-Host "Permissions found for $SharedMailboxName"
    } else {
        Write-Host "No permissions found for $SharedMailboxName"
    }

    # Iterate through each user
    foreach ($user in $users) {
        if ($user.Department -in $sharedmailbox.Dept.LookupValue) { #changed -eq to -in since a sharedmailbox can have multiple depts.
            if ($sharedmailbox.AutoMapping -eq 'True'){$automap = $true}
            else {$automap = $false}


            if ($AllPermissions -and $AllPermissions.User -contains $user.Email) {
                #($automap -!= ){
                    #Check if they have automaping -eq to the value if they aren't the same reset it
                    $warningOutput = ""
                    Add-MailboxPermission -Identity $SharedMailbox.Email -User $user.Email -AccessRights $PermissionType -AutoMapping $automap -WarningVariable warningOutput
                    if ($warningOutput -ne ''){
                    Remove-MailboxPermission -Identity $SharedMailbox.Email -User $user.Email -AccessRights $permissionType -Confirm:$false 
                    Add-MailboxPermission -Identity $SharedMailbox.Email -User $user.Email -AccessRights $PermissionType -AutoMapping $automap
                    }
                }
            }
        }
    }