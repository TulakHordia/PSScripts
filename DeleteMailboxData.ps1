Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName ttadmin@TCSM.onmicrosoft.com

# Define the user's mailbox
$UserMailbox = "input_user"

# Create a Compliance Search for all emails in the mailbox
$Search = New-ComplianceSearch -Name "DeleteMailbox_$UserMailbox" `
    -ExchangeLocation $UserMailbox `
    -ContentMatchQuery "received>01/01/1900"

# Start the search
Start-ComplianceSearch -Identity $Search.Identity

# Wait for the search to complete (optional: check status with Get-ComplianceSearch)
Start-Sleep -Seconds 60

# Purge all items in the mailbox
New-ComplianceSearchAction -SearchName "DeleteMailbox_$UserMailbox" -Purge -PurgeType HardDelete

# Verify purge action
Get-ComplianceSearchAction -SearchName "DeleteMailbox_$UserMailbox" | Format-List
