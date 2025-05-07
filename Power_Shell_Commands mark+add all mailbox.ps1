$creds = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $creds -Authentication Basic -AllowRedirection

import-pssession $session

#Connect-MsolService -Credential $creds





#install Packages Exhange Online
Install-Module exchangeonlinemanagement

#Allow Remote PS to Echange Online
Set-ExecutionPolicy Unrestricted 

#Connec to Exhcange Online
connect-exchangeonline  




#​Set-MailboxFolderPermission –identity <mailbox>:\Calendar –user <user> –AccessRights <permission>

#פקודה לבדיקת הרשאות לתיבה קיימת
Get-MailboxFolderPermission -Identity carmela@reality-fund.com:\$CAL
Get-EXOMailboxPermission -Identity adva@reality-fund.com:\Calendar

$CAL = "לוח שנה"
Get-MailboxFolderPermission -Identity rehamb:\$CAL

#פקודה להוספת הרשאות לוח שנה
add-MailboxFolderPermission –identity abira:\Calendar –user  –AccessRights 'editor'
add-EXOMailboxPermission –identity nirg:\Calendar –user rinatab –AccessRights 'editor'

#פקודה להסרת הרשאות לוח שנה
Remove-MailboxFolderPermission –identity gala:\Calendar –user lian
Remove-EXOMailboxPermission –identity gala:\Calendar –user lian

#פקודה לשינוי הרשאות לוח שנה קיימות

Set-MailboxFolderPermission –identity nirg:\Calendar –user rinatab –AccessRights 'editor'
Set-EXOMailboxPermission –identity nirg:\Calendar –user rinatab –AccessRights 'editor'


פקודה להפוך תיבה למשותפת
Set-Mailbox marlina -Type shared

אם הלוח שנה בעברית יש לרשום קודם הגדרת משתנה $CAL = "לוח שנה"
לאחר מכן יש להריץ את הפקודות באופן הבא, לדוגמא: add-MailboxFolderPermission –identity israelh@d.maof.co.il:\$Cal –user adigi –AccessRights editor

Set-MailboxFolderPermission –identity israelh@d.maof.co.il:\$Cal –user adigi –AccessRights editor

https://www.michev.info/Blog/Post/1976/managing-outlook-delegates-via-powershell
Add-MailboxFolderPermission huku:\calendar -User vasil -AccessRights Editor -SharingPermissionFlags Delegate

Get-ADGroupMember groupname | Select-Object name | Export-Csv -Path c:\XXX.csv

Get-User -Identity "yossi" | Format-List פקודה להצגת כל הערכים רשומים ליוזר
Set-User -Identity Jill -Name "Jill" פקודה לשנות ערך NAME

פקודה שתציג את יש ניתוב לתיבה המבוקשת:
Get-Mailbox username  | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward  
Get-EXOMailbox username  | select UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward




# לקבל תוצאה 
Get-PublicFolderClientPermission -Identity \"יומן אירועים"  

# להוסיף הרשאה
Add-PublicFolderClientPermission -Identity \"יומן אירועים" -User "XXX” -AccessRights Editor 

# להוסיף הרשאה קיימת 

Set-PublicFolderClientPermission 


To get the address list in the domain, you can use the PowerShell cmdlet Get-Recipient:
Get-Recipient| Select-Object Name,PrimarySmtpAddress, Phone

To exclude from the list the entries hidden from the address book (HiddenFromAddressLists attribute). User the Export-CSV cmdlet in order to export results to the CSV file:
Get-Recipient -RecipientPreviewFilter $filter | Where-Object {$_.HiddenFromAddressListsEnabled -ne $true} | Select-Object Name,PrimarySmtpAddress, Phone | Export-CSV c:\Temp\GAL.csv -NoTypeInformation


$credential = Get-Credential
 
$Session = New-PSSession -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -ConfigurationName Microsoft.Exchange -Credential $credential -Authentication Basic -AllowRedirection
Import-PSSession $Session




# הענקת הרשאה גורפת ל Calendar עבור כל יומני העובדים
 
$userRequiringAccess = "input_user"
$accessRight = "editor"
 
$mailboxes = Get-mailbox
$userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -erroraction SilentlyContinue
         
    if ($accessRights.accessRights -notmatch $accessRight -and $mailbox.primarysmtpaddress -notcontains $userRequiringAccess.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) Calendar" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\calendar" -User $userRequiringAccess.PrimarySmtpAddress
        if ($accessRights.accessRights -match $accessRight) {
            Write-Host "Successfully added $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Green
        }
        else {
            Write-Host "Could not add $accessRight permissions on $($mailbox.displayname)'s calendar for $($userrequiringaccess.displayname)" -ForegroundColor Red
        }
    }else{
        Write-Host "Permission level already exists for $($userrequiringaccess.displayname) on $($mailbox.displayname)'s calendar" -foregroundColor Green
    }
}
Remove-PSSession $Session




# הענקת הרשאה גורפת למשתמש ליומן בעברית של כל יומני העובדים - $CAL = "לוח שנה"

 
$userRequiringAccess = "input_user"
$accessRight = "editor"
 
$mailboxes = Get-mailbox
$userRequiringAccess = Get-mailbox $userRequiringAccess
foreach ($mailbox in $mailboxes) {
    $accessRights = $null
    $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\$CAL" -User $userRequiringAccess.PrimarySmtpAddress -erroraction SilentlyContinue
         
    if ($accessRights.accessRights -notmatch $accessRight -and $mailbox.primarysmtpaddress -notcontains $userRequiringAccess.primarysmtpaddress -and $mailbox.primarysmtpaddress -notmatch "DiscoverySearchMailbox") {
        Write-Host "Adding or updating permissions for $($mailbox.primarysmtpaddress) $CAL" -ForegroundColor Yellow
        try {
            Add-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\$CAL" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }
        catch {
            Set-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\$CAL" -User $userRequiringAccess.PrimarySmtpAddress -AccessRights $accessRight -ErrorAction SilentlyContinue    
        }        
        $accessRights = Get-MailboxFolderPermission "$($mailbox.primarysmtpaddress):\$CAL" -User $userRequiringAccess.PrimarySmtpAddress
        if ($accessRights.accessRights -match $accessRight) {
            Write-Host "Successfully added $accessRight permissions on $($mailbox.displayname)'s $CAL for $($userrequiringaccess.displayname)" -ForegroundColor Green
        }
        else {
            Write-Host "Could not add $accessRight permissions on $($mailbox.displayname)'s $CAL for $($userrequiringaccess.displayname)" -ForegroundColor Red
        }
    }else{
        Write-Host "Permission level already exists for $($userrequiringaccess.displayname) on $($mailbox.displayname)'s $CAL" -foregroundColor Green
    }
}
Remove-PSSession $Session
