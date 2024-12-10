# Content: Create Shared-Mailboxes in Exchange Online
# Author: Sebastian Stecher | @HanseSecure | https://hansesecure.de
# Date: 12/2024

echo "installing PowerShellGet... (needs to run only once, comment out on subsequent runs)"
Install-Module PowerShellGet -Force -AllowClobber
echo "installing ExchangeOnlineManagement..., allow if prompted to trust PSGallery (needs to run only once, comment out on subsequent runs)"
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
echo "installed, printing version of ExchangeOnlineManagement..."
Import-Module ExchangeOnlineManagement
Get-Module ExchangeOnlineManagement
$user = Read-Host "Please enter your Exchange Admin Principal Name, for example: admin@hansesecure.com"
echo "trying to connect to Exchange..."
Connect-ExchangeOnline -UserPrincipalName $user
$amount = Read-Host "Please enter the total amount of shared mailboxes to create as an integer"
$boxName = Read-Host "Please enter the alias prefix. Boxes will be created with aliases according to {prefix}{i} with {i} counting up from 1 to the given amount. For example: phishingtest_ as an alias will create the mailboxes phishingtest_1@your-domain.com until phishingtest_{i}@your-domain.com"
$group = Read-Host "Please enter the user OR a mail enabled security group that should have access to the created mailboxes (if it doesn't exist, create it now)"
$csvfile = "$env:userprofile\downloads\test_mails.csv"
echo "making a .csv file: $csvfile ..."
echo "First Name,Last Name,Email,Position" | Out-File -Encoding "UTF8" $csvfile
for($i=1; $i -le $amount; $i++){
    $currentBox = "$boxName$i"
    $created = New-Mailbox -Shared -Name $currentBox -DisplayName $currentBox -Alias $currentBox | select -exp PrimarySMTPAddress
    echo $created
    echo "Phishing,Test,$created,$i" | Out-File -Encoding "UTF8" -Append $csvfile
    Set-Mailbox -Identity $currentBox -GrantSendOnBehalfTo $group
    Add-MailboxPermission -Identity $currentBox -User $group -AccessRights FullAccess -InheritanceType All
}
echo "disconnecting from Exchange..."
Disconnect-ExchangeOnline