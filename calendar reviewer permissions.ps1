## Launch PS and log into Office 365 ##

## Test the policy change with the following script ##

##  Variables - SearchBase  This is the DN of the OU Path which needs to be changed for each OU that we are running this against ##
##              Access Rights - This is what we are providing the permission for (Default is AvailabilityOnly) ##

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights reviewer
}

## Live script is below and configured ##

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights reviewer
}

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights reviewer
}

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights reviewer
}


$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Get-mailboxfolderpermission -identity ($path) -user Default
}


###REVERSE PROCEDURE####

AvailabilityOnly

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights AvailabilityOnly
}

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights AvailabilityOnly
}

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights AvailabilityOnly
}

$ouusers = Get-ADUser -Filter * -SearchBase “OU=**,OU=**,OU=**,OU=**,DC=**,DC=**,DC=**”
Foreach ($Mailbox in $ouusers)
{
$path = $Mailbox.userprincipalname + “:\” + (Get-MailboxFolderStatistics $Mailbox.userprincipalname | Where-Object { $_.Foldertype -eq “Calendar” } | Select-Object -First 1).Name
Set-mailboxfolderpermission -identity ($path) -user Default -Accessrights AvailabilityOnly
