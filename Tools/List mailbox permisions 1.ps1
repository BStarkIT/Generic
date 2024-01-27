$memberList = $memberList.user

$user = $line.User.RawIdentity.split('\')

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue
$memberList = Get-MailboxPermission solemnappeals@scotcourts.gov.uk | select user, isinherited
$memberList = $memberList.user
Result = @()
foreach ($line in $memberList)
{
##$user = $line.RawIdentity.Split('\')[1]
##$Result += Get-Mailbox -Identity $user | select name, alias, primarysmtpaddress

$Result += Get-Mailbox -Identity $line.RawIdentity | select name, alias, primarysmtpaddress
}

$Result | Out-GridView
