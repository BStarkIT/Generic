$Users = Get-Mailbox -ResultSize Unlimited -Filter {AcceptMessagesOnlyFrom -eq 'Helpdesk'} 
foreach($User in $Users)
{
Set-mailbox -Identity $User -AcceptMessagesOnlyFrom $Null
}