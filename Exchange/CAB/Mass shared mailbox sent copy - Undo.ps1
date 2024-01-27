<#
.SYNOPSIS
This PowerShell script is to remove the addition of copies of sent emails from the sent items folder of the shared mailbox

.NOTES
written by Brian Stark 
Date 15/03/2023
Proof Read and approved by: Matt McGowan
Date: 15/03/2023

#>
$Mailboxes = get-exoMailbox -RecipientTypeDetails SharedMailbox | Select-Object Identity
ForEach ($Share in $Mailboxes) {
    set-mailbox $Share -MessageCopyForSendOnBehalfEnabled $False
    set-mailbox $Share -MessageCopyForSentAsEnabled $False
}
