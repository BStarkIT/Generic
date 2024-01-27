$shared = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Select-Object PrimarySMTPAddress
foreach ($share in $shared) {
    $mailbox1 = $share -replace '@{PrimarySmtpAddress=', ''
    $mailbox = $mailbox1 -replace '}', ''
    $messageTrackingLog = Get-MessageTrackingLog -Recipients $mailbox -ResultSize Unlimited | Select-Object sender, timestamp | Sort-Object timestamp -Descending
    Start-Transcript -Path "C:\PS\Mailtrack.txt" -Append
    if ($NULL -eq $messageTrackingLog) {
        Write-Output "There were no messages sent to to '$mailbox' shared mailbox"
    }
    else {
        $sender = $messageTrackingLog[0].Sender
        $timeStamp = $messageTrackingLog[0].TimeStamp

        Write-Output "Last message received to '$mailbox' was from '$sender' on '$timeStamp'"
    }
}