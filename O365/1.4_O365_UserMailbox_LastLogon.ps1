$SharedMailboxUser =  Get-Mailbox -ResultSize unlimited -RecipientTypeDetails UserMailbox | Select-Object Alias -ExpandProperty Alias

foreach($alias in $SharedMailboxUser)
    {
    $Stats = get-mailboxstatistics -identity $alias | Select-Object LastLogonTime, LastLoggedOnUserAccount, LastLogoffTime
    $MailDetails = Get-Mailbox -Identity $alias | Select-Object Alias, PrimarySmtpAddress, DisplayName, WhenMailboxCreated, HiddenFromAddressListsEnabled, recipienttypedetails
   
    $Alias = $MailDetails.Alias
    $PrimarySMTPAddress  = $MailDetails.PrimarySmtpAddress
    $DisplayName = $MailDetails.DisplayName
    $WhenMailCreated = $MailDetails.WhenMailboxCreated
    $HiddenFromAddressListsEnabled = $MailDetails.HiddenFromAddressListsEnabled
    $RecipientTypeDetails = $MailDetails.RecipientTypeDetails

    [string]$LastLogonTime = $Stats.LastLogonTime
    $LastLoggedOnUserAccount = $Stats.LastLoggedOnUserAccount
    [string]$LastLogoffTime = $Stats.LastLogoffTime
    $LastLoggedOnUser = ($LastLoggedOnUserAccount -split "\\")[-1]

    $Details = [pscustomObject]@{
   
    'Alias' = $Alias
    'PrimarySMTPAddress'  = $PrimarySMTPAddress
    'RecipientTypeDetails' = $RecipientTypeDetails
    'DisplayName' = $DisplayName
    'WhenMailCreated' = ($WhenMailCreated -split " ")[0]
    'HiddenFromAddressListsEnabled' = $HiddenFromAddressListsEnabled
    'LastLogonTime' = ($LastLogonTime -split " ")[0]    
    'LastLogoffTime' = ($LastLogoffTime -split " ")[0]
    'LastLoggedOnUserAccount' = $LastLoggedOnUserAccount
    'LastLoggedOnUser' = $LastLoggedOnUser
    }

    $Details | Export-Csv -Path C:\scripts\O365\reports\O365_UserMailboxes_LastLogon_02032023.csv -Append -NoTypeInformation
    }
