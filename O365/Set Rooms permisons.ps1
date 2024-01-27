##
## Defaults permisions to all rooms
##
$users = Get-Mailbox -Resultsize Unlimited
foreach ($user in $users) { 
    Write-Host -ForegroundColor green "Setting permission for $($user.alias)..." 
    Set-MailboxFolderPermission -Identity "$($user.alias):\calendar" -User Default -AccessRights Reviewer 
}
## 
## Adds permisions for new Office manager
##
add-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User NEW.USER@Realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User NEW.USER@Realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User NEW.USER@Realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User NEW.USER@realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User NEW.USER@realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User NEW.USER@realise.com -AccessRights Editor
add-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User NEW.USER@realise.com -AccessRights Editor
## 
## re-applies normal settings to other office managers
##
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Default -AccessRights Author
set-MailboxFolderPermission -Identity quayhouse1@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Default -AccessRights Author
set-MailboxFolderPermission -Identity quayhouse2@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
Set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Default -AccessRights Author
set-MailboxFolderPermission -Identity quayhouse3@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Default -AccessRights Author
set-MailboxFolderPermission -Identity quayhouse4@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Default -AccessRights Reviewer
set-MailboxFolderPermission -Identity quayhouseBoardroom@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Default -AccessRights Reviewer
set-MailboxFolderPermission -Identity quayhouse5@realise.com:\calendar -User Anonymous -AccessRights None 
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Vivienne.Lynch@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Grant.Crain@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Emma.Murray@Realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Eileen.Goddard@realise.com -AccessRights Editor
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Default -AccessRights Reviewer
set-MailboxFolderPermission -Identity quayhouse6@realise.com:\calendar -User Anonymous -AccessRights None 
