    Param
    (
        [String]$OutCsvFilePath = "C:\Intel\AllADUsersLastLogonTimeList.csv",
        [Int]$PageSize = 100
    ) 
$Filter = "(&(objectCategory=User))" 
##
## Ensure you have a folder in C:\ called "Intel"        
##
##
    $Domain = New-Object System.DirectoryServices.DirectoryEntry  
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher 
    $Searcher.SearchRoot = "LDAP://$($Domain.DistinguishedName)" 
    $Searcher.PageSize = 1000 
    $Searcher.SearchScope = "Subtree" 
    $Searcher.Filter = $Filter 
    $Searcher.PropertiesToLoad.Add("Name") | Out-Null 
    $Searcher.PropertiesToLoad.Add("LastLogonTimeStamp") | Out-Null 
 
    $Results = $Searcher.FindAll() 
     
     "SamAccountName,LastLogonTimeStamp" | out-file $OutCsvFilePath -encoding ascii -append  
     Foreach($Result in $Results) 
     { 
            $Name = $Result.Properties.Item("Name") 
            $LastLogonTimeStamp = $Result.Properties.Item("LastLogonTimeStamp")           
            If ($LastLogonTimeStamp.Count -eq 0) 
            { 
                $LastLogonTimeStamp = "Never Logon" 
            } 
            Else 
            { 
                $Time = [DateTime]$LastLogonTimeStamp.Item(0) 
                $LastLogonTimeStamp = $Time.AddYears(1600).ToString("yyyy/MM/dd") 
            } 
             
            $Name.trim().replace(","," ") + "," + $LastLogonTimeStamp.trim() | out-file $OutCsvFilePath -encoding ascii -append                                                
     }
