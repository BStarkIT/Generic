$Path="C:\scripts\Exchange\rooms\"
$date=Get-Date -f "_dd_MM_yyyy"
$CSVOutputFilePath = $Path+"RoomsExport"+$date+"_Export.csv"
$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
#$CSVOutputFilePath = $Path+"RoomsExport"+$date+"_Export.csv"
#$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize unlimited
#$rooms = Get-Mailbox "1st_Floor_ISD_Meeting_Room@scotcourts.com"
foreach($room in $rooms){
      $CalendarProcessingObject=$null
      $OU=""
      $BookInPolicy=""
      $SplittedDN=""
      $SplittedDN=$room.DistinguishedName.split(",")
      $OU=$SplittedDN[3]+","+$SplittedDN[2]+","+$SplittedDN[1]
      $CalendarProcessingObject=Get-CalendarProcessing $room.Alias
      $ResourceDelegates = $CalendarProcessingObject.ResourceDelegates
      $AutomateProcessing = $CalendarProcessingObject.AutomateProcessing
      $AllBookInPolicy = $CalendarProcessingObject.AllBookInPolicy
      $BookInPolicy = $CalendarProcessingObject.BookInPolicy
      $filter = ""
      $BookInPolicyGroupObject = $null
      $Members = $null
      if ($BookInPolicy){
            $filter = [string]$BookInPolicy + "*"
            $BookInPolicyGroupObject = Get-ADGroup -Filter {LegacyExchangeDN -like $filter}                       
            if ($BookInPolicyGroupObject){
                  $Members = Get-ADGroupMember $BookInPolicyGroupObject -Recursive
                  foreach ($Member in $Members) {
                        $ADuser=$null
                        if ($Member.samaccountname){
                              $ADuser=Get-ADUser $Member.samaccountname -Properties mail
                              $Member | Add-Member -MemberType NoteProperty -Name mail -Value $ADuser.mail -Force
                        }else{

                Write-Output $Member
                if($Member){
                              $Member | Add-Member -MemberType NoteProperty -Name mail -Value "null" -Force
                    }   
                        }
                  }

                  if ($Members){
                        $Members = ($Members | where-object{$_.mail -like "*@*"} | Select -ExpandProperty mail) -join "~"
                  }else{
                        $Members = "null"
                  }           
            }
      }
      
      $AllRequestOutofPolicy = $CalendarProcessingObject.AllRequestOutofPolicy
      $BookingWindowInDays = $CalendarProcessingObject.BookingWindowInDays
      $MaximumDurationInMinutes = $CalendarProcessingObject.MaximumDurationInMinutes
      
      $room | Add-Member -MemberType NoteProperty -Name OU -Value $OU -Force
      $room | Add-Member -MemberType NoteProperty -Name ResourceDelegates -Value $ResourceDelegates -Force
      $room | Add-Member -MemberType NoteProperty -Name AutomateProcessing -Value $AutomateProcessing -Force
      $room | Add-Member -MemberType NoteProperty -Name AllBookInPolicy -Value $AllBookInPolicy -Force
      $room | Add-Member -MemberType NoteProperty -Name BookInPolicy -Value $BookInPolicy -Force
      $room | Add-Member -MemberType NoteProperty -Name BookInPolicyGroupName -Value $BookInPolicyGroupObject.Name -Force
      $room | Add-Member -MemberType NoteProperty -Name BookInPolicyGroupMembers -Value $Members -Force
      $room | Add-Member -MemberType NoteProperty -Name AllRequestOutofPolicy -Value $AllRequestOutofPolicy -Force
      $room | Add-Member -MemberType NoteProperty -Name BookingWindowInDays -Value $BookingWindowInDays -Force
      $room | Add-Member -MemberType NoteProperty -Name MaximumDurationInMinutes -Value $MaximumDurationInMinutes -Force
}

#Change Columns order
$RoomsOutput = $Rooms | select DisplayName,primarysmtpaddress,OU,RecipientTypeDetails,customattribute2,ResourceCapacity,ResourceDelegates,HiddenFromAddressListsEnabled,AutomateProcessing,AllBookInPolicy,BookInPolicy,BookInPolicyGroupName,BookInPolicyGroupMembers,AllRequestOutofPolicy,BookingWindowInDays,MaximumDurationInMinutes
$RoomsOutput | Export-Csv -Path $CSVOutputFilePath -NoTypeinformation