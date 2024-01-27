$Clone = 'ACL_Contempt Of Court_Read&Write'
$Target = 'ACL_Contempt Of Court_Read'
get-ADuser -identity $Clone -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $Target