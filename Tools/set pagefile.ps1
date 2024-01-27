$computersys = Get-WmiObject Win32_ComputerSystem -EnableAllPrivileges
$computersys.AutomaticManagedPagefile = $False
$computersys.Put()
$pagefile = Get_WmiObject -Query "Select * From Win32_PageFileSetting Where Name='c:\\pagefile.sys'"
$pagefile.Delete()
Set-WMIInstance -class Win32_PageFileSetting -Arguments @{name="d:\pagefile.sys";InitialSize = 4096;MaximumSize =4096}