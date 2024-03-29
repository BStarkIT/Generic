Function Add-NewPortOpening {
    <#
.Synopsis
Adds a rule to the Windows firewall to allow inbound on a specific port.
.Description
Adds a rule to the Windows firewall to allow inbound on a specific port.
.Parameter Port
The port number you want to allow inbound
.Parameter RemoteAddress
Optional parameter to scope the rule to only allow from a particular IP or subnet. 
See https://docs.microsoft.com/en-us/powershell/module/netsecurity/new-netfirewallrule?view=winserver2012-ps
.Example
Add-NewPortOpening -Port 5986
Adds a rule to the Windows firewall to allow inbound for port 5986.
.Example
Add-NewPortOpening -Port 5986 -RemoteAddress 10.0.0.0/8
Adds a rule to the Windows firewall to allow inbound for port 5986 only from hosts on the 10.0.0.0/8 network.
.Example
Add-NewPortOpening -Port 5986 -RemoteAddress 10.0.0.0/8, 192.168.0.0/16
Adds a rule to the Windows firewall to allow inbound for port 5986 only from hosts on the 10.0.0.0/8 network or the 192.168.0.0/16 network.
.Notes
Version History:
2019-05-05: Initial
#>

    [Cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $True, Position = 0)]   
        [ValidateRange(1, 65535)]
        [int]$Port,

        [Parameter(Mandatory = $False, Position = 0)]   
        [String[]]$RemoteAddress
    )
    
    Begin {
        ####################<Default Begin Block>####################
        # Force verbose because Write-Output doesn't look well in transcript files
        $VerbosePreference = "Continue"
        
        [String]$Logfile = $PSScriptRoot + '\PSLogs\' + (Get-Date -Format "yyyy-MM-dd") +
        "-" + $MyInvocation.MyCommand.Name + ".log"
        
        Function Write-Log {
            Param
            (
                [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
                [PSObject]$InputObject,
                
                [Parameter(Mandatory = $False, Position = 1)]
                [Validateset("Black", "Blue", "Cyan", "Darkblue", "Darkcyan", "Darkgray", "Darkgreen", "Darkmagenta", "Darkred", `
                        "Darkyellow", "Gray", "Green", "Magenta", "Red", "White", "Yellow")]
                [String]$Color = "Green"
            )
            
            $ConvertToString = Out-String -InputObject $InputObject -Width 100
            
            If ($($Color.Length -gt 0)) {
                $previousForegroundColor = $Host.PrivateData.VerboseForegroundColor
                $Host.PrivateData.VerboseForegroundColor = $Color
                Write-Verbose -Message "$(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"): $ConvertToString"
                $Host.PrivateData.VerboseForegroundColor = $previousForegroundColor
            }
            Else {
                Write-Verbose -Message "$(Get-Date -Format "yyyy-MM-dd hh:mm:ss tt"): $ConvertToString"
            }
            
        }

        Function Start-Log {
            # Create transcript file if it doesn't exist
            If (!(Test-Path $Logfile)) {
                New-Item -Itemtype File -Path $Logfile -Force | Out-Null
            }
        
            # Clear it if it is over 10 MB
            [Double]$Sizemax = 10485760
            $Size = (Get-Childitem $Logfile | Measure-Object -Property Length -Sum) 
            If ($($Size.Sum -ge $SizeMax)) {
                Get-Childitem $Logfile | Clear-Content
                Write-Verbose "Logfile has been cleared due to size"
            }
            Else {
                Write-Verbose "Logfile was less than 10 MB"   
            }
            Start-Transcript -Path $Logfile -Append 
            Write-Log "####################<Function>####################"
            Write-Log "Function started on $env:COMPUTERNAME"

        }
        
        Function Stop-Log {
            Write-Log "Function completed on $env:COMPUTERNAME"
            Write-Log "####################</Function>####################"
            Stop-Transcript
       
            # Now we will clean up the transcript file as it contains filler info that needs to be removed...
            $Transcript = Get-Content $Logfile -raw

            # Create a tempfile
            $TempFile = $PSScriptRoot + "\PSLogs\temp.txt"
            New-Item -Path $TempFile -ItemType File | Out-Null
			
            # Get all the matches for PS Headers and dump to a file
            $Transcript | 
            Select-String '(?smi)\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*([\S\s]*?)\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*\*' -AllMatches | 
            ForEach-Object { $_.Matches } | 
            ForEach-Object { $_.Value } | 
            Out-File -FilePath $TempFile -Append

            # Compare the two and put the differences in a third file
            $m1 = Get-Content -Path $Logfile
            $m2 = Get-Content -Path $TempFile
            $all = Compare-Object -ReferenceObject $m1 -DifferenceObject $m2 | Where-Object -Property Sideindicator -eq '<='
            $Array = [System.Collections.Generic.List[PSObject]]@()
            foreach ($a in $all) {
                [void]$Array.Add($($a.InputObject))
            }
            $Array = $Array -replace 'VERBOSE: ', ''

            Remove-Item -Path $Logfile -Force
            Remove-Item -Path $TempFile -Force
            # Finally, put the information we care about in the original file and discard the rest.
            $Array | Out-File $Logfile -Append -Encoding ASCII
            
        }
        
        Start-Log

        Function Set-Console {
            <# 
        .Synopsis
        Function to set console colors just for the session.
        .Description
        Function to set console colors just for the session.
        This function sets background to black and foreground to green.
        Verbose is DarkCyan which is what I use often with logging in scripts.
        I mainly did this because darkgreen does not look too good on blue (Powershell defaults).
        .Notes
        2017-10-19: v1.0 Initial script 
        #>
        
            Function Test-IsAdmin {
                <#
                .Synopsis
                Determines whether or not the user is a member of the local Administrators security group.
                .Outputs
                System.Bool
                #>

                [CmdletBinding()]
    
                $Identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
                $Principal = new-object System.Security.Principal.WindowsPrincipal(${Identity})
                $IsAdmin = $Principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
                Write-Output -InputObject $IsAdmin
            }

            $console = $host.UI.RawUI
            If (Test-IsAdmin) {
                $console.WindowTitle = "Administrator: Powershell"
            }
            Else {
                $console.WindowTitle = "Powershell"
            }
            $Background = "Black"
            $Foreground = "Green"
            $Messages = "DarkCyan"
            $Host.UI.RawUI.BackgroundColor = $Background
            $Host.UI.RawUI.ForegroundColor = $Foreground
            $Host.PrivateData.ErrorForegroundColor = $Messages
            $Host.PrivateData.ErrorBackgroundColor = $Background
            $Host.PrivateData.WarningForegroundColor = $Messages
            $Host.PrivateData.WarningBackgroundColor = $Background
            $Host.PrivateData.DebugForegroundColor = $Messages
            $Host.PrivateData.DebugBackgroundColor = $Background
            $Host.PrivateData.VerboseForegroundColor = $Messages
            $Host.PrivateData.VerboseBackgroundColor = $Background
            $Host.PrivateData.ProgressForegroundColor = $Messages
            $Host.PrivateData.ProgressBackgroundColor = $Background
            Clear-Host
        }
        Set-Console

        ####################</Default Begin Block>####################
        
    }

    Process {
        Try {
            If ( $RemoteAddress -eq '') {
                [string]$Port = $Port.ToString()
                $Params = @{
                    'DisplayName' = "Allow " + $Port + " In"
                    'Description' = "Allow " + $Port + " In"
                    'Profile'     = "Any"
                    'Direction'   = "Inbound"
                    'LocalPort'   = $Port
                    'Protocol'    = "TCP"
                    'Action'      = "Allow"
                    'Enabled'     = "True"
                }
                New-NetFirewallRule @Params | Out-Null
            }
            Else {
                [string]$Port = $Port.ToString()
                $Params = @{
                    'DisplayName'   = "Allow " + $Port + " In"
                    'Description'   = "Allow " + $Port + " In"
                    'Profile'       = "Any"
                    'Direction'     = "Inbound"
                    'LocalPort'     = $Port
                    'Protocol'      = "TCP"
                    'Action'        = "Allow"
                    'Enabled'       = "True"
                    'RemoteAddress' = $RemoteAddress
                }
                New-NetFirewallRule @Params | Out-Null 
            }
        }
        Catch {
            Write-Error $($_.Exception.Message)
        }
    }

    End {
        Stop-log
    }
}

<#######</Body>#######>
<#######</Script>#######>