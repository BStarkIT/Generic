<#
.SYNOPSIS
    Retrieve size information from child items of specified path(s).

.DESCRIPTION
    Retrieve size information from child items of specified path(s).

.Parameter Directory
    "C:\Test\Directory"

.NOTES
    Author: JBear 6/1/2018
#>

param(

    [Parameter(ValueFromPipeline=$true)]
    [String[]]$Directory = $null,

    [Parameter(ValueFroMPipeline=$true)]
    [String[]]$ComputerName = $env:COMPUTERNAME,

    [Parameter(DontShow)]
    [String]$JobThrottleCount = 10
)

if($Directory -eq $null) {

    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $Result = $Dialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{ TopMost = $true }))

    if($Result -eq 'OK') {

        Try {

            $Directory = $Dialog.SelectedPath
        }

        Catch {

            $Directory = $null
          Break
        }
    }

    else {

        #Shows upon cancellation of Save Menu
        Write-Host -ForegroundColor Yellow "Notice: No file(s) selected."
        Break
    }
}

function MeasureSize {

    foreach($Computer in $ComputerName) {

        $i=0
        $j=0

        foreach($Dir in $Directory) {

                $DirMod = "\\$Computer\$(($Dir).Replace(':','$'))"

                #Retrieve object 'name' from each $Computer path
                $Names = Get-ChildItem -LiteralPath $DirMod -Directory | Select-Object Name

            ForEach($Name in $Names) {

                While(@(Get-Job -State Running).Count -ge $JobThrottleCount) {

                    Start-Sleep 1
                }

                Write-Progress -Activity "Begin Measuring Processes..." -Status ("Percent Complete:" + "{0:N0}" -f ((($i++) / $Names.count) * 100) + "%") -CurrentOperation "Processing $((Split-Path -Path $Name.Name -Leaf))..." -PercentComplete ((($j++) / $Names.count) * 100)

                Start-Job {

                  $Drive = "$using:DirMod\" + $using:Name.Name

                    #Measure file lengths (bytes) for each $ServerShare recursively to retrieve full directory size
                  $DirSize = (Get-ChildItem $Drive -Recurse -ErrorAction "SilentlyContinue" -Force | Where {-NOT $_.PSIscontainer}  | Measure-Object -Property Length -Sum)

                    [PSCustomObject] @{

                        Drive = $Drive
                        MB = "{0:N2}" -f $($DirSize.Sum/1MB) + " MB"
                        GB = "{0:N2}" -f $($DirSize.Sum/1GB) + " GB"
                    }     
                } -Name 'Measure Directory'
            }

            $Jobs = Get-Job | Where { $_.State -eq "Running" }
            $Total = $Jobs.Count
            $Running = $Jobs.Count

            While($Running -gt 0) {

                Write-Progress -Activity "Retrieving Metrics Data... (Awaiting Results: $(($Running)))..." -Status ("Percent Complete:" + "{0:N0}" -f ((($Total - $Running) / $Total) * 100) + "%") -PercentComplete ((($Total - $Running) / $Total) * 100) -ErrorAction SilentlyContinue
                $Running = (Get-Job | Where { $_.State -eq "Running" }).Count
            }  
        }
    }
}

#Call main function
MeasureSize | Receive-Job -Wait -AutoRemoveJob | Select Drive, MB, GB
