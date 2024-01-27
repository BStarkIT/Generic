<#
.SUMMARY
Operational readiness test for Windows 10 deployments 
v0.98 - update tests for OU path and client authenticaation certificates.
#>

$TestScriptVersion = '0.98'
$BIOSSettings = Get-CimInstance -Namespace root\HP\InstrumentedBIOS -Class HP_BIOSSetting
$ComputerSystem = Get-CimInstance -Class Win32_ComputerSystem
$HWModel = ($BIOSSettings | Where-Object Name -eq 'Product Name').Value
$SerialNumber = ($BIOSSettings | Where-Object Name -eq 'Serial Number').Value
$SystemBIOSVersion = ($BIOSSettings | Where-Object Name -eq 'System BIOS Version').Value
$CCMClientVersion = (Get-CimInstance -Namespace root\ccm -Class SMS_Client).ClientVersion

$OPGMachine= 'N'
$prompt = Read-Host "Is this an OPG machine (Y/N)? Press ENTER to accept default [$($OPGMachine)]"
if ($prompt -eq "") {} else {
    $OPGMachine = $prompt
    }


Write-Output "   Test Script Version v$TestScriptVersion`n"
Write-Output "   Name:`t`t$env:COMPUTERNAME"
Write-Output "   Model:`t`t$HWModel"
Write-Output "   Serial #:`t`t$SerialNumber"
Write-Output "   System Bios Ver:`t$SystemBIOSVersion`n"
Write-output "   SCCM Client Ver:`t$CCMClientVersion`n"

$IsLaptop = $false
$IsDesktop = $false
If ($ComputerSystem.PCSYstemType -eq 1) {$IsDesktop = $true} else {$IsLaptop = $true}

Describe "Deployment QA" {

    Context "BIOS Settings." {
        
        It "has a BIOS setup password set." {
            (Get-CimInstance -Namespace root\HP\InstrumentedBIOS -Class HP_BIOSPassword | 
            Where-Object {$_.name -eq 'Setup Password'}).IsSet | Should Be '1'
        }

        If ($IsLaptop) {      
            $AutoswitchingSettingName = 'LAN / WLAN Auto Switching'
            switch ($HWModel) {
                'HP EliteBook 840 G3'
                {
                    It "has BIOS setting '$AutoswitchingSettingName' set to 'Enabled'." {
                        ($BIOSSettings | Where-Object Name -eq "$AutoswitchingSettingName").Currentvalue | Should Be 'Enabled'
                    }
                }
                'HP EliteBook 820 G3'
                {
                    It "has BIOS setting '$AutoswitchingSettingName' set to 'Enabled'." {
                        ($BIOSSettings | Where-Object Name -eq "$AutoswitchingSettingName").Currentvalue | Should Be 'Enabled'
                    }
                }
                default
                {
                    It "has BIOS setting '$AutoswitchingSettingName' set to 'Enable'." {
                        ($BIOSSettings | Where-Object Name -eq "$AutoswitchingSettingName").Currentvalue | Should Be 'Enable'
                    }
                }       
            }

            It "has Video Memory Size set to 512 MB." {
                ($BIOSSettings | Where-Object Name -eq 'Video Memory Size').CurrentValue | Should Be '512 MB'
            }
        }   
    } #context

    Context 'SCCM Client Health' {
        It "has an SCCM client installed." {
            (Get-CimInstance -NameSpace Root\CCM -Class Sms_Client).ClientVersion -like "5.00.*" | Should Be True
        }

        $ServicesList =@('SMS Agent Host','Windows Management Instrumentation','Configuration Manager Remote Control')

        foreach ($Service in $ServicesList) {
            It "has the required service '$Service' running." {
                $(Get-Service -DisplayName $Service).status | Should Be 'Running'
            }
        }

        It "has been assignd to site 'SCT'" {
            $sms = New-Object   -COMObject 'Microsoft.SMS.Client'
            $sms.GetAssignedSite() | Should Be 'SCT'
        }

        It 'is not in running provisioning mode.' {
            (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\CcmExec' -Name "ProvisioningMode").ProvisioningMode | Should Be 'false'
        }
        
    } #context

    Context "Active Directory and Group Policy" {
        If ($IsLaptop) {
            It "complies with the PC naming convention for a laptop. (begins with 'LT' or 'TST-LT' followed by 6 digits." {
                $env:computername | Should Match "^(TST-LT|LT)\d{6}"    
            }
        }

        If ($IsDesktop) {
            It "complies with the PC naming convention for a desktop. (begins with 'DT' or 'PM' or 'TST-DT' followed by 6 digits." {
                $env:computername | Should Match "^(TST-DT|DT|PM)\d{6}"    
            }
        }

        It "is joined to the domain." { 
            $ComputerSystem.PartOfDomain | Should Be 'True'
        }
        #generalise valid OU path
        $SOEVersion='2.6'
        If ($IsDesktop) {$HWTypeOU = 'Desktop'} else {$HWTypeOU = 'Portable'}
        If ($env:computername -match "TST-*") {
            $HWEnvironment = ' (Testing)'}
        else 
        {
            $HWEnvironment = $null
        }
        If ($null -eq $HWEnvironment) {
            #production
            $OU = "OU=Azure AD,OU=SOE PCs $SOEVersion,OU=$HWTypeOU,OU=PCs$HWEnvironment,OU=SCTS,DC=scotcourts,DC=local"
        }
        else {
            #testing
            $OU = "OU=$HWTypeOU,OU=CIS 1.5 Applied,OU=SOE - PC - Beta $SOEVersion,OU=PCs$HWEnvironment,OU=SCTS,DC=scotcourts,DC=local"
        }
            
        It "is in the correct OU ($OU)." {
            $testou = ((gpresult /r /scope:computer) -match "OU=$HWTypeOU")
            $testOU -replace "    CN=$env:COMPUTERNAME,",'' | should be "$OU"
        }
    } #context


    context 'LAPS' {
        It 'has a LAPS client DLL registered' {
            (Test-Path -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GPExtensions\{D76B9641-3288-4f75-942D-087DE603E3EA}') | Should Be True
        }
    }

    Context "Bitlocker Encryption Status" {
        It 'has a TPM present.' {
            (Get-TPM).TPMPresent | Should Be 'True'
        }

        It "has a TPM that is ready." {
            (Get-TPM).TPMReady | Should Be 'True'
        }

        $VolumeStatus = (Get-BitLockerVolume -MountPoint C).VolumeStatus
        It "has a Bitlocker status for drive C - that is 'EncryptionInProgress' or 'FullyEncrypted'." {
             
            ($VolumeStatus -eq 'EncryptionInProgress') -or ($VolumeStatus -eq 'FullyEncrypted')| Should Be True
        }

        It "has Bitlocker Protection Status of volume C - 'On'." {
            (Get-BitLockerVolume -MountPoint C).ProtectionStatus | Should Be 'On'
        }
    } #context

    Context "Certificates" {
        It "has a valid 'SSL Secured Remote Desktop' certificate." {
            $SMSCert = Get-Childitem -Path cert:\LocalMachine\My |
            Where-Object {($_.EnhancedKeyUsageList -like 'SSL Secured Remote Desktop (1.3.6.1.4.1.311.54.1.2)') -and  ($_.NotAfter -gt $(Get-Date))} 
            ($SMSCert | Measure-Object).count | Should BeGreaterThan 0
        }
        It "has two valid 'Client Authentication' certificates." {
            $SMSCerts = Get-Childitem cert:\LocalMachine\My | 
            Where-Object {($_.EnhancedKeyUsageList -like 'Client Authentication (1.3.6.1.5.5.7.3.2)') -and  ($_.NotAfter -gt $(Get-Date))} 
            ($SMSCerts | Measure-Object).count | Should Be '2'
        }
        It "has a valid 'SMS Encryption Certificate' certificate." {
            $SMSCerts = Get-Childitem cert:\LocalMachine\SMS | 
            Where-Object {($_.EnhancedKeyUsageList -like '1.3.6.1.4.1.311.101.2') -and  ($_.FriendlyName -eq 'SMS Encryption Certificate') -and ($_.NotAfter -gt $(Get-Date))} 
            ($SMSCerts | Measure-Object).count | Should Be '1'
        }
        It "has a valid 'SMS Signing Certificate' certificate." {
            $SMSCerts = Get-Childitem cert:\LocalMachine\SMS | 
            Where-Object {($_.EnhancedKeyUsageList -like '1.3.6.1.4.1.311.101') -and  ($_.FriendlyName -eq 'SMS Signing Certificate')  -and  ($_.NotAfter -gt $(Get-Date))} 
            ($SMSCerts | Measure-Object).count | Should Be '1'
        }
    } #context

    Context "Network Adapter" {
        It "is currently using the 'Ethernet' adapter" {
            ((Get-NetAdapter -Name Ethernet).status -eq 'Up') | Should Be True
        }

    
        If ((Get-NetAdapter -Name Ethernet ).InterfaceDescription -like "Intel*") {
            It "found Intel NIC advanced property 'Wait on Link' - 'On'." {
                (Get-NetAdapterAdvancedProperty -RegistryKeyword WaitAutoNegComplete).registryValue | Should Be '1'
            }
        }   
    } #context

    Context "Hardware Drivers" {
        It "has no missing drivers." {
            $MissingDriverCount = (Get-CimInstance Win32_PNPEntity |
                Where-Object {($_.configManagerErrorCode -ne 0) -and ($_.configManagerErrorCode -ne 22) -and ($_.Manufacturer -eq $null)} | 
                Measure-Object).count
            $MissingDriverCount | Should Be 0

        }
    } #context

    Context "Base Applications" {
        $Applist = @('Microsoft Office Professional Plus 2016',
                     '7-Zip',
                     'Cisco Jabber',
                     'Google Chrome',
                     'Java 8 Update',
                     'MDOP MBAM',
                     'Adobe Acrobat Reader'
                     )
        If (-not $IsDesktop) {$AppList += 'Big-IP Edge Client'}
        If ($OPGMachine -ieq 'Y') {$AppList += 'PDF-Viewer';$AppList += 'PDF-XChange 4'}
        foreach ($app in $AppList) {
            It "has '$App' installed" {
                $x64Install = (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall" | 
                    Where-Object {$_.GetValue("DisplayName") -like "*$App*"} | Measure-Object).count
                $x86Install = (Get-ChildItem  "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" |
                    Where-Object {$_.GetValue("DisplayName") -like "*$App*"} | Measure-Object).count
                ($x64Install -or $x86Install) | should be 'True'       
            } #It
        } #foreach
    } #context
    
    Context "SCTS Standard Court Apps" {
        $StandardAppList = @('COPII LIVE (X64) V2',
                'Java Runtime [2021]',
                'COPII - CMS Data_Docs folder',
                'CS Solutions - JP System',
                'CS Solutions - Jury System',
                'DVLA Portal',
                'Objective Connect',
                'PAWS',
                'QFM-Concept Reach',
                'ICMS Azure',
                'Webex 43.4.0.25959 (x64)',
                'Docbox'
                )
        $InstalledApplications = (Get-CimInstance -Namespace "root\ccm\ClientSDK" -ClassName CCM_Application | where-object InstallState -eq 'Installed').name
        Foreach ($StandardApp in $StandardAppList) {
            It "has '$StandardApp' installed" {
                ($InstalledApplications -contains $StandardApp) | should be 'True'
            } #It
        } #foreach
    } #context

    Context "Apps Removed" {
        $AppsList ="Microsoft.MicrosoftStickyNotes","Microsoft.3DBuilder","microsoft.windowscommunicationsapps","Microsoft.MicrosoftOfficeHub","Microsoft.SkypeApp","Microsoft.MicrosoftSolitaireCollection","Microsoft.Office.OneNote","Microsoft.People","Microsoft.XboxApp", "Microsoft.Messaging", "Microsoft.Microsoft3DViewer", "Microsoft.OneConnect", "Microsoft.XboxSpeechToTextOverlay", "Microsoft.XboxIdentityProvider", "Microsoft.XboxGamingOverlay", "Microsoft.XboxGameOverlay", "Microsoft.Xbox.TCUI", "Microsoft.StorePurchaseApp"
        ForEach ($App in $AppsList) {
            It "has had '$app' removed." {
                (Get-AppxProvisionedPackage -online | Where-Object {$_.Displayname -eq $App}).PackageName | Should Be $null
            } #It
        } #foreach
    } #context
}
<#
        $Applist = @('COPII LIVE (X64) V2' {871A7010-FEDF-43AE-97E7-900925899C11},
        'Java Runtime [2021]' test-path %Windir%\Sun\Java\Deployment\ and %Windir%\Sun\Java\Deployment\exception.sites has modification date 01/11/2022 07:41:00 or newer ,
        'COPII - CMS Data_Docs folder' test-path C:\DATA\docs - this must be a folder,
        'CS Solutions - JP System' test-path  C:\JPSystem\JRJPs_FE_2016.accdb must have modification 01/06/2021 16:10:48 or newer,
        'CS Solutions - Jury SYstem' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\CS Solutions\Jury System AND test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\CS Solutions\Jury System\Jury System via IE11.lnk ,
        'DVLA Portal' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\DVLA Portal\DVLA Portal.url,
        'Objective Connect' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Objective Connect\Objective Connect.url,
        'PAWS' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\PAWS\PerTemps e-Services.url,
        'QFM - Concept Reach' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\QFM - Concept Reach\QFM - Concept Reach.url,
        'ICMS Azure' test-path C:\ProgramData\Microsoft\Windows\Start Menu\Programs\ICMS\ICMS (Azure).url,
        'Webex 43.4.0.25959 (x64)',
        'Docbox' C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Docbox\Docbox.url)
        #>