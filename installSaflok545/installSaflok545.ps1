<#
.Synopsis
   Helper Script to install SAFLOK SYSTEM 6000 automaticlly
.DESCRIPTION
   This script fully installation for SAFLOK System 6000 automatically
.EXAMPLE
   .\install.ps1 -installRoot [c|d]
.NOTES
    Author: renzo.xie@dormakaba.com
    Saflok version: v5.45, Marriott ONLY
    Create Date: 16 April 2019
    Modified Date: 12 AUG 2019
    *** New Update ***
    digital polling config compatible with new Cert
    add mini requirement for powershell (5.0 or above)
#>
# ---------------------------
# ==> [SECTION 01, VARIALBLES]
# ---------------------------
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit 
}

Function Write-Colr {
    Param ([String[]]$Text,[ConsoleColor[]]$Colour,[Switch]$NoNewline=$false)
    For ([int]$i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -Foreground $Colour[$i] -NoNewLine }
    If ($NoNewline -eq $false) { Write-Host '' }
}

Function Logging ($state, $message) {
    $part1 = $cname;$part2 = ' ';$part3 = $state;$part4 = ": ";$part5 = "$message"
    Switch ($state)
    {
        ERROR {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Red,Red,Red}
        WARN  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Magenta,Magenta,Magenta}
        INFO  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Yellow,Yellow,Yellow}
        ""   {Write-Colr -Text $part1,$part2,$part5 -Colour White,White,Cyan}
        default { Write-Colr -Text $part1,$part2,$part5 -Colour White,White,White}
    }
}
Function Stop-Script {
    Start-Sleep -Second 300   
    exit
}

If ($PSVersionTable.PSVersion.Major -lt 5) {
    Logging "WARN" "Your PowerShell installation is not version 5.0 or greater."
    Logging "WARN" "This script requires PowerShell version 5.0 or greater to function."
    Logging "WARN" "You can download PowerShell version 5.0 at: https://www.microsoft.com/en-us/download/details.aspx?id=50395"
    Stop-Script
} # Mini Powershell version requirement
# ----------------------------------------------------------------------------------------------
# GENERAL
$cname = '[dormakaba]'
$time = Get-Date -Format 'yyyy/MM/dd HH:mm'
$shareName = 'SaflokData'
$hotelChain = 'Marriott Projects ONLY'
[double]$winOS = [string][environment]::OSVersion.Version.major + '.' + [environment]::OSVersion.Version.minor
# ----------------------------------------------------------------------------------------------
# SCRIPT LOCATION
$scriptPath = $PSScriptRoot
# ----------------------------------------------------------------------------------------------
# OBJECT for saflokVersions
$saflokVersions = @(
    [pscustomobject]@{name='scriptVersion'; version='1.6'}
    [pscustomobject]@{name='mainVersion'; version='5.45'}
    [pscustomobject]@{name='progVersion'; version='5.4.0.0'}
    [pscustomobject]@{name='pmsVersion'; version='5.1.0.0'}
    [pscustomobject]@{name='msgrVersion'; version='5.2.0.0'}
    [pscustomobject]@{name='lensVersion'; version='4.7.0.0'}
    [pscustomobject]@{name='wsPmsExeBeforePathVersion'; version='4.7.1.15707'}
    [pscustomobject]@{name='gatewayExeVersion'; version='4.7.2.22694'}
    [pscustomobject]@{name='hmsExeVersion'; version='4.7.1.22400'}
    [pscustomobject]@{name='kdsExeVersion'; version='4.7.0.26769'}
    [pscustomobject]@{name='wsPmsExeVersion'; version='4.7.2.22767'}
    [pscustomobject]@{name='pollingExeVersion'; version='4.5.0.28705'}
)
$scriptVersion = $saflokVersions[0].version
$saflokVersion = $saflokVersions[1].version
$installDrive
$menuOption = 99
Clear-Host
Logging "" "+---------------------------------------------------------"
Logging "" "| $time"
Write-Colr -Text $cname, " |"," $hotelChain" -Colour White,cyan,Yellow
Logging "" "| SCRIPT VERSON: $scriptVersion" 
IF ($winOS -le 6.1) {Logging "" "| INSTALLING ON OS: WINDOWS 7 / SERVER 2008 R2 OR LOWER"}
				Else {Logging "" "| INSTALLING ON OS: WINDOWS 10 / SERVER 2012 or GREATER"}
Logging "" "+---------------------------------------------------------"
Logging "" "| WELCOME TO SAFLOK SYSTEMS INSTALL SCRIPT"
Logging "" "| SAFLOK VERSION: $saflokVersion"
Logging "" "+---------------------------------------------------------"
Logging " " ""
Logging " " "1 - Install to drive C"
Logging " " "2 - Install to drive D"
Logging " " "0 - Exit"
Logging " " ""
$menuOption = Read-Host "$cname Please select option from above list"
Logging "" "+---------------------------------------------------------"
Switch ($menuOption) {
	1 {$script:installDrive = 'C';$menuOption = 99}
	2 {$script:installDrive = 'D';$menuOption = 99}
	0 {Clear-Host;Exit}
	Default {
		Logging " " "Please enter a valid option."
	}
}
# ----------------------------------------------------------------------------------------------
# DRIVE INFO
$installDrive = $installDrive.ToUpper()
$driveLetter = $installDrive
$installDrive = $driveLetter + ':\'
$deviceID = $driveLetter + ':'
$cdDrive = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 5} | Select-Object DeviceID
If ($cdDrive -match $deviceID) {
    Write-Host ''
    Logging "ERROR" "The drive $driveLetter is not a valid location."
    Write-Host ''
    Stop-Script
}

# ----------------------------------------------------------------------------------------------
# packageFolders
$packageFolders =  Get-ChildItem ($scriptPath) | Select-Object Name | Sort-Object -Property Name
$absPackageFolders = @()
For ($i=0; $i -lt ($packageFolders.Length -1); $i++) {
    $absPackageFolders += $scriptPath + '\' + $packageFolders[$i].Name
}
# [01/02] ISS_FOR_Drive & files
# ----------------------------------------------------------------------------------------------
If ($driveLetter -eq 'C') {$iss4Drive = '01.ISS_FOR_' + $driveLetter}
Elseif ($driveLetter -eq 'D') {$iss4Drive = '02.ISS_FOR_' + $driveLetter}
Else {    Logging "ERROR" "Please input C or D for the drive letter."; Stop-Script}
Switch ($iss4Drive)
{
    01.ISS_FOR_C {$issFolder = $absPackageFolders[1]}
    02.ISS_FOR_D {$issFolder = $absPackageFolders[2]}
} 
$issFiles = Get-ChildItem $issFolder | Select-Object Name | Sort-Object -Property Name
$progISS = $issFolder + '\' + $issFiles[0].Name            # [0] Programsetup.iss
$pmsISS = $issFolder + '\' + $issFiles[1].Name             # [1] PMSsetup.iss
$msgrISS = $issFolder + '\' + $issFiles[2].Name            # [2] MSGRsetup.iss
$lensISS = $issFolder + '\' + $issFiles[3].Name            # [3] LENSsetup.iss
$patchLensISS = $issFolder + '\' + $issFiles[4].Name       # [4] patchLens474.iss
$patchPollingISS = $issFolder + '\' + $issFiles[5].Name    # [5] patchPolling450.iss
# ----------------------------------------------------------------------------------------------
# [03] saflokProgram
$progExe = $absPackageFolders[3] + '\' + 'setup.exe'
# ----------------------------------------------------------------------------------------------
# [04] saflokPMS
$pmsExe = $absPackageFolders[4] + '\' + 'setup.exe'
# ----------------------------------------------------------------------------------------------
# [05] saflokMessenger
$msgrExe = $absPackageFolders[5]  + '\' + 'setup.exe'
# ----------------------------------------------------------------------------------------------
# [06] saflokLENS [FOLDER]
$lensExe = $absPackageFolders[6] + '\' + 'AutoPlay\Install Script\Lens\en\setup.exe'
$sqlExprExe = $absPackageFolders[6] + '\' + 'AutoPlay\Install Script\Lens\en\ISSetupPrerequisites\{C38620DE-0463-4522-ADEA-C7A5A47D1FF6}\SQLEXPR_x86_ENU.exe'
# ----------------------------------------------------------------------------------------------
# [07]PATCH FILES
$patchExeFiles = Get-ChildItem ($absPackageFolders[7]) | Select-Object Name | Sort-Object -Property Name
$lensPatchExe = $absPackageFolders[7] + '\' + $patchExeFiles[1].Name     # [1] LENSPatch4.7.4
$pollingPatchExe = $absPackageFolders[7] + '\' + $patchExeFiles[0].Name  # [0] DigitalKeysPollingPatch
# ----------------------------------------------------------------------------------------------
# [08] Web service PMS tester [FOLDER]
$webServiceTester = $absPackageFolders[8]
# ----------------------------------------------------------------------------------------------
# [09] ConfigFiles Folder
$configFiles = Get-ChildItem ($absPackageFolders[9]) | Select-Object Name | Sort-Object -Property Name
$pollingConfig = $absPackageFolders[9] + '\' + $configFiles[0].Name     # [0] DigitalKeysPollingService.exe.config
$lensPmsConfig = $absPackageFolders[9] + '\' + $configFiles[1].Name     # [1] LENS_PMS.exe.config
# ----------------------------------------------------------------------------------------------
# Installed Location
$installedFolder = New-Object -TypeName psobject    # different if version is 6.2 or above
$installedFolder | Add-Member -MemberType NoteProperty -Name kabaInstFolder -Value 'Program Files (x86)\KABA'
$installedFolder | Add-Member -MemberType NoteProperty -Name lensInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens'
$installedFolder | Add-Member -MemberType NoteProperty -Name wsTesterInstFolder -Value 'Program Files (x86)\KABA\08.Web_Service_PMS_Tester'
$installedFolder | Add-Member -MemberType NoteProperty -Name wsPmsInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens\PMS Service'
$installedFolder | Add-Member -MemberType NoteProperty -Name saflokV4InstFolder -Value 'SaflokV4'
$installedFolder | Add-Member -MemberType NoteProperty -Name hubGateWayInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens\HubGatewayService'
$installedFolder | Add-Member -MemberType NoteProperty -Name hmsInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens\HubManagerService'
$installedFolder | Add-Member -MemberType NoteProperty -Name kdsInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens\KeyDeliveryService'
$installedFolder | Add-Member -MemberType NoteProperty -Name digitalKeyPollingInstFolder -Value 'Program Files (x86)\KABA\Messenger Lens\DigitalKeysPollingSoftware'
$installedFolder | Add-Member -MemberType ScriptMethod -name absLensInstFolder -value {$installDrive + $installedFolder.lensInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absWsPmsInstFolder -value {$installDrive + $installedFolder.wsPmsInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absSaflokV4InstFolder -value {$installDrive + $installedFolder.saflokV4InstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absKabaInstFolder -value {$installDrive + $installedFolder.kabaInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absHubGateWayInstFolder -value {$installDrive + $installedFolder.hubGateWayInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absHmsInstFolder -value {$installDrive + $installedFolder.hmsInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absKdsInstFolder -value {$installDrive + $installedFolder.kdsInstFolder}
$installedFolder | Add-Member -MemberType ScriptMethod -name absDigitalKeyPollingInstFolder -value {$installDrive + $installedFolder.digitalKeyPollingInstFolder}
# ----------------------------------------------------------------------------------------------
# Absolute installed folder
$lensInstFolder = $installDrive + $installedFolder.lensInstFolder
$wsPmsInstFolder = $installDrive + $installedFolder.wsPmsInstFolder
$kabaInstFolder = $installDrive + $installedFolder.kabaInstFolder
$wsTesterInstFolder = $installDrive + $installedFolder.wsTesterInstFolder
$digitalPollingFolder = $installDrive + $installedFolder.digitalKeyPollingInstFolder
# ----------------------------------------------------------------------------------------------
# GUI exe file in SAFLOKV4 FOLDER
$saflokClient = $installedFolder.absSaflokV4InstFolder() + '\' + 'Saflok_Client.exe'
$saflokMsgr = $installedFolder.absSaflokV4InstFolder() + '\' + 'Saflok_MsgrServer.exe'
$saflokIRS = $installedFolder.absSaflokV4InstFolder()  + '\' + 'Saflok_IRS.exe'
$shareFolder = $installedFolder.absSaflokV4InstFolder() + '\' + 'SaflokData'
# ----------------------------------------------------------------------------------------------
# EXE files for version inforamtion
$gatewayExe = $installedFolder.absHubGateWayInstFolder() + '\' + 'LENS_Gateway.exe'
$hmsExe = $installedFolder.absHmsInstFolder() + '\' + 'LENS_HMS.exe'
$wsPmsExe = $installedFolder.absWsPmsInstFolder() + '\' + 'LENS_PMS.exe'
$kdsExe = $installedFolder.absKdsInstFolder() + '\' +  'Kaba_KDS.exe'
# ----------------------------------------------------------------------------------------------
# CONFIG FILES - INSTALLED
$lensPmsConfigFileInst = $installedFolder.absWsPmsInstFolder() + '\' + 'LENS_PMS.exe.config'
$hh6ConfigFile = $installedFolder.absSaflokV4InstFolder()  + '\' + 'KabaSaflokHH6.exe.config'
# digital folder and exe installed
# ----------------------------------------------------------------------------------------------
$digitalPollingExe = $installedFolder.absDigitalKeyPollingInstFolder() + '\' + 'DigitalKeysPollingService.exe'
$pollingConfigInst = $installedFolder.absDigitalKeyPollingInstFolder() + '\' + 'DigitalKeysPollingService.exe.config'
# ----------------------------------------------------------------------------------------------
# Polling log
$pollingLog = 'C:\ProgramData\DormaKaba\Server\Polling\logs.log'
# Saflok Services
# ----------------------------------------------------------------------------------------------
$serviceNames = [ordered]@{
    deviceMng = 'DeviceManagerService';
    kIpEncoderSrv = 'KIPEncoderService';
    firebirdSrv = 'FirebirdGuardianDefaultInstance';
    saflokLauncherSrv = 'SaflokServiceLauncher';
    saflokCRSSrv = 'SaflokCRS';
    saflokSchedulerSrv = 'SaflokScheduler';
    saflokIRSSrv = 'SaflokIRS';
    saflokMSGRSrv = 'SaflokMSGR';
    saflokDHSP2MSGR = 'SAFLOKDHSPtoMSGRTranslator';
    saflokMSGR2DHSP = 'SAFLOKMSGRtoDHSPTranslator';
    hubGatewaySrv = 'MessengerNet_Hub Gateway Service'; #[10]
    hubManagerSrv = 'MNet_HMS';                         #[11]
    pmsSrv = 'MNet_PMS Service';                        #[12]
    utilitySrv = 'MessengerNet_Utility Service';
    kdsSrv = 'Kaba_KDS';                                #[14]
    virtualEncoderSrv = 'VirtualEncoderService'
    pollingSrv = 'Kaba Digital Keys Polling Service'    #[16]
}
# ----------------------------------------------------------------------------------------------
# Logging Messages
$mesgNoPkg ="package does not exist, operation exit."
$mesgInstalled = "has already been installed."
$mesgDiffVer = "There is another version exist, please uninstall it first."
$mesgComplete = "installation is complete."
$mesgFailed = "installation failed!"
# ----------------------
# [SECTION 02, FUNCTIONS]
# ----------------------
# check installed version
Function Test-Folder ($folder) {
    Test-Path -Path $folder -PathType Any
} # Test Folder return boolean
Function Get-InstVersion {
    Param ([String[]]$pName)
    [String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName}).Version
}
Function Get-FileVersion ($testFIle) {
    (Get-Item $testFile).VersionInfo.FileVersion
}
Function Assert-IsInstalled ($pName) {
    $findIntallByName = [String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName})
    $condition = ($null -ne $findIntallByName)
    ($true, $false)[!$condition]
} # return Boolean value
Function Update-Status ($pName) {
    Assert-IsInstalled $pName | Out-Null; If (Assert-IsInstalled $pName) { $script:isInstalled = $true}
    $packageFolder | Out-Null
    $curVersion | Out-Null
    $exeExist | Out-Null
} 
Function Install-Prog ($pName,$packageFolder,$curVersion,$exeExist,$destVersion,$exeFile,$issFile) {
    If ($isInstalled) {
        #check current installed version
        If ($curVersion -eq $destVersion) {Logging "INFO" "$pName $mesgInstalled"} Else {Logging "ERROR" "$mesgDiffVer - $pName";Stop-Script}
    } Else {
        If ($packageFolder -eq $false) {Logging "ERROR" "$pName $mesgNoPkg";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $true)){Logging "ERROR" "$mesgDiffVer";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $false)) {
            Logging " " "Processing installation for $pName, Please wait ..."
            Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
            $installed = Assert-IsInstalled $pName
            If ($installed) {Logging " " "$pName $mesgComplete"} Else {Logging "ERROR" "$pName $mesgFailed";Stop-Script}
        }
    }
}
Function Update-FileVersion ($targetFile, $verToMatch) {
    If (Test-Folder $targetFile) {
        Get-FileVersion $targetFile | Out-Null
        If ((Get-FileVersion $targetFile) -eq $verToMatch) {$script:isInstalled = 1}
        Else {$script:isInstalled = 0}
    }
}
Function Install-LensPatch ($targetFile,$destVersion,$pName,$exeFile,$issFile) {
    If ($isInstalled) {
        Logging "INFO" "$pName $mesgInstalled"
    } Else {
        Logging " " "Processing $pName, Please wait ..."
        Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait; Start-Sleep 3
        Update-FileVersion $targetFile $destVersion
        If ((Get-FileVersion $targetFile) -eq $destVersion) {
            Logging " " "$pName $mesgComplete"
        } Else {
            Logging "ERROR" "$pName $mesgFailed"
            Stop-Script
        }
    }
} # Install Messenger Lens Patch
Function Install-DigitalPolling ($pName, $targetFile,$exeFile,$issFile) {
    If (Test-Folder $targetFile) {
        Logging "INFO" "$pName $mesgInstalled"
    } Else {
        Logging " " "Processing installation for $pName, Please wait ..."
        Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
        If (Test-Folder $targetFile){Logging " " "$pName $mesgComplete"}
        Else {Logging "Error" "$pName $mesgFailed";Stop-Script}
    }
}# Install Digital Polling Service
Function Update-Copy ($srcPackage,$instFolder0,$instFolder1) {
    $testSrcPackage  = Test-Folder $srcPackage; $testSrcPackage | Out-Null
    $testInstFolder0 = Test-Folder $instFolder0; $testInstFolder0 | Out-Null
    $testInstFolder1 = Test-Folder $instFolder1; $testInstFolder1 | Out-Null
    If (($testInstFolder0 -eq $false) -and ($testInstFolder1)) { $script:fileCopied = 1 }
} # Update copied files
Function Install-PmsTester ($srcPackage,$instParent,$instFolder0,$instFolder1) {
    $testSrcPackage  = Test-Folder $srcPackage
    $testInstParent = Test-Folder $instParent
    $testInstFolder0 = Test-Folder $instFolder0
    $testInstFolder1 = Test-Folder $instFolder1
    If ($testSrcPackage -eq $False) { Logging "ERROR" "Package files missing!"; Stop-Script } 
    If ($testInstParent -eq $False) { Logging "ERROR" "Messenger LENS has not been installed yet!" ; Stop-Script } 
    If ($fileCopied) {
        Logging "INFO" "Web Service PMS Tester $mesgInstalled."
    } Else {
        If (($testSrcPackage -eq $true) -and ($testInstParent -eq $true)) {Logging " " "Installing Web Service PMS Tester ..." }
        If (($testInstFolder0 -eq $true) -and ($testInstFolder1 -eq $false))  { 
            Rename-Item -Path $instFolder0 -NewName $instFolder1 -force -ErrorAction SilentlyContinue
            Logging " " "Web Service PMS Tester $mesgComplete."
        } Elseif (($testInstFolder0 -eq $false) -and ($testInstFolder1 -eq $false)) {
            Copy-Item $srcPackage -Destination $instParent -Recurse -Force -ErrorAction SilentlyContinue
            Rename-Item -Path $instFolder0 -NewName $instFolder1 -force -ErrorAction SilentlyContinue
            Logging " " "Web Service PMS Tester $mesgComplete."
        } Elseif (($testInstFolder0 -eq $true) -and ($testInstFolder1 -eq $true))  {
            Remove-Item -Path $instFolder0 -Force -Recurse
            Logging " " "Web Service PMS Tester $mesgComplete."
        } Else {
            Logging " " "Web Service PMS Tester $mesgComplete."
        }
    }
} # Install Web Service PMS Tester
Function New-Share {
    param([string]$shareName,[string]$shareFolder)
	net share $shareName=$shareFolder "/GRANT:Everyone,FULL" /REMARK:"Saflok Database Folder Share"
} # Share Folder for windows 2008 R2 or lower
Function Install-SC {
    Param([String]$exeFile,[String]$agrFile,[String]$service)
    Start-Process -NoNewWindow -FilePath $exeFile $service -ArgumentList " $argFile" -Wait -RedirectStandardOutput Out-Null
} # Install SC
Function Install-Sql ($pName,$packageFolder,$exeFile,$argFile) {
    If ($isInstalled -eq $true) {
        Logging "INFO" "$pName $mesgInstalled"
    } Else {
        If ($packageFolder -eq $false) {Logging "ERROR" "$pName $mesgNoPkg";Stop-Script}
        If ($packageFolder -eq $true) {
            Logging " " "Processing installation for $pName."
            Logging "INFO" "The installer is 116M+, this could take a while, please wait... "
            Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " $argFile" -Wait
            $installed = Assert-IsInstalled $pName
            If ($installed) {
				Logging " " "$pName $mesgComplete"
			} Else {
				Logging "ERROR" "$pName $mesgFailed"
				Logging "ERROR" "Reboot system and try the script again"
				Logging "ERROR" "If still same, please contact your SAFLOK representative."
				Stop-Script
			}
        }
    }
} # Install SQL
Function Update-SqlPasswd {
    Param([string]$login,[string]$passwd)
    $ServerNameList = 'localhost\lenssql'
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
    $objSQLConnection = New-Object System.Data.SqlClient.SqlConnection
    foreach($ServerName in $ServerNameList)
    {
        Try {
            $objSQLConnection.ConnectionString = "Server=$ServerName;Integrated Security=SSPI;"
            $objSQLConnection.Open() | Out-Null
            $objSQLConnection.Close()
        } Catch {
            Logging "ERROR" "Fail"
            $errText =  $Error[0].ToString()
            If ($errText.Contains("network-related")) {
                Logging "ERROR" "Connection Error. Check server name, port, firewall."
            }
            Logging "ERROR" "$errText"
            continue
        }
        $srv = New-Object "Microsoft.SqlServer.Management.Smo.Server" $ServerName
        $SQLUser = $srv.Logins | Where-Object {$_.Name -eq "$login"};
        $SQLUser.ChangePassword($passwd);
        $SQLUser.PasswordPolicyEnforced = 1;
        $SQLUser.Alter();
        $SQLUser.Refresh();
    }
} # update SQL password
Function Set-ServiceRecovery{
    [alias('Set-Recovery')]
    param
    (
        [string] [Parameter(Mandatory=$true)] $ServiceDisplayName,
        [string] $action1 = "restart",
        [int] $time1 =  50000, # in milliseconds
        [string] $action2 = "restart",
        [int] $time2 =  50000, # in milliseconds
        [string] $actionLast = "restart",
        [int] $timeLast = 50000, # in milliseconds
        [int] $resetCounter = 86400 # in seconds
    )
    #$recoveryServices
    $action = $action1 + "/" + $time1 + "/" + $action2 + "/" + $time2 + "/" + $actionLast + "/" + $timeLast
    $output = sc.exe $serverPath failure $service actions= $action reset= $resetCounter | Out-Null
    Return $output
} # set service recovery

# -----------------------
# ==> [SECTION 03, HEADER]
Logging "INFO" "You chose drive $driveLetter"
Logging " " ""
Logging " " "By installing you accept licenses for the packages."
$confirmation = Read-Host "$cname Do you want to run the script? [Y] Yes  [N] No"
$confirmation = $confirmation.ToUpper()
If ($confirmation -eq 'Y' -or $confirmation -eq 'YES') {
    Logging " " ""
    $psDrive = Get-Psdrive | Where-Object {$_.Name -eq $driveLetter -and ($_.Free -eq $null -or $_.Free -eq 0)}
    If ($psDrive) {Logging "ERROR" "The drive $driveLetter is not a valid location."; Stop-Script}
    # -------------------------------------------------------------------
    # install Saflok client
    $isInstalled = 0
    $pName = "Saflok Program"
    $packageFolder = Test-Folder ($absPackageFolders[3])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokClient
    $destVersion = ($saflokVersions[2].version)
    Update-Status $pName
    Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $progExe $progISS
    # -------------------------------------------------------------------
    # install Saflok PMS
    $pName = "Saflok PMS"
    $isInstalled = 0
    $packageFolder = Test-Folder ($absPackageFolders[4])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokIRS
    $destVersion = ($saflokVersions[3].version)
    Update-Status "Saflok PMS"
    Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $pmsExe $pmsISS
    # -------------------------------------------------------------------
    # install Saflok Messenger
    $pName = "Saflok Messenger Server"
    $isInstalled = 0
    $packageFolder = Test-Folder ($absPackageFolders[5])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokMsgr
    $destVersion = ($saflokVersions[4].version)
    Update-Status "$pName"
    Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $msgrExe $msgrISS
    # -------------------------------------------------------------------
    # copy database to saflokdata folder
    $srcHotelData = ($absPackageFolders[0])
    $srcGdb = (Get-ChildItem -Path $srcHotelData).Name
    If ((($srcGdb -match '^SAFLOKDATAV2.GDB$').Count -eq 0) -or (($srcGdb -match '^SAFLOKLOGV2.GDB$').Count -eq 0) -or ($null -eq $srcGdb)) {
        Logging "WARN" "Please copy database files to this folder:" 
        Logging "WARN" "$srcHotelData"
        Logging "WARN" "Try this script again after database files have been loaded."
        Write-Host ''
        Stop-Script
    }  
    $instGdb = (Get-ChildItem -Path $shareFolder).Name
    If ((($instGdb -match '^SAFLOKDATAV2.GDB$').Count -eq 0) -or (($instGdb -match '^SAFLOKLOGV2.GDB$').Count -eq 0)) {
        If ((Get-Service -Name FirebirdGuardianDefaultInstance).Status -eq "Running") { # stop firebird service
            Stop-Service -Name FirebirdGuardianDefaultInstance -Force -ErrorAction SilentlyContinue 
            Copy-Item -Path $srcHotelData\*.gdb -Destination $shareFolder        # copy database
        } Else {
            Copy-Item -Path $srcHotelData\*.gdb -Destination $shareFolder 
        }
    } 
    # -------------------------------------------------------------------
    # share database folder
    If ((Test-Path $shareFolder) -and ($winOS -le 6.1)) {
        If (!(net share | findstr "SaflokData")) {
            New-Share -shareName $shareName -shareFolder $shareFolder | Out-Null; Start-Sleep -S 1
        } Else {
            Logging "INFO" "The Saflok database folder already been shared."; Start-Sleep -S 1
        }
    } Elseif (!(Get-SmbShare | where-Object {$_.Name -eq $shareName}) -and ($winOS -gt 6.1) ) {
        New-SmbShare -Name $shareName -Path $shareFolder -FullAccess "everyone" -Description "Saflok database folder share" | Out-Null; Start-Sleep -S 1
    } Elseif ((Get-SmbShare | where-Object {$_.Name -eq $shareName}) -and ($winOS -gt 6.1) ) {
        Logging "INFO" "The Saflok database folder already been shared.";Start-Sleep -S 1
    } Else {
        Logging "ERROR" "Share folder does not exist"
        Logging "ERROR" "Please contact your system administrator"
        Start-Sleep -S 2
        Write-Host''
        Stop-Script
    }
    # -------------------------------------------------------------------
    # start firebird service
    $fbService = $serviceNames['firebirdSrv']
    $fbSvcStat = (Get-Service | Where-Object {$_.Name -eq $fbService}).Status
    If ($fbSvcStat -eq "Stopped"){Start-Service -Name $fbService;Start-Sleep -S 2}
    # -------------------------------------------------------------------
    # start saflok launcher service
    If (Get-Service | Where-Object {$_.Name -eq $serviceNames['saflokLauncherSrv'] -and $_.Status -eq "Stopped"}){
        Start-Service -Name ($serviceNames['saflokLauncherSrv'])
    }
    # -------------------------------------------------------------------
    # IIS FEATURES, requirement for messenger lens
    $msgrInstalled = Assert-IsInstalled "Saflok Messenger Server"
    If (!$msgrInstalled) {
        Logging "WARN" "Please install Saflok messenger first."; Stop-Script
    }
    Logging "" "Checking System Environment Parameter ..."
    If ($winOS -le 6.1) {
        Logging " " "Configuring IIS features for Messenger LENS, please wait ..."
        $iisFeatures = 'IIS-WebServerRole','IIS-WebServer','IIS-CommonHttpFeatures','IIS-HttpErrors','IIS-ApplicationDevelopment','IIS-RequestFiltering',`
            'IIS-NetFxExtensibility','IIS-HealthAndDiagnostics','IIS-HttpLogging','IIS-RequestMonitor','IIS-Performance','WAS-ProcessModel',`
            'WAS-NetFxEnvironment','WAS-ConfigurationAPI','IIS-ISAPIExtensions','IIS-ISAPIFilter','IIS-StaticContent','IIS-DefaultDocument',`
            'IIS-DirectoryBrowsing','IIS-ASPNET','IIS-ASP','IIS-HttpCompressionStatic','IIS-ManagementConsole','NetFx3','WCF-HTTP-Activation','WCF-NonHTTP-Activation'
        For ([int]$i=0; $i -lt ($iisFeatures.Length - 1); $i++) {
            $feature = $iisFeatures[$i]
            DISM /online /enable-feature /featurename:$feature | Out-Null
            Start-Sleep -S 1
            Logging " " "Enabled features $feature"
        }
    } # win7 / win server 2008 R2 or lower
    Else {
        $disabledFeatures = @()
        $iisFeatures = 'NetFx4Extended-ASPNET45','IIS-ASP','IIS-ASPNET45','IIS-NetFxExtensibility45','IIS-WebServerRole','IIS-WebServer', `
            'IIS-CommonHttpFeatures','IIS-HttpErrors','IIS-ApplicationDevelopment','IIS-HealthAndDiagnostics','IIS-HttpLogging','IIS-Security', `
            'IIS-RequestFiltering','IIS-Performance','IIS-WebServerManagementTools','IIS-StaticContent','IIS-DefaultDocument','IIS-DirectoryBrowsing', `
            'IIS-ApplicationInit','IIS-ISAPIExtensions','IIS-ISAPIFilter','IIS-HttpCompressionStatic','IIS-ManagementConsole'
        For ([int]$i=0; $i -lt ($iisFeatures.Length -1 ); $i++) {
            $feature = $iisFeatures[$i]
            If (!((Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -eq $feature}).State -eq "Enabled")) {
                $disabledFeatures += $feature
            }
        }
        If ($disabledFeatures.Count-1 -gt 0){
            Logging "INFO" "Configuring IIS features for Messenger LENS, please wait ..."
            Foreach ($disabled In $disabledFeatures) {
                Enable-WindowsOptionalFeature -Online -FeatureName $disabled -All | Out-Null
                Logging " " "Enabled features $disabled"
            }
        } Else {
            Logging "INFO" "ALL required IIS features have been enabled."
        }
        Start-Sleep -Seconds 1
    } # win10 / win server 2016 or above
    # -------------------------------------------------------------------
    # Microsoft SQL Server 2012
    $pName = 'Microsoft SQL Server 2012'
    $isInstalled = 0
    $packageFolder = Test-Folder ($sqlExprExe)
    $argFile = '/qs /INSTANCENAME="LENSSQL" /ACTION="Install" /Hideconsole /IAcceptSQLServerLicenseTerms="True" '
    $argFile += '/FEATURES=SQLENGINE,SSMS /HELP="False" /INDICATEPROGRESS="True" /QUIETSIMPLE="True" /X86="True" /ERRORREPORTING="False" '
    $argFile += '/SQMREPORTING="False" /SQLSVCSTARTUPTYPE="Automatic" /FILESTREAMLEVEL="0" /FILESTREAMLEVEL="0" /ENABLERANU="True" '
    $argFile += '/SQLCOLLATION="Latin1_General_CI_AS" /SQLSVCACCOUNT="NT AUTHORITY\SYSTEM" /SQLSYSADMINACCOUNTS="BUILTIN\Administrators" '
    $argFile += '/SECURITYMODE="SQL" /ADDCURRENTUSERASSQLADMIN="True" /TCPENABLED="1" /NPENABLED="0" /SAPWD="S@flok2018"'
    Update-Status $pName
    Install-Sql $pName $packageFolder $sqlExprExe $argFile
    If (Assert-IsInstalled 'Microsoft SQL Server 2012') {Update-SqlPasswd -login 'sa' -passwd 'Lens2014'}
    # -------------------------------------------------------------------
    # install Messenger Lens V4.7
    $pName = "Messenger LENS"
    $isInstalled = 0
    $packageFolder = Test-Folder ($absPackageFolders[6])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $wsPmsExe
    $destVersion = ($saflokVersions[5].version)
    Update-Status $pName
    Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $lensExe $lensISS
    # -------------------------------------------------------------------
    # install Messenger Lens Patch v4.74
    $pName = "Messenger Lens Patch"  
    $isInstalled = 0
    $targetFile = $wsPmsExe
    $destVersion = $saflokVersions[10].Version
    Update-FileVersion $wsPmsExe $destVersion
    Install-LensPatch $wsPmsExe $destVersion $pName $lensPatchExe $patchLensISS
    # -------------------------------------------------------------------
    # install digital polling service
    $isInstalled = 0
    $pName = "Marriott digital polling service"
    Install-DigitalPolling $pName $digitalPollingExe $pollingPatchExe $patchPollingISS
    # -------------------------------------------------------------------
    # allow everyone access to lens folder
    If (Test-Folder Container) {
        $acl = Get-Acl -Path $lensInstFolder
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","FullControl","Allow")
        $acl.SetAccessRule($AccessRule); $acl | Set-Acl $lensInstFolder
    }
    # -------------------------------------------------------------------
    # copy config files
    Copy-Item -Path $lensPmsConfig -Destination $wsPmsInstFolder -Force
    Copy-Item -Path $pollingConfig -Destination $digitalPollingFolder -Force
    # SECTION 05, INSTALL WEB SERVICE PMS TESTER
    # -------------------------------------------------------------------
    $fileCopied = 0 
    $newFolder0 = $kabaInstFolder + '\' + $wsTesterInstFolder.Substring($wsTesterInstFolder.Length - 25,25)
    $newFolder1 = $kabaInstFolder + '\' + $wsTesterInstFolder.Substring($wsTesterInstFolder.Length - 22,22)
    Update-Copy $webServiceTester $newFolder0 $newFolder1
    Install-PmsTester $webServiceTester $kabaInstFolder $newFolder0 $newFolder1
    $wsTesterExe = $newFolder1 + '\' + 'MessengerNet WSTestPMS.exe'
    If ((Test-Folder $wsTesterExe)) {
        $TargetFile = $wsTesterExe
        $ShortcutFile = "$env:Public\Desktop\WS_PMS_TESTER.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.IconLocation = "C:\Windows\System32\SHELL32.dll, 12" 
        $Shortcut.Save()
        Start-Sleep -S 1
    } # shotcut for web service pms tester
    # [SECTION 06, CHECK SERVICES STATUS]
    # ----------------------------------------------------------------
    [string[]]$servicesCheck = $serviceNames[0],$serviceNames[1],$serviceNames[14],$serviceNames[10],$serviceNames[11],$serviceNames[12],$serviceNames[13],$serviceNames[15],$serviceNames[16]
    # [0] DeviceManagerService,[1] KIPEncoderService,[14] Kaba_KDS,[10] MessengerNet_Hub Gateway Service,[11] MNet_HMS,[12] MNet_PMS Service,[13] MessengerNet_Utility Service,[15] VirtualEncoderService, [16] digitalPolling
    If (Assert-IsInstalled "Messenger LENS") {
        Foreach ($service In $servicesCheck) { 
            $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
            If ($serviceStatus.Status -eq "stopped") {
                Logging " " "Staring service $service."
                Start-Service -Name $service -ErrorAction SilentlyContinue
                $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
                If ($serviceStatus.Status -eq "running") { Logging " " "$service has been started."}
                Start-Sleep -S 1
            } Else {Logging " " "$service is in running state.";Start-Sleep -S 1}
        } # check service status
        # [SECTION 7, OPEN GUI AND CONFIG FILE]
        # ----------------------------------------------------------------
        Logging "" "+---------------------------------------------------------"
        Logging "" "The following files need to be checked or configure: "
        Logging "" "+---------------------------------------------------------"
        If ($Null -eq (Get-Process | where-object {$_.Name -eq 'Saflok_IRS'}).ID) {
            Start-Process -NoNewWindow -FilePath $saflokIRS; Start-Sleep -S 1
        } # run IRS GUI
        Get-Process -ProcessName notepad* | Stop-Process -Force; Start-Sleep -S 1 # kill all notepad
        If ((Assert-isInstalled "Saflok Program") -and (Test-Folder $hh6ConfigFile)) {
            Logging " " "[ KabaSaflokHH6.exe.config ]"
            Start-Process notepad $hh6ConfigFile -WindowStyle Minimized; Start-Sleep -S 1
        } # hh6 config
        If ((Assert-isInstalled  "Messenger LENS") -and (Test-Folder $lensPmsConfigFileInst)) {
            Logging " " "[ LENS_PMS.exe.config ]"
            Start-Process notepad $lensPmsConfigFileInst -WindowStyle Minimized; Start-Sleep -S 1
        } # PMS config
        If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Folder $pollingConfigInst)) {
            Logging " " "[ DigitalKeysPollingService.exe.config ]"
            Start-Process notepad $pollingConfigInst -WindowStyle Minimized; Start-Sleep -S 1
        } # polling config
        If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Folder $pollingLog)) {
            Logging " " "[ Polling log ]"
            Start-Process notepad $pollingLog -WindowStyle Minimized; Start-Sleep -S 1
            $TargetFile = $pollingLog
            $ShortcutFile = "$env:Public\Desktop\PollingLog.lnk"
            $WScriptShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
            $Shortcut.TargetPath = $TargetFile
            $Shortcut.Save()
            Start-Sleep -S 1
        } # shotcut for polling log
        Write-Colr -Text "$cname ","Please check those config files openning in taksbar area." -Colour White,Magenta
        Logging "" "+---------------------------------------------------------"
        Start-Sleep -Seconds 1
        # ----------------------------------------------------------------
        # [SECTION 08, SET SERVICE RECOVERY]
        [string[]]$recoveryServices = $serviceNames[10],$serviceNames[11],$serviceNames[12],$serviceNames[14],$serviceNames[16]
        Foreach ($service In $recoveryServices){
            If (Get-service -Name $service | where-object {$_.StartType -ne 'Automatic'}) { Set-Service $service -StartupType "Automatic" }
            Set-ServiceRecovery -ServiceDisplayName $service
        } # service recovery 
        # ----------------------------------------------------------------
        # [SECTION 09, FOOTER]
        Logging "" "Installed Version Infomation: " 
        Logging "" "+---------------------------------------------------------"
        $gatewayVer = Get-FileVersion $gatewayExe;$hmsVer = Get-FileVersion $hmsExe;$wsPmsVer = Get-FileVersion $wsPmsExe;$kdsVer = Get-FileVersion $kdsExe;$pollingVer = Get-FileVersion $digitalPollingExe
        If (Test-Path $gatewayExe -PathType Leaf) { Logging " " "Gateway: $gatewayVer"; Start-Sleep -Seconds 1 } 
        If (Test-Path $hmsExe -PathType Leaf) { Logging " " "HMS:     $hmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $wsPmsExe -PathType Leaf) { Logging " " "PMS:     $wsPmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $kdsExe -PathType Leaf) { Logging " " "KDS:     $kdsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $digitalPollingExe -PathType Leaf) { Logging " " "POLLING: $pollingVer"; Start-Sleep -Seconds 1 }
        Logging "" "+---------------------------------------------------------"
        Logging "" "DONE"
        Logging "" "+---------------------------------------------------------"
        Logging "" ""
        Logging "WARN" "The recent program changes indicate a reboot is necessary."
        Write-Host ''
        # clean up script files and SAFLOK folder
        If (Test-Path -Path "$scriptPath\*.*" -Include *.ps1){Remove-Item -Path "$scriptPath\*.*" -Include *.ps1,*.lnk -Force -ErrorAction SilentlyContinue}
        If (Test-path -Path "C:\SAFLOK") { Remove-Item -Path "C:\SAFLOK" -Recurse -Force -ErrorAction SilentlyContinue }   
        Start-Sleep -Second 300
    } # end of MessengerLens is installed 
} # END OF YES
If ($confirmation -eq 'N' -or $confirmation -eq 'NO') {
    Logging " " ""
    Write-Colr -Text $cname," Thank you, Bye!" -Colour White,Gray
    Write-Host ''
    Stop-Script
} # END OF NO