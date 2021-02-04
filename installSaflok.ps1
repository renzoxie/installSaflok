<#
.SYNOPSIS
   PS script helps to install SAFLOK Lodging systems
.DESCRIPTION
   This script fully installation of SAFLOK Lodging Systems for Marriott projects automatically
.EXAMPLE
   .\install.ps1 -inputDrive c -version 5.45
.EXAMPLE
   .\install.ps1 -inputDrive c -version 5.45 -vendor 'dormakaba' -property 'Hotel Name'
.NOTES
    Author: renzoxie@139.com
    Version Option: 5.45, 5.68
    Create Date: 16 April 2019
    Modified Date: 1st Feb 2021
#>
[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$TRUE)]
    [ValidateSet('c','d')]
    [String]$inputDrive,

    [Parameter(Mandatory=$TRUE)]
    [String]$version,
    
    [String]$property = 'vagrant',
    
    [String]$vendor = 'dormakaba'
)

# ---------------------------
# Script location
$scriptPath = $PSScriptRoot
# ---------------------------
# Versions
$scriptVersion = '2.0'
Switch ($version) {
    '5.45' {
        $progVersion = '5.4.0.0'
        $pmsVersion = '5.1.0.0'
        $msgrVersion= '5.2.0.0'
        $lensVersion = '4.7.0.0'
        $gatewayExeVersion = '4.7.2.22694'
        $hmsExeVersion = '4.7.1.22400'
        $kdsExeVersion = '4.7.0.26769'
        $pollingExeVersion = '4.5.0.28705'  
        $wsPmsExeBeforePatchVersion = '4.7.1.15707'
        $wsPmsExeVersion = '4.7.2.22767'
    }
    '5.68' {
        $ver1 = '5.6.0.0'
        $ver2 = '5.6.8.0'
        $wsPmsExeBeforePatchVersion = '5.6.0.0'
        $wsPmsExeVersion = '5.6.7.22261'
        $msgrVersion = '5.6.0.0'
    }
}
# Logging Messages
$mesgNoPkg ="package does not exist, operation exit."
$mesgInstalled = "has already been installed."
$mesgDiffVer = "There is another version exist, please uninstall it first."
$mesgComplete = "installation is complete."
$mesgFailed = "installation failed!"
$mesgNoSource = "Missing source installation folder."
# ---------------------------
# Functions 
# ---------------------------
# Customized color
Function Write-Colr {
    Param ([String[]]$Text,[ConsoleColor[]]$Colour,[Switch]$NoNewline=$false)
    For ([int]$i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -Foreground $Colour[$i] -NoNewLine }
    If ($NoNewline -eq $false) { Write-Host '' }
}
# ---------------------------
# Customized logging
Function Logging ($state, $message) {
    $part1 = $cname;
    $part2 = ' ';
    $part3 = $state;
    $part4 = ": ";
    $part5 = "$message"
    Switch ($state)
    {
        ERROR {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Red,Red,Red}
        WARN  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Magenta,Magenta,Magenta}
        INFO  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Yellow,Yellow,Yellow}
        PROGRESS {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,White,White,White}
        WARNING  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Yellow,Yellow,Yellow}
        ""   {Write-Colr -Text $part1,$part2,$part5 -Colour White,White,Cyan} 
        default { Write-Colr -Text $part1,$part2,$part5 -Colour White,White,White}
   } 
} 
# ---------------------------
# Stop Script
Function Stop-Script {
    Start-Sleep -Seconds 60 
    exit
}
# ---------------------------
# TEST FILE OR FOLDER, RETURN BOOLEAN VALUE
Function Test-Folder ($folder) {
    Test-Path -Path $folder -PathType Any
} 
# ---------------------------
# GET INSTALLED VERSION 
Function Get-InstVersion {
    Param ([String[]]$pName)
    [String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName}).Version
}
# ---------------------------
# Check file versions
Function Get-FileVersion ($testFIle) {
    (Get-Item $testFile).VersionInfo.FileVersion
}
# ---------------------------
# INSTALLED? RETURN BOOLEAN VALUE 
Function Assert-IsInstalled ($pName) {
    $findIntallByName = [String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName})
    $condition = ($null -ne $findIntallByName)
    ($true, $false)[!$condition]
} 
# ---------------------------
# UPDATE INSTALLED STATUS 
Function Update-Status ($pName) {
    Assert-IsInstalled $pName | Out-Null
    If (Assert-IsInstalled $pName) { $script:isInstalled = $true}
    $packageFolder | Out-Null
    $curVersion | Out-Null
    $exeExist | Out-Null
} 
# ---------------------------
# INSTALL PROGRAM 
Function Install-Prog ($pName,$packageFolder,$curVersion,$exeExist,$destVersion,$exeFile,$issFile) {
    If ($isInstalled) {
        If ($curVersion -eq $destVersion) {
            Logging "INFO" "$pName $mesgInstalled"
        } Else {
            Logging "ERROR" "$mesgDiffVer - $pName"
            Stop-Script
        }
    } Else {
        If ($packageFolder -eq $false) {Logging "ERROR" "$pName $mesgNoPkg";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $true)){Logging "ERROR" "$mesgDiffVer";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $false)) {
            Logging "PROGRESS" "Installation for $pName, Please wait ..."
            Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
            $installed = Assert-IsInstalled $pName
            If ($installed) {Logging " " "$pName $mesgComplete";Start-Sleep -S 2} Else {Logging "ERROR" "$pName $mesgFailed";Stop-Script}
        }
    }

}

Function Install-ProgPlusPatch ($pName,$packageFolder,$curVersion,$exeExist,$ver1,$ver2,$exeFile,$issFile,$patchExeFile,$patchIssFile) {
    If ($isInstalled) {
        If ($curVersion -eq $ver2) {
            Logging "INFO" "$pName $mesgInstalled"
        } Elseif ($curVersion -eq $ver1) {
            Logging " " "Processing installation for $pName patch, Please wait ..."
            Start-Process -NoNewWindow -FilePath $patchExeFile -ArgumentList " /s /f1$patchIssFile" -Wait
            $getVersion = Get-InstVersion -pName $pName
            If ($getVersion -eq $ver2) { 
                Logging "INFO" "$pName $mesgInstalled"
            } Else {
                Logging "ERROR" "$pName $mesgFailed"
                Stop-Script
            }
        } Else {
            Logging "ERROR" "$mesgDiffVer - $pName"
            Stop-Script
        }
    } Else {
        If ($packageFolder -eq $false) {Logging "ERROR" "$pName $mesgNoPkg";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $true)){Logging "ERROR" "$mesgDiffVer";Stop-Script}
        If (($packageFolder -eq $true) -and ($exeExist -eq $false)) {
            Logging " " "Processing installation for $pName, Please wait ..."
            Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
            Start-Sleep -Seconds 2
            Start-Process -NoNewWindow -FilePath $patchExeFile -ArgumentList " /s /f1$patchIssFile" -Wait
            $getVersion = Get-InstVersion -pName $pName
            If ($getVersion -eq $ver2) { 
                Logging " " "$pName $mesgComplete"
            } Else {
                Logging "ERROR" "$pName $mesgFailed"
                Stop-Script 
            }
        }
    }
}

# ---------------------------
# Install Lens Patches 5.45
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
} 

# ---------------------------
# Update File Version
Function Update-FileVersion ($targetFile, $verToMatch) {
    If (Test-Folder $targetFile) {
        Get-FileVersion $targetFile | Out-Null
        If ((Get-FileVersion $targetFile) -eq $verToMatch) {$script:isInstalled = 1}
        Else {$script:isInstalled = 0}
    }
}
# ---------------------------
# Install digitalPolling
Function Install-DigitalPolling ($pName, $targetFile,$exeFile,$issFile) {
    If (Test-Folder $targetFile) {
        Logging "INFO" "$pName $mesgInstalled"
    } Else {
        Logging " " "Processing installation for $pName, Please wait ..."
        Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
        If (Test-Folder $targetFile){Logging " " "$pName $mesgComplete"}
        Else {Logging "Error" "$pName $mesgFailed";Stop-Script}
    }
}
# ---------------------------
# Update Copied Flies
Function Update-Copy ($srcPackage,$instFolder0,$instFolder1) {
    $testSrcPackage  = Test-Folder $srcPackage; $testSrcPackage | Out-Null
    $testInstFolder0 = Test-Folder $instFolder0; $testInstFolder0 | Out-Null
    $testInstFolder1 = Test-Folder $instFolder1; $testInstFolder1 | Out-Null
    If (($testInstFolder0 -eq $false) -and ($testInstFolder1)) { $script:fileCopied = 1 }
} 
# ---------------------------
# Install Web Service PMS Tester
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
} 
# ---------------------------
# Share Folder for windows 2008 R2 or lower
Function New-Share {
    param([string]$shareName,[string]$shareFolder)
	net share $shareName=$shareFolder "/GRANT:Everyone,FULL" /REMARK:"Saflok Database Folder Share"
} 
# ---------------------------
# Install SC
Function Install-SC {
    Param([String]$exeFile,[String]$agrFile,[String]$service)
    Start-Process -NoNewWindow -FilePath $exeFile $service -ArgumentList " $argFile" -Wait -RedirectStandardOutput Out-Null
} 
# ---------------------------
# Install SQL
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
} 
# ---------------------------
# update SQL password
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
}
# -----------------------
#$recoveryServices
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
    
    $action = $action1 + "/" + $time1 + "/" + $action2 + "/" + $time2 + "/" + $actionLast + "/" + $timeLast
    $output = sc.exe $serverPath failure $service actions= $action reset= $resetCounter | Out-Null
    Return $output
} 
# ---------------------------
# Mini Powershell version requirement
If ($PSVersionTable.PSVersion.Major -lt 5) {
    Logging "WARNING" "Your PowerShell installation is not version 5.0 or greater."
    Logging "WARNING" "This script requires PowerShell version 5.0 or above."
    Logging "WARNING" "You can download PowerShell version 5.0 at: https://www.microsoft.com/en-us/download/details.aspx?id=50395"
    Logging "WARNING" "Reboot server after installing Powershell 5 or above, try this script again."
    Stop-Script
} 
# ---------------------------
# Header variables
$cname = "[$vendor]"
$hotelName = 'Property: ' + $property.trim().toUpper()
$time = Get-Date -Format 'yyyy/MM/dd HH:mm'
$shareName = 'SaflokData'
[double]$winOS = [string][environment]::OSVersion.Version.major + '.' + [environment]::OSVersion.Version.minor
$osDetail = (Get-CimInstance -ClassName CIM_OperatingSystem).Caption 
# ---------------------------
# MENU OPTION
Clear-Host
Logging "" "+---------------------------------------------------------"
Logging "" "| WELCOME TO SAFLOK LODGING SYSTEMS INSTALLATION"
Logging "" "+---------------------------------------------------------"
Write-Colr -Text $cname, " |"," # IMPORTANT"  -Colour White,cyan,Red
Write-Colr -Text $cname, " |"," # THIS SCRIPT MUST BE RUN AS ADMINISTRATOR" -Colour White,cyan,red
Logging "" "+---------------------------------------------------------"
Logging "" "| $time"
Write-Colr -Text $cname, " |"," $hotelName" -Colour White,cyan,Yellow
Logging "" "| SAFLOK VERSION: $version"
IF ($winOS -le 6.1) {Logging "" "| INSTALLING ON: WINDOWS 7 / SERVER 2008 R2 OR LOWER"}
				Else {Logging "" "| INSTALLING ON: $osDetail"}
Logging "" "+---------------------------------------------------------"
Logging "" "| SCRIPT VERSION: $scriptVersion"
Logging "" "+---------------------------------------------------------"
Logging " " ""

# ---------------------------
# DRIVE INFO
$inputDrive = $inputDrive.Trim().ToUpper()
$driveLetterPattern = "(\w{1})"
# Get drive IDs from localhost, driveType 3="Fixed local disk"
$driveIDs = (Get-WmiObject -class Win32_LogicalDisk -computername localhost -filter "drivetype=3").DeviceID
$driveLetters = @()
for ($i=0; $i -lt $driveIDs.Length; $i++) {
    if ($driveIds[$i] -match $driveLetterPattern) {
        $driveLetters += $matches[0]
    }
}
# ---------------------------
# VALID DRIVE CHARACTER INPUT
if ($inputDrive -IN $driveLetters) {
    $driveIDExist = $True
} else {
    $driveIDExist = $False
}

switch ($driveIDExist) {
    $True {Logging "INFO" "You selected drive $inputDrive"}
    $False {
              Logging "WARNING" 'Please re-run the script again to select the correct drive.'
              Stop-Script
           }
}

# ---------------------------
# Driver letter + :
$installDrive = $inputDrive + ':'

# ---------------------------
# SOURCE FOLDER - INSTALL SCRIPT
$packageFolders =  Get-ChildItem ($scriptPath) | Select-Object Name | Sort-Object -Property Name

# ---------------------------
# validate if source folder exist
Switch ($version) {
    '5.45' {$pattern = "([A-Z]{7}_[A-Z]{9}_[A-Z]{4}_\d{3}_\w{3}\d{4}$)"}
    '5.68' {$pattern = "([A-Z]{9}_[A-Z]{4}_\d{3}_\w{3}\d{4}$)"}
}
Switch ($scriptPath -match $pattern) {
    $True  {
                # -----------------------
                # HEADER Information 
                Logging " " ""
                Logging " " "By installing you accept licenses for the packages."
                $confirmation = Read-Host "$cname Do you want to run the script? [Y] Yes  [N] No"
                $confirmation = $confirmation.ToUpper()
            }
    $False  {
                Logging "ERROR" "$mesgNoSource"
                Stop-Script
            }
}
$absPackageFolders = @()
For ($i=0; $i -lt ($packageFolders.Length-1); $i++) {
    $absPackageFolders += Join-Path $scriptPath $packageFolders[$i].Name
}

# ---------------------------
# Program
$progExe = Join-Path $absPackageFolders[3] 'setup.exe'
# PMS
$pmsExe = Join-Path $absPackageFolders[4] 'setup.exe'
# Messenger
$msgrExe = Join-Path $absPackageFolders[5] 'setup.exe'
# saflokLENS 
$lensSrcFolder = $absPackageFolders[6]
$lensExe = Join-Path $lensSrcFolder '/AutoPlay/Install Script/Lens/en/setup.exe'

# ---------------------------
# SQL 2012 express 
$sqlExprExe =  Join-Path $lensSrcFolder '/AutoPlay/Install Script/Lens/en' | 
               Join-Path -ChildPath 'ISSetupPrerequisites' | 
               Join-Path -ChildPath '{C38620DE-0463-4522-ADEA-C7A5A47D1FF6}' | 
               Join-Path -ChildPath 'SQLEXPR_x86_ENU.exe'

# ---------------------------
# ISS_FOR_Drive & files
Switch ($inputDrive) {
    'C' {$iss4Drive = '01.ISS_FOR_' + 'C'}
    'D' {$iss4Drive = '02.ISS_FOR_' + 'D'}
}
Switch ($iss4Drive)
{
    01.ISS_FOR_C {$issFolder = $absPackageFolders[1]}
    02.ISS_FOR_D {$issFolder = $absPackageFolders[2]}
} 
$issFiles = Get-ChildItem $issFolder | Select-Object Name | Sort-Object -Property Name
$progISS = Join-Path $issFolder $issFiles[0].Name            # [0] Programsetup.iss
Switch ($version) {
    '5.45' {    
            $pmsISS = Join-Path $issFolder $issFiles[1].Name
            $msgrISS = Join-Path $issFolder $issFiles[2].Name           
            $lensISS = Join-Path $issFolder $issFiles[3].Name          
            $patchLensISS = Join-Path $issFolder $issFiles[4].Name    
            $patchPollingISS = Join-Path $issFolder $issFiles[5].Name 
    }
    '5.68' {
            $patchProgISS = Join-Path $issFolder $issFiles[1].Name      
            $pmsISS = Join-Path $issFolder $issFiles[2].Name            
            $patchPmsISS = Join-Path $issFolder $issFiles[3].Name       
            $msgrISS = Join-Path $issFolder $issFiles[4].Name           
            $lensISS = Join-Path $issFolder $issFiles[5].Name          
            $patchLensISS = Join-Path $issFolder $issFiles[6].Name 
    }

}

# ---------------------------
# PATCH FILES
$patchExeFiles = Get-ChildItem ($absPackageFolders[7]) | Select-Object Name | Sort-Object -Property Name
Switch ($version) {
    '5.45' {  
            $pollingPatchExe = Join-Path $absPackageFolders[7]  $patchExeFiles[0].Name   
            $lensPatchExe = Join-Path $absPackageFolders[7]  $patchExeFiles[1].Name       
    }
    '5.68' {
            $progPatchExe = $absPackageFolders[7] + '\' + $patchExeFiles[0].Name     
            $pmsPatchExe = $absPackageFolders[7] + '\' + $patchExeFiles[1].Name     
            $lensPatchExe = $absPackageFolders[7] + '\' + $patchExeFiles[2].Name    
    }
}
# ---------------------------
# Web service PMS tester [FOLDER]
$webServiceTester = $absPackageFolders[8]
Switch ($version) {
    '5.45' { 
            # ---------------------------
            # ConfigFiles Folder & Files
            $configFiles = Get-ChildItem ($absPackageFolders[9]) | Select-Object Name | Sort-Object -Property Name
            $pollingConfig = Join-Path $absPackageFolders[9] $configFiles[0].Name      
            $lensPmsConfig = Join-Path $absPackageFolders[9] $configFiles[1].Name   
    }
}
# ---------------------------
# Absolute installed FOLDER
$kabaInstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'KABA'
$saflokV4InstFolder = Join-Path $installDrive 'SaflokV4'
$lensInstFolder = Join-Path $kabaInstFolder 'Messenger Lens'
$hubGateWayInstFolder = Join-Path $lensInstFolder 'HubGatewayService'
$hmsInstFolder = Join-Path $lensInstFolder 'HubManagerService'
$pmsInstFolder = Join-Path $lensInstFolder 'PMS Service'
$wsTesterInstFolder = Join-Path $kabaInstFolder '08.Web_Service_PMS_Tester'
Switch ($version) {
    '5.45' { 
            $digitalPollingFolder = Join-Path $lensInstFolder 'DigitalKeysPollingSoftware' 
    }
}
$kdsInstFolder = Join-Path $lensInstFolder 'KeyDeliveryService'   
# ---------------------------
# GUI exe file in SAFLOKV4 FOLDER
$saflokClient = Join-Path $saflokV4InstFolder 'Saflok_Client.exe'
$saflokMsgr = Join-Path $saflokV4InstFolder 'Saflok_MsgrServer.exe'
$saflokIRS = Join-Path $saflokV4InstFolder 'Saflok_IRS.exe'
$shareFolder = Join-Path $saflokV4InstFolder  'SaflokData'
# ---------------------------
# Installed EXE files for version information
$gatewayExe = Join-Path $hubGateWayInstFolder 'LENS_Gateway.exe'
$hmsExe = Join-Path $hmsInstFolder 'LENS_HMS.exe'
$wsPmsExe = Join-Path $pmsInstFolder 'LENS_PMS.exe'
Switch ($version) {
    '5.45' { 
            $digitalPollingExe = Join-Path $digitalPollingFolder 'DigitalKeysPollingService.exe'
    }
}
$kdsExe = Join-Path $kdsInstFolder 'Kaba_KDS.exe'
# ---------------------------
# INSTALLED CONFIG FILES
$lensPmsConfigFileInst =  Join-Path $pmsInstFolder 'LENS_PMS.exe.config'
$hh6ConfigFile = Join-Path $saflokV4InstFolder 'KabaSaflokHH6.exe.config'
Switch ($version) {
    '5.45' { 
            $pollingConfigInst  = Join-Path $digitalPollingFolder 'DigitalKeysPollingService.exe.config'
            # ---------------------------
            # Polling log
            $pollingLog = 'C:\ProgramData\DormaKaba\Server\Polling\logs.log'
    }
}
# ---------------------------
# Saflok Services
$serviceNames = [System.Collections.ArrayList]@(
    'DeviceManagerService',
    'KIPEncoderService',
    'FirebirdGuardianDefaultInstance'
    'SaflokServiceLauncher',
    'SaflokCRS',
    'SaflokScheduler',
    'SaflokIRS',
    'SaflokMSGR',
    'SAFLOKDHSPtoMSGRTranslator',
    'SAFLOKMSGRtoDHSPTranslator',
    'MessengerNet_Hub Gateway Service',
    'MNet_HMS',                       
    'MNet_PMS Service',                     
    'MessengerNet_Utility Service', 
    'Kaba_KDS',                            
    'VirtualEncoderService',    
    'Kaba Digital Keys Polling Service'
)
# ---------------------------
# Start to install ==>
If ($confirmation -eq 'Y' -or $confirmation -eq 'YES') {
    Logging " " ""
    # -------------------------------------------------------------------
    # install Saflok client
    $isInstalled = 0
    $pName = "Saflok Program"
    $packageFolder = Test-Folder ($absPackageFolders[3])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokClient
    Update-Status $pName
    Switch ($version) {
        '5.45' {
                $destVersion = $progVersion
                Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $progExe $progISS
        }
        '5.68' {
                Install-ProgPlusPatch $pName $packageFolder $curVersion $exeExist $ver1 $ver2 $progExe $progISS $progPatchExe $patchProgISS
                # -------------------------------------------------------------------    
                # [ clean munit ink ]
                $munit = 'C:\Users\Public\Desktop\Kaba Saflok M-Unit.lnk'
                If (Test-Path -Path $munit){
                    Remove-Item $munit -Force
                }
        }
    }

    # -------------------------------------------------------------------
    # install Saflok PMS
    $pName = "Saflok PMS"
    $isInstalled = 0
    $packageFolder = Test-Folder ($absPackageFolders[4])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokIRS
    Update-Status $pName
        Switch ($version) {
        '5.45' {
                $destVersion = $pmsVersion
                Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $pmsExe $pmsISS
        } 
        '5.68' {
               Install-ProgPlusPatch $pName $packageFolder $curVersion $exeExist $ver1 $ver2 $pmsExe $pmsISS $pmsPatchExe $patchPmsISS
        }
    }
    # -------------------------------------------------------------------
    # install Saflok Messenger
    $pName = "Saflok Messenger Server"
    $isInstalled = 0
    $packageFolder = Test-Folder ($absPackageFolders[5])
    $curVersion = Get-InstVersion -pName $pName
    $exeExist = Test-Folder $saflokMsgr
    $destVersion = $msgrVersion
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
    $fbSvcStat = (Get-Service | Where-Object {$_.Name -eq 'FirebirdGuardianDefaultInstance'}).Status
    If ($fbSvcStat -eq "Stopped"){Start-Service -Name 'FirebirdGuardianDefaultInstance';Start-Sleep -S 2}
    # -------------------------------------------------------------------
    # start Saflok launcher service
    If (Get-Service | Where-Object {$_.Name -eq 'SaflokServiceLauncher' -and $_.Status -eq "Stopped"}){
        Start-Service -Name 'SaflokServiceLauncher'
    }

    # -------------------------------------------------------------------
    # IIS FEATURES, requirement for messenger lens
    $isInstalledMsgr = Assert-IsInstalled "Saflok Messenger Server"
	If ($isInstalledMsgr -ne $True) {
		Logging "WARN" "Please install Saflok messenger before Lens."
		Stop-Script
	} Else {
        $featureState = dism /online /get-featureinfo /featurename:IIS-WebServerRole | findstr /C:'State : '
        If ($featureState -match 'Disabled') {
            Logging "" "Configuring IIS features for Messenger LENS, please wait..."
            Switch ($winOS) {
                {$winOS -le 6.1} {$iisFeatures = 'IIS-WebServerRole','IIS-WebServer','IIS-CommonHttpFeatures',
                    'IIS-HttpErrors','IIS-ApplicationDevelopment','IIS-RequestFiltering','IIS-NetFxExtensibility',
                    'IIS-HealthAndDiagnostics','IIS-HttpLogging','IIS-RequestMonitor','IIS-Performance','WAS-ProcessModel',
                    'WAS-NetFxEnvironment','WAS-ConfigurationAPI','IIS-ISAPIExtensions','IIS-ISAPIFilter','IIS-StaticContent',
                    'IIS-DefaultDocument','IIS-DirectoryBrowsing','IIS-ASPNET','IIS-ASP','IIS-HttpCompressionStatic',
                    'IIS-ManagementConsole','NetFx3','WCF-HTTP-Activation','WCF-NonHTTP-Activation'
                    For ([int]$i=0; $i -lt ($iisFeatures.Length - 1); $i++) {
                        $feature = $iisFeatures[$i]
                        DISM /online /enable-feature /featurename:$feature | Out-Null
                        Start-Sleep -S 1
                        Logging " " "Enabled feature $feature"
                    }
                } 
                {$winOS -ge 6.1}{ $disabledFeatures = @()
                    $iisFeatures = 'NetFx4Extended-ASPNET45','IIS-ASP','IIS-ASPNET45','IIS-NetFxExtensibility45',
                    'IIS-WebServerRole','IIS-WebServer', 'IIS-CommonHttpFeatures','IIS-HttpErrors',
                    'IIS-ApplicationDevelopment','IIS-HealthAndDiagnostics','IIS-HttpLogging','IIS-Security', 
                    'IIS-RequestFiltering','IIS-Performance','IIS-WebServerManagementTools','IIS-StaticContent',
                    'IIS-DefaultDocument','IIS-DirectoryBrowsing', 'IIS-ApplicationInit','IIS-ISAPIExtensions',
                    'IIS-ISAPIFilter','IIS-HttpCompressionStatic','IIS-ManagementConsole'
                    For ([int]$i=0; $i -lt ($iisFeatures.Length -1 ); $i++) {
                        $feature = $iisFeatures[$i]
                        If (!((Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -eq $feature}).State -eq "Enabled")) {
                            $disabledFeatures += $feature
                        }
                    }
                    If ($disabledFeatures.Count-1 -gt 0){
                        Logging "INFO" "Configuring IIS features for Messenger LENS, please wait ..."
                        Foreach ($disabled In $disabledFeatures) {
                            Enable-WindowsOptionalFeature -Online -FeatureName $disabled -All -NoRestart | Out-Null
                            Logging " " "Enabled feature $disabled"
                        }
                    } Else {
                        Logging "INFO" "ALL required IIS features have been enabled."
                    } 
                }
            }
        } Else {
            Logging "INFO" "ALL IIS features that Messenger LENS requires have been enabled."
        }
    }
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
    # install Messenger Lens 
    $pName = "Messenger LENS"
    Update-Status $pName
    If (Assert-IsInstalled $pName) {
        # -------------------------------------------------------------------
        # install Messenger Lens Patch 
        $pName = "Messenger Lens Patch"  
        $destVersion = $wsPmsExeVersion
        Update-FileVersion $wsPmsExe $destVersion
        Install-LensPatch $wsPmsExe $destVersion $pName $lensPatchExe $patchLensISS
    } Else {
        $pName = "Messenger LENS"
        $isInstalled = 0
        $packageFolder = Test-Folder ($absPackageFolders[6])
        $curVersion = Get-InstVersion -pName $pName
        $exeExist = Test-Folder $wsPmsExe
        $destVersion = $wsPmsExeBeforePatchVersion
        Install-Prog $pName $packageFolder $curVersion $exeExist $destVersion $lensExe $lensISS
        $pName = "Messenger Lens Patch"  
        $isInstalled = 0
        $targetFile = $wsPmsExe
        $destVersion = $wsPmsExeVersion
        Update-FileVersion $wsPmsExe $destVersion
        Install-LensPatch $wsPmsExe $destVersion $pName $lensPatchExe $patchLensISS
    }
    # -------------------------------------------------------------------
    # install digital polling service
    If ($version -eq  '5.45') {
        $isInstalled = 0
        $pName = "Marriott digital polling service"
        Install-DigitalPolling $pName $digitalPollingExe $pollingPatchExe $patchPollingISS
    }

    # -------------------------------------------------------------------
    # allow everyone access to Lens folder
    If (Test-Folder Container) {
        $acl = Get-Acl -Path $lensInstFolder
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","FullControl","Allow")
        $acl.SetAccessRule($AccessRule); $acl | Set-Acl $lensInstFolder
    }

    # -------------------------------------------------------------------
    # copy config files
    If ($version -eq '5.45') {
        Copy-Item -Path $lensPmsConfig -Destination $pmsInstFolder -Force
        Copy-Item -Path $pollingConfig -Destination $digitalPollingFolder -Force
    }
    # INSTALL WEB SERVICE PMS TESTER
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
    }
    # ----------------------------------------------------------------
    # CHECK SERVICES STATUS
    $servicesCheck = [System.Collections.ArrayList]@(
                    'DeviceManagerService',
                    'KIPEncoderService',
                    'Kaba_KDS', 
                    'MessengerNet_Hub Gateway Service',
                    'MNet_HMS', 
                    'MNet_PMS Service',  
                    'MessengerNet_Utility Service',
                    'VirtualEncoderService',    
                    'Kaba Digital Keys Polling Service'
    )   

    If (Assert-IsInstalled "Messenger LENS") {
        If ($version -eq '5.68') {
            $servicesCheck.Remove('Kaba Digital Keys Polling Service')
        }
        Foreach ($service In $servicesCheck) { 
            $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
            If ($serviceStatus.Status -eq "stopped") {
                Logging " " "Staring service $service."
                Start-Service -Name $service -ErrorAction SilentlyContinue
                $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
                If ($serviceStatus.Status -eq "running") { Logging " " "$service has been started."}
                Start-Sleep -S 1
            } Else {Logging " " "$service is in running state.";Start-Sleep -S 1}
        }
        # ----------------------------------------------------------------
        # OPEN GUI AND CONFIG FILE
        Logging "" "+---------------------------------------------------------"
        Logging "" "The following files need to be checked or configure: "
        Logging "" "+---------------------------------------------------------"
        If ($Null -eq (Get-Process | where-object {$_.Name -eq 'Saflok_IRS'}).ID) {
            Start-Process -NoNewWindow -FilePath $saflokIRS; Start-Sleep -S 1
        } # run IRS GUI
        Get-Process -ProcessName notepad* | Stop-Process -Force; Start-Sleep -S 1 
        If ((Assert-isInstalled "Saflok Program") -and (Test-Folder $hh6ConfigFile)) {
            Logging " " "[ KabaSaflokHH6.exe.config ]"
            Start-Process notepad $hh6ConfigFile -WindowStyle Minimized; Start-Sleep -S 1
        } # hh6 config
        If ((Assert-isInstalled  "Messenger LENS") -and (Test-Folder $lensPmsConfigFileInst)) {
            Logging " " "[ LENS_PMS.exe.config ]"
            Start-Process notepad $lensPmsConfigFileInst -WindowStyle Minimized; Start-Sleep -S 1
        } # PMS config
        If($version -eq '5.45') {
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
            }
        }
        Write-Colr -Text "$cname ","Please check those config files opening in taskbar area." -Colour White,Magenta
        Logging "" "+---------------------------------------------------------"
        Start-Sleep -Seconds 1
        # ----------------------------------------------------------------
        # SET SERVICE RECOVERY
        $recoveryServices = [System.Collections.ArrayList]@(
                            'MessengerNet_Hub Gateway Service',
                            'MNet_HMS', 
                            'MNet_PMS Service',                     
                            'Kaba_KDS',                               
                            'Kaba Digital Keys Polling Service'
        )
        If ($version -eq '5.68') {
            $recoveryServices.Remove('Kaba Digital Keys Polling Service')
        }
        Foreach ($service In $recoveryServices){
            If (Get-service -Name $service | where-object {$_.StartType -ne 'Automatic'}) { Set-Service $service -StartupType "Automatic" }
            Set-ServiceRecovery -ServiceDisplayName $service
        } 
        # ----------------------------------------------------------------
        # FOOTER
        Logging "" "Installed Version Information: " 
        Logging "" "+---------------------------------------------------------"
        $gatewayVer = Get-FileVersion $gatewayExe;
        $hmsVer = Get-FileVersion $hmsExe;
        $wsPmsVer = Get-FileVersion $wsPmsExe;
        $kdsVer = Get-FileVersion $kdsExe;
        if ($version -eq '5.45') {
            $pollingVer = Get-FileVersion $digitalPollingExe
        }
        If (Test-Path $gatewayExe -PathType Leaf) { Logging " " "Gateway: $gatewayVer"; Start-Sleep -Seconds 1 } 
        If (Test-Path $hmsExe -PathType Leaf) { Logging " " "HMS:     $hmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $wsPmsExe -PathType Leaf) { Logging " " "PMS:     $wsPmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $kdsExe -PathType Leaf) { Logging " " "KDS:     $kdsVer"; Start-Sleep -Seconds 1 }
        If ($version -eq '5.45') {
            If (Test-Path $digitalPollingExe -PathType Leaf) { Logging " " "POLLING: $pollingVer"; Start-Sleep -Seconds 1 }
        }
        Logging "" "+---------------------------------------------------------"
        Logging "" "DONE"
        Logging "" "+---------------------------------------------------------"
        Logging "" ""
        Logging "WARNING" "The recent program changes indicate a reboot is necessary."
        Write-Host ''
        # clean up script files and SAFLOK folder
        If (Test-Path -Path "$scriptPath\*.*" -Include *.ps1){Remove-Item -Path "$scriptPath\*.*" -Include *.ps1 -Force -ErrorAction SilentlyContinue}
        If (Test-path -Path "C:\SAFLOK") { Remove-Item -Path "C:\SAFLOK" -Recurse -Force -ErrorAction SilentlyContinue }   
        Start-Sleep -Second 300
    } Else {
        Logging "ERROR" "Missing Messenger Lens program, Please make sure Messenger Lens is be installed properly. "
        Stop-Script
    }

} # END OF YES
If ($confirmation -eq 'N' -or $confirmation -eq 'NO') {
    Logging " " ""
    Write-Colr -Text $cname," Thank you, Bye!" -Colour White,Gray
    Write-Host ''
    Stop-Script
} # END OF NO

