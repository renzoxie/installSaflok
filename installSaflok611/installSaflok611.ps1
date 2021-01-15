<#
.Synopsis
   Helper Script to install SAFLOK SYSTEM 6000 automatically
.DESCRIPTION
   This script fully installation for SAFLOK System 6000 automatically
.EXAMPLE
   .\install.ps1 
.NOTES
    Author: renzo.xie@dormakaba.com
    Saflok version: v6.11, Marriott ONLY
    Create Date: 16 April 2019
    Modified Date: 30 AUG 2019
#>

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit 
}
# ----------------------------------------------------------------------------------------------
# Script location
$scriptPath = $PSScriptRoot
# ----------------------------------------------------------------------------------------------
# Customized color
Function Write-Colr {
	Param ([String[]]$Text,[ConsoleColor[]]$Colour,[Switch]$NoNewline=$false)
	For ([int]$i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -Foreground $Colour[$i] -NoNewLine }
	If ($NoNewline -eq $false) { Write-Host '' }
}
# ----------------------------------------------------------------------------------------------
# Customized loggings
Function Logging ($state, $message) {
	$part1 = $cname;$part2 = ' ';$part3 = $state;$part4 = ": ";$part5 = "$message"
	Switch ($state)
	{
		ERROR {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Red,Red,Red}
		WARN  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Magenta,Magenta,Magenta}
		INFO  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Yellow,Yellow,Yellow}
		PROGRESS  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,White,White,White}
		""   {Write-Colr -Text $part1,$part2,$part5 -Colour White,White,Cyan}
		default { Write-Colr -Text $part1,$part2,$part5 -Colour White,White,White}
	}
}
Function Stop-Script {
    Start-Sleep -Seconds 300 
    exit
}
# ----------------------------------------------------------------------------------------------
# Mini Powershell version requirement
If ($PSVersionTable.PSVersion.Major -lt 5) {
	Logging "WARN" "Your PowerShell installation is not version 5.0 or greater."
	Logging "WARN" "This script requires PowerShell version 5.0 or greater to function."
	Logging "WARN" "You can download PowerShell version 5.0 at: https://www.microsoft.com/en-us/download/details.aspx?id=50395"
	Stop-Script
} 
# ----------------------------------------------------------------------------------------------
# GENERAL
$cname = '[dormakaba]'
$time = Get-Date -Format 'yyyy/MM/dd HH:mm'
$shareName = 'SaflokData'
$hotelChain = 'MARRIOTT PROJECTS ONLY'
[double]$winOS = [string][environment]::OSVersion.Version.major + '.' + [environment]::OSVersion.Version.minor
# ----------------------------------------------------------------------------------------------
# VERSIONS
$scriptVersion = '1.6'
$saflokVersion = '6.11'
$progVersion = '6.1.1.0'
# ----------------------------------------------------------------------------------------------
# MENU OPTION
$driveLetter
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
	1 {$script:driveLetter = 'C';$menuOption = 99}
	2 {$script:driveLetter = 'D';$menuOption = 99}
	0 {Clear-Host;Exit}
	Default {Logging " " "Please enter a valid option."}
}
# ----------------------------------------------------------------------------------------------
# DRIVE INFO
$installDrive = $driveLetter + ':\'
# ----------------------------------------------------------------------------------------------
# VALID DRIVE CHARACTER INPUT
$deviceID = $driveLetter + ':'
$cdDrive = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 5} | Select-Object DeviceID
If ($cdDrive -match $deviceID) {
	Write-Host ''
	Logging "ERROR" "The drive $driveLetter is not a valid location."
	Write-Host ''
	Stop-Script
}
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - INSTALL SCRIPT
$installScriptFolder = Join-Path (Split-Path $scriptPath) 'Install Script'
$packageFolders =  Get-ChildItem ($installScriptFolder) | Select-Object Name | Sort-Object -Property Name
$absPackageFolders = @()
For ($i=0; $i -lt ($packageFolders.Length); $i++) {
	$absPackageFolders += Join-Path $installScriptFolder $packageFolders[$i].Name
}
$lensSrcFolder = $absPackageFolders[1]
$system6000SrcFolder = $absPackageFolders[2]
$system6000SubFolders = Get-ChildItem ($system6000SrcFolder) | Select-Object Name | Sort-Object -Property Name
$abs6000SubFolders = @()
For ($i=0; $i -lt ($system6000SubFolders.Length); $i++) {
	$abs6000SubFolders += Join-Path $system6000SrcFolder $system6000SubFolders[$i].Name
} 
# Program
$progExe = Join-Path $abs6000SubFolders[2] 'setup.exe'
# PMS
$pmsExe = Join-Path $abs6000SubFolders[1] 'setup.exe'
# Messenger
$msgrExe = Join-Path $abs6000SubFolders[0] 'setup.exe'
# saflokLENS 
$enFolder = Join-Path $lensSrcFolder 'en'
$lensExe = Join-Path $enFolder 'setup.exe'
# ----------------------------------------------------------------------------------------------
# SQL 2016 express 
$sqlExprExe = Join-Path $enFolder 'ISSetupPrerequisites' | `
			  Join-Path -ChildPath '{DEA37953-690E-42ED-B1D0-E75C59D41454}' | `
			  Join-Path -ChildPath 'SQLEXPR_x64_ENU.exe'
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - PLUGINS
$pluginSrcFolder = Join-Path (Split-Path $scriptPath) 'Plugins' 
# ----------------------------------------------------------------------------------------------
# ISS_FOR_Drive & files
If ($driveLetter -eq 'C') {$iss4Drive = 'ISS_FOR_' + $driveLetter}
Elseif ($driveLetter -eq 'D') {$iss4Drive = 'ISS_FOR_' + $driveLetter}
Else { Logging "ERROR" "Please input C or D for the drive letter."; Stop-Script}
Switch ($iss4Drive)
{
	ISS_FOR_C {$issFolder = Join-Path $pluginSrcFolder 'ISS_FOR_C'}
	ISS_FOR_D {$issFolder = Join-Path $pluginSrcFolder 'ISS_FOR_D'}
} 
$issFiles = Get-ChildItem $issFolder | Select-Object Name | Sort-Object -Property Name
$progISS = Join-Path $issFolder $issFiles[0].Name            # [0] Programsetup.iss
$pmsISS = Join-Path $issFolder $issFiles[1].Name             # [1] PMSsetup.iss
$msgrISS = Join-Path $issFolder $issFiles[2].Name            # [2] MSGRsetup.iss
$lensISS = Join-Path $issFolder $issFiles[3].Name            # [3] LENSsetup.iss
# ----------------------------------------------------------------------------------------------
# Absolute installed FOLDER
$kabaInstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'dormakaba'
$saflokV4InstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'SaflokV4'
$lensInstFolder = Join-Path $kabaInstFolder 'Messenger Lens'
$hubGateWayInstFolder = Join-Path $lensInstFolder 'HubGatewayService'
$hmsInstFolder = Join-Path $lensInstFolder 'HubManagerService'
$pmsInstFolder = Join-Path $lensInstFolder 'PMS Service'
# ----------------------------------------------------------------------------------------------
# GUI exe file in SAFLOKV4 FOLDER
$saflokClient = Join-Path $saflokV4InstFolder 'Saflok_Client.exe'
$saflokMsgr = Join-Path $saflokV4InstFolder 'Saflok_MsgrServer.exe'
$saflokIRS = Join-Path $saflokV4InstFolder 'Saflok_IRS.exe'
$shareFolder = Join-Path $saflokV4InstFolder 'SaflokData'
# ----------------------------------------------------------------------------------------------
# EXE files for version inforamtion
$gatewayExe = Join-Path $hubGateWayInstFolder 'LENS_Gateway.exe'
$hmsExe = Join-Path $hmsInstFolder 'LENS_HMS.exe'
$wsPmsExe = Join-Path $pmsInstFolder 'LENS_PMS.exe'
# ----------------------------------------------------------------------------------------------
# Logging Messages
$mesgNoPkg ="package does not exist, operation exit."
$mesgInstalled = "has already been installed."
$mesgDiffVer = "There is another version exist, please uninstall it first."
$mesgComplete = "installation is complete."
$mesgFailed = "installation failed!"
# ----------------------------------------------------------------------------------------------
# TEST FILE OR FOLDER, RETURN BOOLEAN VALUE
Function Test-Folder ($folder) {
	Test-Path -Path $folder -PathType Any
} 
# ----------------------------------------------------------------------------------------------
# GET INSTALLED VERSION 
Function Get-InstVersion {
	Param ([String[]]$pName)
	[String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName}).Version
}
# ----------------------------------------------------------------------------------------------
# INSTALLED? RETURN BOOLEAN VALUE 
Function Assert-IsInstalled ($pName) {
	$findIntallByName = [String](Get-Package -ProviderName Programs -IncludeWindowsInstaller | Where-Object {$_.Name -eq $pName})
	$condition = ($null -ne $findIntallByName)
	($true, $false)[!$condition]
} 
# ----------------------------------------------------------------------------------------------
# UPDATE INSTALLED STATUS 
Function Update-Status ($pName) {
	Assert-IsInstalled $pName | Out-Null
	If (Assert-IsInstalled $pName) { $script:isInstalled = $true}
} 
# ----------------------------------------------------------------------------------------------
# INSTALL PROGRAM 
Function Install-Prog ($pName,$packageFolder,$curVersion,$exeExist,$destVersion,$exeFile,$issFile) {
	If ($isInstalled) {
		If ($curVersion -eq $destVersion) {Logging "INFO" "$pName $mesgInstalled"} Else {Logging "ERROR" "$mesgDiffVer - $pName";Stop-Script}
	} Else {
		If ($packageFolder -eq $false) {Logging "ERROR" "$pName $mesgNoPkg";Stop-Script}
		If (($packageFolder -eq $true) -and ($exeExist -eq $true)){Logging "ERROR" "$mesgDiffVer";Stop-Script}
		If (($packageFolder -eq $true) -and ($exeExist -eq $false)) {
			Logging "PROGRESS" "Installation for $pName, Please wait..."
			Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
			$installed = Assert-IsInstalled $pName
			If ($installed) {Logging " " "$pName $mesgComplete";Start-Sleep -S 2} Else {Logging "ERROR" "$pName $mesgFailed";Stop-Script}
		}
	}
} 
# ----------------------------------------------------------------------------------------------
# SHARE FOLDER FOR WIN2008 OR LOWER 
Function New-Share {
	param([string]$shareName,[string]$shareFolder)
	net share $shareName=$shareFolder "/GRANT:Everyone,FULL" /REMARK:"Saflok Database Folder Share"
} 
# -----------------------
# PROMPT MESSAGE
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
	$packageFolder = Test-Folder $system6000SrcFolder
	$curVersion = Get-InstVersion -pName $pName
	$exeExist = Test-Folder $saflokClient
	Update-Status $pName
	Install-Prog $pName $packageFolder $curVersion $exeExist $progVersion $progExe $progISS
	# -------------------------------------------------------------------
	# install Saflok PMS
	$pName = "Saflok PMS"
	$isInstalled = 0
	$packageFolder = Test-Folder $system6000SrcFolder
	$curVersion = Get-InstVersion -pName $pName
	$exeExist = Test-Folder $saflokIRS
	Update-Status "Saflok PMS"
	Install-Prog $pName $packageFolder $curVersion $exeExist $progVersion $pmsExe $pmsISS
	# -------------------------------------------------------------------
	# install Saflok Messenger
	$pName = "Saflok Messenger Server"
	$isInstalled = 0
	$packageFolder = Test-Folder $system6000SrcFolder
	$curVersion = Get-InstVersion -pName $pName
	$exeExist = Test-Folder $saflokMsgr
	Update-Status "$pName"
	Install-Prog $pName $packageFolder $curVersion $exeExist $progVersion $msgrExe $msgrISS
	# -------------------------------------------------------------------
	# copy database to Saflok data folder
	$srcHotelData = $pluginSrcFolder + '\' + 'hotelData'
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
		If ((Get-Service -Name FirebirdGuardianDefaultInstance).Status -eq "Running") { 
			Stop-Service -Name FirebirdGuardianDefaultInstance -Force -ErrorAction SilentlyContinue 
			Copy-Item -Path $srcHotelData\*.gdb -Destination $shareFolder        
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
	# start saflok launcher service
	If (Get-Service | Where-Object {$_.Name -eq 'SaflokServiceLauncher' -and $_.Status -eq "Stopped"}){
		Start-Service -Name 'SaflokServiceLauncher'
	}
	# -------------------------------------------------------------------
	# IIS FEATURES, requirement for messenger lens
    $isInstalledMsgr = Assert-IsInstalled "Saflok Messenger Server"
	If ($isInstalledMsgr -ne $True) {
		Logging "WARN" "Please install Saflok messenger first."; Stop-Script
	} Else {
        $featureState = dism /online /get-featureinfo /featurename:IIS-WebServerRole | findstr /C:'State : '
        If ($featureState -match 'Disabled') {
            Logging "" "Configuring IIS features for Messenger LENS, please wait..."
            Switch ($winOS) {
                {$winOS -le 6.1} 
                {
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
                }
                {$winOS -ge 6.1}
                {
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
	                    Logging "INFO" "Configuring IIS features for Messenger LENS, please wait..."
	                    Foreach ($disabled In $disabledFeatures) {
		                    Enable-WindowsOptionalFeature -Online -FeatureName $disabled -All -NoRestart | Out-Null
		                    Logging " " "Enabled features $disabled"
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
	# install Messenger Lens
	$pName = "Messenger LENS"
	$isInstalledMesgLens = Assert-IsInstalled "Messenger LENS"
	If ($isInstalledMesgLens -ne $True) {
		Logging "WARN" "Windows will reboot automaticlly after installing Messenger Lens"
		Start-Sleep -Second 5
	}
	$isInstalled = 0
	$packageFolder = Test-Folder $lensSrcFolder
	$curVersion = Get-InstVersion -pName $pName
	$exeExist = Test-Folder $wsPmsExe
	Update-Status $pName
	Install-Prog $pName $packageFolder $curVersion $exeExist $progVersion $lensExe $lensISS
	# -------------------------------------------------------------------
	# FOOTER	
	Logging "" ""
	Logging "" "+---------------------------------------------------------"
	Logging "" "DONE"
	Logging "" "+---------------------------------------------------------"
	Stop-Script
} # END OF YES
If ($confirmation -eq 'N' -or $confirmation -eq 'NO') {
    Logging " " ""
    Write-Colr -Text $cname," Thank you, Bye!" -Colour White,Gray
    Write-Host ''
    Stop-Script
} # END OF NO