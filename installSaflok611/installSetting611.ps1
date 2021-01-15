<#
.Synopsis
   Helper Script to install SAFLOK SYSTEM 6000 automatically
.DESCRIPTION
   This script fully installation for SAFLOK System 6000 automatically
.EXAMPLE
   .\config.ps1 
.NOTES
    Author: renzo.xie@dormakaba.com
    Saflok version: v6.11, Marriott ONLY
    Create Date: 16 April 2019
    Modified Date: 30 AUG 2019
#>
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { 
  Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; Exit 
}
$scriptPath = $PSScriptRoot
$cname = '[dormakaba]'
$saflokVersion = '6.11'
$hotelChain = "Marriott" # Default value - Marriott
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
# GET FILE VERSION
Function Get-FileVersion ($testFile) {
	(Get-Item $testFile).VersionInfo.FileVersion
}
# ----------------------------------------------------------------------------------------------
# UPDATE FILE VERSION  
Function Update-FileVersion ($targetFile, $verToMatch) {
	If (Test-Folder $targetFile) {
		Get-FileVersion $targetFile | Out-Null
		If ((Get-FileVersion $targetFile) -eq $verToMatch) {$script:isInstalled = 1}
		Else {$script:isInstalled = 0}
	}
}
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
# INSTALL MARRIOTT DIGITAL POLLING
Function Install-DigitalPolling ($pName,$targetFile,$exeFile,$issFile) {
	If (Test-Folder $targetFile) {
		Logging "INFO" "$pName $mesgInstalled"
	} Else {
		Logging "PROGRESS" "Installation for $pName, Please wait..."
		Start-Process -NoNewWindow -FilePath $exeFile -ArgumentList " /s /f1$issFile" -Wait
		If (Test-Folder $targetFile){Logging " " "$pName $mesgComplete"}
		Else {Logging "ERROR" "$pName $mesgFailed";Stop-Script}
	}
} 
# ----------------------------------------------------------------------------------------------
# UPDATE FILE COPY STATUS
Function Update-Copy ($srcPackage,$instFolder) {
	$testSrcPackage  = Test-Folder $srcPackage; $testSrcPackage | Out-Null
	$testInstFolder = Test-Folder $instFolder; $testInstFolder | Out-Null
	If ($testInstFolder) { $script:fileCopied = 1 }
} # Update copied files
# ----------------------------------------------------------------------------------------------
# INSTALL WEB SERVICE PMS TESTER
Function Install-PmsTester ($srcPackage,$instParent,$instFolder) {
	$testSrcPackage  = Test-Folder $srcPackage
	$testInstParent = Test-Folder $instParent
	$testInstFolder = Test-Folder $instFolder
	If ($testSrcPackage -eq $False) { Logging "ERROR" "Package files missing!"; Stop-Script } 
	If ($testInstParent -eq $False) { Logging "ERROR" "Messenger LENS has not been installed yet!" ; Stop-Script } 
	If ($fileCopied) {
		Logging "INFO" "Web Service PMS Tester $mesgInstalled"
	} Else {
		If (($testSrcPackage -eq $True) -and ($testInstParent -eq $True)) {Logging "PROGRESS" "Installing Web Service PMS Tester..." }
		If ($testInstFolder -eq $False) {
			Copy-Item $srcPackage -Destination $instParent -Recurse -Force -ErrorAction SilentlyContinue
			Logging " " "Web Service PMS Tester $mesgComplete"		
		} Else {
			Logging " " "Web Service PMS Tester $mesgComplete"
		}
	}
} 
# ----------------------------------------------------------------------------------------------
# SERVER CONTROL FOR SERVICE RECOVERY
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
} 
# ----------------------------------------------------------------------------------------------
# Logging Messages
$mesgNoPkg ="package does not exist, operation exit."
$mesgInstalled = "has already been installed."
$mesgDiffVer = "There is another version exist, please uninstall it first."
$mesgComplete = "installation is complete."
$mesgFailed = "installation failed!"
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - PLUGINS
$pluginSrcFolder = Join-Path (Split-Path $scriptPath) 'Plugins'
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - CONFIG 
If ($hotelChain -like '*Marriott*') {
	$configFolder = Join-Path $pluginSrcFolder 'ConfigFiles' | Join-Path -ChildPath 'Marriott'
	$digitalConfig = Join-Path $configFolder 'DigitalKeysPollingService.exe.config'
	$pmsConfig = Join-Path $configFolder 'LENS_PMS.exe.config'
}
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - Web service PMS Tester
$webServiceTester = Join-Path $pluginSrcFolder 'Web_Service_PMS_Tester'
# ----------------------------------------------------------------------------------------------
# SOURCE FOLDER - Polling Software
$pollingSoftware = Join-Path $pluginSrcFolder 'PollingSoftware'
$digitalPollingExe = Join-Path $pollingSoftware 'DigitalKeysPollingService.exe'
$installOnC = Test-Folder 'C:\Program Files (x86)\dormakaba\Messenger Lens'
$installOnD = Test-Folder 'D:\Program Files (x86)\dormakaba\Messenger Lens'
$saflokV4Folder = 'Program Files (x86)\SaflokV4'
$digitalPolling = 'Program Files (x86)\dormakaba\Messenger Lens\DigitalKeysPollingSoftware\DigitalKeysPollingService.exe'
$digitalPollingConfig = 'Program Files (x86)\dormakaba\Messenger Lens\DigitalKeysPollingSoftware\DigitalKeysPollingService.exe.config'
$lensPms = 'Program Files (x86)\dormakaba\Messenger Lens\PMS Service\LENS_PMS.exe'
$lensPmsConfig = 'Program Files (x86)\dormakaba\Messenger Lens\PMS Service\LENS_PMS.exe.config'
$lensHMS = 'Program Files (x86)\dormakaba\Messenger Lens\HubManagerService\LENS_HMS.exe'
$kabaKds = 'Program Files (x86)\dormakaba\Messenger Lens\KeyDeliveryService\Kaba_KDS.exe'
$digitalPollingService = 'Program Files (x86)\dormakaba\Messenger Lens\DigitalKeysPollingSoftware\DigitalKeysPollingService.exe'
$lensGateway = 'Program Files (x86)\dormakaba\Messenger Lens\HubGatewayService\LENS_Gateway.exe'
If ($installOnC) {
	$issFolder = Join-Path $pluginSrcFolder 'ISS_FOR_C'
	$installDrive = $issFolder.Substring($issFolder.Length - 1, 1) + ':'
    $instDigitalExe = Join-Path $installDrive $digitalPolling
    $instLensPmsExe = Join-Path $installDrive $lensPms
    $pollingConfigInst = Join-Path $installDrive $digitalPollingConfig
    $lensPmsConfigFileInst = Join-Path $installDrive $lensPmsConfig
    $saflokIRS = Join-Path $installDrive $saflokV4Folder | Join-Path -ChildPath 'Saflok_IRS.EXE'
	$instHmsExe = Join-Path $installDrive $lensHMS
	$instKdsExe = Join-Path $installDrive $kabaKds
    $instDigitalPollingExe = Join-Path $installDrive $digitalPollingService
    $instGatewayExe = Join-Path $installDrive $lensGateway
}
If ($installOnD) {
	$issFolder = Join-Path $pluginSrcFolder 'ISS_FOR_D'
    $installDrive = $issFolder.Substring($issFolder.Length - 1, 1) + ':'
    $instDigitalExe = Join-Path $installDrive $digitalPolling
    $instLensPmsExe = Join-Path $installDrive $lensPms
    $pollingConfigInst = Join-Path $installDrive $digitalPollingConfig
    $lensPmsConfigFileInst = Join-Path $installDrive $lensPmsConfig
    $saflokIRS = Join-Path $installDrive $saflokV4Folder | Join-Path -ChildPath 'Saflok_IRS.EXE'
    $instHmsExe = Join-Path $installDrive $lensHMS
    $instKdsExe = Join-Path $installDrive $kabaKds
    $instDigitalPollingExe = Join-Path $installDrive $digitalPollingService
    $instGatewayExe = Join-Path $installDrive $lensGateway
}
$issFiles = Get-ChildItem $issFolder | Select-Object Name | Sort-Object -Property Name
$pollingISS = Join-Path $issFolder $issFiles[4].Name         # [4] polling.iss
# ----------------------------------------------------------------------------------------------
# Absolute installed FOLDER
$kabaInstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'dormakaba'
$lensInstFolder = Join-Path $kabaInstFolder 'Messenger Lens'
# ----------------------------------------------------------------------------------------------
# polling log
$pollingLog = 'C:\ProgramData\DormaKaba\Server\Polling\logs.log'
Logging "" "+---------------------------------------------------------"
Logging "" "| WELCOME TO CONFIGURATION SCRIPT"
Write-Colr -Text $cname, " |"," $hotelChain" -Colour White,cyan,Yellow
Logging "" "| SAFLOK VERSION: $saflokVersion"
Logging "" "+---------------------------------------------------------"
Logging " " ""
Logging " " "By installing you accept licenses for the packages."
$confirmation = Read-Host "$cname Do you want to run the script? [Y] Yes  [N] No"
If ($confirmation -eq 'Y' -or $confirmation -eq 'YES') {
	# -------------------------------------------------------------------
	# Install Digital Polling Service
	$isInstalled = 0
	$pName = "Marriott digital polling service"
	Install-DigitalPolling $pName $instDigitalExe $digitalPollingExe $pollingISS
	# -------------------------------------------------------------------
	# Allow everyone access to lens folder
	If (Test-Folder $lensInstFolder) {
		$acl = Get-Acl -Path $lensInstFolder
		$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","FullControl","Allow")
		$acl.SetAccessRule($AccessRule); $acl | Set-Acl $lensInstFolder
	}
	# -------------------------------------------------------------------
	# Copy config files
	$wsPmsInstFolder = Split-Path($instLensPmsExe)
	If (Test-Folder -folder $pmsConfig) {
		Copy-Item -Path $pmsConfig -Destination $wsPmsInstFolder -Force
    }
	$digitalPollingFolder = Split-Path($instDigitalExe) 
	If (Test-Folder -folder $digitalConfig) {
		Copy-Item -Path $digitalConfig -Destination $digitalPollingFolder -Force
    }	
	# -------------------------------------------------------------------
	# INSTALL WEB SERVICE PMS TESTER
	$fileCopied = 0 
	$wsptInstFolder = Join-Path $lensInstFolder 'Web_Service_PMS_Tester'
	Update-Copy $webServiceTester $wsptInstFolder
	Install-PmsTester $webServiceTester $lensInstFolder $wsptInstFolder
	$wsTesterExe = Join-Path $wsptInstFolder 'MessengerNet WSTestPMS.exe'
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
	If (Assert-IsInstalled "Messenger LENS") {
		# ----------------------------------------------------------------
		# CHECK SERVICES STATUS
		[string[]]$servicesCheck = 'DeviceManagerService','KIPEncoderService','Kaba_KDS','MessengerNet_Hub Gateway Service','MessengerNet_Utility Service','VirtualEncoderService','Kaba Digital Keys Polling Service'
		Foreach ($service In $servicesCheck) { 
			$serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
			If ($serviceStatus.Status -eq "stopped") {
				Logging " " "Staring service $service."
				Start-Service -Name $service -ErrorAction SilentlyContinue
				$serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
				If ($serviceStatus.Status -eq "running") { Logging " " "$service has been started."}
				Start-Sleep -S 1
			} Else {Logging "INFO" "$service is in running state.";Start-Sleep -S 1}
		} 
		# ----------------------------------------------------------------
		# FILES NEED TO BE CHECKED AND MODIFIED
		Logging "" "+---------------------------------------------------------"
		Logging "" "The following files need to be checked or modified: "
		Logging "" "+---------------------------------------------------------"
		# ----------------------------------------------------------------
		# RUN SAFLOK IRS
		If ($Null -eq (Get-Process | where-object {$_.Name -eq 'Saflok_IRS'}).ID) {
			Start-Process -NoNewWindow -FilePath $saflokIRS; Start-Sleep -S 1
		} 
		# ----------------------------------------------------------------
		# KILL ALL NOTEPAD PROCESS
		Get-Process -ProcessName notepad* | Stop-Process -Force; Start-Sleep -S 1
		# ----------------------------------------------------------------
		# PMS CONFIG
		If ((Assert-isInstalled  "Messenger LENS") -and (Test-Folder $lensPmsConfigFileInst)) {
			Logging " " "==> LENS_PMS.exe.config"
			Start-Process notepad $lensPmsConfigFileInst -WindowStyle Minimized; Start-Sleep -S 1
		} 
		# ----------------------------------------------------------------
		# DIGITAL POLLING CONFIG	
		If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Folder $pollingConfigInst)) {
			Logging " " "==> DigitalKeysPollingService.exe.config"
			Start-Process notepad $pollingConfigInst -WindowStyle Minimized; Start-Sleep -S 1
		} 
		# ----------------------------------------------------------------
		# SHORTCUT FOR POLLING LOG	
		If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Folder $pollingLog)) {
			Logging " " "==> Polling log"
			Start-Process notepad $pollingLog -WindowStyle Minimized; Start-Sleep -S 1
			$TargetFile = $pollingLog
			$ShortcutFile = "$env:Public\Desktop\PollingLog.lnk"
			$WScriptShell = New-Object -ComObject WScript.Shell
			$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
			$Shortcut.TargetPath = $TargetFile
			$Shortcut.Save()
			Start-Sleep -S 1
		} 
		Write-Colr -Text "$cname ","Please check those files opening in taskbar area." -Colour White,Magenta
		Logging "" "+---------------------------------------------------------"
		Start-Sleep -Seconds 1
		# ----------------------------------------------------------------	
		# SET SERVICE RECOVERY
		[string[]]$recoveryServices = 'MNet_HMS','Kaba_KDS','Kaba Digital Keys Polling Service'
		Foreach ($service In $recoveryServices){
			If (Get-service -Name $service | where-object {$_.StartType -ne 'Automatic'}) { Set-Service $service -StartupType "Automatic" }
			Set-ServiceRecovery -ServiceDisplayName $service
		} 
		# ----------------------------------------------------------------
        # FILES VERSION
        Logging "" "Installed Version Info:" 
        Logging "" "+---------------------------------------------------------"
        $gatewayVer = Get-FileVersion $instGatewayExe;$hmsVer = Get-FileVersion $instHmsExe;$wsPmsVer = Get-FileVersion $instLensPmsExe;$kdsVer = Get-FileVersion $instKdsExe;$pollingVer = Get-FileVersion $instDigitalPollingExe
        If (Test-Path $instGatewayExe -PathType Leaf) { Logging " " "Gateway: $gatewayVer"; Start-Sleep -Seconds 1 } 
        If (Test-Path $instHmsExe -PathType Leaf) { Logging " " "HMS:     $hmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $instLensPmsExe -PathType Leaf) { Logging " " "PMS:     $wsPmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $instKdsExe -PathType Leaf) { Logging " " "KDS:     $kdsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $instDigitalPollingExe -PathType Leaf) { Logging " " "POLLING: $pollingVer"; Start-Sleep -Seconds 1 }
        Logging "" "+---------------------------------------------------------"
        Logging "" "DONE"
        Logging "" "+---------------------------------------------------------"
        Write-Host ''
        # ----------------------------------------------------------------
        # CLEAN UP
        If (Test-Path -Path "$scriptPath\*.*" -Include *.ps1){Remove-Item -Path "$scriptPath\*.*" -Include *.ps1,*.lnk -Force -ErrorAction SilentlyContinue}
        If (Test-path -Path "C:\SAFLOK") { Remove-Item -Path "C:\SAFLOK" -Recurse -Force -ErrorAction SilentlyContinue }   
    } # end of MessengerLens is installed 
	Stop-Script
}# END OF YES

If ($confirmation -eq 'N' -or $confirmation -eq 'NO') {
    Logging " " ""
    Write-Colr -Text $cname," Thank you, Bye!" -Colour White,Gray
    Write-Host ''
    Stop-Script
} # END OF NO