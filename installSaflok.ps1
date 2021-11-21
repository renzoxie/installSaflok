﻿<#
    .NOTES
    =========================================================================
    Author: renzoxie@139.com
    Version Option: 5.45, 5.68, 6.11
    Create Date: 16 April 2019
    Modified Date: 1st Feb 2021
    Run the script as administrator
    =========================================================================

    .SYNOPSIS
    SAFLOK Lodging Server silent installation script

    .DESCRIPTION
    This script fully installation of SAFLOK Lodging Systems for projects with online and BLE systems automatically

    .EXAMPLE
    .\install.ps1 -inputDrive c -version 5.45

    .EXAMPLE
    .\install.ps1 -inputDrive c -version 5.45 -vendor 'dormakaba' -property 'Hotel Name'

    .LINK
    For deployments of this script, please see https://github.com/renzoxie/installSaflok
#>
[CmdletBinding()]
Param (
    [Parameter(Mandatory=$True)]
    [String]$inputDrive,

    [Parameter(Mandatory=$True)]
    [String]$version,

    [Parameter(Mandatory=$False)]
    [String]$property = 'vagrant',

    [Parameter(Mandatory=$False)]
    [String]$vendor = 'KABA'
)

# ---------------------------
# support for TLS 1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12  

# ---------------------------
# Versions
$scriptVersion = '2.2'
$miniPsRequire = '5.1' -AS [decimal]
$psVer = [string]$psversiontable.PSVersion.Major + '.' + [string]$psversiontable.PSVersion.Minor -AS [Decimal]
Switch ($version) {
    '5.45' {
        $progVersion = '5.4.0.0'
        $progPatchedVersion = '5.4.0.0'
        $pmsVersion = '5.1.0.0'
        $pmsPatchedVersion = '5.1.0.0'
        $msgrVersion = '5.2.0.0'
        $msgrPatchedVersion = '5.2.0.0'
        $msgrLensVersion = '4.7.0.0'
        $wsPmsExeBefPatchVersion = '4.7.1.15707'
        $wsPmsExeAftPatchVersion = '4.7.2.22767'
    }
    '5.68' {
        $progVersion = '5.6.0.0'
        $progPatchedVersion = '5.6.8.0'
        $pmsVersion = '5.6.0.0'
        $pmsPatchedVersion = '5.6.8.0'
        $msgrVersion = '5.6.0.0'
        $msgrPatchedVersion = '5.6.0.0'
        $msgrLensVersion = '5.6.0.0'
        $wsPmsExeBefPatchVersion = '5.6.0.0'
        $wsPmsExeAftPatchVersion = '5.6.7.22261'
    }
    '6.11' {
        $progVersion = '6.1.1.0'
        $progPatchedVersion = '6.1.1.0'
        $pmsVersion = '6.1.1.0'
        $pmsPatchedVersion = '6.1.1.0'
        $msgrVersion = '6.1.1.0'
        $msgrPatchedVersion = '6.1.1.0'
        $msgrLensVersion = '6.1.1.0'
        $wsPmsExeBefPatchVersion = '6.1.1.20121'
    }
}

# ---------------------------
# Program List
$saflokPrograms = [System.Collections.ArrayList]@(
    'Saflok Program';                   
    'Saflok PMS';                        #[1]
    'Saflok Messenger Server';           #[2]
    'Messenger Lens';                    #[3]
    'Marriott digital polling service'   #[4]
)

# ---------------------------
# Version options
$versionOptions = [System.Collections.ArrayList]@('5.45'; '5.68'; '6.11')

# ---------------------------
# DRIVE INFO
$inputDrive = $inputDrive.Trim().ToUpper()
$driveLetterPattern = "(\w{1})"

# ---------------------------
# valid if choco installed
$testchoco = powershell choco -v
If (-not($testchoco)) { 
    Set-ExecutionPolicy Bypass -Scope Process -Force 
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1')) | Out-Null
}

# ---------------------------
# Logging Messages
$lang = (Get-WinSystemLocale).Name
switch ($lang) {
    'zh-CN' {
        $mesgNoPkg ="安装包不存在，准备退出"
        $mesgInstalled = "已安装"
        $mesgDiffVer = "已安装其他版本程序，请卸载其他版本后再运行本脚本"
        $mesgFailed = "程序安装失败，准备退出"
        $mesgNoSource = "未找到源安装文件夹"
        $mesgToInstall = "程序正在安装，请稍等..."
        $mesgConfigIIS = "正在检查Messenger Lens所需IIS组件状态..."
        $mesgIISEnabled = "所有Messenger LENS所需组件已准备就绪"
        $mesgSucc = "程序安装成功"
        $mesgNotExe = "非执行文件，无法过去文件版本"
        $mesgNoFile = "文件不存在"
        $mesgLensNotInstalled = "请先安装Messenger LENS。"
        $mesgHugeInstaller = "程序较大，安装可能需要较长时间，请耐心等待..."
        $mesgReboot = "重启系统，然后尝试再次执行此脚本"
        $mesgContactSaflok = "如果重启后问题依旧, 请联系Saflok技术支持"
        $mesgConnectError = "数据库连接错误，请检查服务器名称，端口，防火墙配置"
        $mesgPsVer = "当前系统PowerShell版本是"
        $mesgPSMiniRequire = "此脚本需要PowerShell $miniPsRequire或者更高版本"
        $mesgPSDownloadUrl = "微软官网下载更高版本PowerShell网址: https://docs.microsoft.com/en-us/powershell/"
        $mesgRootAndTryAgain =  "请先安装Powershell V5以上版本, 重启系统后再次运行此脚本"
        $prompProperty = "酒店名称："
        $prompWelcome = "| 欢迎使用SAFLOK酒店系统安装脚本"
        $prompImprtant = " # 重要"
        $prompRunAsAdmin = " # 此脚本必须以管理员身份运行"
        $prompSafVer = "| SAFLOK 版本号: $version"
        $prompWinOS = "| 操作系统:"
        $prompScriptVer = "| 脚本版本号: $scriptVersion"
        $prompVerNoMatch = "输入版本与安装包不否，准备退出"
        $prompAccept = "选择安装即代表接受软件的许可协议"
        $prompDoU = "是否继续安装?(输入[Y]es/[N]o)"
        $mesgCheckPre = "正在检查依赖程序..."
        $mesgFinished = "完成安装"
        $prompStartUpdate = "开始安装更新"
        $prompFinishUpdate = "完成完整更新"
        $prompStartInstall = "开始安装..."
        $prompChoseDrive = "SAFLOK程序将安装在 ["+ $inputDrive+"] 盘"
        $prompChkConfig = "| 需要手动配置以下文件中的部分参数: "
        $mesgCouldNotIns = "无法将程序安装在 ["+$inputDrive+"] 盘，准备退出"
        $mesgRerun4Drive = "请重新运行本脚本，请选择正确的磁盘位置"
        $mesgCp2dFolder = "请将数据库文件拷贝至此文件夹:"
        $mesgTryScriptAgain = "数据库文件放置在上述文件夹后，请再次运行此脚本"
        $mesgVerNotCorrect = "输入的版本号不正确"
        $mesgRerun4Ver = "请重新运行本脚本，请选择正确的版本号"
        $mesgReboot = "准备重启系统，重启后请重新运行本脚本"
        $mesgDotnet462 = "需要先手动安装 .Net Framework V4.6.2"
        $mesgLongUpdate = "安装此更新大约需要半小时左右时间，等待时间较长"
		#-----------------
        # IIS
		$addinFeature = "正在添加组件"
		$enabledFeature = "已启用"
        #-----------------
        # Share Folder
        $shareFolderNotExist = "文件夹不存在"
        $sharingFolder = "正在处理文件夹共享..."
        $shareFolderDone = "成功共享文件夹"
        $shareAlreadyShared = "数据库文件夹已共享"
        #-----------------
        # Messenger Lens
        $mesgLensBefMessenger = "安装Lens之前，请先安装Saflok messenger程序"
        $mesgFailedEnableIIS = "添加IIS组件失败"
        #-----------------       
        # Services
        $mesgStartinService = "正在启动服务"
        $mesgServiceStarted = "服务已启动"
        $mesgServiceRunning = "服务正处于运行状态中"
        #-----------------       
        # config
        $prompIVI = "已安装版本信息: " 
        $mesgCheckConfig = "提示: 配置文件已最小化，请在任务栏处打开"
        $mesgtks = " 感谢安装Saflok酒店系统"
        $mesgNeedReboot = "# 友情提示: 安装过程序之后需要重启系统"
        $mesgMissingLens = "必须先安装Messenger Lens程序"
        $mesgbye = " 后会有期！"
        $mesgDone = "完成"
    }
    Default {
        $mesgNoPkg ="package does not exist, operation exit"
        $mesgInstalled = "already installed"
        $mesgDiffVer = "There is another version exist, please uninstall it first"
        $mesgFailed = "installation failed"
        $mesgNoSource = "Missing source installation folder"
        $mesgToInstall = "will now be installed, Please wait..."
        $mesgConfigIIS = "Checking IIS features status for Messenger Lens..."
        $mesgIISEnabled = "ALL IIS features Messenger LENS requires are ready"
        $mesgSucc = "install was successful"
        $mesgNotExe = "Can not get file version as it is not an executable file"
        $mesgNoFile = "Oops, File does not exist"
        $mesgLensNotInstalled = "Messenger LENS has not been installed yet"
        $mesgHugeInstaller = "The installer is greater than 100M, this could take more than 5 minutes, please wait..."
        $mesgReboot = "Reboot system and try the script again"
        $mesgContactSaflok = "If still same, please contact your SAFLOK representative"
        $mesgConnectError = "Connection Error. Check server name, port, firewall"
        $mesgPsVer = "Your current PowerShell version is"
        $mesgPSMiniRequire = "This script requires PowerShell version $miniPsRequire or above"
        $mesgPSDownloadUrl = "You can download newer version PowerShell at: https://docs.microsoft.com/en-us/powershell/"
        $mesgRootAndTryAgain =  "Reboot server after installing Powershell 5 or above, run this script again"
        $prompProperty = "Property: "
        $prompWelcome = "| WELCOME TO SAFLOK LODGING SYSTEMS INSTALLATION"
        $prompImprtant = " # IMPORTANT"
        $prompRunAsAdmin = " # THIS SCRIPT MUST BE RUN AS ADMINISTRATOR"
        $prompSafVer = "| SAFLOK VERSION: $version"
        $prompWinOS = "| INSTALLING ON:"
        $prompScriptVer = "| SCRIPT VERSION: $scriptVersion"
        $prompVerNoMatch = "Version input does NOT match corresponding source package"
        $prompAccept = "By installing you accept licenses for the packages"
        $prompDoU = "Do you want to run the script?([Y]es/[N]o)"
        $prompChkConfig = "| The following files need to be checked or configure: "
        $prompChoseDrive = "You chose drive ["+ $inputDrive+"]"
        $mesgCheckPre = "Checking prerequisites"
        $mesgFinished = "Finished installing"
        $prompStartUpdate = "Starting Update"
        $prompFinishUpdate = "Finished Update"
        $prompStartInstall = "Start checking installation..."
        $mesgCouldNotIns = "We could not install on drive "+$inputDrive
        $mesgRerun4Drive = "Please re-run the script again to input a correct drive"
        $mesgCp2dFolder = "Please copy database files to this folder:"
        $mesgTryScriptAgain = "Try this script again after database files have been loaded"
        $mesgVerNotCorrect = "The version number specified is NOT correct"
        $mesgRerun4Ver = "Please re-run the script again to input a correct version"
        $mesgReboot = "server will reboot in 15 seconds, rerun the script after server start up"
        $mesgDotnet462 = "Need to install .Net Framework V4.6.2 manually in advance"
        $mesgLongUpdate = "Install this update will take around 30 mins to complete"
        #-----------------
        # Share Folder
        $shareFolderNotExist = "Folder need to be shared does not exist"
        $sharingFolder = "Processing folder share"
        $shareFolderDone = "Folder share completed"
        $shareAlreadyShared = "The Saflok database folder already been shared"
		#-----------------
        # IIS
		$addinFeature = "Adding feature"
		$enabledFeature = "enabled"

        #-----------------
        # Messenger Lens
        $mesgLensBefMessenger = "Please install Saflok messenger before Lens"
        $mesgFailedEnableIIS = "Oops, fail to add IIS features for Messenger Lens"
        #-----------------       
        # Services
        $mesgStartinService = "Staring service"
        $mesgServiceStarted = "has been started"
        $mesgServiceRunning = "is in running state"
        #-----------------       
        # config
        $prompIVI = "Installed Version Information: " 
        $mesgCheckConfig = "Tip: Please check those config files opening in taskbar area"
        $mesgtks = " Thanks for installing Saflok"
        $mesgNeedReboot = "# NOTE: The recent program changes indicate a reboot is necessary."
        $mesgMissingLens = "Missing Messenger Lens program, Please make sure Messenger Lens is be installed properly"
        $mesgbye = " Thank you, Bye!"
        $mesgDone = "DONE"
    }
}

# ---------------------------
# Functions
# ---------------------------

# ---------------------------
# TEST INTERNET ACCESS
Function Test-Choco {
    Param (
        [string]$url
    )
    $result = !($null -eq (New-Object System.Net.WebClient).DownloadString($url))
    return $result
}

# Customized color
Function Write-Colr {
    [CmdletBinding()]
    Param ([String[]]$Text,[ConsoleColor[]]$Colour,[Switch]$NoNewline=$false)

    For ([int]$i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -Foreground $Colour[$i] -NoNewLine }
    If ($NoNewline -eq $false) { Write-Host '' }
}

# ---------------------------
# Customized logging
Function Logging {
    [CmdletBinding()]
    Param(
        [string]$state,
        [string[]]$message
    )

    $part1 = $cname;
    $part2 = ' ';
    $part3 = $state;
    $part4 = ": ";
    $part5 = "$message"
    Switch ($state)
    {
        'ERRO' {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Red,Red,Red}
        'WARN'  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Magenta,Magenta,Magenta}
        'INFO'  {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Yellow,Yellow,Yellow}
        'PROG' {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,White,White,White}
        'SUCC' {Write-Colr -Text $part1,$part2,$part3,$part4,$part5 -Colour White,White,Green,Green,Green}
        "" {Write-Colr -Text $part1,$part2,$part5 -Colour White,White,Cyan}
   }
}

# ---------------------------
# Stop Script
Function Stop-Script {
    [CmdletBinding()]
    Param(
        [int]$seconds = 60
    )
    Start-Sleep -Seconds $seconds
    exit
}

# ---------------------------
# Check file versions
Function Get-FileVersion {
    Param (
        $testFile
    )

 Switch (Test-Path $testFile -PathType Any) {
        $True {
            If ($testFile -like "*.exe") {
                (Get-Item $testFile).VersionInfo.FileVersion
            } Else {
                Logging "ERRO" "$mesgNotExe"
            }
        }
        $False {
            Write-Warning -Message "$mesgNoFile"
        }
    }
}

# ---------------------------
# GET INSTALLED VERSION
Function Get-InstVersion {
    [CmdletBinding()]
    Param (
        [String]$pName
    )
	$verInfo = (Get-Package -ProviderName "Programs" | Where-Object {$_.Name -eq $pName}).Version
    return $verInfo
}

# ---------------------------
# IF Installed, return Boolean
Function Assert-IsInstalled {
    [CmdletBinding()]
    Param (
        [String]$pName
    )
    $findIntallByName = (Get-Package -ProviderName "Programs" | Where-Object {$_.Name -eq $pName}).Name
    $condition = ($null -ne $findIntallByName)
    return $condition
}

# ---------------------------
# UPDATE INSTALLED STATUS
Function Update-Status {
    [CmdletBinding()]
    Param (
        [String]$pName
    )

    begin {}
    process {
        Switch (Assert-IsInstalled -pName $pName) {
            $true { $script:isInstalled = $true}
            $false { $script:isInstalled = $false}
        }
        $script:updateVersion = Get-InstVersion -pName $pName
    }
    end {}
}

# ---------------------------
# INSTALL BY CHOCO
Function Install-Choco {
    param(
        [string]$pName
    )
   Start-Process -Wait choco -ArgumentList "install $pName","-y" -ErrorAction SilentlyContinue
}

# ---------------------------
# INSTALL PROGRAM
Function Install-Prog {
    [CmdletBinding()]
    Param (
        [string]$pName,
        [string]$progVersion,
        [string]$progPatchedVersion,
        [string]$exeProgFile,
        [string]$exe2Install,
        [string]$iss2Install
    )

    begin {
        Update-Status -pName $pName
        $isExeExist = Test-Path $exeProgFile -PathType Any
    }
    process {
        Switch ($isInstalled)  {
            $true {
                If ($progVersion -ne $progPatchedVersion) {
                    If ($updateVersion -eq $progPatchedVersion) {
                        Logging "INFO" "$pName $mesgInstalled"
                        Start-Sleep -Seconds 2
                    } Elseif ($updateVersion -eq $progVersion) {
                        continue
                    } Else {
                        Logging "ERRO" "$mesgDiffVer"
                        Stop-Script 5
                    }
                } else {
                    Logging "INFO" "$pName $mesgInstalled"
                    Start-Sleep -Seconds 2
                }
            }
            $false {
                If ($isExeExist -eq 0) {
                    Logging "PROG" "$pName $mesgToInstall"
                    Start-Process -NoNewWindow -FilePath $exe2Install -ArgumentList " /s /f1$iss2Install " -Wait
                    Update-Status -pName $pName
                    If ($isInstalled) {
                        Logging "SUCC" "$pName $mesgSucc"
                        Start-Sleep -Seconds 2
                    } Else {
                        Logging "ERRO" "$pName $mesgFailed"
                        Stop-Script 5
                    }
                }
            }
        }
    }
    end {}
}

# ---------------------------
# INSTALL PROGRAM Patch
Function Install-Patch {
    [CmdletBinding()]
    Param (
        $pName,
        $progVersion,
        $progPatchedVersion,
        $patchExeFile,
        $patchIssFile
    )

    begin {
        Update-Status -pName $pName
    }
    process {
        If ($updateVersion -eq $progPatchedVersion)   {
            Logging "INFO" "$pName Patch $mesgInstalled "
            Start-Sleep -Seconds 2
        } Elseif ($updateVersion -eq $progVersion)   {
            Logging "PROG" "$pName patch $mesgToInstall"
            Start-Process -NoNewWindow -FilePath $patchExeFile -ArgumentList " /s /f1$patchIssFile" -Wait
            Start-Sleep -Seconds 2
            Update-Status -pName $pName
            If ($updateVersion -eq $progPatchedVersion) {
                Logging "SUCC" "$pName Patch $mesgSucc"
                Start-Sleep -Seconds 2
            } Else {
                Logging "ERRO" "$pName Patch $mesgFailed"
                Stop-Script 5
            }

        } Else {
            Logging "ERRO" "$mesgDiffVer"
            Stop-Script 5
        }

    }
    end {}
}

# ---------------------------
# Install Lens Patch -New
Function Install-LensPatch {
    [CmdletBinding()]
    Param (
        [String]$targetFile,
        [String]$destVersion,
        [String]$pName,
        [String]$patchExeFile,
        [String]$patchIssFile
    )

    begin {
        Update-Status -pName $pName
        $wsPmsExeVersion = Get-FileVersion -testFile $targetFile
    }
    process {
		If ($isInstalled) {
			Switch ($wsPmsExeVersion -eq $destVersion) {
				$true {
					Logging "INFO" "$pName Patch $mesgInstalled"
					Start-Sleep -Seconds 2
				}
				$false {
					If ($wsPmsExeVersion -eq $wsPmsExeBefPatchVersion) {
						Logging "PROG" "$pName Patch $mesgToInstall"
						Start-Process -NoNewWindow -FilePath $patchExeFile -ArgumentList " /s /f1$patchIssFile" -Wait
						Start-Sleep 5
						$wsPmsExeVersion = Get-FileVersion -testFile $targetFile
						try {
							Switch ($wsPmsExeVersion -eq $destVersion) {
								$true {
									Logging "SUCC" "$pName Patch $mesgSucc"
									Start-Sleep -Seconds 2
								}
								$false { 
									Logging "ERRO" "$pName Patch $mesgFailed"
									Stop-Script 5
								}
							}
						}
						catch {$Error[0]}
					} Else {
						Logging "ERRO" "$mesgDiffVer"
						Stop-Script 5
					}
				}
			}
		} Else {
			Logging "ERRO" "$pName $mesgFailed"
			Stop-Script 5
		}
    }
    end {}
}

# ---------------------------
# Install digitalPolling
Function Install-DigitalPolling {
    [CmdletBinding()]
    Param (
        [String]$pName,
        [String]$exe2Install,
        [String]$iss2Install
    )
    
    begin {
        $targetFile = "$digitalPollingExe"
        $isExeExist = Test-Path $targetFile -PathType Any
    }

    process {
        If($isExeExist) {
            Logging "INFO" "$pName $mesgInstalled"
        } Else {
            Logging "PROG" "$pName $mesgToInstall"
            Start-Process -NoNewWindow -FilePath $exe2Install -ArgumentList " /s /f1$iss2Install" -Wait
            $targetFile = "$digitalPollingExe"
            $isExeExist = Test-Path $targetFile -PathType Any
            If ($isExeExist){
                Logging "SUCC" "$pName $mesgSucc"
                Start-Sleep -Seconds 2
            } Else {
                Logging "ERRO" "$pName $mesgFailed"
                Stop-Script 5
            }
        }

    }

    end {}
}

# ---------------------------
# Update Copied Flies
Function Update-Copy {
    [CmdletBinding()]
    Param (
        [String]$srcPackage,
        [String]$instFolder0,
        [String]$instFolder1
    )

    begin {
        $fileCopied = 0
        Test-Path $srcPackage -PathType Any | Out-Null
        $testInstFolder0 = Test-Path $instFolder0 -PathType Any | Out-Null
        $testInstFolder1 = Test-Path $instFolder1 -PathType Any | Out-Null
    }
    process {
        If (($testInstFolder0 -eq $false) -and ($testInstFolder1 -eq $true)) {
            $script:fileCopied = 1
        }
    }
    end {}

}

# ---------------------------
# Install Web Service PMS Tester
Function Install-PmsTester {
    [CmdletBinding()]
    Param (
        [String]$srcPackage,
        [String]$instParent,
        [String]$instFolder0,
        [String]$instFolder1
    )

    begin {
        $pName = 'Web Service PMS Tester'
        $testSrcPackage  = Test-Path $srcPackage -PathType Any
        $testInstParent = Test-Path $instParent -PathType Any
        $testInstFolder0 = Test-Path $instFolder0 -PathType Any
        $testInstFolder1 = Test-Path $instFolder1 -PathType Any
    }
    process {
        If ($testSrcPackage -eq 0) {
            Logging "ERRO" "$mesgNoSource"
            Stop-Script 5
        }
        If ($testInstParent -eq $False) {
            Logging "ERRO" "$mesgLensNotInstalled"
            Stop-Script 5
        }
        Switch ($testInstFolder1) {
            $true {$fileCopied = $true}
            $false {$fileCopied = $false}
        }
        If ($fileCopied) {
            Logging "INFO" "$pName $mesgInstalled"
            Start-Sleep -Seconds 2
        } Else {
            If (($testSrcPackage -eq $true) -and ($testInstParent -eq $true)) {
                Logging "PROG" "$pName $mesgToInstall"
            }
            If (($testInstFolder0 -eq $true) -and ($testInstFolder1 -eq $false))  {
                Rename-Item -Path $instFolder0 -NewName $instFolder1 -force -ErrorAction SilentlyContinue
                Logging "SUCC" "$pName $mesgSucc"
                Start-Sleep -Seconds 2
            } Elseif (($testInstFolder0 -eq $false) -and ($testInstFolder1 -eq $false)) {
                Copy-Item $srcPackage -Destination $instParent -Recurse -Force -ErrorAction SilentlyContinue
                Rename-Item -Path $instFolder0 -NewName $instFolder1 -force -ErrorAction SilentlyContinue
                Logging "SUCC" "$pName $mesgSucc"
                Start-Sleep -Seconds 2
            } Elseif (($testInstFolder0 -eq $true) -and ($testInstFolder1 -eq $true))  {
                Remove-Item -Path $instFolder0 -Force -Recurse
                Logging "SUCC" "$pName $mesgSucc"
                Start-Sleep -Seconds 2
            } Else {
                Logging "SUCC" "$pName $mesgSucc"
                Start-Sleep -Seconds 2
            }
        }
    }
    end {}
}

# ---------------------------
# Install SQL
Function Install-Sql {
    [CmdletBinding()]
    Param (
        [String]$pName,
        [String]$packageFolder,
        [String]$exe2Install,
        [String]$argFile
    )

    begin {
        Update-Status -pName $pName
        $isPkgFolderExist = Test-Path $packageFolder -PathType Any -IsValid
    }

    Process {
        If ($isInstalled) {
            Logging "INFO" "$pName $mesgInstalled"
        } Else {
            Switch ($isPkgFolderExist) {
                $false {
                    Logging "ERRO" "$pName $mesgNoPkg"
                    Stop-Script 5
                }
                $true {
                    Logging "PROG" "$pName $mesgToInstall"
                    Logging "INFO" "$mesgHugeInstaller"
                    Start-Process -NoNewWindow -FilePath $exe2Install -ArgumentList " $argFile" -Wait
                    Start-Sleep -Seconds 2
                    Update-Status -pName $pName
                    If ($isInstalled) {
                        Logging "SUCC" "$pName $mesgSucc"
                        Start-Sleep -Seconds 2
			        } Else {
				        Logging "ERRO" "$pName $mesgFailed"
				        Logging "ERRO" "$mesgReboot"
				        Logging "ERRO" "$mesgContactSaflok"
				        Stop-Script 5
			        }
                }

            }
        }
    }

    end {}
}

# ---------------------------
# update SQL password 
Function Update-SqlPasswd {
    Param(
        [string]$login,
        [string]$passwd
    )

    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.Smo") | Out-Null
    $objSQLConnection = New-Object System.Data.SqlClient.SqlConnection
    Switch ($version -lt 6) {
        $True  {$SqlServerName = 'localhost\LENSSQL'}
        $False {$SqlServerName = 'localhost\LENSSQL2016'}
    }
    Try {
        $objSQLConnection.ConnectionString = "Server=$SqlServerName;Integrated Security=SSPI;"
        $objSQLConnection.Open() | Out-Null
        $objSQLConnection.Close()
    } Catch {
        $errText =  $Error[0].ToString()
        If ($errText.Contains("network-related")) {
            Logging "ERRO" "$mesgConnectError"
        }
        Logging "ERRO" "$errText"
        continue
    }
    $srv = New-Object "Microsoft.SqlServer.Management.Smo.Server" $SqlServerName
    $SQLUser = $srv.Logins | Where-Object {$_.Name -eq "$login"};
    $SQLUser.ChangePassword($passwd);
    $SQLUser.PasswordPolicyEnforced = 1;
    $SQLUser.Alter();
    $SQLUser.Refresh();
}

# -----------------------
# $recoveryServices
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
# Header variables
$cname = "[$vendor]"
$property = $property.trim().toUpper()   
$time = Get-Date -Format 'yyyy/MM/dd HH:mm'
$Global:shareName = 'SaflokData'
# Windows OS version in decimal
#$osversion = (Get-CimInstance -ClassName CIM_OperatingSystem).version.split(".") -AS [array]
#$winOS = ($osversion[0] + '.' + $osversion[1]) -AS [decimal]
# Windows OS information
$osDetail = (Get-CimInstance -ClassName CIM_OperatingSystem).Caption

# ---------------------------
# Mini Powershell version requirement
If ($psVer -lt $miniPsRequire) {
    Logging "INFO" "$mesgPsVer $psVer"
    Logging "INFO" "$mesgPSMiniRequire"
    If (Test-Choco -url 'https://chocolatey.org/install.ps1') {
        Logging "PROG" "PowerShell version 5.1 $mesgToInstall" 
        start-sleep -seconds 4
        Try {
            Install-Choco -pName 'Powershell'
        }
        catch { $_;stop-script -seconds 2}
        Logging "WARN" "$mesgReboot"
        Start-Sleep -Seconds 15
        Restart-Computer -Force
    } Else {
        Logging "INFO" "$mesgPSDownloadUrl"
        logging "INFO" "$mesgRootAndTryAgain"
    }
}

# ---------------------------
# MENU OPTION
Clear-Host
Logging "" "+---------------------------------------------------------"
Logging "" $prompWelcome
Logging "" "+---------------------------------------------------------"
Write-Colr -Text $cname, " |",$prompImprtant  -Colour White,cyan,Red
Write-Colr -Text $cname, " |",$prompRunAsAdmin -Colour White,cyan,red
Logging "" "+---------------------------------------------------------"
Logging "" "| $time"
Write-Colr -Text "$cname"," | ","$prompProperty","$property" -Colour White,cyan,cyan,Green
Logging "" "$prompSafVer"
Logging "" "$prompWinOS $osDetail"
Logging "" "+---------------------------------------------------------"
Logging "" "$prompScriptVer"
Logging "" "+---------------------------------------------------------"
Logging "" ""

# Get drive IDs from localhost
$driveIDs = (Get-WmiObject -class Win32_LogicalDisk -computername localhost -filter "drivetype=3").DeviceID
$driveLetters = @()
for ($i=0; $i -lt $driveIDs.Length; $i++) {
    if ($driveIds[$i] -match $driveLetterPattern) {
        $driveLetters += $matches[0]
    }
}

# ---------------------------
# Driver letter + :
$installDrive = $inputDrive + ':'

# ---------------------------
# SOURCE FOLDER - INSTALL SCRIPT
$packageFolders =  Get-ChildItem ($PSScriptRoot) | Select-Object Name | Sort-Object -Property Name

# ---------------------------
# if source folder exist
Switch ($version) {
    '5.45' {$pattern = "([A-Z]{8}_[A-Z]{9}_[A-Z]{4}_\d{3}_\w{3}\d{4}$)"}
    '5.68' {$pattern = "([A-Z]{9}_[A-Z]{4}_\d{3}_\w{3}\d{4}$)"}
    '6.11' {$pattern = "([A-Z]{8}_[A-Z]{9}_[A-Z]{4}_\d{3}_\w{3}\d{4}$)"}
}
$ver2Int = $version.Replace(".", "")
$verNoFromRootPath = $PSScriptRoot.Substring($PSScriptRoot.Length -11,3)
Switch ($PSScriptRoot -match $pattern) {
    $True  {
                If ([int]$ver2Int -ne [int]$verNoFromRootPath) {
                    Logging "ERRO" "$prompVerNoMatch"
                    Stop-Script 5
                } Else {
                    # -----------------------
                    # HEADER Information
                    Logging "" ""
                    Write-Colr -Text "$cname ","$prompAccept" -Colour White,White
                    $confirmation = Read-Host "$cname $prompDoU"
                    $confirmation = $confirmation.ToUpper()
                    Logging "" ""
                }
            }
    $False  {
                Logging "ERRO" "$mesgNoSource"
                Stop-Script 5
            }
}

$absPackageFolders = @()
For ($i=0; $i -lt ($packageFolders.Length-1); $i++) {
    $absPackageFolders += Join-Path $PSScriptRoot $packageFolders[$i].Name
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
# SQL Express
Switch ($version -gt 6) {
    $false {
        #2012 
        $sqlExprExe =  Join-Path $lensSrcFolder '/AutoPlay/Install Script/Lens/en' |
                       Join-Path -ChildPath 'ISSetupPrerequisites' |
                       Join-Path -ChildPath '{C38620DE-0463-4522-ADEA-C7A5A47D1FF6}' |
                       Join-Path -ChildPath 'SQLEXPR_x86_ENU.exe'    
    }
    $True {
        #2016
        $sqlExprExe =  Join-Path $lensSrcFolder '/AutoPlay/Install Script/Lens/en' |
                       Join-Path -ChildPath 'ISSetupPrerequisites' |
                       Join-Path -ChildPath '{DEA37953-690E-42ED-B1D0-E75C59D41454}' |
                       Join-Path -ChildPath 'SQLEXPR_x64_ENU.exe'     
    }
}

# ---------------------------
# ISS_FOR_Drive & files [FILE]
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
$progISS = Join-Path $issFolder $issFiles[0].Name
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
    '6.11' {
            $pmsISS = Join-Path $issFolder $issFiles[1].Name
            $msgrISS = Join-Path $issFolder $issFiles[2].Name
            $lensISS = Join-Path $issFolder $issFiles[3].Name
            $patchPollingISS = Join-Path $issFolder $issFiles[4].Name
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
            $progPatchExe = Join-Path $absPackageFolders[7]  $patchExeFiles[0].Name
            $pmsPatchExe = Join-Path $absPackageFolders[7]  $patchExeFiles[1].Name
            $lensPatchExe = Join-Path $absPackageFolders[7] $patchExeFiles[2].Name
    }
    '6.11' {
            $pollingPatchExe = Join-Path $absPackageFolders[7]  $patchExeFiles[0].Name
    }
}

# ---------------------------
# Web service PMS tester [FOLDER]
$webServiceTester = $absPackageFolders[8]
if($version -ne '5.68') {
	# ConfigFiles Folder & Files
	$configFiles = Get-ChildItem ($absPackageFolders[9]) | Select-Object Name | Sort-Object -Property Name
	$pollingConfig = Join-Path $absPackageFolders[9] $configFiles[0].Name
	$lensPmsConfig = Join-Path $absPackageFolders[9] $configFiles[1].Name 
}

# ---------------------------
# Absolute installed FOLDER
Switch ($version -gt 6) {
    $true {
        $saflokV4InstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'SaflokV4'
        $kabaInstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'dormakaba'
    }
    $false {
        $saflokV4InstFolder = Join-Path $installDrive 'SaflokV4'  
        $kabaInstFolder = Join-Path $installDrive 'Program Files (x86)' | Join-Path -ChildPath 'KABA'
    }

}

$lensInstFolder = Join-Path $kabaInstFolder 'Messenger Lens'
$hubGateWayInstFolder = Join-Path $lensInstFolder 'HubGatewayService'
$hmsInstFolder = Join-Path $lensInstFolder 'HubManagerService'
$pmsInstFolder = Join-Path $lensInstFolder 'PMS Service'
$wsTesterInstFolder = Join-Path $kabaInstFolder '08.Web_Service_PMS_Tester'
if ($version -ne '5.68') {
    $digitalPollingFolder = Join-Path $lensInstFolder 'DigitalKeysPollingSoftware'
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
If ($version -ne '5.68') {
     $digitalPollingExe = Join-Path $digitalPollingFolder 'DigitalKeysPollingService.exe'
}
$kdsExe = Join-Path $kdsInstFolder 'Kaba_KDS.exe'
# ---------------------------
# INSTALLED CONFIG FILES
$lensPmsConfigFileInst =  Join-Path $pmsInstFolder 'LENS_PMS.exe.config'
$hh6ConfigFile = Join-Path $saflokV4InstFolder 'KabaSaflokHH6.exe.config'
IF ($version -ne '5.68') {
    $pollingConfigInst  = Join-Path $digitalPollingFolder 'DigitalKeysPollingService.exe.config'
    # ---------------------------
    # Polling log
    $pollingLog = 'C:\ProgramData\DormaKaba\Server\Polling\logs.log'
}

# ---------------------------
# Start to install
If ($confirmation -eq 'Y' -or $confirmation -eq 'YES') {
    # ---------------------------
    # VALID DRIVE CHARACTER INPUT
    if ($inputDrive -IN $driveLetters) {
        Logging "" "$prompChoseDrive"
        Start-Sleep -seconds 2
    } else {
        Logging "ERRO" "$mesgCouldNotIns"
        Logging "WARN" "$mesgRerun4Drive"
        Stop-Script 5
    }

    # ---------------------------
    # VALID Version INPUT
    if ($version -IN $versionOptions) {
        Logging "" "$prompStartInstall"
        Start-Sleep -seconds 2
    } else {
        Logging "ERRO" "$mesgVerNotCorrect"
        Logging "WARN" "$mesgRerun4Ver "
        Stop-Script 5
    }

    If ($version -eq 6.11) {
        # ---------------------------
        # check and install hot-fix
        Logging "" ""
        Start-Sleep -Seconds 1
        Logging "PROG" "$mesgCheckPre"
        Start-Sleep -Seconds 2
        $dotNetVersion = (Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release
        # check if v4.62 is installed
        If(($dotNetVersion -ge '394802') -or  ($dotNetVersion -eq '394806')) {
            Logging "INFO" ".Net Framework V4.6.2 $mesgInstalled"   
        } Else {
            # check hotfix KB2919355
            Switch ($null -eq (get-hotfix | where {$_.HotFixID -eq 'KB2919355' })) {
                # need install update KB2919355
                $true {
                    If (Test-Choco -url 'https://chocolatey.org/install.ps1') {
                        Logging "PROG" "KB2919355 $mesgToInstall"
                        Logging "INFO" "$mesgLongUpdate"
                        Start-Sleep -Seconds 5
                        Install-Choco -pName 'KB2919355' 
                        Start-Sleep -Seconds 5
                        Logging "WARN" "$mesgReboot"
                        start-sleep -Seconds 15
                        Restart-Computer -Force
                    } Else {
                        Logging "INFO" "$mesgDotnet462"
                        stop-script -seconds 5
                    }
                }
                $false {
                    If (Test-Choco -url 'https://chocolatey.org/install.ps1') {
                        Logging "PROG" ".Net Framework V4.6.2 $mesgToInstall"
                        Logging "INFO" "$mesgLongUpdate"
                        Start-Sleep -Seconds 5
                        Install-Choco -pName 'netfx-4.6.2'
                        Logging "INFO" "$mesgFinished .NetFramework V4.6.2"
                        start-sleep -Seconds 5
                        Logging "WARN" "$mesgReboot"
                        start-sleep -Seconds 15
                        Restart-Computer -Force
                    } Else {
                        Logging "INFO" "$mesgDotnet462" 
                        stop-script -seconds 5
                    }
                }
            }
        } 
    }

    # -------------------------------------------------------------------
    # install Saflok client
    $pName = $saflokPrograms[0]
    Install-Prog -pName $pName -progVersion $progVersion -progPatchedVersion $progPatchedVersion -exeProgFile $saflokClient -exe2Install $progExe -iss2Install $progISS
    # -------------------------------------------------------------------
    # install Saflok Program Patch
    if ($version -eq '5.68') {
        $pName = $saflokPrograms[0]
        Install-Patch -pName $pName -progVersion $progVersion -progPatchedVersion $progPatchedVersion -patchExeFile $progPatchExe -patchIssFile $patchProgISS
        # -------------------------------------------------------------------
        # [ clean munit ink ]
        $munit = 'C:\Users\Public\Desktop\Kaba Saflok M-Unit.lnk'
        If (Test-Path -Path $munit){
            Remove-Item $munit -Force
        }
    }

    # -------------------------------------------------------------------
    # install Saflok PMS
    $pName = $saflokPrograms[1]
    Install-Prog -pName $pName -progVersion $pmsVersion -progPatchedVersion $pmsPatchedVersion -exeProgFile $saflokIRS `
                -exe2Install $pmsExe -iss2Install $pmsISS
    # -------------------------------------------------------------------
    # install Saflok PMS Patch
    if ($version -eq '5.68') {
        $pName = $saflokPrograms[1]
        Install-Patch -pName $pName -progVersion $pmsVersion -progPatchedVersion $pmsPatchedVersion -patchExeFile $pmsPatchExe `
                    -patchIssFile $patchPmsISS
    }

    # -------------------------------------------------------------------
    # install Saflok Messenger
    $pName = $saflokPrograms[2]
    Install-Prog -pName $pName -progVersion $msgrVersion -progPatchedVersion $msgrPatchedVersion -exeProgFile $saflokMsgr `
                    -exe2Install $msgrExe -iss2Install $msgrISS
    # -------------------------------------------------------------------
    # copy database to saflokdata folder
    $srcHotelData = ($absPackageFolders[0])
    $srcGdb = (Get-ChildItem -Path $srcHotelData).Name
    If ((($srcGdb -match '^SAFLOKDATAV2.GDB$').Count -eq 0) -or (($srcGdb -match '^SAFLOKLOGV2.GDB$').Count -eq 0) -or ($null -eq $srcGdb)) {
        Logging "WARN" "$mesgCp2dFolder"
        Logging "WARN" "$srcHotelData"
        Logging "WARN" "$mesgTryScriptAgain"
        Write-Host ''
        Stop-Script 5
    }
    $instGdb = (Get-ChildItem -Path $shareFolder).Name
    If ((($instGdb -match '^SAFLOKDATAV2.GDB$').Count -eq 0) -or (($instGdb -match '^SAFLOKLOGV2.GDB$').Count -eq 0)) {
        If ((Get-Service -Name FirebirdGuardianDefaultInstance).Status -eq "Running") { 
            # stop firebird service
            Stop-Service -Name FirebirdGuardianDefaultInstance -Force -ErrorAction SilentlyContinue
            # copy database
            Copy-Item -Path $srcHotelData\*.gdb -Destination $shareFolder        
        } Else {
            Copy-Item -Path $srcHotelData\*.gdb -Destination $shareFolder
        }
    }

    # -------------------------------------------------------------------
    # share database folder
    try {
        $isShareFolderExist = Test-Path -Path $shareFolder -IsValid
        If(!($isShareFolderExist)) {
            Logging "ERRO" "$shareFolderNotExist"
            Stop-Script -seconds 5
        } Else {
            $getSmbShare = Get-SmbShare | Where-Object {$_.Name -eq $shareName}
            Switch ($null -ne $getSmbShare) {
                $True {
                    Logging "INFO" "$shareAlreadyShared"
                    Start-Sleep -S 2
                }
                $False {
                    Logging "PROG" "$sharingFolder"
                    New-SmbShare -Name $shareName -Path $shareFolder -FullAccess "everyone" -Description "Saflok database folder share" | Out-Null
                    Start-Sleep -S 2
                    Logging "SUCC" "$shareFolderDone"
                }
            }          
        }
    }
    catch {
        Logging "WARN" "$Error[0]"
        Stop-Script -seconds 5
    }

    # -------------------------------------------------------------------
    # start firebird service
    If (Get-Service | Where-Object {$_.Name -eq 'FirebirdGuardianDefaultInstance'-and $_.Status -eq "Stopped"}) {
        Start-Service -Name 'FirebirdGuardianDefaultInstance' -ErrorAction SilentlyContinue
        Start-Sleep -S 2
    }
    
    # -------------------------------------------------------------------
    # start Saflok launcher service
    If (Get-Service | Where-Object {$_.Name -eq 'SaflokServiceLauncher' -and $_.Status -eq "Stopped"}){
        Start-Service -Name 'SaflokServiceLauncher' -ErrorAction SilentlyContinue
        Start-Sleep -S 2
    }

    # -------------------------------------------------------------------
    # IIS FEATURES, requirement for messenger lens
    # Removed the support for win7 and win server 2008, as they were retired
    Update-Status -pName $saflokPrograms[2]
	If ($isInstalled -ne $True) {
		Logging "WARN" "$mesgLensBefMessenger"
		Stop-Script 5
	} Else {
        Logging "INFO" "$mesgConfigIIS"
        Start-Sleep -Seconds 3
        # ---------------------------
        # IIS features Messenger Lens requires
        $iisFeatures = [System.Collections.ArrayList]@(
            'IIS-WebServerRole', 'IIS-WebServer', 'IIS-CommonHttpFeatures', 'IIS-HttpErrors', 'IIS-ApplicationDevelopment',
        	'IIS-RequestFiltering', 'IIS-HealthAndDiagnostics', 'IIS-HttpLogging', 'IIS-Performance', 'IIS-ISAPIExtensions',
        	'IIS-ISAPIFilter', 'IIS-StaticContent', 'IIS-DefaultDocument', 'IIS-DirectoryBrowsing', 'IIS-ASP',
            'IIS-ManagementConsole', 'IIS-HttpCompressionStatic', 'NetFx4Extended-ASPNET45', 'IIS-ASPNET45',
        	'IIS-NetFxExtensibility45', 'IIS-Security', 'IIS-WebServerManagementTools', 'IIS-ApplicationInit'
        )
        # arrays to collect features in disabled state
        $disabledFeatures = @()
		$totalFeatures = $iisFeatures.count
		Try {
            For ($i=0; $i -lt $totalFeatures; $i++) {

		        $feature = $iisFeatures[$i]
		        $featureList =  Get-WindowsOptionalFeature -Online |
		        Select-Object -Property @{Name='Name'; expression = {$_.FeatureName}}, @{Name='State'; expression = {$_.State}} |
		        Where-Object {$_.Name -eq $feature} -ErrorAction SilentlyContinue
		        
                # get feature name and state
		        $featureName = $featureList.Name
		        $featureState = $featureList.State
				# add feature in disabledFeatures
		        if ((Get-WindowsOptionalFeature -Online | Where-Object {$_.FeatureName -eq $feature}).State -eq "Disabled") {
		            $disabledFeatures += $feature
		        }
		    }

		    If ($disabledFeatures.length -eq 0) {
                    Write-Colr -Text "$cname ", "$mesgIISEnabled" -Colour White,Gray
		            Start-Sleep -Seconds 1
            } Else {
		        foreach ($disabled in $disabledFeatures) {
		            Logging "PROG" "$addinFeature $disabled"
		            Enable-WindowsOptionalFeature -Online -FeatureName $disabled -All -NoRestart | Out-Null
		            Logging "SUCC" "$disabled $enabledFeature"
		            Start-Sleep -Seconds 1
		        }
		    } 
		}

		catch {
		    Write-Warning -Message "$mesgFailedEnableIIS"
            Stop-Script 2
		}
    }

    # -------------------------------------------------------------------
    # Microsoft SQL Server
    Switch ($version -gt 6) {
        $False {$pName = 'Microsoft SQL Server 2012'
            $argFile = '/qs /INSTANCENAME="LENSSQL" /ACTION="Install" /Hideconsole /IAcceptSQLServerLicenseTerms="True" '
            $argFile += '/FEATURES=SQLENGINE,SSMS /HELP="False" /INDICATEPROGRESS="True" /QUIETSIMPLE="True" /X86="True" /ERRORREPORTING="False" '
            $argFile += '/SQMREPORTING="False" /SQLSVCSTARTUPTYPE="Automatic" /FILESTREAMLEVEL="0" /FILESTREAMLEVEL="0" /ENABLERANU="True" '
            $argFile += '/SQLCOLLATION="Latin1_General_CI_AS" /SQLSVCACCOUNT="NT AUTHORITY\SYSTEM" /SQLSYSADMINACCOUNTS="BUILTIN\Administrators" '
            $argFile += '/SECURITYMODE="SQL" /ADDCURRENTUSERASSQLADMIN="True" /TCPENABLED="1" /NPENABLED="0" /SAPWD="P@ssw0rd2021"'
        }
        $True {
            $pName = 'Microsoft SQL Server 2016 (64-bit)'
            $argFile = '/qs /QUIETSIMPLE /IAcceptSQLServerLicenseTerms /ACTION=install /FEATURES=SQL /INSTANCENAME="LENSSQL2016" '
            $argFile += '/SQLSVCACCOUNT="NT Authority\System" /SQLSYSADMINACCOUNTS="BUILTIN\Administrators" '
            $argFile += '/AGTSVCACCOUNT="NT Authority\System" /SECURITYMODE=SQL /SAPWD="P@ssw0rd2021"'      
        }
    }
    Install-Sql -pName $pName -packageFolder $sqlExprExe -exe2Install $sqlExprExe -argFile $argFile
    
    Switch ($version -gt 6) {
        $false {
            If (Assert-IsInstalled 'Microsoft SQL Server 2012') {
                Try { Update-SqlPasswd -login 'sa' -passwd 'Lens2014' }
                catch {$ERROR[0]}
            }            
        }
        $True {
            If (Assert-IsInstalled 'Microsoft SQL Server 2016 (64-bit)') {
                Try { Update-SqlPasswd -login 'sa' -passwd 'S@flok2018' }
                catch { $ERROR[0] }    
            }
        }  
    }
    # -------------------------------------------------------------------
    # install Messenger Lens
    $pName = $saflokPrograms[3]
    Install-Prog -pName $pName -progVersion $msgrLensVersion -progPatchedVersion $msgrLensVersion `
                -exeProgFile $wsPmsExe -exe2Install $lensExe -iss2Install $lensISS
    If (($version -eq '6.11') -and ((Assert-IsInstalled -pName $pName) -eq $False)) {Logging "WARN" "$mesgReboot"}
    # -------------------------------------------------------------------
    # install Messenger Lens patch
    If ($version -ne '6.11') {
        $pName = $saflokPrograms[3]
        Install-LensPatch -targetFile $wsPmsExe -destVersion $wsPmsExeAftPatchVersion -pName $pName -patchExeFile `
                          $lensPatchExe -patchIssFile $patchLensISS
    }

    # -------------------------------------------------------------------
    # install digital polling service
    If ($version -ne '5.68') {
        $isInstalled = 0
        $pName = $saflokPrograms[4]
        Install-DigitalPolling -pName $pName -exe2Install $pollingPatchExe -iss2Install $patchPollingISS
    }

    # -------------------------------------------------------------------
    # allow everyone access to Lens folder
    If (Test-Path -Path $lensInstFolder -PathType Container) {
        $acl = Get-Acl -Path $lensInstFolder
        $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("everyone","FullControl","Allow")
		try {
			$acl.SetAccessRule($AccessRule)
        	$acl | Set-Acl $lensInstFolder
		}
		catch {Logging "WARN" "$ERROR[0]"}
    }

    # -------------------------------------------------------------------
    # copy config files
    If ($version -eq '5.45') {
        Copy-Item -Path $lensPmsConfig -Destination $pmsInstFolder -Force
        Copy-Item -Path $pollingConfig -Destination $digitalPollingFolder -Force
    }
    # INSTALL WEB SERVICE PMS TESTER
    # -------------------------------------------------------------------
    $newFolder0 = $kabaInstFolder + '\' + $wsTesterInstFolder.Substring($wsTesterInstFolder.Length - 25,25)
    $newFolder1 = $kabaInstFolder + '\' + $wsTesterInstFolder.Substring($wsTesterInstFolder.Length - 22,22)
    Update-Copy $webServiceTester $newFolder0 $newFolder1
    Install-PmsTester $webServiceTester $kabaInstFolder $newFolder0 $newFolder1
    $wsTesterExe = Join-Path $newFolder1 'MessengerNet WSTestPMS.exe'
    If ((Test-Path -path $wsTesterExe -PathType Leaf)) {
        $TargetFile = $wsTesterExe
        $ShortcutFile = "$env:Public\Desktop\WS_PMS_TESTER.lnk"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.TargetPath = $TargetFile
        $Shortcut.IconLocation = "C:\Windows\System32\SHELL32.dll, 12"
        $Shortcut.Save()
        Start-Sleep -S 2
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
	Update-Status -pName $saflokPrograms[3]
    If ($isInstalled) {
        If ($version -eq '5.68') {
            $servicesCheck.Remove('Kaba Digital Keys Polling Service')
        }
        Foreach ($service In $servicesCheck) {
            $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
            If ($serviceStatus.Status -eq "stopped") {
                Logging "" "$mesgStartinService $service."
                Start-Service -Name $service -ErrorAction SilentlyContinue
                $serviceStatus = Get-Service | Where-Object {$_.Name -eq $service}
                If ($serviceStatus.Status -eq "running") { Logging "" "$service $mesgServiceStarted"}
                Start-Sleep -S 2
            } Else {
                Logging "INFO" "$service $mesgServiceRunning"
                Start-Sleep -S 2
            }
        }
        # ----------------------------------------------------------------
        # OPEN GUI AND CONFIG FILE
        Logging "" "+---------------------------------------------------------"
        Logging "" "$prompChkConfig"
        Logging "" "+---------------------------------------------------------"
        If ($Null -eq (Get-Process | where-object {$_.Name -eq 'Saflok_IRS'}).ID) {
            Start-Process -NoNewWindow -FilePath $saflokIRS; Start-Sleep -S 1
        } # run IRS GUI
        Get-Process -ProcessName notepad* | Stop-Process -Force; Start-Sleep -S 1
        If ((Assert-isInstalled $saflokPrograms[0]) -and (Test-Path -path $hh6ConfigFile -PathType Leaf)) {
            Logging "" "[ KabaSaflokHH6.exe.config ]"
            Start-Process notepad $hh6ConfigFile -WindowStyle Minimized; Start-Sleep -S 1
        } # hh6 config
        If ((Assert-isInstalled  $saflokPrograms[3]) -and (Test-Path -Path $lensPmsConfigFileInst -PathType Leaf)) {
            Logging "" "[ LENS_PMS.exe.config ]"
            Start-Process notepad $lensPmsConfigFileInst -WindowStyle Minimized; Start-Sleep -S 1
        } # PMS config
        If($version -eq '5.45') {
            If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Path -Path $pollingConfigInst -PathType Leaf)) {
                Logging "" "[ DigitalKeysPollingService.exe.config ]"
                Start-Process notepad $pollingConfigInst -WindowStyle Minimized; Start-Sleep -S 1
            } # polling config
            If ((Test-Path -Path $digitalPollingExe -PathType Leaf) -and (Test-Path -Path $pollingLog -PathType Leaf)) {
                Logging "" "[ Polling log ]"
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
        Write-Colr -Text "$cname ","$mesgCheckConfig" -Colour White,Gray
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
        Write-Colr -Text "$cname ","$prompIVI" -Colour White,White
        Logging "" "+---------------------------------------------------------"
        $gatewayVer = Get-FileVersion $gatewayExe;
        $hmsVer = Get-FileVersion $hmsExe;
        $wsPmsVer = Get-FileVersion $wsPmsExe;
        $kdsVer = Get-FileVersion $kdsExe;
        if ($version -eq '5.45') {
            $pollingVer = Get-FileVersion $digitalPollingExe
        }
        If (Test-Path $gatewayExe -PathType Leaf) { Logging "" "Gateway: $gatewayVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $hmsExe -PathType Leaf) { Logging "" "HMS:     $hmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $wsPmsExe -PathType Leaf) { Logging "" "PMS:     $wsPmsVer"; Start-Sleep -Seconds 1 }
        If (Test-Path $kdsExe -PathType Leaf) { Logging "" "KDS:     $kdsVer"; Start-Sleep -Seconds 1 }
        If ($version -eq '5.45') {
            If (Test-Path $digitalPollingExe -PathType Leaf) { Logging "" "POLLING: $pollingVer"; Start-Sleep -Seconds 1 }
        }
        Logging "" "+---------------------------------------------------------"
        Write-Colr -Text "$cname ","$mesgDone" -Colour White,Green
        Logging "" "+---------------------------------------------------------"
        Logging "" ""
        Write-Colr -Text $cname,"$mesgtks" -Colour White,Green
        Write-host ""
        Write-host "$mesgNeedReboot" -ForegroundColor Gray
        Write-Host "";Write-Host ""
        # clean up 
        If (Test-Path -Path "$PSScriptRoot\*.*" -Include *.ps1){Remove-Item -Path "$PSScriptRoot\*.*" -Include *.ps1 -Force -ErrorAction SilentlyContinue}
        If (Test-path -Path "C:\SAFLOK") { Remove-Item -Path "C:\SAFLOK" -Recurse -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Second 2
    } Else {
        Logging "ERRO" "$mesgMissingLens"
        Stop-Script 5
    }

}
If ($confirmation -eq 'N' -or $confirmation -eq 'NO') {
    Logging "" ""
    Write-Colr -Text $cname,"$mesgbye" -Colour White,Gray
    Write-Host ''
    Stop-Script 5
}