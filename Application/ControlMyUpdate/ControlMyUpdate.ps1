##########################################################################################
# You running this script/function means you will not blame the author(s) if this breaks your stuff. 
# This script/function is provided AS IS without warranty of any kind. Author(s) disclaim all 
# implied warranties including, without limitation, any implied warranties of merchantability or of 
# fitness for a particular purpose. The entire risk arising out of the use or performance of the sample 
# scripts and documentation remains with you. In no event shall author(s) be held liable for any damages 
# whatsoever (including, without limitation, damages for loss of business profits, business interruption, 
# loss of business information, or other pecuniary loss) arising out of the use of or inability to use 
# the script or documentation. Neither this script/function, nor any part of it other than those parts 
# that are explicitly copied from others, may be republished without author(s) express written permission. 
# Author(s) retain the right to alter this disclaimer at any time.
##########################################################################################



##########################################################################################
# Name: ControlMyUpdate.ps1
# Version: 2.1
# Date: 18.05.2021
# Created by: Grischa Ernst gernst@vmware.com
# Contributor: Camille Debay
#
# Description
# - This Script will provide you a granular control over Window Updates - especially if you using Windows Update for Business
# - This includes
#     1. Download updates at any time, to reduce the installation time
#     2. Install updates during a weekly Maintenance Window
#     3. Download and install updates
# - Updates that require a reboot, will reboot during Maintenance Window - but will not automatically reboot if no MW is configured
# - Registry values in the "HKLM:\SOFTWARE\ControlMyUpdate" hive are used to configure the script
#
# Caution
# 1. Please use the GUI version 2.0 or higher with the script version 1.0 or higher.
#
###################################
# Registry values:
# Name            Value                          Description
# DirectDownload 
#                 True                           Will download update independent of the Maintenance Window
#                 False                          Will not start download before installation
#
# HiddenUpdates
#                 e.g. KB4023057,KB5003173       Hide specific updates to make sure the updates are not installed
#
#
# UnHiddenUpdates
#                 e.g. KB4023057,KB5003173       Un-hide specific updates to make sure the updates are available for installation again
#
#
# LastInstallationDate 
#                 e.g. 18/05/2021 03:24          Last time where update installation was tried
#
#
# Maintenance Window
#                 True                           Will install updates only during Maintenance Window
#                 False                          Will install updates whenever possible
# 
#
# MWDay
#                 e.g. 1 for Monday
#                 e.g. 2 for Tuesday             Will install updates only on a specific day of the week
#
#
# MWStartTime
#                 08:00                          Will start the installation only after this time - make sure you are using 24 hours format
#
#
# MWStopTime
#                 09:30                          Will not install updates after this time - make sure you are using 24 hours format
#
#
# UseMicrosoftUpdate
#                 True                           Will force to use Microsoft Update - will ignore WSUS settings
#                 False                          Will use configured settings
#
# ReportOnly
#                 True                           Will use the normal Windows Update method to install, but will create registry entries for tracking update installation
#                 False                          Script will run as normal
#
# NotifyUser
#                 True                           Will generate a toast notification
#                 False                          Will not notify the end-user 
#
# ToastTitle
#                 e.g. Windows Update            Provides a title for the toast notification
#
# ToastText
#                 e.g. Restart required          Provides a description in the notification
#
# EmergencyKB
#                 e.g. KB12345                   This update will be installed even if the device is outside of the maintenance window and if the device is outside of the scan interval
#
# Update Source   MU,WSUS or Default             Will set the update source to the selected option 
#
# Retry count     0-99                           Configure how often updates getting downloaded and installed if an error appears 
#
##########################################################################################
#                                    Changelog 
#
# 2.1 - added retry count for error handling + pending reboot detection bugfix
# 2.0 - removed the requirement of PSWindowsUpdate Module - now using native Windows Update API
# 1.2 - Bugfixes 
# 1.1 - Added delivery optimization statistics
# 1.0 - Added scan time randomization + bug fixes + reporting only mode + decoupled scan interval with the script running interval
# 0.4 - added reporting functionality 
# 0.3 - added multi day Maintenance Window support
# 0.2 - changed reboot behavior and fixed Maintenance Window detection + small bug fixes
# 0.1 - Initial creation
##########################################################################################

##########################################################################################
#                                    Param 
#
param(
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Path for logs")][String] $LogPath,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Output Version of the script")][Switch] $ScriptVersion,
    [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Verbosity of logging. Default: Info")][ValidateSet("Info", "Debug", "Trace")][String] $ScriptLogLevel = "Info"
)

$ScriptCurrentVersion = "2.1"

if ($ScriptVersion.IsPresent) {
    Return $ScriptCurrentVersion
    break
}
##########################################################################################
#                                    Functions

Function Write-Log {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Logging Level")][alias("Level")][ValidateSet("Info", "Error", "Debug", "Trace")][String] $LogLevel,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Log Message")][alias("Message")][String] $LogMessage
    )

    Process {
        if ($LogEnabled) {
            $TimeStamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss.fff"
            Switch ($LogLevel) {
                "Info" { Add-content -Path "$($logpath)" -value "INFO: $($LogMessage)  `$`$<$Component><$($TimeStamp)><thread=$($thread)>" -Encoding UTF8 }
                "Debug" { if (($ScriptLogLevel -eq "Debug") -or ($ScriptLogLevel -eq "Trace")) { Add-content -Path "$($logpath)" -value "DEBUG: $($LogMessage)  `$`$<$Component><$($TimeStamp)><thread=$($thread)>" -Encoding UTF8 } }
                "Trace" { if ($ScriptLogLevel -eq "Trace") { Add-content -Path "$($logpath)" -value "TRACE: $($LogMessage)  `$`$<$Component><$($TimeStamp)><thread=$($thread)>" -Encoding UTF8 } }
                "Error" { Add-content -Path "$($logpath)" -value "ERROR: $($LogMessage)  `$`$<$Component><$($TimeStamp)><thread=$($thread)>" -Encoding UTF8 }
            }
            Remove-Variable -Name LogMessage, LogLevel -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        }
    }
}


function Test-MaintenanceWindow {
    
    Write-Log -LogLevel Trace -LogMessage "Function: Test-MaintenanceWindow: Start"
    Write-Log -LogLevel Info -LogMessage "Check if device is in maintenance window"
	
    #Get current time
    $CurrentDate = Get-Date
    $CurrentDay = $CurrentDate.DayOfWeek.value__
    $CurrentHour = Get-Date -Format HH 
    $CurrentMinute = Get-Date -Format mm

    #Get start time
    $MWStartHour = $Settings.MWStartTime.Substring(0, 2)
    $MWStartMinute = $Settings.MWStartTime.Remove(0, 3)
    Write-Log -LogLevel Debug -LogMessage "Start Time: Hour: $($MWStartHour) / Minute: $($MWStartMinute)"
    
    #Get stop time
    $MWStopHour = $Settings.MWStopTime.Substring(0, 2)
    $MWStopMinute = $Settings.MWStopTime.Remove(0, 3)
    Write-Log -LogLevel Debug -LogMessage "Stop Time: Hour: $($MWStopHour) / Minute: $($MWStopMinute)"

    #Check if installation day was set - if not, updates will be installed everyday 
    if ($Settings.MWDay) {
        if ($Settings.MWDay -like "*,*") {
            Write-Log -LogLevel Debug -LogMessage "Multiple target day found. $($Settings.MWDay)"
            $TargetDay = $Settings.MWDay.Split(",")
        }
        else {
            Write-Log -LogLevel Debug -LogMessage "1 Target day found $($Settings.MWDay)"
            $TargetDay = $Settings.MWDay
        }
    }
    else {
        Write-Log -LogLevel Debug -LogMessage "Everyday as target day"
        $TargetDay = $CurrentDay
    }

    Write-Log -LogLevel Debug -LogMessage "TargetDay: $($TargetDay)"

    #Check if current time is in Maintenance Window
    Clear-Variable IsInMaintenanceWindow -Force -ErrorAction SilentlyContinue
    $CurrentDayIsInMW = $false

    [Boolean] $IsInMaintenanceWindow = $false
    foreach ($MWDay in $TargetDay) {
        Write-Log -LogLevel Debug -LogMessage "processing: $($MWDay)"  
        if ( $CurrentDay -eq $MWDay ) {
            Write-Log -LogLevel Debug -LogMessage "Current day defined in maintenance window"
            $CurrentDayIsInMW = $true
        }   
        
        if ($CurrentDayIsInMW -eq $true) {
            if (($CurrentDay -eq ($TargetDay[$TargetDay.Count - 1])) -and ($TargetDay.Count -gt 1)) {
                $CurrentDayIsLastDay = $True
            }
            else { $CurrentDayIsLastDay = $False }
            
            Write-Log -LogLevel Trace -LogMessage "If current day in Maintenance Window: True"
            if ($MWStopHour -lt $MWStartHour) {
                Write-Log -LogLevel Trace -LogMessage "If Checking Stop Hour is smaller than Start Hour. Over midnight MW."
                if ($CurrentHour -ge $MWStartHour -or $CurrentHour -le $MWStopHour) { 
                    Write-Log -LogLevel Debug -LogMessage "Checking Maintenance Windows timeframe"
                    if ($CurrentHour -ge $MWStartHour -and $CurrentMinute -ge $MWStartMinute) {                       
                        Write-Log -LogLevel Trace -LogMessage "Stop Hour at 24"
                        $MWStopHour = "24"
                        $MWStopMinute = "00"
                    }
                    if ($CurrentHour -ge $MWStartHour -or $CurrentHour -le $MWStopHour) { 
                        if ($CurrentHour -le $MWStopHour -and $CurrentMinute -le $MWStopMinute) {                            
                            Write-Log -LogLevel Trace -LogMessage "Stop Hour at 0"
                            $MWStopHour = "00"
                            $MWStopMinute = "00"
                        }
                    }
                }
                if ($CurrentDayIsLastDay -eq $true) {
                    $MWStartHour = "00"
                    $MWStartMinute = "00"
                }

            }

            if ($CurrentHour -ge $MWStartHour -and $CurrentHour -le $MWStopHour) {
                Write-Log -LogLevel Debug -LogMessage "Current Hour within maintenance window timeframe"
                if ((($CurrentHour -eq $MWStartHour -and $CurrentMinute -gt $MWStartMinute) -and ($CurrentHour -eq $MWStopHour -and $CurrentMinute -lt $MWStopMinute))) {
                    Write-Log -LogLevel Debug -LogMessage "Same Start and Stop hour, using minutes to check Maintenance Windows"
                    $IsInMaintenanceWindow = $true
                }
                elseif ($CurrentHour -ge $MWStartHour -and $CurrentHour -lt $MWStopHour) {
                    Write-Log -LogLevel Debug -LogMessage "Checking if time is between start and stop hour"
                    $IsInMaintenanceWindow = $true
                }
                elseif (($CurrentHour -ge $MWStartHour -and $CurrentHour -eq $MWStopHour) -and $CurrentMinute -lt $MWStopMinute) {
                    Write-Log -LogLevel Debug -LogMessage "CHecking Maintenance Windows within the last hour"
                    $IsInMaintenanceWindow = $true
                }
            }
            else {
                Write-Log -LogLevel Debug -LogMessage "Current time not within defined maintenance window"
                $IsInMaintenanceWindow = $false
            }
        }
        else {
            Write-Log -LogLevel Debug -LogMessage "Current time not within defined maintenance window"
            $IsInMaintenanceWindow = $false
        }
    }
    Write-Log -LogLevel Debug -LogMessage "IsInMaintenanceWindow: $($IsInMaintenanceWindow)"
    
    if ($IsInMaintenanceWindow -eq $true) { Write-Log -LogLevel Info -LogMessage "Device is in maintenance window" }
    if ($IsInMaintenanceWindow -eq $false) { Write-Log -LogLevel Info -logMessage "Device is outside of the maintenance window" }

    Write-Log -LogLevel Trace -LogMessage "Function: Test-MaintenanceWindow: End"
    return $IsInMaintenanceWindow
}

function Test-PendingReboot {
    param(
        [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = "Enable Automatic Reboot")][bool]$AutomaticReboot = $false
    )
    Write-Log -LogLevel Trace -LogMessage "Function: Test-PendingReboot: Start"
    Write-Log -LogLevel Debug -LogMessage "AutomaticReboot: $($AutomaticReboot)"

    $PendingRestart = $false
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { $PendingRestart = $true }
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { $PendingRestart = $true }

    Write-Log -LogLevel Debug -LogMessage "PendingRestart: $($PendingRestart)"

    if ($PendingRestart -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Pending Restart - Device require reboot"

        if ($AutomaticReboot -eq $true) {
            Write-Log -LogLevel Debug -LogMessage "Automatic reboot activated. Running shutdown command"
            shutdown.exe /r /f /t 120
        }
    }
    
    Write-Log -LogLevel Trace -LogMessage "Function: Test-PendingReboot: End"
    return $PendingRestart
}

function Write-UpdateStatus {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "KB Number to update")][String] $CurrentUpdate,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Force status change update")][String] $StatusChange,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Force status change update")][String] $UpdateTitle,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Status of the update to recorded")][String] $Status
    )
    Write-Log -LogLevel Trace -LogMessage "Function: Write-UpdateStatus: Start"
    Write-Log -LogLevel Debug -LogMessage "CurrentUpdate: $($CurrentUpdate)"
    Write-Log -LogLevel Debug -LogMessage "Status: $($Status)"

    $LastUpdate = Get-Date -Format s
    
    if (!(Test-Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)")) {
        New-Item "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -Force
    }

    if ($UpdateTitle) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "Title" -Value $UpdateTitle -Force | Out-Null
    }

    if (!(Get-Item "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)\$($Status)" -ErrorAction SilentlyContinue) -or $StatusChange -eq $True) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "$Status" -Value $LastUpdate -Force | Out-Null
    }

    Write-Log -LogLevel Trace -LogMessage "Function: Write-UpdateStatus: End"
}

function Get-InstalledWindowsUpdates {
    Write-Log -LogLevel Trace -LogMessage "Function: Get-InstalledWindowsUpdates: Start"
    Write-Log -LogLevel Info -LogMessage "Searching for installed updates"

    $UpdateCollection = [System.Collections.ArrayList]@()
    $RegExKB = "KB(\d+)"

    $MUSession = New-Object -ComObject "Microsoft.Update.Session"
    $UpdateSearcher = $MUSession.CreateUpdateSearcher()
    $TotalHistoryCount = $UpdateSearcher.GetTotalHistoryCount()

    Write-Log -LogLevel Debug -LogMessage "Found $($TotalHistoryCount) with Update Searcher"   
    If ($TotalHistoryCount -gt 0) {
        Write-Log -LogLevel Debug -LogMessage "Query History for each update"
        $QueryHistory = $UpdateSearcher.QueryHistory(0, $TotalHistoryCount)

        foreach ($item in $QueryHistory) {
            $RegExResult = [regex]::Match($item.Title, $RegExKB)
            if ( $RegExResult.Success ) {
                Write-Log -LogLevel Debug -LogMessage "Found KB Number: $($RegExResult.Value)"
                $UpdateCollection.Add($RegExResult.Value) | Out-Null
            }
        }
    }
   
    #Get Windows Updates from WMI
    $WMIKBs = Get-WmiObject win32_quickfixengineering |  Select-Object HotFixID -ExpandProperty HotFixID
    Write-log -LogLevel Debug -LogMessage "WMI KB List: $($WMIKBs)"
   
    #Get Windows Updates from DISM
    $DISMKBList = dism /online /get-packages | findstr KB
   
    $DISMKBNumbers = [regex]::Matches($DISMKBList, $RegExKB).Value

    Write-log -LogLevel Debug -LogMessage "DISM KB:$($DISMKBNumbers)"
   
    $InstalledKBs = ($UpdateCollection + $WMIKBs + $DISMKBNumbers) | Sort-Object -Unique

    Write-Log -LogLevel Info -LogMessage "Following updates are installed: $($InstalledKBs)"
    Write-Log -LogLevel Trace -LogMessage "Function: Get-InstalledWindowsUpdates: End"
    return $InstalledKBs
}

function Search-AllUpdates {
    param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Define Windows Update Source")][ValidateSet("Default", "MU", "WSUS")] $UpdateSource = "Default",
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Ignore Hide/UnHide Status to get a full list")][Switch] $IgnoreHideStatus
    )
    Write-Log -LogLevel Trace -LogMessage "Function: Search-AllUpdates: Start"
    Write-Log -LogLevel Debug -LogMessage "Update Source: $($UpdateSource)"

    #Create Update Session
    $updateSession = New-Object -ComObject 'Microsoft.Update.Session'
    $updateSession.ClientApplicationID = "ControlMyUpdate"
    $updateSearcher = $updateSession.CreateUpdateSearcher()

    Switch ($UpdateSource) {
        "MU" {
            Write-Log -LogLevel Debug -LogMessage "Create Service Manager"
            # Create a new Service Manager Object
            $ServiceManager = New-Object -ComObject 'Microsoft.Update.ServiceManager'
            $ServiceManager.ClientApplicationID = "ControlMyUpdate"
            $ServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
            
            Write-Log -LogLevel Debug -LogMessage "Configure Update Searcher to MU and service manager"
            #configure the Update Searcher
            $updateSearcher.ServerSelection = "3"
            $updateSearcher.ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d"
        }

        "WSUS" {
            Write-Log -LogLevel Debug -LogMessage "Configure Update Searcher to WSUS"
            $updateSearcher.ServerSelection = "1"
        }

        "Default" {
            Write-Log -LogLevel Debug -LogMessage "Configure Update Searcher to Default"
            $updateSearcher.ServerSelection = "0"
        }
    }


    Write-Log -LogLevel Debug -LogMessage "Search for update with update searcher"
    if ( $IgnoreHideStatus ) {
        $SearchFilter = "( IsHidden = 0 )"
        $updates = ($updateSearcher.Search($SearchFilter))
    }
    else {
        $updates = ($updateSearcher.Search($Null))
    }

    switch -exact ($updates.ResultCode) {
        0 { $Status = 'NotStarted' }
        1 { $Status = 'InProgress' }
        2 { $Status = 'Succeeded' }
        3 { $Status = 'SucceededWithErrors' }
        4 { $Status = 'Failed' }
        5 { $Status = 'Aborted' }
        default { $Status = "Unknown result code [$($_)]" }
    }
    Write-Log -LogLevel Trace -LogMessage "Update search result status: $($Status)"

    
    foreach ($item in $Updates.Updates) {
        if ($item.IsDownloaded -eq $false -and $item.IsInstalled -eq $False) {
            Write-UpdateStatus -CurrentUpdate "KB$($item.KBArticleIDs)" -Status "Available" -UpdateTitle ($item.title) | Out-Null
        }
    }


    Write-Log -LogLevel Trace -LogMessage "Function: Search-AllUpdates: End"
    return $($Updates.Updates)
}

function Save-WindowsUpdate {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")] $DownloadUpdateList
    )
    $Component = "Update Download"
    Write-Log -LogLevel Trace -LogMessage "Function: Save-WindowsUpdate: Start"
    Write-Log -LogLevel Debug -LogMessage "DownloadUpdateList: $($DownloadUpdateList.KBArticleIDs)"

    Write-Log -LogLevel Debug -LogMessage "Updates available. Checking Delivery Optimization service"
    if ((Get-Service DoSvc).Status -eq "Stopped") {
        Write-Log -LogLevel Debug "Delivery Optimization service stopped. Starting service."
        Start-Service DoSvc
    }
    if ((Get-Service DoSvc).Status -eq "Running") {
        Write-Log -LogLevel Debug -LogMessage "Delivery Optimization service Running. Restarting to re-initialize DO"
        Restart-Service DoSvc -Force
    }

    Write-Log -LogLevel Trace -LogMessage "Foreach DownloadUpdateList : start"
    foreach ($Update in $DownloadUpdateList) {
        Write-Log -LogLevel Debug -LogMessage "Processing update: $($Update.Title)"
        $updateSession = New-Object -ComObject 'Microsoft.Update.Session'
        $updatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
        
        #Add update to download collection
        $updatesToDownload.Add($Update) | Out-Null
        
        $a = 0
        do
        {  
              $a++

            #Initialize download
            Write-Log -LogLevel Debug -LogMessage "Initializing download"
            $downloader = $updateSession.CreateUpdateDownloader()
            $downloader.Updates = $updatesToDownload
        
            Write-Log -LogLevel Info -LogMessage "Downloading $($Update.Title)"
            Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Download started" -StatusChange $True

            #Start download        
             
          
            
            Write-Log -LogLevel Info -LogMessage "Starting ($($a).) download"            
            $downloadResult = $downloader.Download()

            #Writing logs
            if ($downloadResult.ResultCode -eq 2) {
                Write-Log -LogLevel Info -LogMessage "Download successful for KB$($Update.KBArticleIDs)"
                Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Download completed" -StatusChange $True
            }
            else {
                Write-Log -LogLevel Error -LogMessage "Download error for KB$($Update.KBArticleIDs)"
                Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Download failed" -StatusChange $True
            }
        
        } Until ($a -ge $retrycount -or $downloadResult.ResultCode -eq 2)  
        
        
        Clear-Variable updatesToDownload -Force -ErrorAction SilentlyContinue
    }
    Write-Log -LogLevel Trace -LogMessage "Foreach available update: End"
    Write-Log -LogLevel Trace -LogMessage "Function: Save-WindowsUpdate: End"
}

function Update-InstallationStatus {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")]$AvailableUpdates
    )
    
    Write-Log -LogLevel Trace -LogMessage "Function: Get-NonInstalledUpdates: Start"
    Write-Log -LogLevel Debug -LogMessage "AvailableUpdates: $($AvailableUpdates)"

    Write-Log -LogLevel Trace -LogMessage "Get-InstalledWindowsUpdates"
    $InstalledKBs = Get-InstalledWindowsUpdates
    Write-Log -LogLevel Debug -LogMessage "Installed KBs returned: $($InstalledKBs)"

    $AvailableKBs = $AvailableUpdates.KBArticleIDs

    $RegistryKBList = Get-ChildItem "$($RegistryRootPath)\Status\KBs"
    Write-Log -LogLevel Debug -LogMessage "Item from registry: $($RegistryKBList)"

    foreach ($RegKey in $RegistryKBList) {
        $StatusKeys = $RegKey | Where-Object { $_.property -like "Installation Status*" -and $_.property -ne "Installation Status : Installed" }
       
        if ($statuskeys) {           
            $NonInstalledKB = $StatusKeys.name.Split('\')[-1]
            $KBCheck = $InstalledKBs | Where-Object { $_ -eq $NonInstalledKB }
            $AvailableKBCheck = $AvailableKBs | Where-Object { $_ -eq $($NonInstalledKB.replace("KB", "")) }
            

            $KeyNames = $StatusKeys.property | Where-Object { $_ -like "Installation Status*" -and $_ -ne "Installation Status : Installed" }
            
            foreach ($Item in $KeyNames) {
                Remove-ItemProperty "$($RegistryRootPath)\Status\KBs\$($NonInstalledKB)" -Name $Item -Force
            }

            if ($KBCheck) {
                Write-UpdateStatus -CurrentUpdate "$($NonInstalledKB)" -Status "Installation Status : Installed" -StatusChange $True
                Write-Log -LogLevel Info -LogMessage "Status Change for $($NonInstalledKB) - now Installed"

            }
            elseif ($AvailableKBCheck) {
                Write-UpdateStatus -CurrentUpdate "$($NonInstalledKB)" -Status "Installation Status : Pending Installation" -StatusChange $True
                Write-Log -LogLevel Info -LogMessage "Status Change for $($NonInstalledKB) - now pending for installation"

            }
            elseif (!$KBCheck -and !$AvailableKBCheck) {
                Write-UpdateStatus -CurrentUpdate "$($NonInstalledKB)" -Status "Installation Status : Not Applicable" -StatusChange $True
                Write-Log -LogLevel Info -LogMessage "Status Change for $($NonInstalledKB) - not applicable anymore"
            }
            Clear-Variable KBCheck, AvailableKBCheck -Force
        }
    }

    Write-Log -LogLevel Trace -LogMessage "Function: Get-NonInstalledUpdates: End"
    return $NoninstalledUpdates
}

function Set-WindowsUpdateBlockStatus {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "KB ID Number to hide")][String] $KBArticleID,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Status of the update")][ValidateSet("Blocked", "UnBlocked")][String] $Status,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")] $AllUpdates
    )

    Write-Log -LogLevel Trace -LogMessage "Function: Set-WindowsUpdateBlockStatus: Start"
    Write-Log -LogLevel Debug -LogMessage "ArticleID: $($KBArticleID)"
    Write-Log -LogLevel Debug -LogMessage "Status: $($Status)"
    Write-Log -LogLevel Debug -LogMessage "AllUpdates: $($AllUpdates)"
    
    if ($KBArticleID -like "KB*") {
        $KBArticleID = $KBArticleID.Remove(0, 2)
    }

    if ($Status -eq "Blocked") {
        Write-Log -LogLevel Trace -LogMessage "If Blocked"
        Write-Log -LogLevel Trace -LogMessage "Foreach AllUpdates : start"
        foreach ($KB in $AllUpdates) {
            Write-Log -LogLevel Debug -LogMessage "KB: $($KB)"
            if ($KBArticleID -eq $kb.KBArticleIDs) {
                Write-Log -LogLevel Info -LogMessage "KB$($KBArticleID) is now BLOCKED"
                $KB.IsHidden = $true
            }
        }
    }

    if ($Status -eq "UnBlocked") {
        Write-Log -LogLevel Trace -LogMessage "If UnBlocked"
        Write-Log -LogLevel Trace -LogMessage "Foreach AllUpdates : start"
        foreach ($KB in $AllUpdates) {
            Write-Log -LogLevel Debug -LogMessage "KB: $($KB)"
            if ($KBArticleID -eq $kb.KBArticleIDs) {
                Write-Log -LogLevel Info  -LogMessage "KB$($KBArticleID) is now UNBLOCKED"
                $KB.IsHidden = $False
            }
        }
    }
    Write-Log -LogLevel Trace -LogMessage "Function: Set-WindowsUpdateBlockStatus: End"
}

function Update-DeliveryOptimizationStats {
    Write-Log -LogLevel Trace -LogMessage "Function: Update-DeliveryOptimizationStats: Start"
    Write-Log -LogLevel Info -LogMessage "Writing delivery optimization statistics to registry"

    #Overall
    Write-Log -LogLevel Trace -LogMessage "Running Get-DeliveryOptimizationPerfSnap"
    $OverallStats = Get-DeliveryOptimizationPerfSnap

    $TotalDownloaded = [math]::Round($OverallStats.TotalBytesDownloaded / 1MB)
    Write-Log -LogLevel Debug -LogMessage "TotalDownloaded: $($TotalDownloaded)"
    $TotalUploaded = [math]::Round($OverallStats.TotalBytesUploaded / 1MB)    
    Write-Log -LogLevel Debug -LogMessage "TotalUploaded: $($TotalUploaded)"

    #All time
    Write-Log -LogLevel Trace -LogMessage "Running Get-DeliveryOptimizationStatus"
    $Stats = Get-DeliveryOptimizationStatus

    $DownloadedFromPeers = [math]::Round(($Stats.BytesFromPeers | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "DownloadedFromPeers: $($DownloadedFromPeers)"
    $DownloadedFromHTTP = [math]::Round(($Stats.BytesFromHttp | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "DownloadedFromHTTP: $($DownloadedFromHTTP)"
    $DownloadedFromCacheServer = [math]::Round(($Stats.BytesFromCacheServer | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "DownloadedFromCacheServer: $($DownloadedFromCacheServer)"
    $DownloadedGroupPeers = [math]::Round(($Stats.BytesFromGroupPeers | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "DownloadedGroupPeers: $($DownloadedGroupPeers)"
    $DownloadedFromInternet = [math]::Round(($Stats.BytesFromInternetPeers | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "DownloadedFromInternet: $($DownloadedFromInternet)"

    if ($TotalDownloaded -ne "0") {
        $DownloadedFromPeersPercentage = ($DownloadedFromPeers / $TotalDownloaded).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromPeersPercentage: $($DownloadedFromPeersPercentage)"
        $DownloadedFromHTTPPercentage = ($DownloadedFromHTTP / $TotalDownloaded).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromHTTPPercentage: $($DownloadedFromHTTPPercentage)"
        $DownloadedFromCacheServerPercentage = ($DownloadedFromCacheServer / $TotalDownloaded).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromCacheServerPercentage: $($DownloadedFromCacheServerPercentage)"
        $DownloadedFromGroupPeersPercentage = ($DownloadedGroupPeers / $TotalDownloaded).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromGroupPeersPercentage: $($DownloadedFromGroupPeersPercentage)"
        $DownloadedFromInternetPercentage = ($DownloadedFromInternet / $TotalDownloaded).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromInternetPercentage: $($DownloadedFromInternetPercentage)"
    }
    else {
        $DownloadedFromPeersPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromPeersPercentage: $($DownloadedFromPeersPercentage)"
        $DownloadedFromHTTPPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromHTTPPercentage: $($DownloadedFromHTTPPercentage)"
        $DownloadedFromCacheServerPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromCacheServerPercentage: $($DownloadedFromCacheServerPercentage)"
        $DownloadedFromGroupPeersPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromGroupPeersPercentage: $($DownloadedFromGroupPeersPercentage)"
        $DownloadedFromInternetPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "DownloadedFromInternetPercentage: $($DownloadedFromInternetPercentage)"
    }


    #This Month
    $StatsCurrentMonth = Get-DeliveryOptimizationPerfSnapThisMonth

    $MonthDownloadedFromPeers = [math]::Round(($StatsCurrentMonth.DownloadLanBytes | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "MonthDownloadedFromPeers: $($MonthDownloadedFromPeers)"
    $MonthDownloadedFromHTTP = [math]::Round(($StatsCurrentMonth.DownloadHttpBytes | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "MonthDownloadedFromHTTP: $($MonthDownloadedFromHTTP)"
    $MonthDownloadedFromCacheHost = [math]::Round(($StatsCurrentMonth.DownloadCacheHostBytes | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "MonthDownloadedFromCacheHost: $($MonthDownloadedFromCacheHost)"
    $MonthDownloadedFromInternet = [math]::Round(($StatsCurrentMonth.DownloadInternetBytes | Measure-Object -sum).sum / 1MB)
    Write-Log -LogLevel Debug -LogMessage "MonthDownloadedFromInternet: $($MonthDownloadedFromInternet)"

    $MonthTotalBytes = [math]::Round(($StatsCurrentMonth.DownloadHttpBytes + $StatsCurrentMonth.DownloadLanBytes + $StatsCurrentMonth.DownloadCacheHostBytes + $StatsCurrentMonth.DownloadInternetBytes) / 1MB)
    Write-Log -LogLevel Debug -LogMessage "MonthTotalBytes: $($MonthTotalBytes)"

    if ($MonthTotalBytes -ne "0") {
        $MonthlyPeerDownloadPercentage = ($MonthDownloadedFromPeers / $MonthTotalBytes).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyPeerDownloadPercentage: $($MonthlyPeerDownloadPercentage)"
        $MonthlyHTTPDownloadPercentage = ($MonthDownloadedFromHTTP / $MonthTotalBytes).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyHTTPDownloadPercentage: $($MonthlyHTTPDownloadPercentage)"
        $MonthlyCacheHostDownloadPercentage = ($MonthDownloadedFromCacheHost / $MonthTotalBytes).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyCacheHostDownloadPercentage: $($MonthlyCacheHostDownloadPercentage)"
        $MonthlyInternetDownloadPercentage = ($MonthDownloadedFromInternet / $MonthTotalBytes).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyInternetDownloadPercentage: $($MonthlyInternetDownloadPercentage)"
    }
    else {
        $MonthlyPeerDownloadPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyPeerDownloadPercentage: $($MonthlyPeerDownloadPercentage)"
        $MonthlyHTTPDownloadPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyHTTPDownloadPercentage: $($MonthlyHTTPDownloadPercentage)"
        $MonthlyCacheHostDownloadPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyCacheHostDownloadPercentage: $($MonthlyCacheHostDownloadPercentage)"
        $MonthlyInternetDownloadPercentage = (0).ToString("P")
        Write-Log -LogLevel Debug -LogMessage "MonthlyInternetDownloadPercentage: $($MonthlyInternetDownloadPercentage)"
    }

    #Check if Registry exists 
    if ((Test-Path "$($RegistryRootPath)\Status\DO") -eq $false) {
        New-Item "$($RegistryRootPath)\Status\DO" -Force | Out-Null
    }


    #Overall
    Write-Log -LogLevel Debug -LogMessage "Writing Total Download/Upload MB Statistics"
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Downloaded MB" -Value "$($TotalDownloaded)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Uploaded MB" -Value "$($TotalUploaded)" -Force | Out-Null

    #Total Percentage 
    Write-Log -LogLevel Debug -LogMessage "Writing Percentage Statistics"
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Peer Download Percentage" -Value "$($DownloadedFromPeersPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total HTTP Download Percentage" -Value "$($DownloadedFromHTTPPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Cache Host Download Percentage" -Value "$($DownloadedFromCacheServerPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Group Peers Download Percentage" -Value "$($DownloadedFromGroupPeersPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Internet Download Percentage" -Value "$($DownloadedFromInternetPercentage)" -Force | Out-Null


    #Total MB
    Write-Log -LogLevel Debug -LogMessage "Writing Total MB Statistics"
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Peer Download MB" -Value "$($DownloadedFromPeers)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total HTTP Download MB" -Value "$($DownloadedFromHTTP)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Cache Host Download MB" -Value "$($DownloadedFromCacheServer)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Group Peers Download MB" -Value "$($DownloadedGroupPeers)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Total Internet Download MB" -Value "$($DownloadedFromInternet)" -Force | Out-Null

    #Monthly MB
    Write-Log -LogLevel Debug -LogMessage "Writing Monthly MB Statistics"
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Peer Download MB" -Value "$($MonthDownloadedFromPeers)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly HTTP Download MB" -Value "$($MonthDownloadedFromHTTP)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Cache Host Download MB" -Value "$($MonthDownloadedFromCacheHost)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Internet Download MB" -Value "$($MonthDownloadedFromInternet)" -Force | Out-Null

    #Monthly Percentage
    Write-Log -LogLevel Debug -LogMessage "Writing Monthly Percentage Statistics"
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Peer Download Percentage" -Value "$($MonthlyPeerDownloadPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly HTTP Download Percentage" -Value "$($MonthlyHTTPDownloadPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Cache Host Download Percentage" -Value "$($MonthlyCacheHostDownloadPercentage)" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status\DO" -PropertyType "String" -Name "Monthly Internet Download Percentage" -Value "$($MonthlyInternetDownloadPercentage)" -Force | Out-Null

    Write-Log -LogLevel Trace -LogMessage "Function: Update-DeliveryOptimizationStats: End"
}

function Install-SpecificWindowsUpdate {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Updates To Install")] $SelectedUpdate
    )

    $Component = "UPDATE INSTALLATION"

    Write-Log -LogLevel Trace -LogMessage "Function: Install-SpecificWindowsUpdate: Start"
    Write-Log -LogLevel Trace -LogMessage "Foreach SelectedUpdate : start"


    foreach ($InstallUpdate in $SelectedUpdate) {
        Write-Log -LogLevel Info -LogMessage "Starting installation of $($InstallUpdate.Title)"
        Write-UpdateStatus -CurrentUpdate "KB$($InstallUpdate.KBArticleIDs)" -Status "Start installation" -StatusChange $True

        $updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
        $updatesToInstall.Add($InstallUpdate) | Out-Null 

        $installer = New-Object -ComObject 'Microsoft.Update.Installer'
        $installer.Updates = $updatesToInstall        
        $installResult = $installer.Install()
        
        Write-Log -LogLevel Debug -LogMessage "Install result code: $($installResult.ResultCode)"
        switch -exact ($installResult.ResultCode) {
            0 { $InstallStatus = 'NotStarted' }
            1 { $InstallStatus = 'InProgress' }
            2 { $InstallStatus = 'Installed' }
            3 { $InstallStatus = 'InstalledWithErrors' }
            4 { $InstallStatus = 'Failed' }
            5 { $InstallStatus = 'Aborted' }
            6 { $InstallStatus = 'NoUpdatesNeeded' }
            7 { $InstallStatus = 'RebootRequired' }
            default { $InstallStatus = "Unknown result code [$($_)]" }
        }

        if ($($installResult.ResultCode) -ne 2) {
            Write-Log -LogLevel Error -LogMessage "Update Status: $($InstallStatus)"
            Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Installation Status : $($InstallStatus)" -StatusChange $True
        }
        else {
            Write-Log -LogLevel Info -LogMessage "Update Status: $($InstallStatus)"
            
            if (($installResult.RebootRequired) -eq $True) {
                Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Installation Status : Reboot required" -StatusChange $True
                Write-Log -LogLevel Info -LogMessage "KB$($Update.KBArticleIDs) requires reboot"

            }
            else { Write-UpdateStatus -CurrentUpdate "KB$($Update.KBArticleIDs)" -Status "Installation Status : $($InstallStatus)" -StatusChange $True }
        
        }
    }
    Write-Log -LogLevel Trace -LogMessage "Function: Install-SpecificWindowsUpdate: End"
}

function New-WindowsUpdateScan {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Last Scan Time")][DateTime] $LastScanTime
    )
    Write-Log -LogLevel Trace -LogMessage "Function: New-WindowsUpdateScan: Start"
    if ($Settings.ScanRandomization -ge "1") {
        $NextScanTimeTemp = $LastScanTime.AddHours($Settings.ScanInterval).AddMinutes((Get-Random -Minimum 0 -Maximum $Settings.ScanRandomization))
    }
    else {
        $NextScanTimeTemp = $LastScanTime.AddHours($Settings.ScanInterval)
    }
    Write-Log -LogLevel Debug -LogMessage "NextScanTimeTemp: $($NextScanTimeTemp)"

    $NextScanTime = (Get-Date $NextScanTimeTemp -Format s)
    
    #Get all updates with the configured update source
    Write-Log -LogLevel Debug -LogMessage "Scan for all updates with the configured update source"
    $AllUpdatesFound = Search-AllUpdates -UpdateSource $($Settings.UpdateSource)

    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastScanTime' -Value $(Get-Date -Format s) | Out-Null
    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'NextScanTime' -Value $NextScanTime | Out-Null

    Write-Log -LogLevel Trace -LogMessage "Function: New-WindowsUpdateScan: End"
    return $AllUpdatesFound
}

function Show-Notification {
    [cmdletbinding()]
    Param (
        [string]
        $ToastTitle,
        [string]
        [parameter(ValueFromPipeline)]
        $ToastText
    )

    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

    $RawXml = [xml] $Template.GetXml()
    ($RawXml.toast.visual.binding.text | Where-Object { $_.id -eq "1" }).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
    ($RawXml.toast.visual.binding.text | Where-Object { $_.id -eq "2" }).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

    $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $SerializedXml.LoadXml($RawXml.OuterXml)

    $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
    $Toast.Tag = "Windows Update"
    $Toast.Group = "Windows Update"

    $toast.Priority = "High"

    $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Windows Update")
    $Notifier.Show($Toast);
}

##########################################################################################
#                                   Define variables

$LogFileName = "WindowsUpdate_{0:yyyyMM}.log" -f (Get-Date)
$RegistryRootPath = "HKLM:\SOFTWARE\ControlMyUpdate"

#define the log folder if needed - otherwise delete it or set it to ""
if ($LogPath) { 
    if ((Test-Path $LogPath) -eq $false) { New-Item $LogPath -ItemType Directory -Force }
    $LogPath = "$($LogPath)\$($LogFileName)"
    $LogEnabled = $true
}

###############################################
# Script Header
###############################################
$thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
$Component = "SCRIPT START"
Write-Log -LogLevel Info -LogMessage "----------------------------------------------------------"
Write-Log -LogLevel Info -LogMessage "|           ControlMyUpdate - v$($ScriptCurrentVersion)   "
Write-Log -LogLevel Info -LogMessage "----------------------------------------------------------"

if ( $PSBoundParameters.ContainsKey("ScriptLogLevel") ) {
    Write-Log -LogLevel Debug -LogMessage "ScriptLogLevel specified, using value from command line."
}
elseif ( (Get-ItemProperty $RegistryRootPath -ErrorAction SilentlyContinue).PSObject.Properties.Name -contains "ScriptLogLevel" ) {
    $ScriptLogLevel = Get-ItemPropertyValue -Path $RegistryRootPath -Name "ScriptLogLevel"
    Write-Log -LogLevel Debug -LogMessage "ScriptLogLevel specified, using value from registry."
}

Write-Log -LogLevel Info -LogMessage "Logging Level : $($ScriptLogLevel)"

$Component = "SCRIPT VALIDATION"
Try {
    Write-Log -LogLevel Trace -LogMessage "Set ExecutionPolicy for current process"
    Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
}
Catch {
    Write-Log -LogLevel Trace -LogMessage "Set-Execution for process failed"
}

Write-Log -LogLevel Trace -LogMessage "Registry test path root key"
$RegistryTest = Test-Path $RegistryRootPath

if ($RegistryTest) {
    Write-Log -LogLevel Trace -LogMessage "Registry try write in key"
    try {
        Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastScriptStartTime' -Value $(Get-Date -Format s) | Out-Null
    }
    catch {
        Write-Log -LogLevel Error -LogMessage "Error writing in registry. Script will now exit."
        Write-Log -LogLevel Error -LogMessage "Error detail : $($Error[0])"
        break
    }
    $SettingsRegistryPath = "$($RegistryRootPath)\Settings"
    if ( Test-Path -Path $SettingsRegistryPath ) {
        Write-Log -LogLevel Info -LogMessage "Registry settings successfully detected"
        $Settings = Get-ItemProperty -Path "$($RegistryRootPath)\Settings" -ErrorAction SilentlyContinue

        if(!$settings.retrycount)
        {
            $retrycount = '3'
        }
        else{ $retrycount = $settings.retrycount }
    }
    else {
        Write-Log -LogLevel Error -LogMessage "No registry settings detected"
        break
    }    
}

#Create Status registry keys
if (!(Test-Path "$($RegistryRootPath)\Status")) {
    Write-Log -LogLevel Trace -LogMessage "Add Status Reg Key"
    New-Item "$($RegistryRootPath)\Status" -Force | Out-Null
}
if (!(Test-Path "$($RegistryRootPath)\Status\KBs")) {
    Write-Log -LogLevel Trace -LogMessage "Add Status\KBs Reg Key"
    New-Item "$($RegistryRootPath)\Status\KBs" -Force | Out-Null
}

if (!$settings.LastScanTime) {
    Write-Log -LogLevel Debug -LogMessage "No LastScanTime defined. Set LastScanTime to now"
    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastScanTime' -Value $(Get-Date -Format s)
}


$Component = "Emergency update installation"

#install emergency updates
if ($settings.EmergencyKB) {

    $InstalledKBs = Get-InstalledWindowsUpdates

    if ($settings.EmergencyKB -like "KB*") { $EmergencyKB = $settings.EmergencyKB.Remove(0, 2) } 
    else { $EmergencyKB = $settings.EmergencyKB }

    if ($InstalledKBs -eq "KB$($EmergencyKB)") {
        Write-Log -LogLevel Info -LogMessage "Emergency Update KB$($EmergencyKB) already installed"
        Set-ItemProperty "$($RegistryRootPath)\Settings" -Name "EmergencyKB" -Value ""
    }
    else {       
        Write-Log -LogLevel Info -LogMessage "Emergency Update KB$($EmergencyKB) NOT installed"

        $FullUpdateList = Search-AllUpdates -UpdateSource $($Settings.UpdateSource) -IgnoreHideStatus

        $EmergencyUpdate = $FullUpdateList | Where-Object { $_.KBArticleIDs -eq $EmergencyKB }
    
        if ($EmergencyUpdate) {
        
            Write-Log -LogLevel Info -LogMessage "Emergency Update KB$($EmergencyKB) getting installed"

            if ($EmergencyUpdate.IsDownloaded -eq $false) {
                Save-WindowsUpdate -DownloadUpdateList $EmergencyUpdate
            }
            Install-SpecificWindowsUpdate -SelectedUpdate $EmergencyUpdate 
        }
    }
    
}



$Component = "UPDATE HIDE/UNHIDE"

#Block unwanted updates

if (($settings.HiddenUpdates) -or ($settings.UnHiddenUpdates)) {
    $FullUpdateList = Search-AllUpdates -UpdateSource $($Settings.UpdateSource) -IgnoreHideStatus

    If ($settings.HiddenUpdates) {
        Write-Log -LogLevel Trace -LogMessage "Foreach settings.HiddenUpdates : start"
        foreach ($Item in $($settings.HiddenUpdates)) {
            Write-Log -LogLevel Debug -LogMessage "Hide | Processing Update: $($Item)"
            Set-WindowsUpdateBlockStatus -AllUpdates $FullUpdateList -KBArticleID $Item -Status Blocked
        }
    }

    #Allow hidden update again
    If ($settings.UnHiddenUpdates) {
        Write-Log -LogLevel Trace -LogMessage "Foreach settings.UnHiddenUpdates : start"
        foreach ($Item in $($settings.UnHiddenUpdates)) {
            Write-Log -LogLevel Debug -LogMessage "unHide | Processing Update: $($Item)"
            Set-WindowsUpdateBlockStatus -AllUpdates $FullUpdateList -KBArticleID $Item -Status UnBlocked
        }
    }
}


$Component = "UPDATE SCAN"
#Get last scan time 
if (!$Settings.NextScanTime) {
    Write-Log -LogLevel Info -LogMessage "Script First Run - Searching for updates"
    $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date)
}
else {
    $CurrentTime = Get-Date -Format s
    Write-Log -LogLevel Debug -LogMessage "CurrentTime: $($CurrentTime)"
    $NextScanTime = (Get-Date $($Settings.NextScanTime) -Format s)
    Write-Log -LogLevel Debug -LogMessage "NextScanTime: $($NextScanTime)"

    if ($CurrentTime -ge $NextScanTime -or (($Settings.MaintenanceWindow -eq $true) -and [bool](Test-MaintenanceWindow) -eq $true)) {
        Write-Log -LogLevel Info -LogMessage "Scan Interval reached - Searching for updates"
        $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date $Settings.LastScanTime)
    }
    else {
        Write-Log -LogLevel Info -LogMessage "Scan Interval not reached"
    }
}

#Check detected but not installed updates
if ($AllUpdates) {
    Write-Log -LogLevel Info -LogMessage "Pending Update count: $($AllUpdates.Count)"
    Update-InstallationStatus -AvailableUpdates $AllUpdates
}

# Check if Maintenance Window is enabled
if ($Settings.MaintenanceWindow -eq $true) {
    Write-Log -LogLevel Info -LogMessage "Maintenance Window Setting detected"
    if ((Test-MaintenanceWindow) -eq $true) {
        $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date)
    }
}
    

# run download and installation if updates are available 
if ( ($AllUpdates.Count -gt 0) -and ($Settings.ReportOnly -ne "True") ) {
    $NotDownloadedUpdates = $AllUpdates | Where-Object IsDownloaded -eq $false
    # If Download is independent of installation - Download all Updates
    If ($Settings.DirectDownload -eq $True) {
        Write-Log -LogLevel Debug -LogMessage "Direct Download Enabled"
        Write-Log -LogLevel Info -LogMessage "Start Download process"
        #get all updates that are not downloaded
        
        Write-Log -LogLevel Debug -LogMessage "NotDownloadedUpdates: $($NotDownloadedUpdates.KBArticleIDs)"

        Write-Log -LogLevel Debug -LogMessage "Trigger update download"
        Save-WindowsUpdate -DownloadUpdateList $NotDownloadedUpdates

        Write-Log -LogLevel Info -LogMessage "Download process finished"
    }

    # Install Updates
    # Check if Maintenance Window is enabled
    if ($Settings.MaintenanceWindow -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Maintenance Window Setting detected"

        if ((Test-MaintenanceWindow) -eq $true) {
            Write-Log -LogLevel Debug -LogMessage "Device in maintenance window"
            Test-PendingReboot -AutomaticReboot $true | Out-Null

            Write-Log -LogLevel Debug -LogMessage "Trigger update download"
            Save-WindowsUpdate -DownloadUpdateList $NotDownloadedUpdates

            foreach ($Update in $AllUpdates) {
                if ((Test-MaintenanceWindow) -eq $true) {
                    Write-Log -LogLevel Debug -LogMessage "Device in MW. Installing Update: $($Update.Title)"
                    Install-SpecificWindowsUpdate -SelectedUpdate $Update    
                    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastInstallationDate' -Value $(Get-Date -Format s)                   
                }
                else {
                    Write-Log -LogLevel Debug -LogMessage "Device outside maintenance window. Stopping update installation"
                    break
                }
            }

            if ((Test-MaintenanceWindow) -eq $true) {
                Write-Log -LogLevel Debug -LogMessage "Device in MW. Reboot processing."
                Test-PendingReboot -AutomaticReboot $true | Out-Null
            }
            else {
                Write-Log -LogLevel Debug -LogMessage "Device outside maintenance window. Skipping reboot"
            }
        }
        else {
            Write-Log -LogLevel Debug -LogMessage "Device outside maintenance window"
        }
    }
    else { # If no maintenance window is configured, just install the updates at any time
        Write-Log -LogLevel Info -LogMessage "Starting installation without Maintenance Window"

        Write-Log -LogLevel Debug -LogMessage "Trigger update download"
        Save-WindowsUpdate -DownloadUpdateList $NotDownloadedUpdates

        Clear-Variable Update -Force -ErrorAction SilentlyContinue

        foreach ($Update in $AllUpdates) {
            Write-Log -LogLevel Debug -LogMessage "Installing Update: $($Update.Title)"
            Install-SpecificWindowsUpdate -SelectedUpdate $Update
            Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastInstallationDate' -Value $(Get-Date -Format s)                   
        }
    }
}

$Component = "REBOOT CHECK"
#Check if  Reboot pending   
$RebootRequired = Test-PendingReboot -AutomaticReboot $false
Write-Log -LogLevel Debug -LogMessage "RebootRequired: $($RebootRequired)"

if ($RebootRequired -eq $true) {
    Write-Log -LogLevel Trace -LogMessage "RebootRequired: True"
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "PendingReboot" -Value "True" -Force | Out-Null

    if ( $Settings.NotifyUser ) {
        Write-Log -LogLevel Info -LogMessage "Notifying User of pending reboot"
        $RebootNotification = Get-ItemPropertyValue "$($RegistryRootPath)\Status" -Name 'RebootNotificationCreated' -ErrorAction Ignore 

        if ($RebootNotification -eq $False -or !($RebootNotification)) {
            Show-Notification -ToastTitle $Settings.ToastTitle -ToastText $Settings.ToastText 
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootNotificationCreated" -Value "True" -Force | Out-Null    
        }
    }
}
else {
    Write-Log -LogLevel Trace -LogMessage "RebootRequired: False"
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "PendingReboot" -Value "False" -Force | Out-Null
    if ( $Settings.NotifyUser ) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootNotificationCreated" -Value "False" -Force | Out-Null
    }
}

$Component = "UPDATE STATISTICS"
Write-Log -LogLevel Info -LogMessage "Update statistics"


$LastInstallationDate = Get-Date (Get-ItemPropertyValue -Path "$($RegistryRootPath)\Settings" -Name LastInstallationDate ) -Format s -ErrorAction SilentlyContinue
$NextInstallationDate = Get-Date (Get-ItemPropertyValue -Path "$($RegistryRootPath)\Settings" -Name NextScanTime ) -Format s -ErrorAction SilentlyContinue


if ((Get-Date $LastInstallationDate) -lt (Get-Date $NextInstallationDate)) {
    $AllUpdates = Search-AllUpdates -UpdateSource ($Settings.UpdateSource)
}


#Resetting counter
$UpdateCount = 0
$UpdateRollupsCount = 0
$DefinitionUpdatesCount = 0
$UpgradesCount = 0
$SecurityUpdatesCount = 0
$CriticalUpdatesCount = 0 

if ($AllUpdates.Count -gt 0) {
    foreach ($update in $AllUpdates) {
        $UpdateCategory = $Update.Categories._NewEnum.name | Select-Object -First 1

        switch ($UpdateCategory) {
            "Updates" { $UpdateCount = $UpdateCount + 1 }
            "Update Rollups" { $UpdateRollupsCount = $UpdateRollupsCount + 1 }
            "Definition Updates" { $DefinitionUpdatesCount = $DefinitionUpdatesCount + 1 }
            "Upgrades" { $UpgradesCount = $UpgradesCount + 1 }
            "Security Updates" { $SecurityUpdatesCount = $SecurityUpdatesCount + 1 }
            "Critical Updates" { $CriticalUpdatesCount = $CriticalUpdatesCount + 1 }
        }

        Write-UpdateStatus -CurrentUpdate "KB$($update.KBArticleIDs)" -Status "Available"
    }
      
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Missing Updates" -Value $($AllUpdates.count) -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Open Pending Updates" -Value "True" -Force | Out-Null
}
else {
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Missing Updates" -Value 0 -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Open Pending Updates" -Value "False" -Force | Out-Null
}

#Reporting the current installation count to the registry
Write-Log -LogLevel Debug -LogMessage "Write Updates category count in registry"
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Updates" -Value $UpdateCount -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Update Rollups" -Value $UpdateRollupsCount -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Definition Updates" -Value $DefinitionUpdatesCount -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Feature Upgrades" -Value $UpgradesCount -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Security Updates" -Value $SecurityUpdatesCount -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Pending Critical Updates" -Value $CriticalUpdatesCount -Force | Out-Null

 

#Report all installed updates 
$InstalledUpdates = Get-InstalledWindowsUpdates
Write-Log -LogLevel Debug -LogMessage "Write installed update list in registry"
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Installed KBs" -Value $InstalledUpdates -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Installed KBs" -Value $($InstalledUpdates.Count) -Force | Out-Null

#Update Delivery Optimization statistics 
Update-DeliveryOptimizationStats

Write-Log -LogLevel Info -LogMessage "###############################################################################" 