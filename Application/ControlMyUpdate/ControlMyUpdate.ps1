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
# Version: 2.3.1
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
# 1. Please use the GUI version 2.0 or higher with the script version 2.0 or higher.
# 2. Be aware that the GUI was not updated with the changes in Version 2.3 - please change the restart settings manually
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
#
# EmergencyKB
#                 e.g. KB12345                   This update will be installed even if the device is outside of the maintenance window and if the device is outside of the scan interval
#
# Update Source   MU,WSUS or Default             Will set the update source to the selected option 
#
# Retry count     
#                 0-99                           Configure how often updates getting downloaded and installed if an error appears 
#
# RunConnectionTests
#               True                            Device will try to reach all URLs in the ConnectionCheck.csv file
#               False                           No connection tests
#
# UninstallKBs
#               True                            KB's that are blocked will also be uninstalled - if already installed
#               False                           KB's will only be blocked and will stay installed if already installed
#

##########################################################################################
# Reboot settings in version 2.3 and above: 
# NoReboot 
#               True                            Device will not reboot
#               False                           Device will reboot if required
#
# ForceRebootWithNoUser
#               True                            For devices without maintenance window, the device will reboot ASAP if no user is logged on
#               False                           No automatic reboot if no user is logged on
#
# BlockRebootWithUser 
#               True                            Will block all reboots if a user is logged in to the device - will ignore the user session if Grace Period End is reached
#               False                           Will reboot the device regardless of an open user session
#
# MWAutomaticReboot
#               True                            Will force the reboot outside of the MW after Deadline is reached - only if no user is logged in
#               False                           Will wait for the next MW to reboot the device
#
# RebootGracePeriod
#               0-30                            Device will wait till x days are reached to force the reboot - independent of the current user session
#
# ForceReboot
#               True                            Force Reboot after x days
#               False                           Do not force the reboot after x days 

##########################################################################################
# Notification settings in version 2.3 and above: 
#
# NotifyUser
#                 True                           Will generate a toast notification
#                 False                          Will not notify the end-user 
#
# NotificationInterval
#                 1-99                            Hours between the Reboot notifications
#
# ForceRebootNotification
#                 True                          Only show the  last notification before the device gets rebooted
#                 False                        Either show all notifications - if NotifyUser = True; or no notifications - if NotifyUser = False

##########################################################################################
# Removed Settings: 
#

# NoMWAutomaticReboot
#               True                            Will reboot the device automatic if a pending reboot is detected and device is in MW
#               False                           Will not reboot the device automatic during the MW
#
# NoMWAutoRebootInterval
#               1-24                            If no MW is configured the device will force a reboot after X days - only possible if "NotifyUser" is set to true
#               0                               Disabled
#
# NotifyEnduserOutsideOfMW
#               True                            Will show the reboot notification if a user is logged on to the device and the device has a MW configured
#               False                           Will not show any notification
#
#
# MWAutoRebootInterval
#               1-28                            Days till the device gets forced rebooted outside of the MW
#               0                               Disabled
#

#
# MWBlockRebootWithUser
#               True                            Will block the reboot during MW if a user is logged in to the device
#               False                           Will execute the reboot during MW if a user is logged in
#
# MWForceRebootOnlyDuringMW
#               True                            Will only reboot the device during MW - will ignore logged in user sessions
#               False                           Will not reboot during MW if a use is logged on
#
# CMUAutoRebootInterval_OnlyMW
#               1-28                            Days till the device gets forced rebooted inside of the MW
#               0                               Disabled
#
##########################################################################################
#                                    Changelog 
#
# 2.3.1 - Fixed pending reboot registry status
# 2.3   - Re-desing of reboot handler
# 2.2.5 - Bugfixing Update Categories (OR instead of AND if selected more than one category)
# 2.2.4 - Bugfixing Update Categories
# 2.2.3 - New Feature
#           - Forced reboot only during configured MW
# 2.2.2 - Bugfixes:
#           - Block reboot during MW with user logged in
# 2.2.1 - Bugfixes:
#           - GUID for FeaturePacks fixed
#       - New Features:
#           - Block reboot during MW if an user is logged in to the system
# 2.2   - Bugfixes: 
#           - Reboot notification loop if user is no admin user
#           - Connection test can now be disabled
#           - removed the notification sound
#       - New Features:
#           - Optional Prompt for end users if MW is configured
#           - Search only specific categories
#           - Disable Driver downloads
#           - Per Day Maintenance Windows
#       - Improvements: 
#           - NotificationInterval - The time between the end user notification is now configurable 
#           - Notification Message is now based on the registry -> customizable for the admin
# 2.1.3 - added option for connection test (enabled/disabled)
#       - connection test will only run before Windows Update search is triggered 
#       - Added the option to force a reboot if no user is logged on (for non MW devices)
#       - Added Update Rollback
# 2.1.2 - bugfix notification configuration
# 2.1.1 - bugfix for single update installation
# 2.1 - added the following features:
#       - retry count for error handling (for download and installation of updates) 
#       - pending reboot detection bugfix (if no update was pending, reboot was not triggered during the MW)
#       - New toast notification feature from Damien van Baeys - https://github.com/damienvanrobaeys/Intune-Proactive-Remediation-scripts/tree/main/Reboot%20warning
#           - Removed toast title and text
#       - Improved update detection and reporting
#       - Added Windows Update and Delivery Optimization service connection test
#       - Added force reboot function for non MW devices
#       - Added automatic reboot configuration option for MW devices
#       - Added "old" CSP support natively (detection if 32 or 64 bit registry hive is used)
#       - Bugfix for logging function (no date / time in CMtrace after a specific date)
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

$ScriptCurrentVersion = "2.3.1"

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
            $Time = Get-Date -Format "HH:mm:ss.ffffff"
            $Date = Get-Date -Format "MM-dd-yyyy"
 
            if ($ErrorMessage -ne $null) { $Type = 3 }
            if ($Component -eq $null) { $Component = " " }

            switch ($LogLevel) {
                "Info" { [int]$Type = 1 }
                "Warning" { [int]$Type = 2 }
                "Error" { [int]$Type = 3 }
                "Debug" { [int]$Type = 4 }
                "Trace" { [int]$Type = 5 }
                default { [int]$Type = 1 }
            }

            Switch ($LogLevel) {
                "Info" { $Log = "<![LOG[$LogMessage $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"$thread`" >" }
                "Debug" { if (($ScriptLogLevel -eq "Debug") -or ($ScriptLogLevel -eq "Trace")) { $Log = "<![LOG[$LogMessage $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"$thread`">" } }
                "Trace" { if ($ScriptLogLevel -eq "Trace") { $Log = "<![LOG[$LogMessage $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"$thread`" >" } }
                "Error" { $Log = "<![LOG[$LogMessage $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"$thread`" >" }
            }
                
            if ($log) { $Log | Out-File -Append -Encoding UTF8 -FilePath $LogPath }
        }
        Remove-Variable -Name LogMessage, LogLevel, Log -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }    
}

function Test-MaintenanceWindow {
    $Component = "TEST MAINTENANCE WINDOW"
    Write-Log -LogLevel Trace -LogMessage "Function: Test-MaintenanceWindow: Start"
    Write-Log -LogLevel Info -LogMessage "Check if device is in maintenance window"
	
    #Get current time
    $CurrentDate = Get-Date
    $CurrentDay = $CurrentDate.DayOfWeek.value__
    $CurrentHour = Get-Date -Format HH 
    $CurrentMinute = Get-Date -Format mm

    #Per Day MW
    if ($settings.EnablePerDayMW -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Per Day MW enabled"
        $CurrentDay = Get-Date -Format "dddd"

        if ($settings.EnablePerDayMW -eq $true) {
            $StartTime = Get-ItemPropertyValue -Path"$($RegistryRootPath)\Settings"  -Name "MWPerDay$($Currentday)StartTime"
            $EndTime = Get-ItemPropertyValue -Path "$($RegistryRootPath)\Settings"  -Name "MWPerDay$($Currentday)EndTime"        
        }


        #Get start time
        $MWStartHour = $StartTime.Substring(0, 2)
        $MWStartMinute = $StartTime.Remove(0, 3)
        
        #Get stop time
        $MWStopHour = $EndTime.Substring(0, 2)
        $MWStopMinute = $EndTime.Remove(0, 3)
        
        Clear-Variable IsInMaintenanceWindow -Force -ErrorAction SilentlyContinue
        [Boolean] $IsInMaintenanceWindow = $false

        if ($CurrentHour -ge $MWStartHour -and $CurrentHour -le $MWStopHour) {
            Write-Log -LogLevel Info -LogMessage  "Current Hour within maintenance window timeframe"
            if ((($CurrentHour -eq $MWStartHour -and $CurrentMinute -gt $MWStartMinute) -and ($CurrentHour -eq $MWStopHour -and $CurrentMinute -lt $MWStopMinute))) {
                Write-Log -LogLevel Info -LogMessage "Same Start and Stop hour, using minutes to check Maintenance Windows"
                $IsInMaintenanceWindow = $true
            }
            elseif ($CurrentHour -ge $MWStartHour -and $CurrentHour -lt $MWStopHour) {
                Write-Log -LogLevel Info -LogMessage "Checking if time is between start and stop hour"
                $IsInMaintenanceWindow = $true
            }
            elseif (($CurrentHour -ge $MWStartHour -and $CurrentHour -eq $MWStopHour) -and $CurrentMinute -lt $MWStopMinute) {
                Write-Log -LogLevel Info -LogMessage  "CHecking Maintenance Windows within the last hour"
                $IsInMaintenanceWindow = $true
            }
        }
        else {
            Write-Log -LogLevel Info -LogMessage  "Current time not within defined maintenance window"
            $IsInMaintenanceWindow = $false
        }
            
    }
    
    #Simple MW
    else {
        #Get start time
        $MWStartHour = $Settings.MWStartTime.Substring(0, 2)
        $MWStartMinute = $Settings.MWStartTime.Remove(0, 3)
        Write-Log -LogLevel Info -LogMessage "Start Time: Hour: $($MWStartHour) / Minute: $($MWStartMinute)"
        
        #Get stop time
        $MWStopHour = $Settings.MWStopTime.Substring(0, 2)
        $MWStopMinute = $Settings.MWStopTime.Remove(0, 3)
        Write-Log -LogLevel Info -LogMessage "Stop Time: Hour: $($MWStopHour) / Minute: $($MWStopMinute)"

        #Check if installation day was set - if not, updates will be installed everyday 
        if ($Settings.MWDay) {
            if ($Settings.MWDay -like "*,*") {
                Write-Log -LogLevel Info -LogMessage "Multiple target day found. $($Settings.MWDay)"
                $TargetDay = $Settings.MWDay.Split(",")
            }
            else {
                Write-Log -LogLevel Info -LogMessage "1 Target day found $($Settings.MWDay)"
                $TargetDay = $Settings.MWDay
            }
        }
        else {
            Write-Log -LogLevel Info -LogMessage "Everyday as target day"
            $TargetDay = $CurrentDay
        }

        Write-Log -LogLevel Info -LogMessage "TargetDay: $($TargetDay)"

        #Check if current time is in Maintenance Window
        Clear-Variable IsInMaintenanceWindow -Force -ErrorAction SilentlyContinue
        $CurrentDayIsInMW = $false

        [Boolean] $IsInMaintenanceWindow = $false
        foreach ($MWDay in $TargetDay) {
            Write-Log -LogLevel Info -LogMessage "processing: $($MWDay)"  
            if ( $CurrentDay -eq $MWDay ) {
                Write-Log -LogLevel Info -LogMessage "Current day defined in maintenance window"
                $CurrentDayIsInMW = $true
            }   
            
            if ($CurrentDayIsInMW -eq $true) {
                if (($CurrentDay -eq ($TargetDay[$TargetDay.Count - 1])) -and ($TargetDay.Count -gt 1)) {
                    $CurrentDayIsLastDay = $True
                }
                else { $CurrentDayIsLastDay = $False }
                
                Write-Log -LogLevel Info -LogMessage "If current day in Maintenance Window: True"
                if ($MWStopHour -lt $MWStartHour) {
                    Write-Log -LogLevel Info -LogMessage "If Checking Stop Hour is smaller than Start Hour. Over midnight MW."
                    if ($CurrentHour -ge $MWStartHour -or $CurrentHour -le $MWStopHour) { 
                        Write-Log -LogLevel Info -LogMessage "Checking Maintenance Windows timeframe"
                        if ($CurrentHour -ge $MWStartHour -and $CurrentMinute -ge $MWStartMinute) {                       
                            Write-Log -LogLevel Info -LogMessage "Stop Hour at 24"
                            $MWStopHour = "24"
                            $MWStopMinute = "00"
                        }
                        if ($CurrentHour -ge $MWStartHour -or $CurrentHour -le $MWStopHour) { 
                            if ($CurrentHour -le $MWStopHour -and $CurrentMinute -le $MWStopMinute) {                            
                                Write-Log -LogLevel Info -LogMessage "Stop Hour at 0"
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
                    Write-Log -LogLevel Info -LogMessage "Current Hour within maintenance window timeframe"
                    if ((($CurrentHour -eq $MWStartHour -and $CurrentMinute -gt $MWStartMinute) -and ($CurrentHour -eq $MWStopHour -and $CurrentMinute -lt $MWStopMinute))) {
                        Write-Log -LogLevel Info -LogMessage "Same Start and Stop hour, using minutes to check Maintenance Windows"
                        $IsInMaintenanceWindow = $true
                    }
                    elseif ($CurrentHour -ge $MWStartHour -and $CurrentHour -lt $MWStopHour) {
                        Write-Log -LogLevel Info -LogMessage "Checking if time is between start and stop hour"
                        $IsInMaintenanceWindow = $true
                    }
                    elseif (($CurrentHour -ge $MWStartHour -and $CurrentHour -eq $MWStopHour) -and $CurrentMinute -lt $MWStopMinute) {
                        Write-Log -LogLevel Info -LogMessage "CHecking Maintenance Windows within the last hour"
                        $IsInMaintenanceWindow = $true
                    }
                }
                else {
                    Write-Log -LogLevel Info -LogMessage "Current time not within defined maintenance window"
                    $IsInMaintenanceWindow = $false
                }
            }
            else {
                Write-Log -LogLevel Info -LogMessage "Current time not within defined maintenance window"
                $IsInMaintenanceWindow = $false
            }
        }
    }
    Write-Log -LogLevel Info -LogMessage "IsInMaintenanceWindow: $($IsInMaintenanceWindow)"
    
    if ($IsInMaintenanceWindow -eq $true) { Write-Log -LogLevel Info -LogMessage "Device is in maintenance window" }
    if ($IsInMaintenanceWindow -eq $false) { Write-Log -LogLevel Info -logMessage "Device is outside of the maintenance window" }

    Write-Log -LogLevel Trace -LogMessage "Function: Test-MaintenanceWindow: End"
    return $IsInMaintenanceWindow
}

function Start-RebootNotification {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $false, HelpMessage = "Grace Period for user deferral restarts")][bool]$ForceRebootNotification = $false
    )

    $Component = "REBOOT NOTIFICATION"
    Write-Log -LogLevel Trace -LogMessage "Function: Start-RebootNotification : Start"

    if ($ForceRebootNotification -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Start forced Reboot Notification to reboot"
        Start-ScheduledTask -TaskName "Control My Update - Reboot Notification" 
    }
    elseif ($NotifyUser -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Start Pending Reboot Notification"
        
        $RebootNotificationCreated = Get-ItemPropertyValue -Path "$($RegistryRootPath)\Status" -Name "RebootNotificationCreated" -ErrorAction SilentlyContinue

        if (!($RebootNotificationCreated)) { 
            $NotificationDate = Get-Date -Format s    
            Write-Log -LogLevel Info -LogMessage "Create first end user notification"                
            Start-ScheduledTask -TaskName "Control My Update - Reboot Notification" 
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootNotificationCreated" -Value "True" -Force | Out-Null    
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootNotificationDate" -Value $NotificationDate -Force | Out-Null   
        }
        elseif ($RebootNotificationCreated -eq $True) {
            $LastNotificationTime = Get-Date (Get-ItemPropertyValue -Path "$($RegistryRootPath)\Status" -Name "RebootNotificationDate") -Format s
            $IntervalTimerTemp = (Get-Date).AddHours( - $($Settings.NotificationInterval))
            $IntervalTimer = Get-Date ($IntervalTimerTemp) -format s
            if ($LastNotificationTime -le $IntervalTimer) {
                $NotificationDate = Get-Date -Format s   
                Write-Log -LogLevel Info -LogMessage "Create Notification since interval is reached"
                Start-ScheduledTask -TaskName "Control My Update - Reboot Notification" 
                New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootNotificationDate" -Value $NotificationDate -Force | Out-Null  
            }
        }
        
    }

    Write-Log -LogLevel Trace -LogMessage "Function:  Start-RebootNotification : End"
}

function Test-GracePeriod {

    $Component = "TEST GRACE PERIOD"
    Write-Log -LogLevel Trace -LogMessage "Function: Test-GracePeriod: Start"

    if ($Settings.RebootGracePeriod -eq '0') {
        New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "ShowDismissButton" -Value "False" -Force | Out-Null   
        $GracePeriodEnd = $true
    }
    elseif ($Settings.RebootGracePeriod -ne '0') {
        if (!(Get-ItemProperty "$($RegistryRootPath)\Status" -Name 'RebootDetectionDate' -ErrorAction Ignore)) {
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootDetectionDate" -Value (Get-Date -Format s) -Force | Out-Null   
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "ShowDismissButton" -Value "True" -Force | Out-Null   
            
            $GracePeriodEnd = $false
        }
        else {
            
            $RebootDetectionDate = Get-Date (Get-ItemPropertyValue "$($RegistryRootPath)\Status" -Name 'RebootDetectionDate' -ErrorAction Ignore)
    
            if ($RebootDetectionDate -le (Get-Date).AddDays( - ($Settings.RebootGracePeriod))) {
                Write-Log -LogLevel Info -LogMessage "Force reboot notification"

                New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "ShowDismissButton" -Value "False" -Force | Out-Null   
               
                $GracePeriodEnd = $true
            }
            else { $GracePeriodEnd = $false }
        }
    }

    Write-Log -LogLevel Info -LogMessage "Device Grace Period for reboot ended: $($GracePeriodEnd)"
    

    Write-Log -LogLevel Trace -LogMessage "Function: Test-GracePeriod: End"
    Return $GracePeriodEnd
}

function Start-RebootExecution {
    $Component = "START REBOOT EXECUTION"
    Write-Log -LogLevel Trace -LogMessage "Function: Start-RebootExecution: Start"

    $NoReboot = $False

    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "PendingReboot" -Value "True" -Force | Out-Null

    if ($NoReboot -eq $true) {
        $DoReboot = $False
    }
    else {
        
        [bool]$GracePeriodResult = Test-GracePeriod

        #Check Maintenance Window use cases
        If ($MaintenanceWindow -eq $true) {
            [bool]$DeviceInMaintenanceWindow = Test-MaintenanceWindow

            if (($GracePeriodResult -eq $false) -and ($MWAutomaticReboot -eq $True) -and ($MaintenanceWindow -eq $True) -and ($DeviceInMaintenanceWindow -eq $True)) {
                Write-Log -LogLevel Info -LogMessage "Pending Reboot - Device in MW - with Automatic Reboot enabled and no grace period"

                #reboot the device if pending reboot and no user is logged in
                if ((!(Get-Process explorer -ErrorAction SilentlyContinue) -and $($BlockRebootWithUser) -eq $true)) {          
                    Write-Log -LogLevel Info -LogMessage "No User logged in - Block Reboot with User session enabled"      
                    $DoReboot = $true
                }
                #Do not reboot the device if pending reboot and user is logged in
                elseif (((Get-Process explorer -ErrorAction SilentlyContinue) -and $($BlockRebootWithUser) -eq $true)) {    
                    Write-Log -LogLevel Info -LogMessage "User logged in - Block Reboot with User session disabled"                 
                    $DoReboot = $false
                }
                #Reboot the device regardless if a user is logged in or not
                elseif ($BlockRebootWithUser -eq $false) {        
                    Write-Log -LogLevel Info -LogMessage "Block Reboot with User session disabled - ignore user session"             
                    $DoReboot = $true
                }
            }
            #Force Reboot the device after x days if the device is in MW
            elseif (($GracePeriodResult -eq $true) -and $MaintenanceWindow -eq $True -and ($DeviceInMaintenanceWindow -eq $True) -and $ForceReboot -eq $false) {
                Write-Log -LogLevel Info -LogMessage "Device is in MW - with Grace Period end reached"     
                $DoReboot = $true
            }
            #Force Reboot the device after x days regardless of the device is in MW or not
            elseif (($GracePeriodResult -eq $true) -and $MaintenanceWindow -eq $True -and $ForceReboot -eq $true) {
                Write-Log -LogLevel Info -LogMessage "Device is in MW - with Grace Period end reached"  
                $DoReboot = $true
            }
            if ((!(Get-Process explorer -ErrorAction SilentlyContinue) -and $MaintenanceWindow -eq $True -and ($DeviceInMaintenanceWindow -eq $True) -and $($ForceRebootwithNoUser) -eq $true)) {                
                #reboot the device if pending reboot and no user is logged in
                Write-Log -LogLevel Info -LogMessage "No User logged in - Force Reboot"  
                $DoReboot = $true
            }
        }
        else {
        
            Write-Log -LogLevel Info -LogMessage "Device has no MW configured"  
            #reboot the device if pending reboot and no user is logged in
            if ((!(Get-Process explorer -ErrorAction SilentlyContinue) -and $($BlockRebootWithUser) -eq $true) -and ($GracePeriodResult -eq $true)) {                
                Write-Log -LogLevel Info -LogMessage "No User logged in and grace period ended - Force Reboot" 
                $DoReboot = $true
            }
            elseif (((Get-Process explorer -ErrorAction SilentlyContinue) -and $($BlockRebootWithUser) -eq $true) -and ($GracePeriodResult -eq $true)) {                
                Write-Log -LogLevel Info -LogMessage "User logged in and grace period ended - No Force Reboot due to BlockRebootWithUser configuration" 
                $DoReboot = $False
            }
            #reboot the device if pending reboot and  user is logged in
            elseif (((Get-Process explorer -ErrorAction SilentlyContinue) -and $($BlockRebootWithUser) -eq $false) -and ($GracePeriodResult -eq $true)) {                
                Write-Log -LogLevel Info -LogMessage "User logged in and grace period ended - Force Reboot"              
                $DoReboot = $true
            }
            if ((!(Get-Process explorer -ErrorAction SilentlyContinue) -and $($ForceRebootwithNoUser) -eq $true)) {        
                Write-Log -LogLevel Info -LogMessage "No User logged in - Force Reboot"               
                $DoReboot = $true
            }
        }

    }

    #execute reboot
    if ($DoReboot -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Automatic reboot activated. Running shutdown command"
        Start-RebootNotification -ForceRebootNotification $true
        shutdown.exe /r /f /t 120
        New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "ShowDismissButton" -Value "False" -Force | Out-Null 
    }
    else {
        Write-Log -LogLevel Info -LogMessage "Automatic reboot NOT triggered."
        Start-RebootNotification -ForceRebootNotification $False
    }

    Write-Log -LogLevel Trace -LogMessage "Function: Start-RebootExecution: End"

}

function Test-PendingReboot {
    param(
        [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = "Enable Automatic Reboot")][bool]$AutomaticReboot = $false
    )
    $Component = "TEST PENDING REBOOT"
    Write-Log -LogLevel Trace -LogMessage "Function: Test-PendingReboot: Start"
    Write-Log -LogLevel Info -LogMessage "AutomaticReboot: $($AutomaticReboot)"

    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "ShowDismissButton" -Value "True" -Force | Out-Null 

    $PendingRestart = $false
    if (Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -EA Ignore) { $PendingRestart = $true }
    if (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -EA Ignore) { $PendingRestart = $true }
    if ((New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired -eq $true) { $PendingRestart = $true }

    Write-Log -LogLevel Info -LogMessage "PendingRestart: $($PendingRestart)"

    if ($PendingRestart -eq $true) {
        if (!(Get-ItemProperty "$($RegistryRootPath)\Status" -Name 'RebootDetectionDate' -ErrorAction Ignore)) {
            New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "RebootDetectionDate" -Value (Get-Date -Format s) -Force | Out-Null   
        }
        else { $RebootDetectionDate = Get-Date (Get-ItemPropertyValue "$($RegistryRootPath)\Status" -Name 'RebootDetectionDate') -Format s }

        Start-RebootExecution
        Write-Log -LogLevel Info -LogMessage "Pending Restart - Device require reboot"
    } 
    else {
        if ((Get-ItemProperty "$($RegistryRootPath)\Status" -Name 'RebootDetectionDate' -ErrorAction Ignore)) {
            Remove-ItemProperty -Path "$($RegistryRootPath)\Status" -Name "RebootDetectionDate" -Force | Out-Null   
        }
	New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "PendingReboot" -Value "False" -Force | Out-Null
    }
    
    Write-Log -LogLevel Trace -LogMessage "Function: Test-PendingReboot: End"
    return $PendingRestart
}

function Write-UpdateStatus {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "KB Number to update")][String] $CurrentUpdate,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Force status change update")][String] $StatusChange,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Update Title")][String] $UpdateTitle,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Update Category")][String] $UpdateCategory,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Status of the update to recorded")][String] $Status,
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "UpdateID")][String] $UpdateID
    )
    $Component = "WRITE UPDATE STATUS"

    Write-Log -LogLevel Trace -LogMessage "Function: Write-UpdateStatus: Start"
    Write-Log -LogLevel Debug -LogMessage "CurrentUpdate: $($CurrentUpdate)"
    Write-Log -LogLevel Debug -LogMessage "Status: $($Status)"

    $LastUpdate = Get-Date -Format s
    
    if (!(Test-Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)")) {
        New-Item "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -Force | Out-Null
    }

    if ($UpdateTitle) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "Title" -Value $UpdateTitle -Force | Out-Null
    }
    if ($UpdateCategory) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "Category" -Value $UpdateCategory -Force | Out-Null
    }

    if ($UpdateID) {
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "UpdateID" -Value $UpdateID -Force | Out-Null
    }

    if ($Status) {
        if (!(Get-Item "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)\$($Status)" -ErrorAction SilentlyContinue) -or $StatusChange -eq $True) {
            New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "$Status" -Value $LastUpdate -Force | Out-Null
        }
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "Last Status" -Value $Status -Force | Out-Null
        New-ItemProperty -Path "$($RegistryRootPath)\Status\KBs\$($CurrentUpdate)" -PropertyType "String" -Name "Last Status Timestamp" -Value $LastUpdate -Force | Out-Null
    }
    Write-Log -LogLevel Trace -LogMessage "Function: Write-UpdateStatus: End"
}

function Get-InstalledWindowsUpdates {
    param (
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")]$AllUpdates
    )
    $Component = "INSTALLED UPDATE VALIDATION"
    
    Write-Log -LogLevel Trace -LogMessage "Function: Get-InstalledWindowsUpdates: Start"
    Write-Log -LogLevel Info -LogMessage "Searching for installed updates"

    #Get WUA installation status
    $WUAUpdateKBs = ($AllUpdates | Where-Object { $_.IsInstalled -eq $true }).KBArticleIDs

    $UpdateCollection = @()
    foreach ($temp in $WUAUpdateKBs) {    
        $KBName = "KB$($temp)"
        $UpdateCollection += $KBName
    }  

    #Get Windows Updates from WMI
    $WMIKBs = Get-WmiObject win32_quickfixengineering |  Select-Object HotFixID -ExpandProperty HotFixID
    Write-log -LogLevel Debug -LogMessage "WMI KB List: $($WMIKBs)"
   
    
    #Get Windows Updates from DISM
    $RegExKB = "KB(\d+)"
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
    $Component = "SEARCH ALL UPDATES"
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
        $HiddenFilter = "IsHidden = 0"
        $updates = ($updateSearcher.Search($SearchFilter))
    }
    if ($settings.UpdateCategories -ne "All") {
        $CategorySettings = $settings.UpdateCategories.Split(",")
           
        foreach ($Category in $CategorySettings) {
            if ($CategoryFilter) {
                $CategoryFilter = "$($CategoryFilter) OR CategoryIDs contains '$($Category)'"
            }
            else { $CategoryFilter = "CategoryIDs contains '$($Category)'" }
        }
    }
    if ($InstallDrivers -eq $False) {
        $TypeFilter = "Type != 'Driver'"
    }
    
    if (($TypeFilter -and $CategoryFilter) -or ($TypeFilter -and $CategoryFilter -and $HiddenFilter) -or ($HiddenFilter -and $CategoryFilter) -or ($HiddenFilter -and $TypeFilter)) {
        if ($HiddenFilter -and $CategoryFilter -and !$TypeFilter) { $SearchFilter = "$($HiddenFilter) AND $($CategoryFilter)" }
        elseif ($HiddenFilter -and $CategoryFilter -and $TypeFilter) { $SearchFilter = "$($HiddenFilter) AND $($TypeFilter) AND $($CategoryFilter)" }
        elseif (!$HiddenFilter -and $CategoryFilter -and $TypeFilter) { $SearchFilter = "$($TypeFilter) AND $($CategoryFilter)" }
        elseif ($HiddenFilter -and !$CategoryFilter -and $TypeFilter) { $SearchFilter = "$($HiddenFilter) AND $($TypeFilter)" }
    }
    elseif ($TypeFilter) { $SearchFilter = "$TypeFilter" }
    elseif ($HiddenFilter) { $SearchFilter = "$HiddenFilter" }
    elseif ($CategoryFilter) { $SearchFilter = "$CategoryFilter" }
    
    Write-Log -Loglevel Info -LogMessage "Search Filter: $($SearchFilter)"

    if ($SearchFilter) {
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
    Write-Log -LogLevel Info -LogMessage "Update search result status: $($Status)"

    
    foreach ($item in $Updates.Updates) {
        if ($item.IsDownloaded -eq $false -and $item.IsInstalled -eq $False) {
            $Status = "Available"
            Write-Log -LogLevel Info -LogMessage "KB$($item.KBArticleIDs) $($Status)"
            Write-UpdateStatus -CurrentUpdate "KB$($item.KBArticleIDs)" -Status $Status -UpdateTitle ($item.title) -UpdateCategory ($item.Categories._NewEnum.name) -UpdateID ($item.identity._NewEnum) | Out-Null
        }
        else {
            Write-Log -LogLevel Debug -LogMessage "KB$($item.KBArticleIDs) $($Status)"
            Write-UpdateStatus -CurrentUpdate "KB$($item.KBArticleIDs)" -UpdateTitle ($item.title) -UpdateCategory ($item.Categories._NewEnum.name) -UpdateID ($item.identity._NewEnum) | Out-Null 
        }

        Clear-Variable Status -Force -ErrorAction SilentlyContinue
        
    }

    Write-Log -LogLevel Trace -LogMessage "Function: Search-AllUpdates: End"
    return $($Updates.Updates)
}

function Save-WindowsUpdate {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")] $DownloadUpdateList
    )
    $Component = "DOWNLOAD UPDATE"
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
        do {  
            $a++

            #Initialize download
            Write-Log -LogLevel Debug -LogMessage "Initializing download"
            $downloader = $updateSession.CreateUpdateDownloader()
            $downloader.ClientApplicationID = "ControlMyUpdate"
            $downloader.IsForced = $True
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
        [Parameter(Mandatory = $false, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")]$AllUpdates
    )
    $Component = "UPDATE INSTALLATION STATUS"
    Write-Log -LogLevel Trace -LogMessage "Function: Get-NonInstalledUpdates: Start"
    Write-Log -LogLevel Debug -LogMessage "AvailableUpdates: $($AvailableUpdates)"


    Write-Log -LogLevel Trace -LogMessage "Get-InstalledWindowsUpdates"
    $InstalledKBs = Get-InstalledWindowsUpdates
    Write-Log -LogLevel Debug -LogMessage "Installed KBs returned: $($InstalledKBs)"

    $AvailableKBs = ($AllUpdates | Where-Object { $_.IsInstalled -eq $false -and $_.IsHidden -eq $false }).KBArticleIDs

    $RegistryKBList = Get-ChildItem "$($RegistryRootPath)\Status\KBs"
    Write-Log -LogLevel Debug -LogMessage "Item from registry: $($RegistryKBList)"

    foreach ($RegKey in $RegistryKBList) {  
                 
        $DetectedKBs = $RegKey.name.Split('\')[-1]
        $KBCheck = $InstalledKBs | Where-Object { $_ -eq $DetectedKBs }
        $AvailableKBCheck = $AvailableKBs | Where-Object { $_ -eq $($DetectedKBs.replace("KB", "")) }
            

        $KeyNames = $RegKey.property | Where-Object { $_ -like "Installation Status*" }
            
        foreach ($Item in $KeyNames) {
            Remove-ItemProperty "$($RegistryRootPath)\Status\KBs\$($DetectedKBs)" -Name $Item -Force
        }

        if ($KBCheck) {
            Write-UpdateStatus -CurrentUpdate "$($DetectedKBs)" -Status "Installation Status : Installed" -StatusChange $True
            Write-Log -LogLevel Info -LogMessage "Status Change for $($DetectedKBs) - now Installed"
        }
        elseif ($AvailableKBCheck) {
            Write-UpdateStatus -CurrentUpdate "$($DetectedKBs)" -Status "Installation Status : Pending Installation" -StatusChange $True
            Write-Log -LogLevel Info -LogMessage "Status Change for $($DetectedKBs) - now pending for installation"
        }
        elseif (!$KBCheck -and !$AvailableKBCheck -and $RegKey.property -notcontains "Installation Status : Installed" ) {
            Write-UpdateStatus -CurrentUpdate "$($DetectedKBs)" -Status "Installation Status : Not Applicable" -StatusChange $True
            Write-Log -LogLevel Info -LogMessage "Status Change for $($DetectedKBs) - not applicable anymore"
        }
        elseif (!$KBCheck -and !$AvailableKBCheck -and $RegKey.property -contains "Installation Status : Installed" ) {
            Write-UpdateStatus -CurrentUpdate "$($DetectedKBs)" -Status "Installation Status : Uninstalled" -StatusChange $True
            Write-Log -LogLevel Info -LogMessage "Status Change for $($DetectedKBs) - Uninstalled"
        }


        Clear-Variable KBCheck, AvailableKBCheck -Force
    }

    #Write installed updates to registry
    $Reglist = @()
    foreach ($RegKB in $RegistryKBList) {
        $RegName = $RegKB.name.Split('\')[-1]
        $Reglist += $RegName
    }

    foreach ($WUAUpdate in $AllUpdates) {
        if ($Reglist -contains "KB$($WUAUpdate.KBArticleIDs)") {
            Write-UpdateStatus -CurrentUpdate "KB$($WUAUpdate.KBArticleIDs)" -UpdateTitle $WUAUpdate.Title   
        }
    }

    foreach ($KB in ($InstalledKBs | Where-Object { $_ })) {        
        if ($Reglist -notcontains $KB) {
            Write-UpdateStatus -CurrentUpdate $KB -Status "Installation Status : Installed"  -StatusChange $True   
        }        
    }

    Write-Log -LogLevel Trace -LogMessage "Function: Get-NonInstalledUpdates: End"
    # return $DetectedKBs
}

function Set-WindowsUpdateBlockStatus {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "KB ID Number to hide")][String] $KBArticleID,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Status of the update")][ValidateSet("Blocked", "UnBlocked")][String] $Status,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Windows Update ComObject from Search-AllUpdates")] $AllUpdates
    )
    $Component = "SET BLOCK STATUS"
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
    $Component = "DO STATISTICS"
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
function Test-UpdateConnectivity {
    $Component = "TEST CONNECTIVITY"
    $DOConnected = $True
    $WUConnected = $True

    if ((Test-Path "$RegistryRootPath\Status\Connection") -eq $false) {
        New-Item -Path "$RegistryRootPath\Status\Connection" -Force
        New-Item -Path "$RegistryRootPath\Status\Connection\Delivery Optimization" -Force
        New-Item -Path "$RegistryRootPath\Status\Connection\Microsoft Update" -Force
        New-Item -Path "$RegistryRootPath\Status\Connection\Other" -Force
    }

    $ConnectionURLs = Import-Csv "$PSScriptRoot\ConnectionCheck.csv"  

    foreach ($Url in $ConnectionURLs) {
        try {
            $testPort = [System.Net.Sockets.TCPClient]::new()
            $testPort.SendTimeout = 5
            $testPort.Connect($Url.url, $Url.Port)
            $result = $testPort.Connected
        }
        catch {
            $result = $_.Exception.InnerException.Message
        }
            
        Write-Log -LogLevel Debug -LogMessage  $result

        if ($url.UseCase -eq "WU") { $ConnectionRegPath = "Microsoft Update" }
        elseif ($url.UseCase -eq "DO") { $ConnectionRegPath = "Delivery Optimization" }
        else { $ConnectionRegPath = "Other" } 

        New-ItemProperty -Path "$RegistryRootPath\Status\Connection\$($ConnectionRegPath)" -Name "$($Url.URL):$($Url.Port)" -PropertyType String -Value $result -Force | Out-Null
        
        $testPort.Close()
    }
  
    $DOResults = Get-Item -Path "$RegistryRootPath\Status\Connection\Delivery Optimization"
   
    foreach ($Item in $DOResults.Property) {
        if ((Get-ItemPropertyValue -Path "$RegistryRootPath\Status\Connection\Delivery Optimization" -Name $Item) -ne $True) { $DOConnected = $False }
    }

    $WUResults = Get-Item -Path "$RegistryRootPath\Status\Connection\Microsoft Update"
   
    foreach ($Item in $WUResults.Property) {
        if ((Get-ItemPropertyValue -Path "$RegistryRootPath\Status\Connection\Microsoft Update" -Name $Item) -ne $True) { $WUConnected = $False }
    }

    New-ItemProperty -Path "$RegistryRootPath\Status\Connection" -Name "Delivery Optimization Status" -PropertyType String -Value $DOConnected -Force | Out-Null
    New-ItemProperty -Path "$RegistryRootPath\Status\Connection" -Name "Windows Update Status" -PropertyType String -Value $WUConnected -Force | Out-Null

    Write-Log -LogLevel Info -LogMessage "Delivery Optimization connection status: $($DOConnected)"
    Write-Log -LogLevel Info -LogMessage "Microsoft Update connection status: $($WUConnected)"


    Return $DOConnected, $WUConnected
    
}

function Install-SpecificWindowsUpdate {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, HelpMessage = "Updates To Install")] $SelectedUpdate
    )

    $Component = "UPDATE INSTALLATION"

    Write-Log -LogLevel Trace -LogMessage "Function: Install-SpecificWindowsUpdate: Start"
    Write-Log -LogLevel Trace -LogMessage "Foreach SelectedUpdate : start"


    foreach ($InstallUpdate in $SelectedUpdate) {

        $updatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
        $updatesToInstall.Add($InstallUpdate) | Out-Null 

        $b = 0
        do {
            $b++
            Write-Log -LogLevel Info -LogMessage "Starting ($($b).) installation of $($InstallUpdate.Title)"
            Write-UpdateStatus -CurrentUpdate "KB$($InstallUpdate.KBArticleIDs)" -Status "Start installation" -StatusChange $True
    
            #start update installation
            $installer = New-Object -ComObject 'Microsoft.Update.Installer'
            $installer.ClientApplicationID = "ControlMyUpdate"
            $installer.Updates = $updatesToInstall 
            $installer.IsForced = $True         
            $installResult = $installer.Install()
            
            Write-Log -LogLevel Info -LogMessage "Install result code: $($installResult.ResultCode)"
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
    
        }Until($b -ge $retrycount -or $installResult.ResultCode -eq 2 -or $installResult.ResultCode -eq 7)

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

    $Component = "NEW WINDOWS UPDATE SCAN"

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
    $FullUpdateList = Search-AllUpdates -UpdateSource $($Settings.UpdateSource) -IgnoreHideStatus
    $AllUpdatesFound = $FullUpdateList | Where-Object { $_.IsHidden -eq $False }

    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastScanTime' -Value $(Get-Date -Format s) | Out-Null
    Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'NextScanTime' -Value $NextScanTime | Out-Null

    Write-Log -LogLevel Trace -LogMessage "Function: New-WindowsUpdateScan: End"
    return $AllUpdatesFound
}


##########################################################################################
#                                   Define variables

$LogFileName = "WindowsUpdate_{0:yyyyMM}.log" -f (Get-Date)
$RegistryRootPath = "HKLM:\SOFTWARE\ControlMyUpdate"
$64BitRegistryRootPath = "HKLM:\SOFTWARE\WOW6432Node\ControlMyUpdate"

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

Write-Log -LogLevel Trace -LogMessage "Registry test path root key"
$32BitRegistryTest = Test-Path $RegistryRootPath

If (!$32BitRegistryTest) {
    $64BitRegistryTest = Test-Path $64BitRegistryRootPath

    if ($64BitRegistryRootPath) {
        $RegistryRootPath = $64BitRegistryRootPath
        $RegistryTest = $64BitRegistryTest
    }
}
else { $RegistryTest = $32BitRegistryTest }


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



if ($RegistryTest -eq $true) {
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

        if (!$settings.retrycount) {
            $retrycount = '3'
        }
        else { $retrycount = $settings.retrycount }
        
        
        if ($Settings.NoMWAutomaticReboot -eq "True") {
            [bool]$Reboot = $True
        }
        elseif ($settings.MWAutomaticReboot -eq "True") {
            [bool]$Reboot = $True
        }
        else { [bool]$Reboot = $False }           
               
    }
    else {
        Write-Log -LogLevel Error -LogMessage "No registry settings detected"
        break
    }    
}
else {
    Write-Log -LogLevel Error -LogMessage "No registry settings detected"
    break
}   

#Registry mapping - to make sure bool is bool
if ($settings.DirectDownload -eq "True") { [bool]$DirectDownload = $true } else { [bool]$DirectDownload = $false }
if ($Settings.MWAutomaticReboot -eq "True") { [bool]$MWAutomaticReboot = $true } else { [bool]$MWAutomaticReboot = $false }
if ($Settings.NotifyEnduserOutsideOfMW -eq "True") { [bool]$NotifyEnduserOutsideOfMW = $true } else { [bool]$NotifyEnduserOutsideOfMW = $false }
if ($Settings.RunConnectionTests -eq "True") { [bool]$RunConnectionTests = $true } else { [bool]$RunConnectionTests = $false }
if ($Settings.ForceRebootwithNoUser -eq "True") { [bool]$ForceRebootwithNoUser = $true } else { [bool]$ForceRebootwithNoUser = $false }
if ($Settings.ReportOnly -eq "True") { [bool]$ReportOnly = $true } else { [bool]$ReportOnly = $false }
if ($Settings.NotifyUser -eq "True") { [bool]$NotifyUser = $true } else { [bool]$NotifyUser = $false }
if ($Settings.MaintenanceWindow -eq "True") { [bool]$MaintenanceWindow = $true } else { [bool]$MaintenanceWindow = $false }
if ($Settings.UninstallKBs -eq "True") { [bool]$UninstallKBs = $true } else { [bool]$UninstallKBs = $false }
if ($Settings.NoMWAutomaticReboot -eq "True") { [bool]$NoMWAutomaticReboot = $true } else { [bool]$NoMWAutomaticReboot = $false }
if ($Settings.MWBlockRebootWithUser -eq "True") { [bool]$MWBlockRebootWithUser = $true } else { [bool]$MWBlockRebootWithUser = $false }
if ($Settings.MWForceRebootOnlyDuringMW -eq "True") { [bool]$MWForceRebootOnlyDuringMW = $true } else { [bool]$MWForceRebootOnlyDuringMW = $false }
if ($Settings.BlockRebootWithUser -eq "True") { [bool]$BlockRebootWithUser = $true } else { [bool]$BlockRebootWithUser = $false }
if ($Settings.ForceRebootNotification -eq "True") { [bool]$ForceRebootNotification = $true } else { [bool]$ForceRebootNotification = $false }
if ($Settings.ForceReboot -eq "True") { [bool]$ForceReboot = $true } else { [bool]$ForceReboot = $false }
if ($Settings.ForceRebootwithNoUser -eq "True") { [bool]$ForceRebootwithNoUser = $true } else { [bool]$ForceRebootwithNoUser = $false }
if ($Settings.InstallDrivers -eq "True") { [bool]$InstallDrivers = $true } else { [bool]$InstallDrivers = $false }




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

#restart the device if pending reboot and in MW
if ($MaintenanceWindow -eq $True) {
    if (Test-MaintenanceWindow -eq $true) {
        Test-PendingReboot -AutomaticReboot $Reboot      
    }
}
#Check if Force Reboot With No User is enabled an if NO user is currently logged in
if (!(Get-Process explorer -ErrorAction SilentlyContinue) -and $($ForceRebootwithNoUser) -eq $true) {
    #reboot the device if pending reboot and no user is logged in
    if (Test-PendingReboot -eq $true) {

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
            
            if ($UninstallKBs -eq $true) {
                Get-WindowsPackage -Online | ? { $_.ReleaseType -like "*Update*" } |  ForEach-Object { Get-WindowsPackage -Online -PackageName $_.PackageName } |  Where-Object { $_.Description -like "*$($Item)*" } | Remove-WindowsPackage -Online -NoRestart
            }

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

    if ($RunConnectionTests -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Testing connectivity"
        Test-UpdateConnectivity | Out-Null
    }

    Write-Log -LogLevel Info -LogMessage "Script First Run - Searching for updates"
    $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date)
}
else {
    $CurrentTime = Get-Date -Format s
    Write-Log -LogLevel Debug -LogMessage "CurrentTime: $($CurrentTime)"
    $NextScanTime = (Get-Date $($Settings.NextScanTime) -Format s)
    Write-Log -LogLevel Debug -LogMessage "NextScanTime: $($NextScanTime)"

    if ($CurrentTime -ge $NextScanTime -or (($MaintenanceWindow -eq $true) -and [bool](Test-MaintenanceWindow) -eq $true)) {
        
        if ($RunConnectionTests -eq $true) {
            Write-Log -LogLevel Info -LogMessage "Testing connectivity"
            Test-UpdateConnectivity | Out-Null
        }

        Write-Log -LogLevel Info -LogMessage "Scan Interval reached - Searching for updates"
        $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date $Settings.LastScanTime)
    }
    else {
        Write-Log -LogLevel Info -LogMessage "Scan Interval not reached"
    }
}


if (!$AllUpdates) {
    #Checking if device was rebooted
    $bootuptime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $CurrentDate = Get-Date
    $uptime = $CurrentDate - $bootuptime
    $Uptime = New-TimeSpan -Start $bootuptime  -End $CurrentDate 

    Update-InstallationStatus

    # Check if Maintenance Window is enabled
    if ($MaintenanceWindow -eq $true) {

        Write-Log -LogLevel Info -LogMessage "Maintenance Window Setting detected"
        if ((Test-MaintenanceWindow) -eq $true) {
            $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date)
            Update-InstallationStatus -AllUpdates $AllUpdates
        }
    }

    elseif ($Uptime.Days -eq 0 -and $uptime.Hours -eq 0 -and $uptime.Minutes -le "15") {
        Write-Log -LogLevel Info -LogMessage "Update installation status after reboot"

        $AllUpdates = New-WindowsUpdateScan -LastScanTime (Get-Date)       
        Update-InstallationStatus -AllUpdates $AllUpdates
        Clear-Variable AllUpdates, Updates
    }
}

if ($AllUpdates) {
    $AllAvailableUpdates = @($AllUpdates | Where-Object { $_.IsInstalled -eq $false -and $_.IsHidden -eq $false })
    Write-Log -LogLevel Info -LogMessage "Pending Update count: $($AllAvailableUpdates.Count)"
    Update-InstallationStatus -AllUpdates $AllUpdates
}    


# run download and installation if updates are available 
if ( ($AllAvailableUpdates.Count -gt 0) -and ($ReportOnly -ne "True") ) {
    $NotDownloadedUpdates = $AllAvailableUpdates | Where-Object IsDownloaded -eq $false
    # If Download is independent of installation - Download all Updates
    If ($DirectDownload -eq $True) {
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
    if ($MaintenanceWindow -eq $true) {
        Write-Log -LogLevel Info -LogMessage "Maintenance Window Setting detected"

        if ((Test-MaintenanceWindow) -eq $true) {
            Write-Log -LogLevel Debug -LogMessage "Device in maintenance window"
            Test-PendingReboot       

            Write-Log -LogLevel Debug -LogMessage "Trigger update download"
            Save-WindowsUpdate -DownloadUpdateList $NotDownloadedUpdates
            

            foreach ($Update in $AllAvailableUpdates) {
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
                Test-PendingReboot    
            }
            else {
                Write-Log -LogLevel Debug -LogMessage "Device outside maintenance window. Skipping reboot"
            }
        }
        else {
            Write-Log -LogLevel Debug -LogMessage "Device outside maintenance window"
        }
    }
    else {
        # If no maintenance window is configured, just install the updates at any time
        Write-Log -LogLevel Info -LogMessage "Starting installation without Maintenance Window"

        Write-Log -LogLevel Debug -LogMessage "Trigger update download"
        Save-WindowsUpdate -DownloadUpdateList $NotDownloadedUpdates

        Clear-Variable Update -Force -ErrorAction SilentlyContinue
       

        foreach ($Update in $AllAvailableUpdates) {
            Write-Log -LogLevel Debug -LogMessage "Installing Update: $($Update.Title)"
            Install-SpecificWindowsUpdate -SelectedUpdate $Update
            Set-ItemProperty -Path "$($RegistryRootPath)\Settings" -Name 'LastInstallationDate' -Value $(Get-Date -Format s)                   
        }
    }
}


$Component = "REBOOT CHECK"
#Check if  Reboot pending   
$RebootRequired = Test-PendingReboot 
Write-Log -LogLevel Info -LogMessage "RebootRequired: $($RebootRequired)"


$Component = "UPDATE STATISTICS"
Write-Log -LogLevel Info -LogMessage "Update statistics"


$LastInstallationDate = Get-Date (Get-ItemPropertyValue -Path "$($RegistryRootPath)\Settings" -Name LastInstallationDate ) -Format s -ErrorAction SilentlyContinue
$NextInstallationDate = Get-Date (Get-ItemPropertyValue -Path "$($RegistryRootPath)\Settings" -Name NextScanTime ) -Format s -ErrorAction SilentlyContinue


if ((Get-Date $LastInstallationDate) -lt (Get-Date $NextInstallationDate)) {
    $AllUpdates = Search-AllUpdates -UpdateSource ($Settings.UpdateSource) -IgnoreHideStatus
    $AllAvailableUpdates = $AllUpdates | Where-Object { $_.IsInstalled -eq $False -and $_.IsHidden -eq $False }
}


#Resetting counter
$UpdateCount = 0
$UpdateRollupsCount = 0
$DefinitionUpdatesCount = 0
$UpgradesCount = 0
$SecurityUpdatesCount = 0
$CriticalUpdatesCount = 0 

if ($AllAvailableUpdates.Count -gt 0) {
    foreach ($update in $AllAvailableUpdates) {
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
      
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Missing Updates" -Value $($AllAvailableUpdates.count) -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Open Pending Updates" -Value "True" -Force | Out-Null

    $AvailableUpdates = @()
    foreach ($KB in ($AllAvailableUpdates.KBArticleIDs)) {
        $AvailableUpdates += "KB$($KB)"
    }
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Available Update KBs" -Value $AvailableUpdates -Force | Out-Null
}
else {
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Missing Updates" -Value 0 -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Open Pending Updates" -Value "False" -Force | Out-Null
    New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Available Update KBs" -Value "none" -Force | Out-Null
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
$InstalledUpdates = Get-InstalledWindowsUpdates -AllUpdates $AllUpdates
Write-Log -LogLevel Debug -LogMessage "Write installed update list in registry"
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Installed KBs" -Value $InstalledUpdates -Force | Out-Null
New-ItemProperty -Path "$($RegistryRootPath)\Status" -PropertyType "String" -Name "Total Installed KBs" -Value $($InstalledUpdates.Count) -Force | Out-Null

#Update Delivery Optimization statistics 
Update-DeliveryOptimizationStats

Write-Log -LogLevel Info -LogMessage "###############################################################################" 
