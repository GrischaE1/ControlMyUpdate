
##########################################################################################
#                                    Param 
#
#define the log folder if needed - otherwise delete it or set it to ""
param(
		[string]$LogPath,
        [bool]$UpdateScheduledTask,
        [int]$UpdateInterval = '30'
	)

$InstallDir = "C:\Windows\ControlMyUpdate"

#Check if scheduled task is created
$TaskCheck = Get-ScheduledTask -TaskName "Control My Update" -ErrorAction SilentlyContinue

If(!$TaskCheck -or $UpdateScheduledTask -eq $true)
{
    #Create action
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -file ""$($InstallDir)\ControlMyUpdate.ps1"" -LogPath ""$($LogPath)""" 

    #create triggers
    $starttime=(get-date)
    $TimeSpan = New-TimeSpan -Minutes $UpdateInterval

    $trigger = @()
    $trigger += New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval $TimeSpan
    $trigger += New-ScheduledTaskTrigger -AtStartup

    #Use System account
    $User = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

    #Settings to make sure the task can run every time
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -DontStopOnIdleEnd -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    #Create the new scheduled task
    Register-ScheduledTask -TaskName "Control My Update" -Action $action -Trigger $trigger -Principal $user  -Settings $settings -Force
}


#Check if scheduled task is created
$RebootTaskCheck = Get-ScheduledTask -TaskName "Control My Update - Reboot Notification" -ErrorAction SilentlyContinue

If(!$RebootTaskCheck -or $UpdateScheduledTask -eq $true)
{
    #Create action
    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "C:\Windows\ControlMyUpdate\HiddenPowerShell.vbs" 

    #Use System account
    $User = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" 

    #Settings to make sure the task can run every time
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries 

    #Create the new scheduled task
    Register-ScheduledTask -TaskName "Control My Update - Reboot Notification" -Action $action -Principal $user  -Settings $settings -Force
}


New-Item -Path $InstallDir -ItemType Directory -Force
Copy-Item "$PSScriptRoot\ControlMyUpdate.ps1" -Destination $InstallDir -Force
Copy-Item "$PSScriptRoot\ConnectionCheck.csv" -Destination $InstallDir -Force
Copy-Item "$PSScriptRoot\HiddenPowerShell.vbs" -Destination $InstallDir -Force
Copy-Item "$PSScriptRoot\Reboot_Notification.ps1" -Destination $InstallDir -Force
Copy-Item "$PSScriptRoot\RestartScript.cmd" -Destination $InstallDir -Force