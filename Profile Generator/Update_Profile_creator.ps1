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
# Name: Update_Profile_Creator.ps1
# Version: 2.1.1
# Date: 25.02.2022
# Created by: Grischa Ernst gernst@vmware.com
#
# Description
# - This GUI will create three different profiles: 
# 	1. Windows Update - Fully supported settings from 20H2 and above incl. Windows 11
# 	2. Control My Update - New GUI + added some more features 
# 	3. Delivery Optimization - Fully supported settings from 20H2 and above incl. Windows 11
#
# Caution:
# If you are using the Control My Update, you need Version 2.1 or higher
##########################################################################################

##########################################################################################
#                                    Changelog 
#
# 2.1.1 - Added 22H2 support
#		- Added new "Updates automatically download and install at an optimal time determined by the device" installation method
# 2.1 - added the following features:
#       - new GUI settings for CMU
#		- Bug fix MW day output 
#		- Bug fix legacy profile
#		- new settings for updated CMU version
# 2.0 - Initial creation
##########################################################################################



#-------------------------------------------------------------#
#----Initial Declarations-------------------------------------#
#-------------------------------------------------------------#

Add-Type -AssemblyName PresentationCore, PresentationFramework



#-------------------------------------------------------------#
#----Control Event Handlers-----------------------------------#
#-------------------------------------------------------------#

. "$PSScriptRoot\resources\navigationfunctions.ps1"
. "$PSScriptRoot\resources\otherGUIfunctions.ps1"


#-------------------------------------------------------------#
#----Script Execution-----------------------------------------#
#-------------------------------------------------------------#

$Xaml = Get-Content "$($PSScriptRoot)\resources\layout.xml"

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml
#[xml]$xml = Get-Content C:\Users\Grisc\GIT\Development\Development\layout.xml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

#-------------------------------------------------------------#
#----Navigation-----------------------------------------------#
#-------------------------------------------------------------#
# Side Nav bindings
$Tab1BT.Add_Click( { Tab1Click $this $_ })
$Tab2BT.Add_Click( { Tab2Click $this $_ })
$Tab3BT.Add_Click( { Tab3Click $this $_ })
$TabWSUSBT.Add_Click( { TabWSUSClick $this $_ })


# Side Nav bindings
$CMUTab1BT.Add_Click( { CMUTab1Click $this $_ })
$CMUTab2BT.Add_Click( { CMUTab2Click $this $_ })
$CMUTab3BT.Add_Click( { CMUTab3Click $this $_ })
$CMUTab4BT.Add_Click( { CMUTab4Click $this $_ })
$CMUTab5BT.Add_Click( { CMUTab5Click $this $_ })



#Next and previous button bindings
#Next
$NextTab1.Add_Click( { NextTab1Click $this $_ })
$NextTab2.Add_Click( { NextTab2Click $this $_ })
$NextTab3.Add_Click( { NextTab3Click $this $_ })
$NextTabWSUS.Add_Click( { NextTabWSUSClick $this $_ })


$CMUNextTab1.Add_Click( { CMUNextTab1Click $this $_ })
$CMUNextTab2.Add_Click( { CMUNextTab2Click $this $_ })
$CMUNextTab3.Add_Click( { CMUNextTab3Click $this $_ })
$CMUNextTab4.Add_Click( { CMUNextTab4Click $this $_ })

#Previous
$PreviousTab2.Add_Click( { PreviousTab2Click $this $_ })
$PreviousTab3.Add_Click( { PreviousTab3Click $this $_ })
$PreviousTabWSUS.Add_Click( { PreviousTabWSUSClick $this $_ })


$CMUPreviousTab2.Add_Click( { CMUPreviousTab2Click $this $_ })
$CMUPreviousTab3.Add_Click( { CMUPreviousTab3Click $this $_ })
$CMUPreviousTab4.Add_Click( { CMUPreviousTab4Click $this $_ })
$CMUPreviousTab5.Add_Click( { CMUPreviousTab5Click $this $_ })

#-------------------------------------------------------------#
#---General---------------------------------------------------#
#-------------------------------------------------------------#

$ToolCheckBox.Add_Checked({ $GenerateCustomProfiles.IsChecked = $False; $GenerateCustomScript.IsChecked = $true; $TABCustomToolConfig.IsEnabled = $true })
$ToolCheckBox.Add_UnChecked({ $GenerateCustomProfiles.IsChecked = $false; $TABCustomToolConfig.IsEnabled = $false })



#-------------------------------------------------------------#
#----Timing---------------------------------------------------#
#-------------------------------------------------------------#

#-------------------------------------------------------------#
#----Timing---------------------------------------------------#
#-------------------------------------------------------------#

#-------------------------------------------------------------#
#----Releases-------------------------------------------------#
#-------------------------------------------------------------#

#-------------------------------------------------------------#
#----WSUS-----------------------------------------------------#
#-------------------------------------------------------------#


#-------------------------------------------------------------#
#----Generate-------------------------------------------------#
#-------------------------------------------------------------#

$GenerateButton.Add_Click({
		Add-Type -AssemblyName System.Windows.Forms
		$FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{SelectedPath = $PSScriptRoot; ShowNewFolderButton = $false }
		# [void]$FolderBrowser.ShowDialog()

		if ($FolderBrowser.ShowDialog() -eq "OK") {
			if ($GenerateWUProfiles.IsChecked -eq $true) {
				for ($i = 1; $i -le $Ringslider.Value; $i++) {
					$WUProfileParams = @{}	
					$WUProfileParams["TargetPath"] = $FolderBrowser.SelectedPath
					$WUProfileParams["Ringcount"] = $i

					###########################################################################
					# General 

					#$DriversCheckBox 
					if ($DriversCheckBox.IsChecked -eq $true) { $DriversCheckBoxValue = 1 }else { $DriversCheckBoxValue = 0 }
					$WUProfileParams["DriversCheckBox"] = $DriversCheckBoxValue

					#$TelemetryCheckBox 
					if ($TelemetryCheckBox.IsChecked -eq $true) { $TelemetryCheckValue = 1 }
					$WUProfileParams["TelemetryCheckBox"] = $TelemetryCheckValue
				
					#Update Config
					$WUProfileParams["AutoUpdateConfig"] = $AutoUpdateConfig.SelectionBoxItem

					#disable UX Access if tool is used
					if ($ToolCheckBox.IsChecked -eq $true) {
						$WUProfileParams["SetDisableUXWUAccess"] = 1
					}

					###########################################################################
					# Timing
				
					if ($TabWSUSBT.IsEnabled -eq $False) {
						#Cumulative Update deferral
						$CurrentCUDeferral = $CUSlider.Value * $i
						$WUProfileParams["CUSlider"] = $CurrentCUDeferral
					
						#Feature Update deferral
						$CurrentFUDeferral = $FUSlider.Value * $i
						$WUProfileParams["FUSlider"] = $CurrentFUDeferral
					}

					#Restart config
					$WUProfileParams["RestartSlider"] = $RestartSlider.Value
					$WUProfileParams["RestartGraceSlider"] = $RestartGraceSlider.Value
				

					#Update Interval config
					if ($TabWSUSBT.IsEnabled -eq $True) {
						$WUProfileParams["UpdateIntervalSlider"] = $UpdateIntervalSlider.Value
					}

					###########################################################################
					# Releases
				
					#Windows branch config
					if ($RingSlider.Value -gt '3' -and $i -eq '1') {
						$WUProfileParams["WindowsChannel"] = 'Windows Insider build - Fast'
					}
					else { $WUProfileParams["WindowsChannel"] = $WindowsChannel.SelectionBoxItem }			

					#Update Safeguards
					$WUProfileParams["SafeguardCheckBox"] = $SafeguardCheckBox.IsChecked

					#Product Version
					if ($ProductVersionSelection.SelectionBoxItem -ne $Null) {
						$WUProfileParams["ProductVersionSelection"] = $ProductVersionSelection.SelectionBoxItem

						#Release Version
						if ($ProductVersionSelection.SelectionBoxItem -like "*Windows 10") {
							$WUProfileParams["ReleaseVersion"] = $ReleaseVersionW10Box.SelectionBoxItem
						}
						if ($ProductVersionSelection.SelectionBoxItem -like "*Windows 11") {
							$WUProfileParams["ReleaseVersion"] = $ReleaseVersionW11Box.SelectionBoxItem
						}
					}


					###########################################################################
					# WSUS
					if ($TabWSUSBT.IsEnabled -eq $True) {
						#WSUS Server URL
						$WUProfileParams["WSUSURL"] = $WSUSURL.Text
						$WUProfileParams["WSUSaltURL"] = $WSUSaltURL.Text

						#Allow 3rd Party updates
						$WUProfileParams["AllowNonMicrosoftSignedUpdate"] = $AllowNonMicrosoftSignedUpdate.IsChecked

						#Enable TLS
						$WUProfileParams["UseTLSpinning"] = $UseTLSpinning.IsChecked
						$WUProfileParams["AllowNonMSUpdates"] = $AllowNonMSUpdates.IsChecked
						$WUProfileParams["DualScanSelection"] = $DualScanSelection.SelectedItems.text
					}
				
					#Generate the profile
					New-WUProfile @WUProfileParams

				}
			}
			if ($GenerateCustomProfiles.IsChecked -eq $true) {
				#Generate function variables
				$CUProfileParams = @{}  
				$CUProfileParams["TargetPath"] = $FolderBrowser.SelectedPath

				if ($ReportOnly.IsChecked -eq "True") {
					$CUProfileParams["ReportOnly"] = $ReportOnly.IsChecked
				}
				else {
					$CUProfileParams["ToolUpdateSource"] = $ToolUpdateSource.SelectionBoxItem
					$CUProfileParams["DirectDownload"] = $DirectDownload.IsChecked
					$CUProfileParams["EnablemaintenanceWindow"] = $EnablemaintenanceWindow.IsChecked
					$CUProfileParams["EnablePerDayMW"] = $EnablePerDayMW.IsChecked
					if($EnablePerDayMW.IsChecked)
					{
						$CUProfileParams["MWPerDayMondayStartTime"] = $MWPerDayMondayStartTime.text
						$CUProfileParams["MWPerDayMondayEndTime"] = $MWPerDayMondayEndTime.text
						$CUProfileParams["MWPerDayTuesdayStartTime"] = $MWPerDayTuesdayStartTime.text
						$CUProfileParams["MWPerDayTuesdayEndTime"] = $MWPerDayTuesdayEndTime.text
						$CUProfileParams["MWPerDayWednesdayStartTime"] = $MWPerDayWednesdayStartTime.text
						$CUProfileParams["MWPerDayWednesdayEndTime"] = $MWPerDayWednesdayEndTime.text
						$CUProfileParams["MWPerDayThursdayStartTime"] = $MWPerDayThursdayStartTime.text
						$CUProfileParams["MWPerDayThursdayEndTime"] = $MWPerDayThursdayEndTime.text
						$CUProfileParams["MWPerDayFridayStartTime"] = $MWPerDayFridayStartTime.text
						$CUProfileParams["MWPerDayFridayEndTime"] = $MWPerDayFridayEndTime.text
						$CUProfileParams["MWPerDaySaturdayStartTime"] = $MWPerDaySaturdayStartTime.text
						$CUProfileParams["MWPerDaySaturdayEndTime"] = $MWPerDaySaturdayEndTime.text
						$CUProfileParams["MWPerDaySundayStartTime"] = $MWPerDaySundayStartTime.text
						$CUProfileParams["MWPerDaySundayEndTime"] = $MWPerDaySundayEndTime.text
					}
					else
					{
						$CUProfileParams["MWStartTime"] = $MWStartTime.text
						$CUProfileParams["MWEndTime"] = $MWEndTime.text
						$CUProfileParams["MWDay"] = $MWDays.SelectedItems.text
					}
					
					$CUProfileParams["EmergencyKB"] = $EmergencyKB.text
					$CUProfileParams["BlockedKBs"] = $BlockedKBs.text
					$CUProfileParams["UnBlockedKBs"] = $UnBlockedKBs.text
					$CUProfileParams["ScanInterval"] = $CMUUpdateIntervalSlider.Value
					$CUProfileParams["ScanRandomization"] = $ScanRandomizationSlider.Value
					$CUProfileParams["NotifyUser"] = $RebootNotificationCheckBox.IsChecked
					$CUProfileParams["ToastTitle"] = $ToastTitle.text
					$CUProfileParams["ToastMessage"] = $ToastMessage.text
					$CUProfileParams["ToastAdvice"] = $ToastAdvise.text
					$CUProfileParams["NoMWAutoRebootInterval"] = $CMUAutoRebootIntervalSlider_NoMW.Value
					$CUProfileParams["NoMWAutomaticReboot"] = $NoMWAutomaticRebootCheckBox.IsChecked
					$CUProfileParams["ForceRebootwithNoUser"] = $ForceRebootwithNoUser.IsChecked
					$CUProfileParams["UninstallKBs"] = $UninstallKBsCheckBox.IsChecked
					$CUProfileParams["RunConnectionTests"] = $RunConnectionTests.IsChecked
					$CUProfileParams["RetryCount"] = $RetryCountSlider.Value
					$CUProfileParams["CMUCategories_SelectAll"] = $CMUCategories_SelectAll.IsChecked
					$CUProfileParams["CMUCategories_Application"] = $CMUCategories_Application.IsChecked
					$CUProfileParams["CMUCategories_Connectors"] = $CMUCategories_Connectors.IsChecked
					$CUProfileParams["CMUCategories_CriticalUpdates"] = $CMUCategories_CriticalUpdates.IsChecked
					$CUProfileParams["CMUCategories_DefinitionUpdates"] = $CMUCategories_DefinitionUpdates.IsChecked
					$CUProfileParams["CMUCategories_DeveloperKits"] = $CMUCategories_DeveloperKits.IsChecked
					$CUProfileParams["CMUCategories_FeaturePacks"] = $CMUCategories_FeaturePacks.IsChecked
					$CUProfileParams["CMUCategories_Guidance"] = $CMUCategories_Guidance.IsChecked
					$CUProfileParams["CMUCategories_ServicePacks"] = $CMUCategories_ServicePacks.IsChecked
					$CUProfileParams["CMUCategories_Tools"] = $CMUCategories_Tools.IsChecked
					$CUProfileParams["CMUCategories_UpdateRollups"] = $CMUCategories_UpdateRollups.IsChecked
					$CUProfileParams["CMUCategories_Updates"] = $CMUCategories_Updates.IsChecked
					$CUProfileParams["CMUCategories_SecurityUpdates"] = $CMUCategories_Updates.IsChecked
					$CUProfileParams["CMUCategories_Drivers"] = $CMUCategories_Drivers.IsChecked

				}

				New-CustomUpdateProfile @CUProfileParams
			}

			if ($GenerateCustomScript.IsChecked -eq $true) {
				#Generate function variables
				$CUScriptParams = @{}  
				$CUScriptParams["TargetPath"] = $FolderBrowser.SelectedPath

				if ($ReportOnly.IsChecked -eq "True") {
					$CUScriptParams["ReportOnly"] = $ReportOnly.IsChecked
				}
				else {
					$CUScriptParams["ToolUpdateSource"] = $ToolUpdateSource.SelectionBoxItem
					$CUScriptParams["DirectDownload"] = $DirectDownload.IsChecked
					$CUScriptParams["EnablemaintenanceWindow"] = $EnablemaintenanceWindow.IsChecked
					$CUProfileParams["EnablePerDayMW"] = $EnablePerDayMW.IsChecked
					if($EnablePerDayMW.IsChecked)
					{
						$CUScriptParams["MWPerDayMondayStartTime"] = $MWPerDayMondayStartTime.text
						$CUScriptParams["MWPerDayMondayEndTime"] = $MWPerDayMondayEndTime.text
						$CUScriptParams["MWPerDayTuesdayStartTime"] = $MWPerDayTuesdayStartTime.text
						$CUScriptParams["MWPerDayTuesdayEndTime"] = $MWPerDayTuesdayEndTime.text
						$CUScriptParams["MWPerDayWednesdayStartTime"] = $MWPerDayWednesdayStartTime.text
						$CUScriptParams["MWPerDayWednesdayEndTime"] = $MWPerDayWednesdayEndTime.text
						$CUScriptParams["MWPerDayThursdayStartTime"] = $MWPerDayThursdayStartTime.text
						$CUScriptParams["MWPerDayThursdayEndTime"] = $MWPerDayThursdayEndTime.text
						$CUScriptParams["MWPerDayFridayStartTime"] = $MWPerDayFridayStartTime.text
						$CUScriptParams["MWPerDayFridayEndTime"] = $MWPerDayFridayEndTime.text
						$CUScriptParams["MWPerDaySaturdayStartTime"] = $MWPerDaySaturdayStartTime.text
						$CUScriptParams["MWPerDaySaturdayEndTime"] = $MWPerDaySaturdayEndTime.text
						$CUScriptParams["MWPerDaySundayStartTime"] = $MWPerDaySundayStartTime.text
						$CUScriptParams["MWPerDaySundayEndTime"] = $MWPerDaySundayEndTime.text
					}
					else
					{
						$CUScriptParams["MWStartTime"] = $MWStartTime.text
						$CUScriptParams["MWEndTime"] = $MWEndTime.text
						$CUScriptParams["MWDay"] = $MWDays.SelectedItems.text
					}
					$CUScriptParams["EmergencyKB"] = $EmergencyKB.text
					$CUScriptParams["BlockedKBs"] = $BlockedKBs.text
					$CUScriptParams["UnBlockedKBs"] = $UnBlockedKBs.text
					$CUScriptParams["ScanInterval"] = $CMUUpdateIntervalSlider.Value
					$CUScriptParams["ScanRandomization"] = $ScanRandomizationSlider.Value
					$CUScriptParams["NotifyUser"] = $RebootNotificationCheckBox.IsChecked
					$CUScriptParams["ToastTitle"] = $ToastTitle.text
					$CUScriptParams["ToastMessage"] = $ToastMessage.text
					$CUScriptParams["ToastAdvice"] = $ToastAdvise.text
					$CUScriptParams["NoMWAutomaticReboot"] = $NoMWAutomaticRebootCheckBox.IsChecked
					$CUScriptParams["NoMWAutoRebootInterval"] = $CMUAutoRebootIntervalSlider.Value
					$CUScriptParams["MWAutomaticReboot"] = $MWAutomaticRebootCheckBox.IsChecked
					$CUScriptParams["MWAutoRebootInterval"] = $CMUAutoRebootIntervalTextBox_MW.Value
					$CUScriptParams["ForceRebootwithNoUser"] = $ForceRebootwithNoUser.IsChecked
					$CUScriptParams["UninstallKBs"] = $UninstallKBsCheckBox.IsChecked
					$CUScriptParams["RunConnectionTests"] = $RunConnectionTests.IsChecked
					$CUScriptParams["RetryCount"] = $RetryCountSlider.Value
					$CUScriptParams["CMUCategories_SelectAll"] = $CMUCategories_SelectAll.IsChecked
					$CUScriptParams["CMUCategories_Application"] = $CMUCategories_Application.IsChecked
					$CUScriptParams["CMUCategories_Connectors"] = $CMUCategories_Connectors.IsChecked
					$CUScriptParams["CMUCategories_CriticalUpdates"] = $CMUCategories_CriticalUpdates.IsChecked
					$CUScriptParams["CMUCategories_DefinitionUpdates"] = $CMUCategories_DefinitionUpdates.IsChecked
					$CUScriptParams["CMUCategories_DeveloperKits"] = $CMUCategories_DeveloperKits.IsChecked
					$CUScriptParams["CMUCategories_FeaturePacks"] = $CMUCategories_FeaturePacks.IsChecked
					$CUScriptParams["CMUCategories_Guidance"] = $CMUCategories_Guidance.IsChecked
					$CUScriptParams["CMUCategories_ServicePacks"] = $CMUCategories_ServicePacks.IsChecked
					$CUScriptParams["CMUCategories_Tools"] = $CMUCategories_Tools.IsChecked
					$CUScriptParams["CMUCategories_UpdateRollups"] = $CMUCategories_UpdateRollups.IsChecked
					$CUScriptParams["CMUCategories_Updates"] = $CMUCategories_Updates.IsChecked
					$CUScriptParams["CMUCategories_SecurityUpdates"] = $CMUCategories_Updates.IsChecked
					$CUScriptParams["CMUCategories_Drivers"] = $CMUCategories_Drivers.IsChecked
				}

				New-CustomUpdateScript @CUScriptParams
			}

			#########Delivery Optimization
			if ($GenerateDOProfile.IsChecked -eq $true) {
				#Generate function variables
				$DOProfileParams = @{}  
				$DOProfileParams["TargetPath"] = $FolderBrowser.SelectedPath

				switch ($DODownloadMode.SelectionBoxItem) {
					'HTTP only, no peering' { $DODownloadModeValue = 0 }
					'HTTP blended with peering behind the same NAT' { $DODownloadModeValue = 1 }
					'HTTP blended with peering across a private group' { $DODownloadModeValue = 2 }
					'HTTP blended with Internet peering' { $DODownloadModeValue = 3 }
					'Simple download mode with no peering' { $DODownloadModeValue = 4 }
					Default { $DODownloadModeValue = 1 }
				}
				$DOProfileParams["DODownloadMode"] = $DODownloadModeValue
			
				#Add Group Identifier
				if ($DODownloadModeValue -eq '2') {
					if ($DOAutomaticGUID.IsChecked -eq $false) {
						$DOProfileParams["DOGroupID"] = $DOGroupID.Text
					}
					if ($DOAutomaticGUID.IsChecked -eq $true) {
						$DOProfileParams["DOGroupIDSource"] = $DOGroupIDSource.SelectionBoxItem
					}
				}

				$DOProfileParams["DOMinFileSizeToCache"] = $DOMinFileSizeToCache.Text
				$DOProfileParams["DOMaxDownloadBandwidth"] = $DOMaxDownloadBandwidth.Text

				if ($DOUseCacheHost.IsChecked -eq $true) {
					if ($DOUseDHCPCacheHost.IsChecked -eq $true) {
						$DOProfileParams["DOCacheHostSource"] = $DOCacheHostSource.SelectionBoxItem	
					}
					else { $DOProfileParams["DOCacheHost"] = $DOCacheHost.Text }			
				}
			

				New-DOProfile @DOProfileParams 

			}

			#open the Target folder
			Invoke-Item $($FolderBrowser.SelectedPath)
		}
	})
#-------------------------------------------------------------#
#----Control My Update----------------------------------------#
#-------------------------------------------------------------#

$GenerateCustomScriptProfile.Add_Click( { $MenuNavigation.SelectedItem = $TABGenerate })

$EnableMaintenanceWindow.Add_Checked({ $EnablePerDayMW.IsEnabled = $true; $MWAutomaticRebootCheckBox.IsEnabled = $True;  $CMUAutoRebootIntervalTextBox_MW.IsEnabled = $true; $CMUAutoRebootIntervalSlider_MW.IsEnabled = $True; $NoMWAutomaticRebootCheckBox.IsEnabled = $false; $CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $false; $CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $false; $MWStartTime.IsEnabled = $true; $MWEndTime.IsEnabled = $true; $MWDays.IsEnabled = $true; $StartTimeHelpText.Visibility = "Visible"; $StopTimeHelpText.Visibility = "Visible"; $MWAutomaticRebootCheckBox.IsEnabled = $true })
$EnableMaintenanceWindow.Add_UnChecked({ $EnablePerDayMW.IsEnabled = $false; $MWAutomaticRebootCheckBox.IsEnabled = $False;  $CMUAutoRebootIntervalTextBox_MW.IsEnabled = $False; $CMUAutoRebootIntervalSlider_MW.IsEnabled = $False; $NoMWAutomaticRebootCheckBox.IsEnabled = $true; $CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $true; $CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $true; $MWStartTime.IsEnabled = $false; $MWEndTime.IsEnabled = $false; $MWDays.IsEnabled = $false; $StartTimeHelpText.Visibility = "Hidden"; $StopTimeHelpText.Visibility = "Hidden"; $MWAutomaticRebootCheckBox.IsEnabled = $false })


$EnablePerDayMW.Add_UnChecked({
	$PerDayStartTimeHelpText.Visibility = "Hidden"
	$PerDayStartTimeText.Visibility = "Hidden"
	$PerDayStopTimeText.Visibility = "Hidden"
	$PerDayMonday.Visibility = "Hidden"
	$MWPerDayMondayStartTime.Visibility = "Hidden"
	$MWPerDayMondayEndTime.Visibility = "Hidden"
	$PerDayTuesday.Visibility = "Hidden"
	$MWPerDayTuesdayStartTime.Visibility = "Hidden"
	$MWPerDayTuesdayEndTime.Visibility = "Hidden"
	$PerDayWednesday.Visibility = "Hidden"
	$MWPerDayWednesdayStartTime.Visibility = "Hidden"
	$MWPerDayWednesdayEndTime.Visibility = "Hidden"
	$PerDayThursday.Visibility = "Hidden"
	$MWPerDayThursdayStartTime.Visibility = "Hidden"
	$MWPerDayThursdayEndTime.Visibility = "Hidden"
	$PerDayFriday.Visibility = "Hidden"
	$MWPerDayFridayStartTime.Visibility = "Hidden"
	$MWPerDayFridayEndTime.Visibility = "Hidden"
	$PerDaySaturday.Visibility = "Hidden"
	$MWPerDaySaturdayStartTime.Visibility = "Hidden"
	$MWPerDaySaturdayEndTime.Visibility = "Hidden"
	$PerDaySunday.Visibility = "Hidden"
	$MWPerDaySundayStartTime.Visibility = "Hidden"
	$MWPerDaySundayEndTime.Visibility = "Hidden"
	$MWSimpleDaysText.Visibility = "Visible"; $MWSimpleEndText.Visibility = "Visible"; $MWSimpleStartText.Visibility = "Visible"; $MWAutomaticRebootCheckBox.Visibility = "Visible";  $CMUAutoRebootIntervalTextBox_MW.Visibility = "Visible"; $CMUAutoRebootIntervalSlider_MW.Visibility = "Visible"; $NoMWAutomaticRebootCheckBox.IsEnabled = $false; $CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $false; $CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $false; $MWStartTime.Visibility = "Visible"; $MWEndTime.Visibility = "Visible"; $MWDays.Visibility = "Visible"; $StartTimeHelpText.Visibility = "Visible"; $StopTimeHelpText.Visibility = "Visible"; $MWAutomaticRebootCheckBox.Visibility = "Visible"
})

$EnablePerDayMW.Add_Checked({
	$PerDayStartTimeHelpText.Visibility = "Visible"
	$PerDayStartTimeText.Visibility = "Visible"
	$PerDayStopTimeText.Visibility = "Visible"
	$PerDayMonday.Visibility = "Visible"
	$MWPerDayMondayStartTime.Visibility = "Visible"
	$MWPerDayMondayEndTime.Visibility = "Visible"
	$PerDayTuesday.Visibility = "Visible"
	$MWPerDayTuesdayStartTime.Visibility = "Visible"
	$MWPerDayTuesdayEndTime.Visibility = "Visible"
	$PerDayWednesday.Visibility = "Visible"
	$MWPerDayWednesdayStartTime.Visibility = "Visible"
	$MWPerDayWednesdayEndTime.Visibility = "Visible"
	$PerDayThursday.Visibility = "Visible"
	$MWPerDayThursdayStartTime.Visibility = "Visible"
	$MWPerDayThursdayEndTime.Visibility = "Visible"
	$PerDayFriday.Visibility = "Visible"
	$MWPerDayFridayStartTime.Visibility = "Visible"
	$MWPerDayFridayEndTime.Visibility = "Visible"
	$PerDaySaturday.Visibility = "Visible"
	$MWPerDaySaturdayStartTime.Visibility = "Visible"
	$MWPerDaySaturdayEndTime.Visibility = "Visible"
	$PerDaySunday.Visibility = "Visible"
	$MWPerDaySundayStartTime.Visibility = "Visible"
	$MWPerDaySundayEndTime.Visibility = "Visible"
	$MWSimpleDaysText.Visibility = "Hidden"; $MWSimpleEndText.Visibility = "Hidden"; $MWSimpleStartText.Visibility = "Hidden"; $MWAutomaticRebootCheckBox.Visibility = "Hidden";  $CMUAutoRebootIntervalTextBox_MW.Visibility = "Hidden"; $CMUAutoRebootIntervalSlider_MW.Visibility = "Hidden"; $NoMWAutomaticRebootCheckBox.IsEnabled = $false; $CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $false; $CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $false; $MWStartTime.Visibility = "Hidden"; $MWEndTime.Visibility = "Hidden"; $MWDays.Visibility = "Hidden"; $StartTimeHelpText.Visibility = "Hidden"; $StopTimeHelpText.Visibility = "Hidden"; $MWAutomaticRebootCheckBox.Visibility = "Hidden"
})



$NoMWAutomaticRebootCheckBox.Add_Checked({

		$CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $true
		$CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $true
	})

$NoMWAutomaticRebootCheckBox.Add_UnChecked({

		$CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $false
		$CMUAutoRebootIntervalTextBox_NoMW.IsEnabled = $false
	})

$MWAutomaticRebootCheckBox.Add_Checked({

		$CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $true
		$CMUAutoRebootIntervalSlider_MW.IsEnabled = $true
	})

$MWAutomaticRebootCheckBox.Add_UnChecked({

		$CMUAutoRebootIntervalSlider_NoMW.IsEnabled = $false
		$CMUAutoRebootIntervalSlider_MW.IsEnabled = $false
	})

$ReportOnly.Add_Checked({
		$ToolUpdateSource.IsEnabled = $false
		$DirectDownload.IsEnabled = $false
		$EnableMaintenanceWindow.IsEnabled = $false
		$MWStartTime.IsEnabled = $false
		$MWEndTime.IsEnabled = $false
		$MWDays.IsEnabled = $false
		$BlockedKBs.IsEnabled = $false
		$UnBlockedKBs.IsEnabled = $false
		$ScanRandomizationSlider.IsEnabled = $false
		$ScanRandomizationTextBox.IsEnabled = $false
		$EmergencyKB.IsEnabled = $false
		$RebootNotificationCheckBox.IsEnabled = $false
		$ToastTitle.IsEnabled = $false
		$ToastText.IsEnabled = $false
		$CMUTab2BT.IsEnabled = $false
		$CMUTab2BT.Foreground = "#5b5a5c"
		$CMUTab3BT.IsEnabled = $false
		$CMUTab3BT.Foreground = "#5b5a5c"
		$CMUTab4BT.IsEnabled = $false
		$CMUTab4BT.Foreground = "#5b5a5c"
	})

$ReportOnly.Add_UnChecked({
		$CMUTab2BT.IsEnabled = $true
		$CMUTab3BT.IsEnabled = $true
		$CMUTab4BT.IsEnabled = $true
		$ToolUpdateSource.IsEnabled = $true
		$DirectDownload.IsEnabled = $true
		$EnableMaintenanceWindow.IsEnabled = $true
		$BlockedKBs.IsEnabled = $true
		$UnBlockedKBs.IsEnabled = $true
		$ScanRandomizationSlider.IsEnabled = $true
		$ScanRandomizationTextBox.IsEnabled = $true
		$EmergencyKB.IsEnabled = $true
		$RebootNotificationCheckBox.IsEnabled = $true
		$ToastTitle.IsEnabled = $true
		$ToastText.IsEnabled = $true
		$CMUTab2BT.IsEnabled = $true
		$CMUTab2BT.Foreground = "#ffffff"
		$CMUTab3BT.IsEnabled = $true
		$CMUTab3BT.Foreground = "#ffffff"
		$CMUTab4BT.IsEnabled = $true
		$CMUTab4BT.Foreground = "#ffffff"

	})

$RebootNotificationCheckBox.Add_UnChecked({
		$ToastMessage.IsEnabled = $false
		$ToastTitle.IsEnabled = $false
		$ToastAdvise.IsEnabled = $false
		$ToastMessageTextBox.IsEnabled = $false
		$ToastAdviseTextBox.IsEnabled = $false
		$ToastTitleTextBox.IsEnabled  = $false
	})

$RebootNotificationCheckBox.Add_Checked({
		$ToastMessage.IsEnabled = $True
		$ToastTitle.IsEnabled = $True
		$ToastAdvise.IsEnabled = $True
		$ToastMessageTextBox.IsEnabled = $True
		$ToastAdviseTextBox.IsEnabled = $True
		$ToastTitleTextBox.IsEnabled  = $True
	})


$CMUCategories_SelectAll.Add_Checked({ $CMUCategories_SecurityUpdates.IsEnabled = $False; $CMUCategories_Application.IsEnabled = $False;  $CMUCategories_Connectors.IsEnabled = $False; $CMUCategories_CriticalUpdates.IsEnabled = $false; $CMUCategories_DefinitionUpdates.IsEnabled = $false; $CMUCategories_DeveloperKits.IsEnabled = $false; $CMUCategories_Guidance.IsEnabled = $false; $CMUCategories_FeaturePacks.IsEnabled = $false; $CMUCategories_ServicePacks.IsEnabled = $false; $CMUCategories_Tools.IsEnabled = $false; $CMUCategories_UpdateRollups.IsEnabled = $false; $CMUCategories_Updates.IsEnabled = $false })
$CMUCategories_SelectAll.Add_UnChecked({ $CMUCategories_SecurityUpdates.IsEnabled = $true; $CMUCategories_Application.IsEnabled = $True;  $CMUCategories_Connectors.IsEnabled = $True; $CMUCategories_CriticalUpdates.IsEnabled = $True; $CMUCategories_DefinitionUpdates.IsEnabled = $True; $CMUCategories_DeveloperKits.IsEnabled = $True; $CMUCategories_Guidance.IsEnabled = $True; $CMUCategories_FeaturePacks.IsEnabled = $True; $CMUCategories_ServicePacks.IsEnabled = $True; $CMUCategories_Tools.IsEnabled = $True; $CMUCategories_UpdateRollups.IsEnabled = $True; $CMUCategories_Updates.IsEnabled = $True })



#-------------------------------------------------------------#
#----Delivery Optimization------------------------------------#
#-------------------------------------------------------------#

$GenerateDOButton.Add_Click( { $MenuNavigation.SelectedItem = $TABGenerate })

$DODownloadMode.Add_SelectionChanged({
		if ($DODownloadMode.SelectedItem -like "*peering across a private group") {
			$DOGroupIDLabel.Visibility = "Visible"; $DOGroupID.Visibility = "Visible"; $DOGenerateGUID.Visibility = "Visible"; $DOAutomaticGUID.Visibility = "Visible"; $DOAutomaticGUIDLabel.Visibility = "Visible"
		}
		else { $DOGroupIDLabel.Visibility = "Hidden"; $DOGroupID.Visibility = "Hidden"; $DOGenerateGUID.Visibility = "Hidden"; $DOAutomaticGUID.Visibility = "Hidden"; $DOAutomaticGUIDLabel.Visibility = "Hidden"; $DOGroupIDSource.Visibility = "Hidden"; $DOGroupIDSourceLabel.Visibility = "Hidden"; $DOAutomaticGUID.IsChecked = $false }
	})

$DOGenerateGUID.Add_Click({ $DOGroupID.Text = [guid]::NewGuid() })

$DOAutomaticGUID.Add_Checked({ $DOGroupIDSource.Visibility = "Visible"; $DOGroupIDSourceLabel.Visibility = "Visible"; $DOGroupIDLabel.Visibility = "Hidden"; $DOGroupID.Visibility = "Hidden"; $DOGenerateGUID.Visibility = "Hidden" })
$DOAutomaticGUID.Add_UnChecked({ $DOGroupIDSource.Visibility = "Hidden"; $DOGroupIDSourceLabel.Visibility = "Hidden"; $DOGroupIDLabel.Visibility = "Visible"; $DOGroupID.Visibility = "Visible"; $DOGenerateGUID.Visibility = "Visible" })

$DOUseCacheHost.Add_Checked({ $DOUseDHCPCacheHost.Visibility = "Visible"; $DOUseDHCPCacheHostLabel.Visibility = "Visible"; $DOCacheHost.Visibility = "Visible"; $DOCacheHostLabel.Visibility = "Visible"; $DOCacheHostSource.Visibility = "Hidden"; $DOCacheHostSourceLabel.Visibility = "Hidden" })
$DOUseCacheHost.Add_UnChecked({ $DOUseDHCPCacheHost.Visibility = "Hidden"; $DOUseDHCPCacheHostLabel.Visibility = "Hidden"; $DOCacheHost.Visibility = "Hidden"; $DOCacheHostLabel.Visibility = "Hidden"; $DOCacheHostSource.Visibility = "Hidden"; $DOCacheHostSourceLabel.Visibility = "Hidden" })

$DOUseDHCPCacheHost.Add_Checked({ $DOCacheHost.Visibility = "Hidden"; $DOCacheHostLabel.Visibility = "Hidden"; $DOCacheHostSource.Visibility = "Visible"; $DOCacheHostSourceLabel.Visibility = "Visible" })
$DOUseDHCPCacheHost.Add_UnChecked({ $DOCacheHost.Visibility = "Visible"; $DOCacheHostLabel.Visibility = "Visible"; $DOCacheHostSource.Visibility = "Hidden"; $DOCacheHostSourceLabel.Visibility = "Hidden" })

#-------------------------------------------------------------#
#----About----------------------------------------------------#
#-------------------------------------------------------------#



# Function controls

#Update Ring slider configuration
$RingSlider.add_ValueChanged({ 
		if ($RingSlider.Value -ge 4) { $RingDescription.content = "With greater or equal 4 Rings, the first ring will be in Insider Slow" }
		if ($RingSlider.Value -le 3) { $RingDescription.content = "With equal or less than 3 Rings, all rings should be in the GA Channel" }
	})

$ToolCheckBox.Add_Checked({ 
		$AutoUpdateConfig.Visibility = "Hidden"
		$AutoUpdate.Visibility = "Hidden"

		$UpdateIntervalTextBox.Visibility = "Visible"
		$UpdateIntervalTextBlock.Visibility = "Visible"
		$UpdateIntervalSlider.Visibility = "Visible"
		$ScanIntervalHeader.Visibility = "Visible"

	})

$ToolCheckBox.Add_UnChecked({ 
		$AutoUpdateConfig.Visibility = "Visible"
		$AutoUpdate.Visibility = "Visible"

		if ($UpdateSource.SelectedItem -like "*WSUS") {
			$UpdateIntervalTextBox.Visibility = "Visible"
			$UpdateIntervalTextBlock.Visibility = "Visible"
			$UpdateIntervalSlider.Visibility = "Visible"
			$ScanIntervalHeader.Visibility = "Visible"
		}
		else {
			$UpdateIntervalTextBox.Visibility = "Hidden"
			$UpdateIntervalTextBlock.Visibility = "Hidden"
			$UpdateIntervalSlider.Visibility = "Hidden"
			$ScanIntervalHeader.Visibility = "Hidden"
		}
	})

$UpdateSource.Add_SelectionChanged({
		if ($UpdateSource.SelectedItem -like "*WSUS") {
			$UpdateIntervalTextBox.Visibility = "Visible"
			$UpdateIntervalTextBlock.Visibility = "Visible"
			$UpdateIntervalSlider.Visibility = "Visible"
			$ScanIntervalHeader.Visibility = "Visible"
			$TabWSUSBT.IsEnabled = "True"
			$TabWSUSBT.Foreground = "#ffffff"
			$CUSliderLabel.Visibility = "Hidden"
			$FUSliderLabel.Visibility = "Hidden"
			$CUSlider.Visibility = "Hidden"
			$FUSlider.Visibility = "Hidden"
			$FUSliderTextBox.Visibility = "Hidden"
			$CUSliderTextBox.Visibility = "Hidden"
		}
		else {
			if ($ToolCheckBox.IsChecked -eq $false) {
				$UpdateIntervalTextBox.Visibility = "Hidden"
				$UpdateIntervalTextBlock.Visibility = "Hidden"
				$UpdateIntervalSlider.Visibility = "Hidden"
				$ScanIntervalHeader.Visibility = "Hidden"
			}
			$CUSliderLabel.Visibility = "Visible"
			$FUSliderLabel.Visibility = "Visible"
			$CUSlider.Visibility = "Visible"
			$FUSlider.Visibility = "Visible"
			$FUSliderTextBox.Visibility = "Visible"
			$CUSliderTextBox.Visibility = "Visible"
			$TabWSUSBT.IsEnabled = $false
			$TabWSUSBT.Foreground = "#5b5a5c"
		}
	})

$ProductVersionSelection.Add_SelectionChanged({
		if ($ProductVersionSelection.SelectedItem -like "*Windows 10") { $TargetReleaseVersion.Visibility = "Visible"; $ReleaseVersionW10Box.Visibility = "Visible"; $ReleaseVersionW11Box.Visibility = "Hidden" }
		if ($ProductVersionSelection.SelectedItem -like "*Windows 11") { $TargetReleaseVersion.Visibility = "Visible"; $ReleaseVersionW10Box.Visibility = "Hidden"; $ReleaseVersionW11Box.Visibility = "Visible" }
		if ($ProductVersionSelection.SelectedItem -Like "*None") { $TargetReleaseVersion.Visibility = "Hidden"; $ReleaseVersionW11Box.Visibility = "Hidden"; $ReleaseVersionW10Box.Visibility = "Hidden" }
	})

# End side Nav Bindings

$Global:SyncHash = [HashTable]::Synchronized(@{ })
$Jobs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
$initialSessionState = [initialsessionstate]::CreateDefault()

Function Start-RunspaceTask {
	[CmdletBinding()]
	Param([Parameter(Mandatory = $True, Position = 0)][ScriptBlock]$ScriptBlock,
		[Parameter(Mandatory = $True, Position = 1)][PSObject[]]$ProxyVars)

	$Runspace = [RunspaceFactory]::CreateRunspace($InitialSessionState)
	$Runspace.ApartmentState = 'STA'
	$Runspace.ThreadOptions = 'ReuseThread'
	$Runspace.Open()
	ForEach ($Var in $ProxyVars) { $Runspace.SessionStateProxy.SetVariable($Var.Name, $Var.Variable) }
	$Thread = [PowerShell]::Create('NewRunspace')
	$Thread.AddScript($ScriptBlock) | Out-Null
	$Thread.Runspace = $Runspace
	[Void]$Jobs.Add([PSObject]@{ PowerShell = $Thread ; Runspace = $Thread.BeginInvoke() })
}

$JobCleanupScript = {
	Do {
		ForEach ($Job in $Jobs) {
			If ($Job.Runspace.IsCompleted) {
				[Void]$Job.Powershell.EndInvoke($Job.Runspace)
				$Job.PowerShell.Runspace.Close()
				$Job.PowerShell.Runspace.Dispose()
				$Runspace.Powershell.Dispose()

				$Jobs.Remove($Runspace)
			}
		}

		Start-Sleep -Seconds 1
	}
	While ($SyncHash.CleanupJobs)
}

Get-ChildItem Function: | Where-Object { $_.name -notlike "*:*" } | Select-Object name -ExpandProperty name |
ForEach-Object {
	$Definition = Get-Content "function:$_" -ErrorAction Stop
	$SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $Definition
	$InitialSessionState.Commands.Add($SessionStateFunction)
}


$Window.Add_Closed( {
		Write-Verbose 'Halt runspace cleanup job processing'
		$SyncHash.CleanupJobs = $False
	})

$SyncHash.CleanupJobs = $True
function Async($scriptBlock) {
	Start-RunspaceTask $scriptBlock @([PSObject]@{ Name = 'DataContext' ; Variable = $DataContext }, [PSObject]@{Name = "State"; Variable = $State })
}

Start-RunspaceTask $JobCleanupScript @([PSObject]@{ Name = 'Jobs' ; Variable = $Jobs })

$Window.ShowDialog()