Function New-WUProfile
{
    Param(
        $TargetPath,
        #General Tab items
        $Ringcount,
        $DriversCheckBox,
        $AutoUpdateConfig,
        $TelemetryCheckBox,
        $SetDisableUXWUAccess,

        #Timing items
        $CUSlider,
        $FUSlider,
        $RestartSlider,
        $UpdateIntervalSlider,
        $RestartGraceSlider,

        #Release
        $WindowsChannel,
        $SafeguardCheckBox,
        $ProductVersionSelection,
        $ReleaseVersion,

        #WSUS
        $WSUSURL,
        $WSUSaltURL,
        $AllowNonMicrosoftSignedUpdate,
        $UseTLSpinning,
        $AllowUpdateService,
        $DualScanSelection
    )
    $JSONFile = "$PSscriptRoot\WU_URIMapping.json"
    $JsonInformation = Get-Content $JSONFile | ConvertFrom-Json

    $global:FinalInstallationCSP = @()
    $global:FinalUnInstallationCSP = @()

    ###########################################################################
    # General 

        #Exclude 3rd Party drivers
        #$DriversCheckBox 
        Get-JSONData -JSONFile $JsonInformation -SettingName 'ExcludeDrivers' -SettingValue $DriversCheckBox

        #Set Telemetry level
        #$TelemetryCheckBox 
        if($TelemetryCheckBox -and $WindowsChannel -notlike "*Insider*")
        {
            Get-JSONData -JSONFile $JsonInformation -SettingName 'AllowTelemetry' -SettingValue $TelemetryCheckBox
        }

        
        #Auto update behavior
        #$AutoUpdateConfig 
        switch ($AutoUpdateConfig) {
            'Notify the user before downloading the update'  { $AutoUpdateValue = 0}
            'Auto install the update and then notify the user to schedule a device restart' { $AutoUpdateValue = 1}
            'Auto install and restart' { $AutoUpdateValue = 2}
            'Auto install and restart at a specified time' { $AutoUpdateValue = 3}
            'Auto install and restart without end-user control' { $AutoUpdateValue = 4}
            'Turn off automatic updates' { $AutoUpdateValue = 5}
            Default { $AutoUpdateValue = 5}
        }
        Get-JSONData -JSONFile $JsonInformation -SettingName 'AutoUpdate' -SettingValue $AutoUpdateValue


        #Disable UX if tool is used
        if($SetDisableUXWUAccess)
        {
            Get-JSONData -JSONFile $JsonInformation -SettingName 'SetDisableUXWUAccess' -SettingValue $SetDisableUXWUAccess

        }

    ###########################################################################
    # Timing
        #Quality Updates deferral (days)
        #$CUSlider
        if($CUSlider)
        { 
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'CUDeferral' -SettingValue $CUSlider
        }

        #Feature Updates deferral (days)
        #$FUSlider 
        if($FUSlider)
        { 
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'FUDeferral' -SettingValue $FUSlider
        }
    
        #Restart deferral (days)"
        #$RestartSlider 
        if($RestartSlider)
        {
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'RestartDeadlineFU' -SettingValue $RestartSlider
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'RestartDeadlineCU' -SettingValue $RestartSlider
        }

        #Restart grace period (days)"
        #$RestartGraceSlider
        if($RestartGraceSlider)
        {
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'DeadlineGracePeriod' -SettingValue $RestartGraceSlider
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'DeadlineGracePeriodFU' -SettingValue $RestartGraceSlider
        }

        #Update Scan frequency  
        #$UpdateIntervalSlider 
        if($UpdateIntervalSlider)
        {
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'UpdateInterval' -SettingValue $UpdateIntervalSlider
        }

    ###########################################################################
    # Releases
        #Windows Release Branch
        #WindowsChannel
        switch ($WindowsChannel) {
            'Windows Insider build - Fast'  { $WindowsChannelValue = 2; Get-JSONData -JSONFile $JsonInformation  -SettingName 'ManagePreviewBuilds' -SettingValue '2'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowBuildPreview' -SettingValue '1'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowTelemetry' -SettingValue '3'}
            'Windows Insider build - Slow' { $WindowsChannelValue = 4; Get-JSONData -JSONFile $JsonInformation  -SettingName 'ManagePreviewBuilds' -SettingValue '2'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowBuildPreview' -SettingValue '1'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowTelemetry' -SettingValue '3'}
            'Release Windows Insider build' { $WindowsChannelValue = 8; Get-JSONData -JSONFile $JsonInformation  -SettingName 'ManagePreviewBuilds' -SettingValue '2'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowBuildPreview' -SettingValue '1'; Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowTelemetry' -SettingValue '3'}
            'General Availability Channel (Targeted)' { $WindowsChannelValue = 16; Get-JSONData -JSONFile $JsonInformation  -SettingName 'ManagePreviewBuilds' -SettingValue '0';  Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowBuildPreview' -SettingValue '0'}
            Default { $WindowsChannelValue = 16; Get-JSONData -JSONFile $JsonInformation  -SettingName 'ManagePreviewBuilds' -SettingValue '0';  Get-JSONData -JSONFile $JsonInformation  -SettingName 'AllowBuildPreview' -SettingValue '0'}
        }    
        Get-JSONData -JSONFile $JsonInformation  -SettingName 'BranchReadinessLevel' -SettingValue $WindowsChannelValue

        #Disable Windows Safeguards
        #SafeguardCheckBox
        if($SafeguardCheckBox -eq "True"){$SafeguardValue = 0}else{$SafeguardValue = 1}
        Get-JSONData -JSONFile $JsonInformation  -SettingName 'DisableWUfBSafeguards' -SettingValue $SafeguardValue


        #Target Release Version configuration
        if($ProductVersionSelection -like "*Windows*")
        {
            #Windows Product Version
            #ProductVersionSelection
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'ProductVersion' -SettingValue $ProductVersionSelection

            #Release Version
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'TargetReleaseVersion' -SettingValue $ReleaseVersion
        }

    ###########################################################################
    # WSUS
        if($WSUSURL)
        {
            #$WSUSURL
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'UpdateServiceUrl' -SettingValue $WSUSURL

            #$WSUSaltURL
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'UpdateServiceUrlAlternate' -SettingValue $WSUSaltURL

            #$AllowNonMicrosoftSignedUpdate
            if($AllowNonMicrosoftSignedUpdate -eq $true){$AllowNonMicrosoftSignedUpdateValue = 1}else{$AllowNonMicrosoftSignedUpdateValue = 0}
            Get-JSONData -JSONFile $JsonInformation -SettingName 'AllowNonMicrosoftSignedUpdate' -SettingValue $AllowNonMicrosoftSignedUpdateValue


            #$Use TLS pinning
            if($UseTSLpinning -eq $true){$UseTLSpinningValue = 1}else{$UseTLSpinningValue = 0}
            Get-JSONData -JSONFile $JsonInformation -SettingName 'TLSCertPinning' -SettingValue $UseTLSpinningValue

            #$AllowUpdateService
            if($AllowUpdateService -eq $true){$AllowUpdateServiceValue = 1}else{$AllowUpdateServiceValue = 0}
            Get-JSONData -JSONFile $JsonInformation -SettingName 'AllowUpdateService' -SettingValue $AllowUpdateServiceValue

            #$DualScanSelection
            $DualScanSelection
            foreach($Item in $DualScanSelection)
            {
                switch ($Item) {
                    'Driver'  { Get-JSONData -JSONFile $JsonInformation -SettingName 'UpdateSourceForDriver' -SettingValue "1" }
                    'Feature Updates' { Get-JSONData -JSONFile $JsonInformation -SettingName 'UpdateSourceForFeature' -SettingValue "1"}
                    'Quality Updates' { Get-JSONData -JSONFile $JsonInformation -SettingName 'UpdateSourceForQuality' -SettingValue "1"}
                    'Other' { Get-JSONData -JSONFile $JsonInformation -SettingName 'UpdateSourceForOther' -SettingValue "1"}

                    Default { }
                }
            }
        }

    #Add Hidden Settings to CSP
    $HiddenSettings = $JsonInformation | Where-Object {$_.HiddenSetting -eq $True}
    foreach($Setting in $HiddenSettings)
    {
        Get-JSONData -JSONFile $JsonInformation  -SettingName $($Setting.Name) -SettingValue $($Setting.BPValue)
    }

    ###########################################################################
    # Generate XML
     $FinalInstallationCSP | Out-File "$($TargetPath)\Ring-$($Ringcount)_WU_InstallProfile.txt" -Force
     $FinalUnInstallationCSP | Out-File "$($TargetPath)\Ring-$($Ringcount)_WU_UnInstallProfile.txt" -Force
}


Function New-DOProfile
{
    Param(
        $TargetPath,
        $DODownloadMode,
        $DOAutomaticGUID,
        $DOGroupID,
        $DOGroupIDSource,
        $DOMinFileSizeToCache,
        $DOMaxDownloadBandwidth,
        $DOUseCacheHost,
        $DOUseDHCPCacheHost,
        $DOCacheHost,
        $DOCacheHostSource
    )
    $JSONFile = "$PSscriptRoot\DO_URIMapping.json"
    $JsonInformation = Get-Content $JSONFile | ConvertFrom-Json

    $global:FinalInstallationCSP = @()
    $global:FinalUnInstallationCSP = @()


    #Download Mode    
    Get-JSONData -JSONFile $JsonInformation  -SettingName 'DODownloadMode' -SettingValue $DODownloadModeValue

    if($DODownloadMode -eq '2')
    {
        #GroupID
        if($DOGroupID)
        {            
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOGroupID' -SettingValue $DOGroupID
        }

        #Group ID Source
        if($DOGroupIDSource)
        {
            switch ($DOGroupIDSource) {
                'AD Site'  { $DOGroupIDSourceValue = 1}
                'Authenticated domain SID' { $DOGroupIDSourceValue = 2}
                'DHCP Option ID' { $DOGroupIDSourceValue = 3}
                'DNS Suffix' { $DOGroupIDSourceValue = 4}
                'AAD' { $DOGroupIDSourceValue = 5}
                Default { $DOGroupIDSourceValue = 1}
            }
            Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOGroupIDSource' -SettingValue $DOGroupIDSourceValue
        }
    }

    #Minimum Peer Caching Content File Size
    Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOMinFileSizeToCache' -SettingValue $DOMinFileSizeToCache

    #Maximum download bandwidth in KiloBytes/second
    Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOMaxForegroundDownloadBandwidth' -SettingValue $DOMaxDownloadBandwidth
    Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOMaxBackgroundDownloadBandwidth' -SettingValue $DOMaxDownloadBandwidth


    #Use Cache Host config
    if($DOCacheHost)
    {
        Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOCacheHost' -SettingValue $DOCacheHost
    }

    #DHCP Options for Cache Host
    if($DOCacheHostSource)
    {
        switch ($DOCacheHostSource) {
            'DHCP Option ID'  { $DOCacheHostSourceValue = 1}
            'DHCP Option ID Force' { $DOCacheHostSourceValue = 2}
            Default { $DOCacheHostSourceValue = 1}
        }
        Get-JSONData -JSONFile $JsonInformation  -SettingName 'DOCacheHostSource' -SettingValue $DOCacheHostSourceValue
    }

    #Add Hidden Settings to CSP
    $HiddenSettings = $JsonInformation | Where-Object {$_.HiddenSetting -eq $True}
    foreach($Setting in $HiddenSettings)
    {
        Get-JSONData -JSONFile $JsonInformation  -SettingName  $($Setting.Name) -SettingValue $($Setting.BPValue)
    }


    ###########################################################################
    # Generate XML
    $FinalInstallationCSP | Out-File "$($TargetPath)\DO_InstallProfile.txt"-Force
    $FinalUnInstallationCSP | Out-File "$($TargetPath)\DO_UnInstallProfile.txt" -Force
}

function New-CustomUpdateProfile {
      param (
        $TargetPath,
        $ReportOnly,
        $ToolUpdateSource,
        $DirectDownload,
        $EnablemaintenanceWindow,
        $MWStartTime,
        $MWEndTime,
        $MWDay,
        $BlockedKBs,
        $UnBlockedKBs,
        $EmergencyKB,
        $ScanInterval,
        $ScanRandomization,
        $NotifyUser,
        $ToastTitle,
        $ToastText
    )

    #Generate GUID'S
    $guid1 = [guid]::NewGuid()
    $guid2 = [guid]::NewGuid()


    if($ReportOnly)
    {

        $temp = '
                <wap-provisioningdoc id="$($guid1)" name="customprofile">
                    <characteristic type="com.airwatch.winrt.powershellcommand" uuid="$($guid2)">
                    <parm name="PowershellCommand" value="Invoke-Command -ScriptBlock {
                        New-Item HKLM:\SOFTWARE\ControlMyUpdate\Settings -Force;
                        New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ReportOnly -PropertyType String -Value True;
                        New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate -Name ScriptLogLevel -PropertyType String -Value Info
                    }"/>
                    </characteristic>
                </wap-provisioningdoc>
                '

        $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
        $XMLProfile | Out-File "$($TargetPath)\CU_InstallProfile.txt"-Force
    }

    else
    {
        $convertedDays = @()
        $selectedDays = ""
        #Day mapping for maintenance Window
        if($MWDay)
        {
            foreach($day in $MWDay)
            {
                switch ( $day )
                {
                    'Monday'    { $TargetDay =  1   }
                    'Tuesday'   { $TargetDay =  2   }
                    'Wednesday' { $TargetDay =  3   }
                    'Thursday'  { $TargetDay =  4   }
                    'Friday'    { $TargetDay =  5   }           
                    'Saturday'  { $TargetDay =  6   }
                    'Sunday'    { $TargetDay =  7   }
                    default { 'None' }
                }
               
                $convertedDays +=  $TargetDay            
            }
        }
        else{$selectedDays = "None"}
        
        if($convertedDays.count -ge 1)
        {
            [string]$selectedDays = $convertedDays[0]
            ForEach($day in $convertedDays)
            {
                if($day -ne $convertedDays[0])
                {
                    [string]$selectedDays = "$selectedDays,$day"
                }
            }
        }

        if(!$BlockedKBs){$BlockedKBs = '&quot;&quot;'}
        if(!$EmergencyKB){$EmergencyKB = '&quot;&quot;'}
        if(!$UnBlockedKBs){$UnBlockedKBs = '&quot;&quot;'}
        if($MWStartTime -eq "00:00" -and $MWEndTime -eq "00:00"){$MWStartTime = '&quot;&quot;' 
        $MWEndTime = '&quot;&quot;'}
        if($selectedDays -eq "none"){$selectedDays = '&quot;&quot;'}
        else{$selectedDays = "&quot;$($selectedDays)&quot;"}
        if($EnablemaintenanceWindow -eq $false){
            $selectedDays = '&quot;&quot;'
            $MWStartTime = '&quot;&quot;'
            $MWEndTime = '&quot;&quot;'
        }
        switch ( $ToolUpdateSource )
        {
            'Default'    { $selectedUpdateSource =  "Default"   }
            'Microsoft Update'   { $selectedUpdateSource =  "MU"   }
            'WSUS' { $selectedUpdateSource =  "WSUS"   }          
        }
      

        $temp = '
        <wap-provisioningdoc id="$($guid1)" name="customprofile">
            <characteristic type="com.airwatch.winrt.powershellcommand" uuid="$($guid2)">
                <parm name="PowershellCommand" value="Invoke-Command -ScriptBlock {
                    New-Item HKLM:\SOFTWARE\ControlMyUpdate\Settings -Force;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name DirectDownload -PropertyType String -Value $($DirectDownload);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name EmergencyKB -PropertyType String -Value $($EmergencyKB);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name HiddenUpdates -PropertyType String -Value $($BlockedKBs);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name UnHiddenUpdates -PropertyType String -Value  $($UnBlockedKBs);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name LastInstallationDate -PropertyType String -Value &quot;&quot;;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MaintenanceWindow -PropertyType String -Value $($EnableMaintenanceWindow);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWDay -PropertyType String -Value $($selectedDays);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWStartTime -PropertyType String -Value  $($MWStartTime);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWStopTime -PropertyType String -Value $($MWEndTime);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name UpdateSource -PropertyType String -Value $($selectedUpdateSource);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name LastScanTime -PropertyType String -Value &quot;&quot;;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name NextScanTime -PropertyType String -Value &quot;&quot;;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ScanInterval -PropertyType String -Value $($ScanInterval);  
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ScanRandomization -PropertyType String -Value $($ScanRandomization);  
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ReportOnly -PropertyType String -Value False;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name NotifyUser -PropertyType String -Value $($NotifyUser);
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ToastTitle -PropertyType String -Value &quot;$($ToastTitle)&quot;;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ToastText -PropertyType String -Value &quot;$($ToastText)&quot;;
                    New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate -Name ScriptLogLevel -PropertyType String -Value Info
                }"/>
            </characteristic>
        </wap-provisioningdoc>
        '

        $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
        $XMLProfile | Out-File "$($TargetPath)\CU_InstallProfile.txt"-Force
    }
    
        $temp = '
        <wap-provisioningdoc id="$($guid1)" name="customprofile">
            <characteristic type="com.airwatch.winrt.powershellcommand" uuid="$($guid2)">
            <parm name="PowershellCommand" value="Invoke-Command -ScriptBlock {
                Remove-Item HKLM:\SOFTWARE\ControlMyUpdate -Force               
            }"/>
            </characteristic>
        </wap-provisioningdoc>
        '

        $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
        $XMLProfile | Out-File "$($TargetPath)\CU_UnInstallProfile.txt"-Force
}

function New-CustomUpdateScript {
    param (
      $TargetPath,
      $ReportOnly,
      $ToolUpdateSource,
      $DirectDownload,
      $EnablemaintenanceWindow,
      $MWStartTime,
      $MWEndTime,
      $MWDay,
      $BlockedKBs,
      $UnBlockedKBs,
      $EmergencyKB,
      $ScanInterval,
      $ScanRandomization,
      $NotifyUser,
      $ToastTitle,
      $ToastText
  )

  if($ReportOnly)
  {

      $temp = '
                      New-Item HKLM:\SOFTWARE\ControlMyUpdate\Settings -Force
                      New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ReportOnly -PropertyType String -Value True
                      New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate -Name ScriptLogLevel -PropertyType String -Value Info
              '

      $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
      $XMLProfile | Out-File "$($TargetPath)\CU_InstallScript.ps1"-Force
  }

  else
  {
      $convertedDays = @()
      $selectedDays = ""
      #Day mapping for maintenance Window
      if($MWDay)
      {
          foreach($day in $MWDay)
          {
              switch ( $day )
              {
                  'Monday'    { $TargetDay =  1   }
                  'Tuesday'   { $TargetDay =  2   }
                  'Wednesday' { $TargetDay =  3   }
                  'Thursday'  { $TargetDay =  4   }
                  'Friday'    { $TargetDay =  5   }           
                  'Saturday'  { $TargetDay =  6   }
                  'Sunday'    { $TargetDay =  7   }
                  default { 'None' }
              }
             
              $convertedDays +=  $TargetDay            
          }
      }
      else{$selectedDays = "None"}
      
      if($convertedDays.count -ge 1)
      {
          [string]$selectedDays = $convertedDays[0]
          ForEach($day in $convertedDays)
          {
              if($day -ne $convertedDays[0])
              {
                  [string]$selectedDays = "$selectedDays,$day"
              }
          }
      }

      if(!$BlockedKBs){$BlockedKBs = '""'}
      if(!$EmergencyKB){$EmergencyKB = '""'}
      if(!$UnBlockedKBs){$UnBlockedKBs = '""'}
      if($MWStartTime -eq "00:00" -and $MWEndTime -eq "00:00"){$MWStartTime = '""' 
      $MWEndTime = '""'}
      if($selectedDays -eq "none"){$selectedDays = '""'}
      else{$selectedDays = "$($selectedDays)"}
      if($EnablemaintenanceWindow -eq $false){
          $selectedDays = '""'
          $MWStartTime = '""'
          $MWEndTime = '""'
      }
      switch ( $ToolUpdateSource )
      {
          'Default'    { $selectedUpdateSource =  "Default"   }
          'Microsoft Update'   { $selectedUpdateSource =  "MU"   }
          'WSUS' { $selectedUpdateSource =  "WSUS"   }          
      }
    

      $temp = '
                  New-Item HKLM:\SOFTWARE\ControlMyUpdate\Settings -Force
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name DirectDownload -PropertyType String -Value $($DirectDownload)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name EmergencyKB -PropertyType String -Value $($EmergencyKB)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name HiddenUpdates -PropertyType String -Value $($BlockedKBs)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name UnHiddenUpdates -PropertyType String -Value  $($UnBlockedKBs)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name LastInstallationDate -PropertyType String -Value ""
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MaintenanceWindow -PropertyType String -Value $($EnableMaintenanceWindow)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWDay -PropertyType String -Value $($selectedDays)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWStartTime -PropertyType String -Value  $($MWStartTime)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name MWStopTime -PropertyType String -Value $($MWEndTime)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name UpdateSource -PropertyType String -Value $($selectedUpdateSource)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name LastScanTime -PropertyType String -Value ""
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name NextScanTime -PropertyType String -Value ""
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ScanInterval -PropertyType String -Value $($ScanInterval)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ScanRandomization -PropertyType String -Value $($ScanRandomization)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ReportOnly -PropertyType String -Value False
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name NotifyUser -PropertyType String -Value $($NotifyUser)
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ToastTitle -PropertyType String -Value "$($ToastTitle)"
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate\Settings -Name ToastText -PropertyType String -Value "$($ToastText)"
                  New-ItemProperty -Path HKLM:\SOFTWARE\ControlMyUpdate -Name ScriptLogLevel -PropertyType String -Value Info
      '

      $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
      $XMLProfile | Out-File "$($TargetPath)\CU_InstallScript.ps1"-Force
  }
  
      $temp = '
              Remove-Item HKLM:\SOFTWARE\ControlMyUpdate -Force               
      '

      $XMLProfile = $ExecutionContext.InvokeCommand.ExpandString($temp)
      $XMLProfile | Out-File "$($TargetPath)\CU_UnInstallScript.ps1"-Force
}

function Get-JSONData {
    param ( 
        $JSONFile,
        $SettingName,
        $SettingValue
    )
    
   

    $JSONData = $JsonInformation | Where-Object {$_.Name -eq $SettingName}

    $global:FinalInstallationCSP += Add-XMLProfile -XMLURI $($JSONData.URI) -XMLFormat $($JSONData.Format) -XMLdata $SettingValue
    $global:FinalUnInstallationCSP += Remove-XMLProfile -XMLURI $($JSONData.URI)

}


Function Add-XMLProfile
{
    param(
        $XMLURI,
        $XMLFormat,
        $XMLdata
    )

    #generate GUID
    $GUID = [guid]::NewGuid()


    $temp = '<Replace>
    <CmdID>$($GUID)</CmdID>
    <Item>
        <Target>
            <LocURI>$($XMLURI)</LocURI>
        </Target>
        <Meta>
            <Format xmlns="syncml:metinf">$($XMLFormat)</Format>
            <Type>text/plain</Type>
        </Meta>
        <Data>$($XMLData)</Data>
    </Item>
</Replace>
    '
   
    return $ExecutionContext.InvokeCommand.ExpandString($temp)
}

Function Remove-XMLProfile
{
    param(
        $XMLURI
    )

    #generate GUID
    $GUID = [guid]::NewGuid()


    $temp = '<Delete>
    <CmdID>$($GUID)</CmdID>
      <Item>
        <Target>
          <LocURI>$($XMLURI)</LocURI>
        </Target>
      </Item>
</Delete>
    '
   
    return $ExecutionContext.InvokeCommand.ExpandString($temp)
}