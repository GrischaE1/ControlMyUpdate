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
# Name: upload_sensors.ps1
# Version: 0.1
# Date: 24.01.2022
# Created by: Grischa Ernst gernst@vmware.com
#
# Description
# - This script will upload all .ps1 files that are in the source folder ($sourcepath) as sensor
#
# How To
# - Run the script with the parameter
#
# upload_sensors.ps1 -SourcePath "C:\Temp" -APIEndpoint "as137.awmdm.com" -APIUser "APIAdmin" -APIPassword 'Password' -APIKey '123412341234' -OGID "1234"
#
##########################################################################################
#                                    Changelog 
#
# 0.3 - Bug fixing
# 0.2 - Changing the data type due to new sensors
# 0.1 - Initial creation
##########################################################################################

##########################################################################################
#                                    Param 
#

param(
		[string]$SourcePath,
        [string]$APIEndpoint,
        [string]$APIUser,
        [string]$APIPassword,
        [string]$APIKey,
        [string]$OGID
	)

##########################################################################################
#                                    Functions

function Create-UEMAPIHeader
{
    param(
        [string] $APIUser, 
        [string] $APIPassword,
        [string] $APIKey,
        [string] $ContentType  = "json",
        [string] $Accept  = "json",
        [int] $APIVersion  = 1
    )

        #generate API Credentials
        $UserNameWithPassword  =  $APIUser + ":" + $APIPassword
        $Encoding  = [System.Text.Encoding]::ASCII.GetBytes($UserNameWithPassword)
        $EncodedString  = [Convert]::ToBase64String($Encoding)
        $Auth  = "Basic " + $EncodedString

        #generate Header
        $headers  = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $headers.Add("aw-tenant-code", $APIKey)
        $headers.Add("Authorization", $auth)
        $headers.Add("Accept", "application/$($Accept);version=$($APIVersion)")
        $headers.Add("Content-Type", "application/$($ContentType)")
        return $headers

}


function Create-APISensorUploadBody 
{
    param(
         $sensorname,
         $orgGUID,
         $scriptcontent,
         $responsetype = "STRING") 
    
         $APIbody = '{
            "name": "$($sensorname)",
            "description": "",
            "is_read_only": false,
            "organization_group_uuid": "$($orgGUID)",
            "platform": "WIN_RT",
            "query_type": "POWERSHELL",
            "query_response_type": "$($responsetype)",
            "trigger_type": "SCHEDULE",
            "execution_context": "SYSTEM",
            "execution_architecture": "ONLY64BIT",
            "script_data": "$($scriptcontent)",
            "timeout": 0,
            "event_trigger": [
                "UNKNOWN"
            ],
            "schedule_trigger": "UNKNOWN",
            "assigned_smart_groups": []
        }'
    
            $json = $ExecutionContext.InvokeCommand.ExpandString($APIbody) 
            return $json
}

function Get-SmartGroupInfo
{
    param(
        $SmartGroupName)    
    
    #generate UEM header
    $header = Create-UEMAPIHeader -APIUser $APIUser -APIPassword $APIPassword -APIKey $APIKey

    #Get OG UUID
    $url  = "https://$($APIEndpoint)/API/mdm/smartgroups/search?name=$($SmartGroupName.Replace(" ","%20"))"
    $SmartGroupDetails = (Invoke-RestMethod $url -Method 'GET' -Headers $header).SmartGroups 

    return  $SmartGroupDetails
}


function New-SensorAssignmentBody
{
    param(
        $SmartGroupUUID,
        $AssignmentName
    )
       
        $APIbody = '{
            "name": "$($AssignmentName)",
            "smart_group_uuids": [
              "$($SmartGroupUUID)"
            ],
            "trigger_type": "SCHEDULE",
            "event_triggers": []
          }'

            $json = $ExecutionContext.InvokeCommand.ExpandString($APIbody) 
            return $json

}



##########################################################################################
# Start


#generate UEM header
$header = Create-UEMAPIHeader -APIUser $APIUser -APIPassword $APIPassword -APIKey $APIKey

#Get OG UUID
$url  = "https://$($APIEndpoint)/API/system/groups/$($OGID)"
$OGDetails = Invoke-RestMethod $url -Method 'GET' -Headers $header 

$OGGUID = $OGDetails.uuid

#Get the sensor files
$AllFiles = Get-ChildItem $SourcePath | Where-Object {$_.Name -like "*.ps1" -and $_.Name -notlike "*sensor*"}

if($AllFiles)
{
    foreach($file in $AllFiles)
    {
        #Get the script Data and encrypt the data
        $Data = Get-Content $file.FullName -Encoding UTF8 -Raw
        $Bytes = [System.Text.Encoding]::UTF8.GetBytes($Data)
        $Script = [Convert]::ToBase64String($Bytes)
        
        #create JSON body
        if($file.BaseName -like "*_date*")
        {
            $body = Create-APISensorUploadBody -sensorname $($file.BaseName.ToLower()) -orgGUID $OGGUID -scriptcontent $Script -responsetype "DATETIME"
        }
        elseif($file.BaseName -like "*_bool*")
        {
            $body = Create-APISensorUploadBody -sensorname $($file.BaseName.ToLower()) -orgGUID $OGGUID -scriptcontent $Script -responsetype "BOOLEAN"
        }
        elseif($file.BaseName -like "*_count*")
        {
            $body = Create-APISensorUploadBody -sensorname $($file.BaseName.ToLower()) -orgGUID $OGGUID -scriptcontent $Script -responsetype "INTEGER"
        }
        else 
        {
            $body = Create-APISensorUploadBody -sensorname $($file.BaseName.ToLower()) -orgGUID $OGGUID -scriptcontent $Script
        }
        
        #upload sensor data
        $url  = "https://$($APIEndpoint)/API/mdm/devicesensors"
        Invoke-RestMethod $url -Method 'POST' -Headers $header -Body $body  

        Write-Output "$($file.basename) uploaded"
    }
   

    #Get all Sensors for specific OG
    $header = Create-UEMAPIHeader -APIUser $APIUser -APIPassword $APIPassword -APIKey $APIKey
    $url  = "https://$($APIEndpoint)/API/mdm/devicesensors/list/$($OGGUID)"
    $AllSensors = (Invoke-RestMethod $url -Method 'GET' -Headers $header).result_set

    if($AllSensors)
    {
        #Get Smart Group information
        $SmartGroupInfo = Get-SmartGroupInfo -SmartGroupName $SmartGroupName
        $SmartGroupUUID = $SmartGroupInfo.SmartGroupUuid
        $SmartGroupName = $SmartGroupInfo.Name

        #Assign the sensors to the Smart Group
        $header = Create-UEMAPIHeader -APIUser $APIUser -APIPassword $APIPassword -APIKey $APIKey -APIVersion 2

        foreach($file in $AllFiles)
        {
            $CurrentSensor = $AllSensors | Where-Object {$_.name -eq $file.BaseName}
            
            if($CurrentSensor -ne "" -and $CurrentSensor -ne $null)
            {
                $AssignmentName = ("$($file.BaseName) to $($SmartGroupName)").ToLower().Replace(" ","_")
                $Body = New-SensorAssignmentBody -SmartGroupUUID $SmartGroupUUID -AssignmentName $AssignmentName

                #Assign sensors
                $url  = "https://$($APIEndpoint)/API/mdm/devicesensors/$($CurrentSensor.uuid)/assignment"
                Invoke-RestMethod $url -Method 'POST' -Headers $header -Body $body 
            
            }
            Clear-Variable CurrentSensor
        }
    }
}