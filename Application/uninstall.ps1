
##########################################################################################
#                                    Param 
#
#define the log folder if needed - otherwise delete it or set it to ""
param(
        [string]$InstallDir = "C:\Windows\ControlMyUpdate"
	)

#Check if scheduled task is created
$TaskCheck = Get-ScheduledTask -TaskName "Control My Update" -ErrorAction SilentlyContinue

If($TaskCheck)
{
    Unregister-ScheduledTask -TaskName $TaskCheck.TaskName -TaskPath $TaskCheck.TaskPath -Confirm:$false

}

$FolderTest = Get-Item $InstallDir -ErrorAction SilentlyContinue
if($FolderTest)
{
    Remove-Item -Path $InstallDir -recurse -Force 
}

$FolderValidation = Get-Item $InstallDir -ErrorAction SilentlyContinue
$TaskValidation = Get-ScheduledTask -TaskName "Control My Update" -ErrorAction SilentlyContinue

if(!$FolderValidation -and !$TaskValidation)
{
    exit 0
}
else{exit 1234}

