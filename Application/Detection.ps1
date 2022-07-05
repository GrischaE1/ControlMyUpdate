param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="File Hash of the ControlMyUpdate.ps1 file")][String] $FileHash,
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="Version of Script to detect")][String] $ScriptExpectedVersion
)
$InstallDir = "C:\Windows\ControlMyUpdate"

$ScriptCurrentVersion = Invoke-Expression "$($InstallDir)\ControlMyUpdate.ps1 -ScriptVersion"

if ( $ScriptCurrentVersion -eq $ScriptExpectedVersion )
{
	Write-Debug "Script Version Match. Script Current Version : $($ScriptCurrentVersion) / Script Version Expected : $($ScriptExpectedVersion)"

	$InstalledFileHash = (Get-FileHash "$($InstallDir)\ControlMyUpdate.ps1").Hash

	if($FileHash -eq $InstalledFileHash)
	{
		Write-Debug "Script file hash match."
		Exit 0
	}
	else
	{
		Write-Debug "Script file hash doesn't match."
		Exit 4321
	}

	
}
else
{
	Write-Debug "Script Version do not match. Script Current Version : $($ScriptCurrentVersion) / Script Version Expected : $($ScriptExpectedVersion)"
	Exit 1234
}