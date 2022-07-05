$RegPath = "HKLM:\SOFTWARE\ControlMyUpdate\Settings"
$RegKey = "LastInstallationDate"

if((Test-path "$($RegPath)"))
{
    $Value = Get-ItemPropertyValue -Path $RegPath -Name $RegKey
    if($Value)
    {
        return $Value
    }
    else
    {
        return "Path not found"
    }
}
else
{
    return "Path not found"
}