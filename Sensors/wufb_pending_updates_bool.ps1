$RegPath = "HKLM:\SOFTWARE\ControlMyUpdate\Status"
$RegKey = "Open Pending Updates"

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