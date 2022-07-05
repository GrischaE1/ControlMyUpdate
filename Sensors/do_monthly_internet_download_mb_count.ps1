﻿$RegPath = "HKLM:\SOFTWARE\ControlMyUpdate\Status\DO"
$RegKey = "Monthly Internet Download MB"

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