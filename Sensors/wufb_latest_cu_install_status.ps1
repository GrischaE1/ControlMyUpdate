function Get-SecondTuesday ([int]$Month, [int]$Year) {
    [int]$Day = 1
    while((Get-Date -Day $Day -Hour 0 -Millisecond 0 -Minute 0 -Month $Month -Year $Year -Second 0).DayOfWeek -ne "Tuesday") {
        $day++
    }
    $day += 7
    return (Get-Date -Day $Day -Hour 0 -Millisecond 0 -Minute 0 -Month $Month -Year $Year -Second 0)
}



 $RegPath = "HKLM:\SOFTWARE\ControlMyUpdate\Status\KBs"
 $Regitems = Get-ChildItem "$($RegPath)" -Recurse | Where-Object {$_.Property -like "Title"}
 



Foreach($Item in $Regitems)
{
   $Title = Get-ItemPropertyValue "$($Item.PSPath)" -Name "Title"
   

   
   if($Title -like "*Cumulative Update for Windows*")
   {
   
       $KBArticle = Split-Path -Path $item.PSPath -Leaf
       $KBInfo = Get-ChildItem "$($Item.PSParentPath)" -Recurse | Where-Object {$_.Name -like "*$($KBArticle)*"}  | Select-Object Property -ExpandProperty Property 
       $KBInstallstatus = ($KBInfo | Where-Object {$_ -Like "*Installation Status*"}).Replace("Installation Status : ","") 


        $SecondTuesday = Get-SecondTuesday -Month (Get-Date -Format MM) -Year (Get-Date -Format yyyy)

        if((Get-date) -ge $SecondTuesday)
        {
          $CurrentMonth = Get-Date -Format yyyy-MM
          $CurrentCU = $Title | Where-Object {$_ -like "*$($CurrentMonth) Cumulative Update for Windows*"}
  
          if($CurrentCU)
          {
                $CUInstalled = $KBInstallstatus
          }
          else{$CUInstalled = "Not detected"}
        }

        elseif((Get-date) -le $SecondTuesday)
        {
          $TempMonth = (Get-Date).AddMonths(-1)
          $LastMonth = Get-Date $TempMonth -Format yyyy-MM 

          $CurrentCU = $AllTitles | Where-Object {$_ -like "*$($LastMonth) Cumulative Update for Windows*"}
  
          if($CurrentCU)
          {
            $CUInstalled = $KBInstallstatus
          }
          else{$CUInstalled = "Not detected"}
        }

    }
}
return $CUInstalled
