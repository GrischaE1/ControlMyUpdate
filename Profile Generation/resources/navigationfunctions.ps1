#Write your code here

#Nav Tab selection
Function Tab1Click() {
	$TabNav.SelectedItem = $Tab1
	$Tab1BT.Background = "#8c4803"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
}

Function Tab2Click() {
	$TabNav.SelectedItem = $Tab2
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#8c4803"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
}
Function Tab3Click() {
	$TabNav.SelectedItem = $Tab3
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#8c4803"; $TabWSUSBT.Background = "#002f54"
}

Function TabWSUSClick() {
	$TabNav.SelectedItem = $TabWSUS
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#8c4803"

}

Function TabGenerateClick() {
	$TabNav.SelectedItem = $TabGenerate
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
}

Function TabToolClick() {
	$TabNav.SelectedItem = $TabTool
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
}



#Next Button
Function NextTab1Click() {
	if ($Tab2BT.IsEnabled -eq "True") 
	{
		$TabNav.SelectedItem = $Tab2
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#8c4803"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
	}
	else {
		$TabNav.SelectedItem = $Tab3
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#8c4803"; $TabWSUSBT.Background = "#002f54"
	}
}

Function NextTab2Click() {
	if ($Tab3BT.IsEnabled -eq "True") 
	{
		$TabNav.SelectedItem = $Tab3
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#8c4803"; $TabWSUSBT.Background = "#002f54"
	}
	else {
		$TabNav.SelectedItem = $TabWSUS
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#8c4803"
	}
}

Function NextTab3Click() {
	if ($TabWSUSBT.IsEnabled -eq "True") 
	{
		$TabNav.SelectedItem = $TabWSUS
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#8c4803"
	}
	elseif ($ToolCheckBox.IsChecked -eq $true) {
		$MenuNavigation.SelectedItem = $TABCustomToolConfig
	}
	else {
		$MenuNavigation.SelectedItem = $TABGenerate
	}	
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"

	
}

Function NextTabWSUSClick() {
	if ($ToolCheckBox.IsChecked -eq $true) {
		$MenuNavigation.SelectedItem = $TABCustomToolConfig
	}
	else {
		$MenuNavigation.SelectedItem = $TABGenerate
	}	
	$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
}
	

#Previous Tab


Function PreviousTabWSUSClick() {
	if ($Tab3BT.IsEnabled -eq "True") 
	{
		$TabNav.SelectedItem = $Tab3
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#8c4803"; $TabWSUSBT.Background = "#002f54"
	}
	else {
		$TabNav.SelectedItem = $Tab2
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#8c4803"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
	}
}

Function PreviousTab3Click() {
	if ($Tab2BT.IsEnabled -eq "True") 
	{
		$TabNav.SelectedItem = $Tab2
		$Tab1BT.Background = "#002f54"; $Tab2BT.Background = "#8c4803"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
	}
	else {
		$TabNav.SelectedItem = $Tab1
		$Tab1BT.Background = "#8c4803"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
	}
}
Function PreviousTab2Click() {
		$TabNav.SelectedItem = $Tab1
		$Tab1BT.Background = "#8c4803"; $Tab2BT.Background = "#002f54"; $Tab3BT.Background = "#002f54"; $TabWSUSBT.Background = "#002f54"
	}

	Function GenerateCustomScriptProfileClick() {
			$MenuNavigation.SelectedItem = $TABGenerate
	}