$wsh = New-Object -ComObject WScript.Shell
$voice = New-Object -ComObject Sapi.spvoice
$voice.speak("System activated!")	
while (1 -eq 1) {
	cls
	echo Running...
	if ( (Get-Date).Hour -ge 17 -and (Get-Date).Minute -ge 15 ) {
		$voice.speak("Warning! This is a three minute warning!")		
		start-sleep -seconds 150
		$voice.speak("Warning! This is a thirty second warning!")
		start-sleep -seconds 30
		Stop-Computer -ComputerName localhost
		exit
	}
	$wsh.SendKeys('{NUMLOCK}{NUMLOCK}')
	start-sleep -seconds 60
}
