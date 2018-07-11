param([string]$CSVFileName = "WLANProfiles.CSV")

function GetProfileList
{
 process
 {
	Write-Debug "Entering GetProfileList function"
	$command = 'netsh wlan show profile'
	$output = Invoke-Expression $command
	
	$profileList = @()
	foreach($profobj in $output)
	{
		if($profobj -like "*All User Profile*")
		{
			$profobj = $($profobj -Replace "All User Profile     :", "").Trim()
			Write-Debug "Found Profile : $($profobj)"
			$profileList += $profobj
		}
	}

	Write-Debug "Exiting GetProfileList function"
	return $profileList
 }
}

function LookUpProfileInformation
{
	param([string] $profileName)
	
	process
	{
		Write-Debug "Entering GetProfileList function"
		$command = "netsh wlan show profile name='{0}' key=clear" -f $profileName
		$output = Invoke-Expression $command
		#Write-Debug ($output | Format-Table | Out-String)
		Write-Debug ($output | Format-List | Out-String)

		$strout = $($output | Out-String)
		Write-Verbose $strout
		
		#
		$indexSSIDStart = $strout.IndexOf("Name                   :") + $("Name                   :").Length
		Write-Debug "Index of SSID: $($indexSSIDStart)"
		$indexKeyStart = $strout.IndexOf("Key Content            :") + $("Key Content            :").Length
		Write-Debug "Index of Key: $($indexKeyStart)"
		$SSID = $strout.Substring($indexSSIDStart, ($strout.IndexOf("Control options        :")) - $indexSSIDStart)
		if($indexKeyStart -ne 23)
		{
			$Passphrase = $strout.Substring($indexKeyStart, ($strout.IndexOf("Cost settings")) - $indexKeyStart)
		}
		else
		{
			$Passphrase = "N/A"
		}
		
		$keyobj = New-Object PSObject

		$keyobj | Add-Member NoteProperty "SSID" $SSID.Trim()
		$keyobj | Add-Member NoteProperty "PassPhrase" $Passphrase.Trim()

		Write-Debug "Exiting GetProfileList function"
		return $keyobj
	}
}

Clear-Host

$DebugPreference = "SilentyContinue"

$profileData = @()

$list = GetProfileList
foreach($profile in $list)
{
	Write-Host "Lookup Profile Information for $($profile)"
	$keyItem = LookUpProfileInformation -profileName $profile
	
	Write-Host "  SSID      : $($keyItem.SSID)"
	Write-Host "  Passphrase: $($keyItem.PassPhrase)"
	
	$profileData += $keyItem
	
}

Write-Host "`nWrite data to CSV file: $($CSVFileName)"
$profileData | export-csv -NoTypeInformation -Path $CSVFileName

Write-Host "Done!"

exit(0)


