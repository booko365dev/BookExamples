
Function ConnectPsEwsBA()
{
	$ExService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
	$ExService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(`
		$configFile.appsettings.exUserName, $configFile.appsettings.exUserPw)
	$ExService.Url = new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx");
	#$ExService.TraceEnabled = $true
	#$ExService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
	$ExService.AutodiscoverUrl($configFile.appsettings.exUserName, {$true})

	return $ExService
}

Function ConnectPsOnlBA() 
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.exUserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList `
				$configFile.appsettings.exUserName, $securePW
	$mySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri `
				https://outlook.office365.com/powershell-liveid/ -Authentication Basic `
				-AllowRedirection -Credential $myCredentials
	Import-PSSession $mySession -AllowClobber
}
#-----------------------------------------------------------------------------------------

Function CallEWSTest($ExService) {
	$myFolderView = [Microsoft.Exchange.WebServices.Data.FolderView]100
	$allFolders = $ExService.FindFolders(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, `
		$myFolderView)
	foreach ($oneFolder in $allFolders) {
		Write-Host $oneFolder.DisplayName
	}
}

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\exPs.values.config"

##==> EWS Basic Authorization
Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$ExService = ConnectPsEwsBA

CallEWSTest $ExService  #Calling any function

##==> EWS oAuth Authorization
Import-Module .\GenericOauthEWS.ps1 -Force
#Test-EWSConnection -MailboxName $configFile.appsettings.exUserName
$ExService = Connect-Exchange `
				$configFile.appsettings.exUserName "" $configFile.appsettings.exAppId

CallEWSTest $ExService  #Calling any function

##==> Exchange Online PowerShell Basic Authorization
ConnectPsOnlBA

Get-Mailbox  #Calling any cmdlet

$currentSession = Get-PSSession
Remove-PSSession -Session $currentSession

Write-Host "Done"  


