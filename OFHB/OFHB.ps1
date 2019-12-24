
Function LoginPsTeams()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.tmUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.tmUserName, $securePW
	Connect-MicrosoftTeams -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

Function TeamsPsMtpTeamsEnumarate()
{
	Get-Team
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamsGetByDisplayName()
{
	Get-Team -DisplayName "Test Team from PS"
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamsCreate()
{
	New-Team -DisplayName "Test Team from PS" `
			 -Description "Team created with PowerShell" `
			 -Visibility Private
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamsUpdate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-Team -GroupId $myTeam.GroupId `
			 -Description "Team updated with PowerShell" `
			 -Visibility Public
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamsDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-Team -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamsHelp()
{
	Get-TeamHelp
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpChannelsEnumerate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamChannel -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpChannelsCreate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	New-TeamChannel -GroupId $myTeam.GroupId `
					-DisplayName "Test Channel from PS" 
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpChannelsUpdate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-TeamChannel -GroupId $myTeam.GroupId `
					-CurrentDisplayName "Test Channel from PS" `
					-Description "This is a test Channel"
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpChannelsDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamChannel -GroupId $myTeam.GroupId `
					   -DisplayName "Test Channel from PS"
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamUserEnumerate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamUser -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamUserCreate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Add-TeamUser -GroupId $myTeam.GroupId `
				 -User "user@domain.onmicrosoft.com" 
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpTeamUserDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamUser -GroupId $myTeam.GroupId `
					-User "user@domain.onmicrosoft.com"
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpPolicyPackageEnumerate()
{
	Get-CsPolicyPackage
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpPolicyPackageUser()
{
	Get-CsUserPolicyPackage -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}

Function TeamsPsMtpPolicyPackageUserRecommended()
{
	Get-CsUserPolicyPackageRecommendation -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\tmPs.values.config"

LoginPsTeams

#TeamsPsMtpTeamsHelp
#TeamsPsMtpTeamsEnumarate
#TeamsPsMtpTeamsGetByDisplayName
#TeamsPsMtpTeamsCreate
#TeamsPsMtpTeamsUpdate
#TeamsPsMtpTeamsDelete

#TeamsPsMtpChannelsEnumerate
#TeamsPsMtpChannelsCreate
#TeamsPsMtpChannelsUpdate
#TeamsPsMtpChannelsDelete

#TeamsPsMtpTeamUserEnumerate
#TeamsPsMtpTeamUserCreate
#TeamsPsMtpTeamUserDelete

#TeamsPsMtpPolicyPackageEnumerate
#TeamsPsMtpPolicyPackageUser
#TeamsPsMtpPolicyPackageUserRecommended

Write-Host "Done"  


