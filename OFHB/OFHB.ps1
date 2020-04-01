﻿
#gavdcodebegin 01
Function LoginPsTeams()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.tmUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.tmUserName, $securePW
	Connect-MicrosoftTeams -Credential $myCredentials
}
#gavdcodeend 01

#----------------------------------------------------------------------------------------

#gavdcodebegin 02
Function TeamsPsMtpTeamsEnumarate()
{
	Get-Team
	Disconnect-MicrosoftTeams
}
#gavdcodeend 02

#gavdcodebegin 03
Function TeamsPsMtpTeamsGetByDisplayName()
{
	Get-Team -DisplayName "Test Team from PS"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 03

#gavdcodebegin 04
Function TeamsPsMtpTeamsCreate()
{
	New-Team -DisplayName "Test Team from PS" `
			 -Description "Team created with PowerShell" `
			 -Visibility Private
	Disconnect-MicrosoftTeams
}
#gavdcodeend 04

#gavdcodebegin 05
Function TeamsPsMtpTeamsUpdate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-Team -GroupId $myTeam.GroupId `
			 -Description "Team updated with PowerShell" `
			 -Visibility Public
	Disconnect-MicrosoftTeams
}
#gavdcodeend 05

#gavdcodebegin 06
Function TeamsPsMtpTeamsDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-Team -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 06

#gavdcodebegin 07
Function TeamsPsMtpTeamsHelp()
{
	Get-TeamHelp
	Disconnect-MicrosoftTeams
}
#gavdcodeend 07

#gavdcodebegin 08
Function TeamsPsMtpChannelsEnumerate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamChannel -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 08

#gavdcodebegin 09
Function TeamsPsMtpChannelsCreate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	New-TeamChannel -GroupId $myTeam.GroupId `
					-DisplayName "Test Channel from PS" 
	Disconnect-MicrosoftTeams
}
#gavdcodeend 09

#gavdcodebegin 10
Function TeamsPsMtpChannelsUpdate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-TeamChannel -GroupId $myTeam.GroupId `
					-CurrentDisplayName "Test Channel from PS" `
					-Description "This is a test Channel"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 10

#gavdcodebegin 11
Function TeamsPsMtpChannelsDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamChannel -GroupId $myTeam.GroupId `
					   -DisplayName "Test Channel from PS"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 11

#gavdcodebegin 12
Function TeamsPsMtpTeamUserEnumerate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamUser -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 12

#gavdcodebegin 13
Function TeamsPsMtpTeamUserCreate()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Add-TeamUser -GroupId $myTeam.GroupId `
				 -User "user@domain.onmicrosoft.com" 
	Disconnect-MicrosoftTeams
}
#gavdcodeend 13

#gavdcodebegin 14
Function TeamsPsMtpTeamUserDelete()
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamUser -GroupId $myTeam.GroupId `
					-User "user@domain.onmicrosoft.com"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 14

#gavdcodebegin 15
Function TeamsPsMtpPolicyPackageEnumerate()
{
	Get-CsPolicyPackage
	Disconnect-MicrosoftTeams
}
#gavdcodeend 15

#gavdcodebegin 16
Function TeamsPsMtpPolicyPackageUser()
{
	Get-CsUserPolicyPackage -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}
#gavdcodeend 16

#gavdcodebegin 17
Function TeamsPsMtpPolicyPackageUserRecommended()
{
	Get-CsUserPolicyPackageRecommendation -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}
#gavdcodeend 17

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

