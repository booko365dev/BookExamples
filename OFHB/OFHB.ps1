
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

# functions to login in Azure

function Get-AzureTokenApplication
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret,
 
		[Parameter(Mandatory=$False)]
		[String]$TenantName
	)
   
	 $LoginUrl = "https://login.microsoftonline.com"
	 $ScopeUrl = "https://graph.microsoft.com/.default"
	 
	 $myBody  = @{ Scope = $ScopeUrl; `
					grant_type = "client_credentials"; `
					client_id = $ClientID; `
					client_secret = $ClientSecret }

	 $myOAuth = Invoke-RestMethod `
					-Method Post `
					-Uri $LoginUrl/$TenantName/oauth2/v2.0/token `
					-Body $myBody

	return $myOAuth
}

function Get-AzureTokenDelegation
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	 $LoginUrl = "https://login.microsoftonline.com"
	 $ScopeUrl = "https://graph.microsoft.com/.default"

	 $myBody  = @{ Scope = $ScopeUrl; `
					grant_type = "Password"; `
					client_id = $ClientID; `
					Username = $UserName; `
					Password = $UserPw }

	 $myOAuth = Invoke-RestMethod `
					-Method Post `
					-Uri $LoginUrl/$TenantName/oauth2/v2.0/token `
					-Body $myBody

	return $myOAuth
}

# Functions to login in Teams

#gavdcodebegin 001
function PsTeamsMtp_LoginPsTeams
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-MicrosoftTeams -Credential $myCredentials
}
#gavdcodeend 001

#gavdcodebegin 018
function PsTeamsSkype_LoginPsTeams
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	
	Import-Module SkypeOnlineConnector
	$mySession = New-CsOnlineSession -Credential $myCredentials
	Import-PSSession $mySession
}
#gavdcodeend 018

#gavdcodebegin 037
function PsTeamsPnpPowerShell_LoginPsTeams
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteBaseUrl -Credentials $myCredentials
}
#gavdcodeend 037

function PsTeamsCliM365_LoginPsTeams
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 002
function PsTeamsMtp_EnumarateTeams
{
	$allTeams = Get-Team
	foreach($oneTeam in $allTeams) {
		Write-Host $oneTeam.DisplayName
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 002

#gavdcodebegin 003
function PsTeamsMtp_GetTeamsByDisplayName
{
	$oneTeam = Get-Team -DisplayName "Mark 8 Project Team"
	Write-Host $oneTeam.GroupId

	Disconnect-MicrosoftTeams
}
#gavdcodeend 003

#gavdcodebegin 004
function PsTeamsMtp_CreateTeam
{
	New-Team -DisplayName "Test Team from PS" `
			 -Description "Team created with PowerShell" `
			 -Visibility Private
	Disconnect-MicrosoftTeams
}
#gavdcodeend 004

#gavdcodebegin 005
function PsTeamsMtp_UpdateTeam
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-Team -GroupId $myTeam.GroupId `
			 -Description "Team updated with PowerShell" `
			 -Visibility Public
	Disconnect-MicrosoftTeams
}
#gavdcodeend 005

#gavdcodebegin 006
function PsTeamsMtp_DeleteTeam
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-Team -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 006

#gavdcodebegin 007
function PsTeamsMtp_TeamsHelp
{
	Get-TeamHelp
	Disconnect-MicrosoftTeams
}
#gavdcodeend 007

#gavdcodebegin 008
function PsTeamsMtp_EnumerateChannels
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	$allChannels = Get-TeamChannel -GroupId $myTeam.GroupId
	foreach($oneChannel in $allChannels) {
		Write-Host $oneChannel.DisplayName
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 008

#gavdcodebegin 009
function PsTeamsMtp_CreateChannels
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	New-TeamChannel -GroupId $myTeam.GroupId `
					-DisplayName "Test Channel from PS" 

	Disconnect-MicrosoftTeams
}
#gavdcodeend 009

#gavdcodebegin 010
function PsTeamsMtp_UpdateChannels
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-TeamChannel -GroupId $myTeam.GroupId `
					-CurrentDisplayName "Test Channel from PS" `
					-Description "This is a test Channel"

	Disconnect-MicrosoftTeams
}
#gavdcodeend 010

#gavdcodebegin 011
function PsTeamsMtp_DeleteChannels
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamChannel -GroupId $myTeam.GroupId `
					   -DisplayName "Test Channel from PS"

	Disconnect-MicrosoftTeams
}
#gavdcodeend 011

#gavdcodebegin 012
function PsTeamsMtp_EnumerateTeamUser
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	$allUsers = Get-TeamUser -GroupId $myTeam.GroupId
	foreach($oneUser in $allUsers) {
		Write-Host $oneUser.User
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 012

#gavdcodebegin 013
function PsTeamsMtp_CreateTeamUser
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Add-TeamUser -GroupId $myTeam.GroupId `
				 -User "user@domain.onmicrosoft.com"
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 013

#gavdcodebegin 014
function PsTeamsMtp_DeleteTeamUser
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamUser -GroupId $myTeam.GroupId `
					-User "user@domain.onmicrosoft.com"

	Disconnect-MicrosoftTeams
}
#gavdcodeend 014

#gavdcodebegin 015
function PsTeamsMtp_EnumeratePolicyPackage
{
	$allPolicies = Get-CsPolicyPackage
	foreach($onePolicy in $allPolicies) {
		Write-Host $onePolicy.Name
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 015

#gavdcodebegin 016
function PsTeamsMtp_PolicyPackageUser
{
	Get-CsUserPolicyPackage -Identity "user@domain.onmicrosoft.com"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 016

#gavdcodebegin 017
function PsTeamsMtp_PolicyPackageUserRecommended
{
	$allPolicies = Get-CsUserPolicyPackageRecommendation -Identity "admin@guitacadev.onmicrosoft.com"
	foreach($onePolicy in $allPolicies) {
		Write-Host $onePolicy.Name
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 017

#gavdcodebegin 030
function PsTeamsMtp_GetCsTeamTemplateList
{
	$allTemplates = Get-CsTeamTemplateList
	foreach($oneTemplate in $allTemplates) {
		Write-Host(" - " + $oneTemplate.Name + " - " + $oneTemplate.OdataId)
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 030

#gavdcodebegin 031
function PsTeamsMtp_GetCsTeamTemplate
{
	$oneTemplate = Get-CsTeamTemplate -OdataId `
		"/api/teamtemplates/v1.0/com.microsoft.teams.template.ManageAProject/Public/en-US" `
		| ConvertTo-Json
	Write-Host $oneTemplate

	Disconnect-MicrosoftTeams
}
#gavdcodeend 031

#gavdcodebegin 032
function PsTeamsMtp_GetTeamsApp
{
	$allApps = Get-TeamsApp
	foreach($oneApp in $allApps) {
		Write-Host(" - " + $oneApp.DisplayName + " - " + $oneApp.Id)
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 032

#gavdcodebegin 033
function PsTeamsMtp_GetOneTeamsAppByIdOrName
{
	$oneAppById = Get-TeamsApp -Id "fe157fa1-3f58-47ac-b66c-5eaafb55c4ad" | ConvertTo-Json 
	Write-Host $oneAppById

	$oneAppByName = Get-TeamsApp -DisplayName "Analytics 365" | ConvertTo-Json   
	Write-Host $oneAppByName
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 033

#gavdcodebegin 034
function PsTeamsMtp_NewTeamsApp
{
	New-TeamsApp -DistributionMethod "organization" `
				 -Path "C:\Temporary\App01FromDevSite.zip" 
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 034

#gavdcodebegin 035
function PsTeamsMtp_SetTeamsApp
{
	Set-TeamsApp -Id "eed59874-e471-49ca-a01f-7d92bee85fc6" `
				 -Path "C:\Temporary\App01FromDevSite.zip" 
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 035

#gavdcodebegin 036
function PsTeamsMtp_DeleteTeamsApp
{
	Remove-TeamsApp -Id "eed59874-e471-49ca-a01f-7d92bee85fc6"
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 036

#gavdcodebegin 019
function PsTeamsSkype_GetCallingPolicy
{
	#*** LEGACY CODE ***
	Get-CsTeamsCallingPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 019

#gavdcodebegin 020
function PsTeamsSkype_GetCallParkPolicy
{
	#*** LEGACY CODE ***
	Get-CsTeamsCallParkPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 020

#gavdcodebegin 021
function PsTeamsSkype_GetChannelPolicy
{
	#*** LEGACY CODE ***
	Get-CsTeamsChannelsPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 021

#gavdcodebegin 022
function PsTeamsSkype_CreateChannelPolicy
{
	#*** LEGACY CODE ***
	New-CsTeamsChannelsPolicy -Identity myPolicy -AllowPrivateTeamDiscovery $false
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 022

#gavdcodebegin 023
function PsTeamsSkype_AssignChannelPolicy
{
	#*** LEGACY CODE ***
	Grant-CsTeamsChannelsPolicy -Identity user@tenant.OnMicrosoft.com -PolicyName myPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 023

#gavdcodebegin 024
function PsTeamsSkype_ModifyChannelPolicy
{
	#*** LEGACY CODE ***
	Set-CsTeamsChannelsPolicy -Identity myPolicy -AllowPrivateTeamDiscovery $true
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 024

#gavdcodebegin 025
function PsTeamsSkype_ModifyChannelPolicy
{
	#*** LEGACY CODE ***
	Grant-CsTeamsChannelsPolicy -Identity user@tenant.OnMicrosoft.com -PolicyName Default
	Remove-CsTeamsChannelsPolicy -Identity myPolicy -Force
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 025

#gavdcodebegin 026
function PsTeamsSkype_GetTeamsClientConfiguration
{
	#*** LEGACY CODE ***
	Get-CsTeamsClientConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 026

#gavdcodebegin 027
function PsTeamsSkype_GetGuestMessagingConfiguration
{
	#*** LEGACY CODE ***
	Get-CsTeamsGuestMessagingConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 027

#gavdcodebegin 028
function PsTeamsSkype_GetMeetingBroadcastConfiguration
{
	#*** LEGACY CODE ***
	Get-CsTeamsMeetingBroadcastConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 028

#gavdcodebegin 029
function PsTeamsSkype_RemoveGoogleDrive
{
	#*** LEGACY CODE ***
	Set-CsTeamsClientConfiguration -Identity Global -AllowGoogleDrive $false
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 029

#gavdcodebegin 038
function PsTeamsPnpPowerShell_GetAllTeams
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsTeam
}
#gavdcodeend 038

#gavdcodebegin 039
function PsTeamsPnpPowerShell_GetOneTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsTeam -Identity "3216704d-ed17-4286-9206-2fa67b85780c"  # GroupID
	Get-PnPTeamsTeam -Identity "Sales and Marketing"  # DisplayName
	Get-PnPTeamsTeam -Identity "SalesAndMarketing"  # MailNickname
}
#gavdcodeend 039

#gavdcodebegin 040
function PsTeamsPnpPowerShell_NewTeamByName
{
	# Permissions required: Group.ReadWrite.All
	New-PnPTeamsTeam -DisplayName "TeamCreatedWithPnP" `
					 -Visibility Public `
					 -MailNickName "TeamCreatedWithPnPMail" `
					 -AllowUserDeleteMessages $true
}
#gavdcodeend 040

#gavdcodebegin 041
function PsTeamsPnpPowerShell_NewTeamByGroup
{
	# Permissions required: Group.ReadWrite.All
	New-PnPTeamsTeam -GroupId "89e67c39-b5b3-440d-9bcd-ac8b3743dda1" `
					 -AllowUserDeleteMessages $true
}
#gavdcodeend 041

#gavdcodebegin 042
function PsTeamsPnpPowerShell_SetTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTeam -Identity "TeamCreatedWithPnP" `
					 -DisplayName "Team Created With PnP" `
					 -Description "This is a test Team"
}
#gavdcodeend 042

#gavdcodebegin 043
function PsTeamsPnpPowerShell_SetPictureTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTeamPicture -Team "Team Created With PnP" `
							-Path "C:\Temporary\Buggy.png"
}
#gavdcodeend 043

#gavdcodebegin 044
function PsTeamsPnpPowerShell_SetArchivedTeam
{
	# Permissions required: Group.ReadWrite.All or Directory.ReadWrite.All
	Set-PnPTeamsTeamArchivedState -Identity "Team Created With PnP" `
								  -Archived $true `
								  -SetSiteReadOnlyForMembers $true
}
#gavdcodeend 044

#gavdcodebegin 045
function PsTeamsPnpPowerShell_RemoveTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsTeam -Identity "Team Created With PnP" -Force
	#Remove-PnPTeamsTeamm -GroupId "89e67c39-b5b3-440d-9bcd-ac8b3743dda1" `
}
#gavdcodeend 045

#gavdcodebegin 046
function PsTeamsPnpPowerShell_GetAllChannelsTeam
{
	# Permissions required: Group.ReadWrite.All
	Get-PnPTeamsChannel -Team "Team Created With PnP"
}
#gavdcodeend 046

#gavdcodebegin 047
function PsTeamsPnpPowerShell_GetOneChannelTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsChannel -Team "Team Created With PnP" `
						-Identity "19:Gl2G3C9_9cGRxZtYjExQ7rx6aAw1@thread.tacv2"
}
#gavdcodeend 047

#gavdcodebegin 131
function PsTeamsPnpPowerShell_GetOneChannelFilesFolder
{
	# Permissions required: Group.ReadWrite.All
	Get-PnPTeamsChannelFilesFolder -Team "Team Created With PnP" `
								   -Channel "General"
}
#gavdcodeend 131

#gavdcodebegin 048
function PsTeamsPnpPowerShell_AddOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Add-PnPTeamsChannel -Team "Team Created With PnP" `
						-DisplayName "Channel Created With PnP"
}
#gavdcodeend 048

#gavdcodebegin 049
function PsTeamsPnpPowerShell_UpdateOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsChannel -Team "Team Created With PnP" `
						-Identity "Channel Created With PnP" `
						-DisplayName "Channel Updated With PnP" `
						-Description "This is a test Channel" `
						-IsFavoriteByDefault $true
}
#gavdcodeend 049

#gavdcodebegin 050
function PsTeamsPnpPowerShell_SendMessageToOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Submit-PnPTeamsChannelMessage -Team "Team Created With PnP" `
								  -Channel "Channel Updated With PnP" `
								  -Message "<strong>This is a Channel message</strong>" `
								  -ContentType "Html" `
								  -Important
}
#gavdcodeend 050

#gavdcodebegin 051
function PsTeamsPnpPowerShell_GetMessagesFromOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myMessages = Get-PnPTeamsChannelMessage -Team "Team Created With PnP" `
											 -Channel "Channel Updated With PnP" `
											 -IncludeDeleted

	foreach($oneMessage in $myMessages) {
		Write-Host $oneMessage.Id " - " `
				   $oneMessage.Body.Content " - " `
				   $oneMessage.CreatedDateTime
	}
}
#gavdcodeend 051

#gavdcodebegin 132
function PsTeamsPnpPowerShell_GetReplayMessageOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myReplay = Get-PnPTeamsChannelMessageReply -Team "Team Created With PnP" `
									-Channel "Channel Updated With PnP" `
									-Message "1712761789060" `
									-Identity "1712763578181"
	Write-Host $myReplay.Body
}
#gavdcodeend 132

#gavdcodebegin 052
function PsTeamsPnpPowerShell_RemoveOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsChannel -Team "Team Created With PnP" `
						   -Identity "Channel Updated With PnP"
}
#gavdcodeend 052

#gavdcodebegin 053
function PsTeamsPnpPowerShell_GetAllTabsChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myTabs = Get-PnPTeamsTab -Team "Team Created With PnP" `
							   -Channel "General"

	foreach($oneTab in $myTabs) {
		Write-Host $oneTab.Id " - " $oneTab.DisplayName
	}
}
#gavdcodeend 053

#gavdcodebegin 054
function PsTeamsPnpPowerShell_GetOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$oneTab = Get-PnPTeamsTab -Team "Team Created With PnP" `
							  -Channel "General" `
							  -Identity "Notes"

	Write-Host $oneTab.Id
}
#gavdcodeend 054

#gavdcodebegin 055
function PsTeamsPnpPowerShell_AddOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myDocsUrl = $configFile.appsettings.SiteBaseUrl + `
								"/sites/TeamCreatedWithPnPMail/Shared Documents"
	Add-PnPTeamsTab -Team "Team Created With PnP" `
					-Channel "General" `
					-DisplayName "My Documents" `
					-Type "DocumentLibrary" `
					-ContentUrl $myDocsUrl
}
#gavdcodeend 055

#gavdcodebegin 056
function PsTeamsPnpPowerShell_UpdateOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTab -Team "Team Created With PnP" `
					-Channel "General" `
					-Identity "My Documents" `
					-DisplayName "My Documents Library"
}
#gavdcodeend 056

#gavdcodebegin 057
function PsTeamsPnpPowerShell_DeleteOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsTab -Team "Team Created With PnP" `
					   -Channel "General" `
					   -Identity "My Documents Library" `
					   -Force
}
#gavdcodeend 057

#gavdcodebegin 058
function PsTeamsPnpPowerShell_GetAllUsersTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	$myUsers = Get-PnPTeamsUser -Team "Team Created With PnP"

	foreach($oneUser in $myUsers) {
		Write-Host  $oneUser.Id " - " `
					$oneUser.UserPrincipalName  " - " `
					$oneUser.DisplayName " - " `
					$oneUser.UserType
	}
}
#gavdcodeend 058

#gavdcodebegin 059
function PsTeamsPnpPowerShell_GetAllUsersChannelTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	$myUsers = Get-PnPTeamsUser -Team "Team Created With PnP" `
							    -Channel "My Private Channel"

	foreach($oneUser in $myUsers) {
		Write-Host  $oneUser.Id " - " `
					$oneUser.UserPrincipalName  " - " `
					$oneUser.DisplayName " - " `
					$oneUser.UserType
	}
}
#gavdcodeend 059

#gavdcodebegin 060
function PsTeamsPnpPowerShell_GetAllUsersChannelByRoleTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	$myUsers = Get-PnPTeamsUser -Team "Team Created With PnP" `
							    -Channel "My Private Channel" `
								-Role "Owner"

	foreach($oneUser in $myUsers) {
		Write-Host  $oneUser.Id " - " `
					$oneUser.UserPrincipalName  " - " `
					$oneUser.DisplayName " - " `
					$oneUser.UserType
	}
}
#gavdcodeend 060

#gavdcodebegin 061
function PsTeamsPnpPowerShell_AddOneUserTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Add-PnPTeamsUser -Team "Team Created With PnP" `
					 -User "user@domain.onmicrosoft.com" `
					 -Role "Member"
}
#gavdcodeend 061

#gavdcodebegin 129
function PsTeamsPnpPowerShell_AddOneUserChannel
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Add-PnPTeamsUser -Team "Team Created With PnP" `
					 -Channel "My Private Channel" `
					 -User "user@domain.onmicrosoft.com" `
					 -Role "Member"
}
#gavdcodeend 129

#gavdcodebegin 130
function PsTeamsPnpPowerShell_DeleteOneUserChannel
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Remove-PnPTeamsChannelUser -Team "Team Created With PnP" `
							   -Channel "My Private Channel" `
							   -Identity "user@domain.onmicrosoft.com" `
							   -Force
}
#gavdcodeend 130

#gavdcodebegin 062
function PsTeamsPnpPowerShell_DeleteOneUserTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Remove-PnPTeamsUser -Team "Team Created With PnP" `
						-User "user@domain.onmicrosoft.com" `
						-Role "Member"
}
#gavdcodeend 062

#gavdcodebegin 063
function PsTeamsPnpPowerShell_GetAllAppsTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	$myApps = Get-PnPTeamsApp

	foreach($oneApp in $myApps) {
		Write-Host  $oneApp.Id " - " `
					$oneApp.DisplayName " - " `
					$oneApp.DistributionMethod " - " `
					$oneApp.ExternalId
	}
}
#gavdcodeend 063

#gavdcodebegin 064
function PsTeamsPnpPowerShell_GetOneAppTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	$myApp = Get-PnPTeamsApp -Identity "Salesforce"
	#$myApp = Get-PnPTeamsApp -Identity "d6e4a9b6-646c-32fc-88ba-a6dd6150d1f7"

	Write-Host  $myApp.Id " - " `
				$myApp.DisplayName " - " `
				$myApp.DistributionMethod " - " `
				$myApp.ExternalId
}
#gavdcodeend 064

#gavdcodebegin 065
function PsTeamsPnpPowerShell_AddOneAppTeam
{
	# Permissions required: AppCatalog.ReadWrite.All or Directory.ReadWrite.All
	New-PnPTeamsApp -Path "C:\Temporary\App01FromDevSite.zip"
}
#gavdcodeend 065

#gavdcodebegin 066
function PsTeamsPnpPowerShell_UpdateOneAppTeam
{
	# Permissions required: Group.ReadWrite.All
	Update-PnPTeamsApp -Identity "1e67180b-1904-4637-91b5-fa09420953f6" `
					   -Path "C:\Temporary\App01FromDevSite.zip"
}
#gavdcodeend 066

#gavdcodebegin 067
function PsTeamsPnpPowerShell_DeleteOneAppTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsApp -Identity "App01FromDevSite" -Force
	#Remove-PnPTeamsApp -Identity "1e67180b-1904-4637-91b5-fa09420953f6" -Force
}
#gavdcodeend 067

#gavdcodebegin 068
function PsTeamsGraphRest_GetJoinedTeams
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/me/joinedTeams"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allTeams = ConvertFrom-Json –InputObject $myResult
	foreach($oneTeam in $allTeams) {
		$oneTeam.value.displayName 
	}
}
#gavdcodeend 068 

#gavdcodebegin 069
function PsTeamsGraphRest_GetAllTeamsByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName," + `
															"resourceProvisioningOptions"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allTeams = ConvertFrom-Json –InputObject $myResult
	foreach($oneTeam in $allTeams) {
		$oneTeam.value.displayName
	}
}
#gavdcodeend 069 

#gavdcodebegin 070
function PsTeamsGraphRest_GetOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$groepId = "1a348563-cdb8-42c8-9686-7ad64e2db3fd"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $groupId
	#$Url = "https://graph.microsoft.com/v1.0/teams/" + $groepId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 070 

#gavdcodebegin 071
function PsTeamsGraphRest_CreateOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, 
	#								Group.ReadWrite.All, Team.Create

	$teamTemplate = "standard"
	$Url = "https://graph.microsoft.com/v1.0/teams"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	# NOTE: The value of $myBody must be in one code line
	$myBody = '{ 
		"template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(''' + 
			$teamTemplate + ''')", 
			"displayName": "Team created with Graph", 
			"description": "This is a Team created with Graph" }'
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 071

#gavdcodebegin 072
function PsTeamsGraphRest_CreateOneGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, 
	#								Group.ReadWrite.All, Team.Create

	$Url = "https://graph.microsoft.com/v1.0/groups"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName': 'Group Created with Graph', `
				 'mailNickname': 'GroupCreatedWithGraph', `
				 'description': 'This is a Group created with Graph', `
				 'visibility': 'Private', `
				 'groupTypes': ['Unified'], `
				 'mailEnabled': true, `
				 'securityEnabled': false, `
				 'members@odata.bind': [ `
		'https://graph.microsoft.com/v1.0/users/c295c60c-f4cb-4965-9a30-7ec81ea0e192', `
		'https://graph.microsoft.com/v1.0/users/55681f61-1bb6-46c1-b59b-0270d82326d1', `
		'https://graph.microsoft.com/v1.0/users/7e0297ea-75b8-4b49-953a-c4ade11bc132' `
				 ], `
				 'owners@odata.bind': [ `
		'https://graph.microsoft.com/v1.0/users/e5855162-8ea4-40b5-baa6-e00b53a8121b' `
				] }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 072

#gavdcodebegin 073
function PsTeamsGraphRest_CreateOneTeamFromGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, 
	#								Group.ReadWrite.All, Team.Create

	$grpId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$teamTemplate = "standard"
	$Url = "https://graph.microsoft.com/v1.0/teams"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	# NOTE: The value of $myBody must be in one code line
	$myBody = '{ `
	 "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(''' +`
	 $teamTemplate + ''')", `
	 "group@odata.bind": "https://graph.microsoft.com/v1.0/groups(''' + $grpId + ''')" }'
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 073

#gavdcodebegin 074
function PsTeamsGraphRest_UpdateOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1a348563-cdb8-42c8-9686-7ad64e2db3fd"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName': 'Updated with Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 074 

#gavdcodebegin 075
function PsTeamsGraphRest_DeleteOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1a348563-cdb8-42c8-9686-7ad64e2db3fd"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 075

#gavdcodebegin 076
function PsTeamsGraphRest_GetAllChannelsInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allChannels = ConvertFrom-Json –InputObject $myResult
	foreach($oneChannel in $allChannels) {
		Write-Host $oneChannel.value.displayName " - " $oneChannel.value.id
	}
}
#gavdcodeend 076

#gavdcodebegin 077
function PsTeamsGraphRest_GetOneChannelInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:004e76dacbfc4ace9589f2f415c8bf23@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																		$channelId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 077 

#gavdcodebegin 078
function PsTeamsGraphRest_CreateOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Channel.Create

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
		'displayName': 'Channel created with Graph', `
		'description': 'This is a Channel created with Graph' `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 078

#gavdcodebegin 079
function PsTeamsGraphRest_UpdateOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:627b848f36344dd6aea92cd941bd3e26@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																		$channelId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName': 'Channel Updated with Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 079

#gavdcodebegin 080
function PsTeamsGraphRest_DeleteOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:004e76dacbfc4ace9589f2f415c8bf23@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																		$channelId

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 080

#gavdcodebegin 081
function PsTeamsGraphRest_GetAllTabsInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
															$channelId + "/tabs"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allTabs = ConvertFrom-Json –InputObject $myResult
	foreach($oneTab in $allTabs) {
		$oneTab.value.displayName
	}
}
#gavdcodeend 081

#gavdcodebegin 082
function PsTeamsGraphRest_GetOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$tabId = "22bd75b6-987b-465e-b420-a80098e9527b"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
													   $channelId + "/tabs/" + $tabId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 082

#gavdcodebegin 083
function PsTeamsGraphRest_CreateOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, 
	#								Group.ReadWrite.All, Channel.Create

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/tabs"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBind = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + ` 
										"com.microsoft.teamspace.tab.files.sharepoint"
	$myUrl = $configFile.appsettings.SiteBaseUrl + "/sites/TeamCreatedWithGraph/" + `
										"Shared%20Documents"
	$myBody = "{ `
		'displayName': 'Document Library', `
		'teamsApp@odata.bind': '" + $myBind + "', `
		'configuration': { `
			'entityId': '', `
			'contentUrl': '" + $myUrl + "', `
			'removeUrl': null, `
			'websiteUrl': null `
	}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 083

#gavdcodebegin 084
function PsTeamsGraphRest_UpdateOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$tabId = "a0a73e10-3212-4a1b-bf9a-aca4b84e89ce"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
											$channelId + "/tabs/" + $tabId
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ 'displayName': 'My Docs' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)"; `
				   'If-Match' = "$($eTag)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Patch `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 084

#gavdcodebegin 085
function PsTeamsGraphRest_DeleteOneTabFromOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$tabId = "a0a73e10-3212-4a1b-bf9a-aca4b84e89ce"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
											$channelId + "/tabs/" + $tabId

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 085

#gavdcodebegin 086
function PsTeamsGraphRest_GetAllUsersInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId + "/members"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allUsers = ConvertFrom-Json –InputObject $myResult
	foreach($oneUser in $allUsers) {
		$oneUser.value.displayName
	}
}
#gavdcodeend 086

#gavdcodebegin 133
function PsTeamsGraphRest_AddOneUserToOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, 
	#								Group.ReadWrite.All, Channel.Create

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$userId = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId + "/members/`$ref"

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$userUri = "https://graph.microsoft.com/v1.0/directoryObjects/" + $userId
	$myBody = '{ "@odata.id": "' + $userUri + '" }'
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 133

#gavdcodebegin 134
function PsTeamsGraphRest_DeleteOneUserFromOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$userId = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId + "/members/" + `
											$userId + "/`$ref"

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 134

#gavdcodebegin 087
function PsTeamsGraphRest_SendMessageToOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: ChannelMessage.Send

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/messages"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
		'body': {`
			'contentType': 'html', `
			'content': '<strong>This is a Channel message</strong>' `
		}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
											-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 087

#gavdcodebegin 088
function PsTeamsGraphRest_GetAllMessagesChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.Read, Chat.ReadWrite

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/messages"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 088

#gavdcodebegin 089
function PsTeamsGraphRest_SendMessageReplayToOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: ChannelMessage.Send, Group.ReadWrite.All

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$messageId = "1712863079783"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
									$channelId + "/messages/" + $messageId + "/replies"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
		'body': {`
			'contentType': 'html', `
			'content': '<strong>This is a replay to the Channel message</strong>' `
		}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 089

#gavdcodebegin 090
function PsTeamsGraphRest_GetAllReplaysToOneMessagesChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.Read, Chat.ReadWrite

	$teamId = "1dac89ce-2c60-4c24-aa69-41ee6c7e2df1"
	$channelId = "19:jllIFaFLhczhNNPNzVw1o4WzcrAO9H6cJ6EROSz_jiI1@thread.tacv2"
	$messageId = "1712863079783"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
									$channelId + "/messages/" + $messageId + "/replies"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 090

#gavdcodebegin 152
function PsTeamsGraphRest_GetAllMeetings
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$startMeeting = "2024-04-09T01:00:00Z"
	$endMeeting = "2024-04-16T23:59:59Z"
	$Url = "https://graph.microsoft.com/v1.0/me/events/?$filter=" + `
					"start/DateTime ge '" + $startMeeting + "' AND " + `
					"end/DateTime le '" + $endMeeting + "'"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 152

#gavdcodebegin 153
function PsTeamsGraphRest_GetOneMeeting
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$meetingUrl = "https://outlook.office365.com/owa/?itemid=AAMkAG..."
	$Url = "https://graph.microsoft.com/v1.0/me/events/?$filter=" + `
					"joinWebURL eq '" + $startMeeting + "'"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 153

#gavdcodebegin 154
function PsTeamsGraphRest_CreateOneMeeting
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/onlineMeetings"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
			'startDateTime': '2024-04-16T10:00:00Z', `
			'endDateTime': '2024-04-16T11:55:00Z' `	
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 154

#gavdcodebegin 155
function PsTeamsGraphRest_DeleteOneMeeting
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$meetingId = "AAMkAGE0ODQ3...AAENAAC1vtBLB-F9SJ2ZDb7Xo-OrAAGb3qfbAAA="
	$Url = "https://graph.microsoft.com/v1.0/me/onlineMeetings/" + $meetingId

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 155

#gavdcodebegin 156
function PsTeamsGraphRest_GetAllChats
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$Url = "https://graph.microsoft.com/v1.0/me/chats"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 156

#gavdcodebegin 157
function PsTeamsGraphRest_GetOneChat
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-96...8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 157

#gavdcodebegin 158
function PsTeamsGraphRest_GetOneChatMessages
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-96...9c1-2246d8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/messages"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 158

#gavdcodebegin 159
function PsTeamsGraphRest_GetAllChatsMessages
{
	# App Registration type:		Application
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-...46d8b7b9fb@unq.gbl.spaces"
	$messagesTop = "10"
	$beginDate = "2024-04-09T01:00:00Z"
	$endDate = "2024-04-16T23:59:59Z"
	$userId = "9c251cf6-afc8-3720-47b7-a5ff4257ade5"
	$Url = "https://graph.microsoft.com/v1.0/users/chats/" + $userId + `
							"chats/getAllMessages?$top=" + $messagesTop + `
							"$filter=lastModifiedDateTime gt " + $beginDate + `
							" and lastModifiedDateTime lt " + $endDate
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 159

#gavdcodebegin 160
function PsTeamsGraphRest_GetOneChatParticipants
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-4...fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/members"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 160

#gavdcodebegin 161
function PsTeamsGraphRest_GetOneChatOneParticipant
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-96...1-2246d8b7b9fb@unq.gbl.spaces"
	$memberId = "MCMjMCMjYWRlNTYwN...DZmZTVjYy00NjJhLTRhNjAtYjljMS0yMjQ2ZDhiN2I5ZmI="
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + `
										"/members/" + $memberId 
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 161

#gavdcodebegin 162
function PsTeamsGraphRest_AddOneUserToChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...46d8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/members"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
		'@odata.type': '#microsoft.graph.aadUserConversationMember', `
		'user@odata.bind': 'https://graph.microsoft.com/v1.0/users/3ce805...2cf0db9d', `
		'visibleHistoryStartDateTime': '0001-01-01T00:00:00Z', `
		'roles': ['member']
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 162

#gavdcodebegin 163
function PsTeamsGraphRest_DeleteOneUserFromChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-960...6d8b7b9fb@unq.gbl.spaces"
	$memberId = "MCMjMCMjYWRlNTYwN...ZmZTVjYy00NjJhLTRhNjAtYjljMS0yMjQ2ZDhiN2I5ZmI="
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + `
													"members/" + $memberId

	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 163

#gavdcodebegin 164
function PsTeamsGraphRest_SendMessageToChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-9...d8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/messages"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	  'body': { `
		 'content': 'Message sent using MS Graph' `
		} `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 164

#gavdcodebegin 165
function PsTeamsGraphRest_HideChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/hideForUser"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	  'user': { `
		'id': 'acc28fcb-9c61-a8f8-960b-715d2f98a431', `
		'tenantId': 'ade56059-a6c0-45cd-9f73-e4772a8168ca'
	   } `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 165

#gavdcodebegin 166
function PsTeamsGraphRest_UnhideChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/unhideForUser"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	  'user': { `
		'id': 'acc28fcb-9c61-a8f8-960b-715d2f98a431', `
		'tenantId': 'ade56059-a6c0-45cd-9f73-e4772a8168ca'
	   } `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 166

#gavdcodebegin 167
function PsTeamsGraphRest_PinChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...46d8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/pinnedMessages"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	'message@odata.bind': 'https://graph.microsoft.com/v1.0/chats/[cid]/messages/[mid]' `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 167

#gavdcodebegin 168
function PsTeamsGraphRest_GetPinnedChats
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.ReadBasic, Chat.Read, Chat.ReadWrite

	$chatId = "19:acc28fcb-5261-47f8-96...6d8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + "/pinnedMessages"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 168

#gavdcodebegin 169
function PsTeamsGraphRest_UnpinChat
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...46d8b7b9fb@unq.gbl.spaces"
	$messageId = "1713019794330"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + `
						"/pinnedMessages/" + $messageId
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
	
	Write-Host $myResult
}
#gavdcodeend 169

#gavdcodebegin 170
function PsTeamsGraphRest_ReadChatForUser
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + `
														"/markChatReadForUser"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	  'user': { `
		'id': 'acc28fcb-9c61-a8f8-960b-715d2f98a431', `
		'tenantId': 'ade56059-a6c0-45cd-9f73-e4772a8168ca'
	   } `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 170

#gavdcodebegin 171
function PsTeamsGraphRest_UnreadChatForUser
{
	# App Registration type:		Delegation
	# App Registration permissions: OnlineMeetings.Read, OnlineMeetings.ReadWrite

	$chatId = "19:acc28fcb-5261-47f...8b7b9fb@unq.gbl.spaces"
	$Url = "https://graph.microsoft.com/v1.0/me/chats/" + $chatId + `
														"/markChatUnreadForUser"
	# For Application registration use "/users/userId/" instead of "/me/"
	
	$myOAuth = Get-AzureTokenDelegation `
							-ClientID $configFile.appsettings.ClientIdWithAccPw `
							-TenantName $configFile.appsettings.TenantName `
							-UserName $configFile.appsettings.UserName `
							-UserPw $configFile.appsettings.UserPw
	
	$myBody = "{ `
	  'user': { `
		'id': 'acc28fcb-9c61-a8f8-960b-715d2f98a431', `
		'tenantId': 'ade56059-a6c0-45cd-9f73-e4772a8168ca'
	   } `
	}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType
	
	Write-Host $myResult
}
#gavdcodeend 171

#gavdcodebegin 091
function PsTeamsCliM365_GetAllTeams
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team list

	m365 logout
}
#gavdcodeend 091

#gavdcodebegin 092
function PsTeamsCliM365_GetTeamsByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team list --output json --query "[?displayName == 'Sales and Marketing']"

	m365 logout
}
#gavdcodeend 092

#gavdcodebegin 093
function PsTeamsCliM365_GetOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team get --id "3216704d-ed17-4286-9206-2fa67b85780c"
	#m365 teams team get --name "Team Cloned With CLI"

	m365 logout
}
#gavdcodeend 093

#gavdcodebegin 094
function PsTeamsCliM365_CreateOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team add --name "Team Created With CLI" `
						--description "Team Created With the CLI" `
						--wait

	m365 logout
}
#gavdcodeend 094

#gavdcodebegin 095
function PsTeamsCliM365_CloneOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team clone --id "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						  --name "Team Cloned With CLI" `
						  --description "Team Cloned With the CLI" `
						  --partsToClone "apps,tabs,settings,channels,members" `
						  --visibility "public"

	m365 logout
}
#gavdcodeend 095

#gavdcodebegin 096
function PsTeamsCliM365_UpdateOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team set --id "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						--description "Team Updated With the CLI"

	m365 logout
}
#gavdcodeend 096

#gavdcodebegin 097
function PsTeamsCliM365_ArchiveOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team archive --id "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						    --shouldSetSpoSiteReadOnlyForMembers

	m365 logout
}
#gavdcodeend 097

#gavdcodebegin 098
function PsTeamsCliM365_UnarchiveOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team unarchive --id "02e80b46-223e-4dfa-bbe5-c57fd5a28a95"

	m365 logout
}
#gavdcodeend 098

#gavdcodebegin 099
function PsTeamsCliM365_DeleteOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams team remove --name "Team Cloned With CLI" `
						   --force

	m365 logout
}
#gavdcodeend 099

#gavdcodebegin 100
function PsTeamsCliM365_GetAllChannelsOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel list --teamName "Team Created With CLI"

	m365 logout
}
#gavdcodeend 100

#gavdcodebegin 101
function PsTeamsCliM365_GetChannelByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel list --output json `
							--teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
							--query "[?displayName == 'General']"

	m365 logout
}
#gavdcodeend 101

#gavdcodebegin 102
function PsTeamsCliM365_GetOneChannelFromOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel get --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						   --id "19:Ok65JBLI9xaKjIYvXyXhyxxxeak1@thread.tacv2"

	m365 logout
}
#gavdcodeend 102

#gavdcodebegin 103
function PsTeamsCliM365_CreateOneChannelInOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel add --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						   --name "Channel Created With CLI" `
						   --description "Channel Created With the CLI"

	m365 logout
}
#gavdcodeend 103

#gavdcodebegin 104
function PsTeamsCliM365_UpdateOneChannelInOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel set --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						   --name "Channel Created With CLI" `
						   --description "Channel Updated With the CLI PnP"

	m365 logout
}
#gavdcodeend 104

#gavdcodebegin 105
function PsTeamsCliM365_DeleteOneChannelFromOneTeam
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel remove --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						      --name "Channel Created With CLI" `
						      --force

	m365 logout
}
#gavdcodeend 105

#gavdcodebegin 106
function PsTeamsCliM365_GetAllTabs
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams tab list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
					    --channelId "19:Ok65JBLI9xaKjIxxx1-uMATiMsgaeak1@thread.tacv2"

	m365 logout
}
#gavdcodeend 106

#gavdcodebegin 107
function PsTeamsCliM365_GetTabByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams tab list --output json `
						--teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						--channelId "19:Ok65JBLI9xaKjIxxx1-uMATiMsgaeak1@thread.tacv2" `
						--query "[?displayName == 'Files']"

	m365 logout
}
#gavdcodeend 107

#gavdcodebegin 108
function PsTeamsCliM365_GetOneTab
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams tab get --teamName "Team Created with CLI" `
					   --channelName "General" `
					   --name "Files"

	m365 logout
}
#gavdcodeend 108

#gavdcodebegin 109
function PsTeamsCliM365_AddOneTab
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams tab add --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
					   --channelId "19:Ok65JBLI9xaKjxxxM1-uMATiMsgaeak1@thread.tacv2" `
					   --appId "e2acbf5d-6a4f-4d35-a760-503dc0faf314" `
					   --appName "Guitaca Site" `
					   --contentUrl "https://guitaca.com"

	m365 logout
}
#gavdcodeend 109

#gavdcodebegin 110
function PsTeamsCliM365_DeleteOneTab
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams tab remove --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						  --channelId "19:Ok65J9xaKjxxxM1-uMATiMsgaeak1@thread.tacv2" `
					      --tabId "0ef6aadb-ff6c-450f-aa2f-085be8fc1d21" `
						  --force

	m365 logout
}
#gavdcodeend 110

#gavdcodebegin 111
function PsTeamsCliM365_GetAllUsers
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams user list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
					     --role "Owner"

	m365 logout
}
#gavdcodeend 111

#gavdcodebegin 112
function PsTeamsCliM365_GetUserByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams user list --output json `
						 --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						 --role "Owner" `
						 --query "[?displayName == 'Admin']"

	m365 logout
}
#gavdcodeend 112

#gavdcodebegin 136
function PsTeamsCliM365_GetAllUsersChannel
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel member list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
								   --channelName "My Private Channel" `
								   --role "owner"

	m365 logout
}
#gavdcodeend 136

#gavdcodebegin 113
function PsTeamsCliM365_AddOneUser
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams user add --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
					    --userName "user@domain.onmicrosoft.com" `
						--role "Member"

	m365 logout
}
#gavdcodeend 113

#gavdcodebegin 137
function PsTeamsCliM365_AddOneUserToOneChannel
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel member add --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
								  --channelName "My Private Channel" `
								  --userIds "user@domain.onmicrosoft.com" `
								  --owner

	m365 logout
}
#gavdcodeend 137

#gavdcodebegin 114
function PsTeamsCliM365_UpdateOneUser
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel member set --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
								  --channelName "My Private Channel" `
								  --userName "user@domain.onmicrosoft.com" `
								  --role "member"

	m365 logout
}
#gavdcodeend 114

#gavdcodebegin 115
function PsTeamsCliM365_DeleteOneUser
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 entra m365group user remove --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
								     --userName "user@domain.onmicrosoft.com" `
								     --force

	m365 logout
}
#gavdcodeend 115

#gavdcodebegin 138
function PsTeamsCliM365_DeleteOneUserFromOneChannel
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams channel member remove --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
									 --channelName "My Private Channel" `
								     --userName "adelev@guitacadev.onmicrosoft.com" `
								     --force

	m365 logout
}
#gavdcodeend 138

#gavdcodebegin 116
function PsTeamsCliM365_GetAllApps
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app list
	Write-Host ("-------")
	m365 teams app list --distributionMethod "store"
	Write-Host ("-------")
	m365 teams app list --distributionMethod "organization"
	Write-Host ("-------")
	m365 teams app list --distributionMethod "sideloaded"

	m365 logout
}
#gavdcodeend 116

#gavdcodebegin 117
function PsTeamsCliM365_GetAppByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app list --output json `
						--query "[?displayName == 'MailChimp']"

	m365 logout
}
#gavdcodeend 117

#gavdcodebegin 118
function PsTeamsCliM365_AddOneApp
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app install --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
					       --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b"

	m365 logout
}
#gavdcodeend 118

#gavdcodebegin 119
function PsTeamsCliM365_PublishOneApp
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app publish --filePath "C:\Projects\MyApp.zip"

	m365 logout
}
#gavdcodeend 119

#gavdcodebegin 120
function PsTeamsCliM365_UpdateOneApp
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app update --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b" `
						  --filePath "C:\Projects\MyApp.zip"

	m365 logout
}
#gavdcodeend 120

#gavdcodebegin 121
function PsTeamsCliM365_UninstallOneApp
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app uninstall --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					         --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b" `
							 --force

	m365 logout
}
#gavdcodeend 121

#gavdcodebegin 122
function PsTeamsCliM365_DeleteOneApp
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams app remove --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b0" `
						  --force

	m365 logout
}
#gavdcodeend 122

#gavdcodebegin 123
function PsTeamsCliM365_GetAllMessages
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams message list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						    --channelId "19:Ok65JBLxxxWVM1-uMATiMsgaeak1@thread.tacv2"

	m365 logout
}
#gavdcodeend 123

#gavdcodebegin 124
function PsTeamsCliM365_GetMessageByQuery
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams message list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						    --channelId "19:Ok65JxxxWVM1-uMATiMsgaeak1@thread.tacv2" `
							--output json `
							--query "[?id == '1712933217065']"

	m365 logout
}
#gavdcodeend 124

#gavdcodebegin 125
function PsTeamsCliM365_GetOneMessage
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams message get --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						   --channelId "19:Ok65xxxKoOWVM1-uMATiMsgaeak1@thread.tacv2" `
						   --id "1712933217065"

	m365 logout
}
#gavdcodeend 125

#gavdcodebegin 139
function PsTeamsCliM365_SendMessageToOneChannel
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams message send --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
						    --channelId "19:Ok65JBxxxoOWVM1-uMATiMsgaeak1@thread.tacv2" `
						    --message "Message sent by the CLI"

	m365 logout
}
#gavdcodeend 139

#gavdcodebegin 126
function PsTeamsCliM365_GetMessageReplays
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams message reply list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
								  --channelId "19:Ok6LxxxoOWVM1-uMATaeak1@thread.tacv2" `
								  --messageId "1712933217065"

	m365 logout
}
#gavdcodeend 126

#gavdcodebegin 140
function PsTeamsCliM365_GetAllMeetings
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams meeting list --startDateTime "2024-01-01T10:00:00Z" `
							--endDateTime "2024-04-30T23:59:59Z"

	m365 logout
}
#gavdcodeend 140

#gavdcodebegin 141
function PsTeamsCliM365_GetOneMeeting
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams meeting get --userName "user@domain.onmicrosoft.com" `
						   --joinUrl "https://teams.microsoft.com/l/meetup-join/19%..."

	m365 logout
}
#gavdcodeend 141

#gavdcodebegin 142
function PsTeamsCliM365_CreateOneMeeting
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams meeting add --subject "Meeting created with the CLI" `
						   --participantUserNames "user@domain.onmicrosoft.com" `
						   --startTime "2024-04-13T11:00:00Z" `
						   --endTime "2024-04-13T11:55:00Z"

	m365 logout
}
#gavdcodeend 142

#gavdcodebegin 143
function PsTeamsCliM365_AttendanceMeeting
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams meeting attendancereport list --meetingId "MSphY2MyO..."

	m365 logout
}
#gavdcodeend 143

#gavdcodebegin 144
function PsTeamsCliM365_GetAllChats
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat list

	m365 logout
}
#gavdcodeend 144

#gavdcodebegin 145
function PsTeamsCliM365_GetOneChat
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat get --id "19:acc28fcb-5261-47f8-960b-..."

	m365 logout
}
#gavdcodeend 145

#gavdcodebegin 146
function PsTeamsCliM365_GetOneChatParticipants
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat member list --chatId "19:acc28fcb-5261-47f8-960b-..."

	m365 logout
}
#gavdcodeend 146

#gavdcodebegin 147
function PsTeamsCliM365_AddOneChatParticipant
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat member add --chatId "19:acc28fcb-5261-47f8-960b-..." `
							   --userName "user@domain.onmicrosoft.com" `
							   --role "guest" `
							   --includeAllHistory

	m365 logout
}
#gavdcodeend 147

#gavdcodebegin 148
function PsTeamsCliM365_DeleteOneChatParticipant
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat member remove --chatId "19:acc28fcb-5261-47f8-960b-..." `
								  --userName "adelev@guitacadev.onmicrosoft.com" `
								  --force

	m365 logout
}
#gavdcodeend 148

#gavdcodebegin 149
function PsTeamsCliM365_GetChatMessages
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat message list --chatId "19:acc28fcb-5261-47f8-960b-..."

	m365 logout
}
#gavdcodeend 149

#gavdcodebegin 150
function PsTeamsCliM365_SendChatMessageToChat
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat message send --chatId "19:acc28fcb-5261-47f8-960b-..." `
								 --message "Message to Chat sent using the CLI"

	m365 logout
}
#gavdcodeend 150

#gavdcodebegin 151
function PsTeamsCliM365_SendChatMessageToPerson
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams chat message send --userEmails "user@domain.onmicrosoft.com" `
								 --message "Message to user sent using the CLI"

	m365 logout
}
#gavdcodeend 151

#gavdcodebegin 127
function PsTeamsCliM365_GetSettings
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams membersettings list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95"
	Write-Host ("-------")
	m365 teams guestsettings list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95"
	Write-Host ("-------")
	m365 teams messagingsettings list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95"
	Write-Host ("-------")
	m365 teams funsettings list --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95"

	m365 logout
}
#gavdcodeend 127

#gavdcodebegin 128
function PsTeamsCliM365_SetSettings
{
	PsTeamsCliM365_LoginPsTeams
	
	m365 teams funsettings set --teamId "02e80b46-223e-4dfa-bbe5-c57fd5a28a95" `
							   --allowGiphy false

	m365 logout
}
#gavdcodeend 128


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 171 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------------------------ Using Teams cmdlets

#PsTeamsMtp_LoginPsTeams

#PsTeamsMtp_TeamsHelp
#PsTeamsMtp_EnumarateTeams
#PsTeamsMtp_GetTeamsByDisplayName
#PsTeamsMtp_CreateTeam
#PsTeamsMtp_UpdateTeam
#PsTeamsMtp_DeleteTeam

#PsTeamsMtp_EnumerateChannels
#PsTeamsMtp_CreateChannels
#PsTeamsMtp_UpdateChannels
#PsTeamsMtp_DeleteChannels

#PsTeamsMtp_EnumerateTeamUser
#PsTeamsMtp_CreateTeamUser
#PsTeamsMtp_DeleteTeamUser

#PsTeamsMtp_EnumeratePolicyPackage
#PsTeamsMtp_PolicyPackageUser
#PsTeamsMtp_PolicyPackageUserRecommended

#PsTeamsMtp_GetCsTeamTemplateList
#PsTeamsMtp_GetCsTeamTemplate
#PsTeamsMtp_GetTeamsApp
#PsTeamsMtp_GetOneTeamsAppByIdOrName
#PsTeamsMtp_NewTeamsApp
#PsTeamsMtp_SetTeamsApp
#PsTeamsMtp_DeleteTeamsApp

#------------------------ Using Skype For Business cmdlets

#PsTeamsSkype_LoginPsTeams

#PsTeamsSkype_GetCallingPolicy
#PsTeamsSkype_GetCallParkPolicy
#PsTeamsSkype_GetChannelPolicy
#PsTeamsSkype_CreateChannelPolicy
#PsTeamsSkype_AssignChannelPolicy
#PsTeamsSkype_ModifyChannelPolicy
#PsTeamsSkype_ModifyChannelPolicy
#PsTeamsSkype_GetTeamsClientConfiguration
#PsTeamsSkype_GetGuestMessagingConfiguration
#PsTeamsSkype_GetMeetingBroadcastConfiguration
#PsTeamsSkype_RemoveGoogleDrive

#------------------------ Using PnP PowerShell for Teams

#PsTeamsPnpPowerShell_LoginPsTeams

#PsTeamsPnpPowerShell_GetAllTeams
#PsTeamsPnpPowerShell_GetOneTeam
#PsTeamsPnpPowerShell_NewTeamByName
#PsTeamsPnpPowerShell_NewTeamByGroup
#PsTeamsPnpPowerShell_SetTeam
#PsTeamsPnpPowerShell_SetPictureTeam
#PsTeamsPnpPowerShell_SetArchivedTeam
#PsTeamsPnpPowerShell_RemoveTeam

#PsTeamsPnpPowerShell_GetAllChannelsTeam
#PsTeamsPnpPowerShell_GetOneChannelFilesFolder
#PsTeamsPnpPowerShell_GetOneChannelTeam
#PsTeamsPnpPowerShell_AddOneChannelTeam
#PsTeamsPnpPowerShell_UpdateOneChannelTeam
#PsTeamsPnpPowerShell_SendMessageToOneChannelTeam
#PsTeamsPnpPowerShell_GetMessagesFromOneChannelTeam
#PsTeamsPnpPowerShell_GetReplayMessageOneChannelTeam
#PsTeamsPnpPowerShell_RemoveOneChannelTeam
#PsTeamsPnpPowerShell_GetAllTabsChannelTeam
#PsTeamsPnpPowerShell_GetOneTabChannelTeam
#PsTeamsPnpPowerShell_AddOneTabChannelTeam
#PsTeamsPnpPowerShell_UpdateOneTabChannelTeam
#PsTeamsPnpPowerShell_DeleteOneTabChannelTeam

#PsTeamsPnpPowerShell_GetAllUsersTeam
#PsTeamsPnpPowerShell_GetAllUsersChannelTeam
#PsTeamsPnpPowerShell_GetAllUsersChannelByRoleTeam
#PsTeamsPnpPowerShell_AddOneUserTeam
#PsTeamsPnpPowerShell_AddOneUserChannel
#PsTeamsPnpPowerShell_DeleteOneUserChannel
#PsTeamsPnpPowerShell_DeleteOneUserTeam

#PsTeamsPnpPowerShell_GetAllAppsTeam
#PsTeamsPnpPowerShell_GetOneAppTeam
#PsTeamsPnpPowerShell_AddOneAppTeam
#PsTeamsPnpPowerShell_UpdateOneAppTeam
#PsTeamsPnpPowerShell_DeleteOneAppTeam

#------------------------ Using Microsoft Graph for Teams (REST calls)

#PsTeamsGraphRest_GetJoinedTeams
#PsTeamsGraphRest_GetAllTeamsByGroup
#PsTeamsGraphRest_GetOneTeam
#PsTeamsGraphRest_CreateOneTeam
#PsTeamsGraphRest_CreateOneGroup
#PsTeamsGraphRest_CreateOneTeamFromGroup
#PsTeamsGraphRest_UpdateOneTeam
#PsTeamsGraphRest_DeleteOneTeam

#PsTeamsGraphRest_GetAllChannelsInOneTeam
#PsTeamsGraphRest_GetOneChannelInOneTeam
#PsTeamsGraphRest_CreateOneChannel
#PsTeamsGraphRest_UpdateOneChannel
#PsTeamsGraphRest_DeleteOneChannel
#PsTeamsGraphRest_GetAllTabsInOneChannel
#PsTeamsGraphRest_GetOneTabInOneChannel
#PsTeamsGraphRest_CreateOneTabInOneChannel
#PsTeamsGraphRest_UpdateOneTabInOneChannel
#PsTeamsGraphRest_DeleteOneTabFromOneChannel

#PsTeamsGraphRest_GetAllUsersInOneTeam
#PsTeamsGraphRest_AddOneUserToOneTeam
#PsTeamsGraphRest_DeleteOneUserFromOneTeam

#PsTeamsGraphRest_SendMessageToOneChannel
#PsTeamsGraphRest_GetAllMessagesChannel
#PsTeamsGraphRest_SendMessageReplayToOneChannel
#PsTeamsGraphRest_GetAllReplaysToOneMessagesChannel

#PsTeamsGraphRest_GetAllMeetings
#PsTeamsGraphRest_GetOneMeeting
#PsTeamsGraphRest_CreateOneMeeting
#PsTeamsGraphRest_DeleteOneMeeting

#PsTeamsGraphRest_GetAllChats
#PsTeamsGraphRest_GetOneChat
#PsTeamsGraphRest_GetOneChatMessages
#PsTeamsGraphRest_GetAllChatsMessages
#PsTeamsGraphRest_GetOneChatParticipants
#PsTeamsGraphRest_GetOneChatOneParticipant
#PsTeamsGraphRest_AddOneUserToChat
#PsTeamsGraphRest_DeleteOneUserFromChat
#PsTeamsGraphRest_SendMessageToChat
#PsTeamsGraphRest_HideChat
#PsTeamsGraphRest_PinChat
#PsTeamsGraphRest_GetPinnedChats
#PsTeamsGraphRest_UnpinChat
#PsTeamsGraphRest_ReadChatForUser
#PsTeamsGraphRest_UnreadChatForUser

#------------------------ Using PnP CLI M365 for Teams

#PsTeamsCliM365_GetAllTeams
#PsTeamsCliM365_GetTeamsByQuery
#PsTeamsCliM365_GetOneTeam
#PsTeamsCliM365_CreateOneTeam
#PsTeamsCliM365_CloneOneTeam
#PsTeamsCliM365_UpdateOneTeam
#PsTeamsCliM365_ArchiveOneTeam
#PsTeamsCliM365_UnarchiveOneTeam
#PsTeamsCliM365_DeleteOneTeam

#PsTeamsCliM365_GetAllChannelsOneTeam
#PsTeamsCliM365_GetChannelByQuery
#PsTeamsCliM365_GetOneChannelFromOneTeam
#PsTeamsCliM365_CreateOneChannelInOneTeam
#PsTeamsCliM365_UpdateOneChannelInOneTeam
#PsTeamsCliM365_DeleteOneChannelFromOneTeam

#PsTeamsCliM365_GetAllTabs
#PsTeamsCliM365_GetTabByQuery
#PsTeamsCliM365_GetOneTab
#PsTeamsCliM365_AddOneTab
#PsTeamsCliM365_DeleteOneTab

#PsTeamsCliM365_GetAllUsers
#PsTeamsCliM365_GetUserByQuery
#PsTeamsCliM365_GetAllUsersChannel
#PsTeamsCliM365_AddOneUser
#PsTeamsCliM365_AddOneUserToOneChannel
#PsTeamsCliM365_UpdateOneUser
#PsTeamsCliM365_DeleteOneUser
#PsTeamsCliM365_DeleteOneUserFromOneChannel

#PsTeamsCliM365_GetAllApps
#PsTeamsCliM365_GetAppByQuery
#PsTeamsCliM365_AddOneApp
#PsTeamsCliM365_PublishOneApp
#PsTeamsCliM365_UpdateOneApp
#PsTeamsCliM365_UninstallOneApp
#PsTeamsCliM365_DeleteOneApp

#PsTeamsCliM365_GetAllMessages
#PsTeamsCliM365_GetMessageByQuery
#PsTeamsCliM365_GetOneMessage
#PsTeamsCliM365_SendMessageToOneChannel
#PsTeamsCliM365_GetMessageReplays

#PsTeamsCliM365_GetAllMeetings
#PsTeamsCliM365_GetOneMeeting
#PsTeamsCliM365_CreateOneMeeting
#PsTeamsCliM365_AttendanceMeeting

#PsTeamsCliM365_GetAllChats
#PsTeamsCliM365_GetOneChat
#PsTeamsCliM365_GetOneChatParticipants
#PsTeamsCliM365_AddOneChatParticipant
#PsTeamsCliM365_DeleteOneChatParticipant
#PsTeamsCliM365_GetChatMessages
#PsTeamsCliM365_SendChatMessageToChat
#PsTeamsCliM365_SendChatMessageToPerson

#PsTeamsCliM365_GetSettings
#PsTeamsCliM365_SetSettings

Write-Host "Done"  

