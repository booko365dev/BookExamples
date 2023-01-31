
# Functions to login in Azure

Function Get-AzureTokenApplication(){
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

Function Get-AzureTokenDelegation(){
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

#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function LoginPsTeams
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.tmUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.tmUserName, $securePW
	Connect-MicrosoftTeams -Credential $myCredentials
}
#gavdcodeend 001

#gavdcodebegin 018
Function LoginPsTeamsSkype
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.tmUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.tmUserName, $securePW
	
	Import-Module SkypeOnlineConnector
	$mySession = New-CsOnlineSession -Credential $myCredentials
	Import-PSSession $mySession
}
#gavdcodeend 018

#gavdcodebegin 037
Function LoginPsPnPPowerShellManaged
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.tmUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.tmUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.tmUrl -Credentials $myCredentials
}
#gavdcodeend 037

Function LoginPsCLI
{
	m365 login --authType password `
			   --userName $configFile.appsettings.tmUserName `
			   --password $configFile.appsettings.tmUserPw
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 002
Function TeamsPsMtpTeamsEnumarate
{
	Get-Team
	Disconnect-MicrosoftTeams
}
#gavdcodeend 002

#gavdcodebegin 003
Function TeamsPsMtpTeamsGetByDisplayName
{
	Get-Team -DisplayName "Test Team from PS"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 003

#gavdcodebegin 004
Function TeamsPsMtpTeamsCreate
{
	New-Team -DisplayName "Test Team from PS" `
			 -Description "Team created with PowerShell" `
			 -Visibility Private
	Disconnect-MicrosoftTeams
}
#gavdcodeend 004

#gavdcodebegin 005
Function TeamsPsMtpTeamsUpdate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-Team -GroupId $myTeam.GroupId `
			 -Description "Team updated with PowerShell" `
			 -Visibility Public
	Disconnect-MicrosoftTeams
}
#gavdcodeend 005

#gavdcodebegin 006
Function TeamsPsMtpTeamsDelete
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-Team -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 006

#gavdcodebegin 007
Function TeamsPsMtpTeamsHelp
{
	Get-TeamHelp
	Disconnect-MicrosoftTeams
}
#gavdcodeend 007

#gavdcodebegin 008
Function TeamsPsMtpChannelsEnumerate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamChannel -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 008

#gavdcodebegin 009
Function TeamsPsMtpChannelsCreate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	New-TeamChannel -GroupId $myTeam.GroupId `
					-DisplayName "Test Channel from PS" 
	Disconnect-MicrosoftTeams
}
#gavdcodeend 009

#gavdcodebegin 010
Function TeamsPsMtpChannelsUpdate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Set-TeamChannel -GroupId $myTeam.GroupId `
					-CurrentDisplayName "Test Channel from PS" `
					-Description "This is a test Channel"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 010

#gavdcodebegin 011
Function TeamsPsMtpChannelsDelete
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamChannel -GroupId $myTeam.GroupId `
					   -DisplayName "Test Channel from PS"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 011

#gavdcodebegin 012
Function TeamsPsMtpTeamUserEnumerate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Get-TeamUser -GroupId $myTeam.GroupId
	Disconnect-MicrosoftTeams
}
#gavdcodeend 012

#gavdcodebegin 013
Function TeamsPsMtpTeamUserCreate
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Add-TeamUser -GroupId $myTeam.GroupId `
				 -User "user@domain.onmicrosoft.com" 
	Disconnect-MicrosoftTeams
}
#gavdcodeend 013

#gavdcodebegin 014
Function TeamsPsMtpTeamUserDelete
{
	$myTeam = Get-Team -DisplayName "Test Team from PS"
	Remove-TeamUser -GroupId $myTeam.GroupId `
					-User "user@domain.onmicrosoft.com"
	Disconnect-MicrosoftTeams
}
#gavdcodeend 014

#gavdcodebegin 015
Function TeamsPsMtpPolicyPackageEnumerate
{
	Get-CsPolicyPackage
	Disconnect-MicrosoftTeams
}
#gavdcodeend 015

#gavdcodebegin 016
Function TeamsPsMtpPolicyPackageUser
{
	Get-CsUserPolicyPackage -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}
#gavdcodeend 016

#gavdcodebegin 017
Function TeamsPsMtpPolicyPackageUserRecommended
{
	Get-CsUserPolicyPackageRecommendation -Identity user@domain.onmicrosoft.com
	Disconnect-MicrosoftTeams
}
#gavdcodeend 017

#gavdcodebegin 019
Function TeamsPsGetCallingPolicy
{
	Get-CsTeamsCallingPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 019

#gavdcodebegin 020
Function TeamsPsGetCallParkPolicy
{
	Get-CsTeamsCallParkPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 020

#gavdcodebegin 021
Function TeamsPsGetChannelPolicy
{
	Get-CsTeamsChannelsPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 021

#gavdcodebegin 022
Function TeamsPsCreateChannelPolicy
{
	New-CsTeamsChannelsPolicy -Identity myPolicy -AllowPrivateTeamDiscovery $false
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 022

#gavdcodebegin 023
Function TeamsPsAssignChannelPolicy
{
	Grant-CsTeamsChannelsPolicy -Identity user@tenant.OnMicrosoft.com -PolicyName myPolicy
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 023

#gavdcodebegin 024
Function TeamsPsModifyChannelPolicy
{
	Set-CsTeamsChannelsPolicy -Identity myPolicy -AllowPrivateTeamDiscovery $true
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 024

#gavdcodebegin 025
Function TeamsPsModifyChannelPolicy
{
	Grant-CsTeamsChannelsPolicy -Identity user@tenant.OnMicrosoft.com -PolicyName Default
	Remove-CsTeamsChannelsPolicy -Identity myPolicy -Force
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 025

#gavdcodebegin 026
Function TeamsPsGetTeamsClientConfiguration
{
	Get-CsTeamsClientConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 026

#gavdcodebegin 027
Function TeamsPsGetGuestMessagingConfiguration
{
	Get-CsTeamsGuestMessagingConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 027

#gavdcodebegin 028
Function TeamsPsGetMeetingBroadcastConfiguration
{
	Get-CsTeamsMeetingBroadcastConfiguration
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 028

#gavdcodebegin 029
Function TeamsPsRemoveGoogleDrive
{
	Set-CsTeamsClientConfiguration -Identity Global -AllowGoogleDrive $false
	Get-PSSession | Remove-PSSession
}
#gavdcodeend 029

#gavdcodebegin 030
Function TeamsPsGetCsTeamTemplateList
{
	$allTemplates = Get-CsTeamTemplateList
	foreach($oneTemplate in $allTemplates) {
		Write-Host(" - " + $oneTemplate.Name + " - " + $oneTemplate.OdataId)
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 030

#gavdcodebegin 031
Function TeamsPsGetCsTeamTemplate
{
	$oneTemplate = Get-CsTeamTemplate -OdataId `
		"/api/teamtemplates/v1.0/com.microsoft.teams.template.ManageAProject/Public/en-US" `
		| ConvertTo-Json
	Write-Host $oneTemplate

	Disconnect-MicrosoftTeams
}
#gavdcodeend 031

#gavdcodebegin 032
Function TeamsPsGetTeamsApp
{
	$allApps = Get-TeamsApp
	foreach($oneApp in $allApps) {
		Write-Host(" - " + $oneApp.DisplayName + " - " + $oneApp.Id)
	}

	Disconnect-MicrosoftTeams
}
#gavdcodeend 032

#gavdcodebegin 033
Function TeamsPsGetOneTeamsAppByIdOrName
{
	$oneAppById = Get-TeamsApp -Id "ed734a73-73d5-4339-bb60-b078d9fea5a2" | ConvertTo-Json 
	Write-Host $oneAppById

	$oneAppByName = Get-TeamsApp -DisplayName "Analytics 365" | ConvertTo-Json   
	Write-Host $oneAppByName
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 033

#gavdcodebegin 034
Function TeamsPsNewTeamsApp
{
	New-TeamsApp -DistributionMethod "organization" `
				 -Path "C:\Temporary\App01FromDevSite.zip" 
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 034

#gavdcodebegin 035
Function TeamsPsSetTeamsApp
{
	Set-TeamsApp -Id "eed59874-e471-49ca-a01f-7d92bee85fc6" `
				 -Path "C:\Temporary\App01FromDevSite.zip" 
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 035

#gavdcodebegin 036
Function TeamsPsDeleteTeamsApp
{
	Remove-TeamsApp -Id "eed59874-e471-49ca-a01f-7d92bee85fc6"
	
	Disconnect-MicrosoftTeams
}
#gavdcodeend 036

#gavdcodebegin 038
Function TeamsPsPnPGetAllTeams
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsTeam
}
#gavdcodeend 038

#gavdcodebegin 039
Function TeamsPsPnPGetOneTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsTeam -Identity "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"  # GroupID
	Get-PnPTeamsTeam -Identity "Sales and Marketing"  # DisplayName
	Get-PnPTeamsTeam -Identity "SalesAndMarketing"  # MailNickname
}
#gavdcodeend 039

#gavdcodebegin 040
Function TeamsPsPnPNewTeamByName
{
	# Permissions required: Group.ReadWrite.All
	New-PnPTeamsTeam -DisplayName "TeamCreatedWithPnP" `
					 -Visibility Public `
					 -MailNickName "TeamCreatedWithPnPMail" `
					 -AllowUserDeleteMessages $true
}
#gavdcodeend 040

#gavdcodebegin 041
Function TeamsPsPnPNewTeamByGroup
{
	# Permissions required: Group.ReadWrite.All
	New-PnPTeamsTeam -GroupId "89e67c39-b5b3-440d-9bcd-ac8b3743dda1" `
					 -AllowUserDeleteMessages $true
}
#gavdcodeend 041

#gavdcodebegin 042
Function TeamsPsPnPSetTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTeam -Identity "TeamCreatedWithPnP" `
					 -DisplayName "Team Created With PnP" `
					 -Description "This is a test Team"
}
#gavdcodeend 042

#gavdcodebegin 043
Function TeamsPsPnPSetPictureTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTeamPicture -Team "Team Created With PnP" `
							-Path "C:\Temporary\Hulk_logo.jpg"
}
#gavdcodeend 043

#gavdcodebegin 044
Function TeamsPsPnPSetArchivedTeam
{
	# Permissions required: Group.ReadWrite.All or Directory.ReadWrite.All
	Set-PnPTeamsTeamArchivedState -Identity "Team Created With PnP" `
								  -Archived $true `
								  -SetSiteReadOnlyForMembers $true
}
#gavdcodeend 044

#gavdcodebegin 045
Function TeamsPsPnPRemoveTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsTeam -Identity "Team Created With PnP" -Force
	#Remove-PnPTeamsTeamm -GroupId "89e67c39-b5b3-440d-9bcd-ac8b3743dda1" `
}
#gavdcodeend 045

#gavdcodebegin 046
Function TeamsPsPnPGetAllChannelsTeam
{
	# Permissions required: Group.ReadWrite.All
	Get-PnPTeamsChannel -Team "Team Created With PnP"
}
#gavdcodeend 046

#gavdcodebegin 047
Function TeamsPsPnPGetOneChannelTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Get-PnPTeamsChannel -Team "Team Created With PnP" `
						-Identity "19:330cc0611f7648539292e7ea73892e87@thread.tacv2"
}
#gavdcodeend 047

#gavdcodebegin 048
Function TeamsPsPnPAddOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Add-PnPTeamsChannel -Team "Team Created With PnP" `
						-DisplayName "Channel Created With PnP"
}
#gavdcodeend 048

#gavdcodebegin 049
Function TeamsPsPnPUpdateOneChannelTeam
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
Function TeamsPsPnPSendMessageToOneChannelTeam
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
Function TeamsPsPnPGetMessagesFromOneChannelTeam
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

#gavdcodebegin 052
Function TeamsPsPnPRemoveOneChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsChannel -Team "Team Created With PnP" `
						   -Identity "Channel Updated With PnP4"
}
#gavdcodeend 052

#gavdcodebegin 053
Function TeamsPsPnPGetAllTabsChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myTabs = Get-PnPTeamsTab -Team "Team Created With PnP" `
							   -Channel "Channel Created With PnP"

	foreach($oneTab in $myTabs) {
		Write-Host $oneTab.Id " - " $oneTab.DisplayName
	}
}
#gavdcodeend 053

#gavdcodebegin 054
Function TeamsPsPnPGetOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$oneTab = Get-PnPTeamsTab -Team "Team Created With PnP" `
							  -Channel "Channel Created With PnP" `
							  -Identity "Wiki"

	Write-Host $oneTab.Id
}
#gavdcodeend 054

#gavdcodebegin 055
Function TeamsPsPnPAddOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	$myDocsUrl = $configFile.appsettings.tmUrl + "/sites/TeamCreatedWithPnP/MyDocs"
	Add-PnPTeamsTab -Team "Team Created With PnP" `
					-Channel "Channel Created With PnP" `
					-DisplayName "My Documents" `
					-Type "DocumentLibrary" `
					-ContentUrl $myDocsUrl
}
#gavdcodeend 055

#gavdcodebegin 056
Function TeamsPsPnPUpdateOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Set-PnPTeamsTab -Team "Team Created With PnP" `
					-Channel "Channel Created With PnP" `
					-Identity "My Documents" `
					-DisplayName "My Documents Library"
}
#gavdcodeend 056

#gavdcodebegin 057
Function TeamsPsPnPDeleteOneTabChannelTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsTab -Team "Team Created With PnP" `
					   -Channel "Channel Created With PnP" `
					   -Identity "My Documents Library" `
					   -Force
}
#gavdcodeend 057

#gavdcodebegin 058
Function TeamsPsPnPGetAllUsersTeam
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
Function TeamsPsPnPGetAllUsersChannelTeam
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
Function TeamsPsPnPGetAllUsersChannelByRoleTeam
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
Function TeamsPsPnPAddOneUserTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Add-PnPTeamsUser -Team "Team Created With PnP" `
					 -User "user@domain.OnMicrosoft.com" `
					 -Role "Member"
}
#gavdcodeend 061

#gavdcodebegin 062
Function TeamsPsPnPDeleteOneUserTeam
{
	# Permissions required: Group.Read.All or Group.ReadWrite.All
	Add-PnPTeamsUser -Team "Team Created With PnP" `
					 -User "user@domain.OnMicrosoft.com" `
					 -Role "Member"
}
#gavdcodeend 062

#gavdcodebegin 063
Function TeamsPsPnPGetAllAppsTeam
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
Function TeamsPsPnPGetOneAppTeam
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
Function TeamsPsPnPAddOneAppTeam
{
	# Permissions required: AppCatalog.ReadWrite.All or Directory.ReadWrite.All
	New-PnPTeamsApp -Path "C:\Temporary\App01FromDevSite.zip"
}
#gavdcodeend 065

#gavdcodebegin 066
Function TeamsPsPnPUpdateOneAppTeam
{
	# Permissions required: Group.ReadWrite.All
	Update-PnPTeamsApp -Identity "1e67180b-1904-4637-91b5-fa09420953f6" `
					   -Path "C:\Temporary\App01FromDevSite.zip"
}
#gavdcodeend 066

#gavdcodebegin 067
Function TeamsPsPnPDeleteOneAppTeam
{
	# Permissions required: Group.ReadWrite.All
	Remove-PnPTeamsApp -Identity "App01FromDevSite" -Force
	#Remove-PnPTeamsApp -Identity "1e67180b-1904-4637-91b5-fa09420953f6" -Force
}
#gavdcodeend 067

#gavdcodebegin 068
Function TeamsPsGraphGetJoinedTeams
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/me/joinedTeams"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphGetAllTeamsByGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$Url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName," + `
															"resourceProvisioningOptions"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphGetOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.Read.All, Group.ReadWrite.All

	$groepId = "607afc8e-c9eb-4aa2-90b0-104044ebb2f7"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $groupId
	#$Url = "https://graph.microsoft.com/v1.0/teams/" + $groepId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 070 

#gavdcodebegin 071
Function TeamsPsGraphCreateOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Team.Create

	$teamTemplate = "standard"
	$Url = "https://graph.microsoft.com/v1.0/teams"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	# NOTE: The value of $myBody must be in one code line
	$myBody = '{ 
		"template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(''' + 
			$teamTemplate + ''')", 
		"displayName": "Team created with Graph AAA", 
		"description": "This is a Team created with Graph" }'
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 071

#gavdcodebegin 072
Function TeamsPsGraphCreateOneGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Team.Create

	$Url = "https://graph.microsoft.com/v1.0/groups"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphCreateOneTeamFromGroup
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Team.Create

	$grpId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$teamTemplate = "standard"
	$Url = "https://graph.microsoft.com/v1.0/teams"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	# NOTE: The value of $myBody must be in one code line
	$myBody = '{ "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates(''' +`
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
Function TeamsPsGraphUpdateOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphDeleteOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "607afc8e-c9eb-4aa2-90b0-104044ebb2f7"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 075

#gavdcodebegin 076
Function TeamsPsGraphGetAllChannelsInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allChannels = ConvertFrom-Json –InputObject $myResult
	foreach($oneChannel in $allChannels) {
		$oneChannel.value.displayName
	}
}
#gavdcodeend 076

#gavdcodebegin 077
Function TeamsPsGraphGetOneChannelInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:2c68dd98958346f6b53f76d02b3822ee@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																			$channelId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 077 

#gavdcodebegin 078
Function TeamsPsGraphCreateOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Channel.Create

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphUpdateOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:2c68dd98958346f6b53f76d02b3822ee@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																			$channelId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphDeleteOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:2c68dd98958346f6b53f76d02b3822ee@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																			$channelId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 080

#gavdcodebegin 081
Function TeamsPsGraphGetAllTabsInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
																$channelId + "/tabs"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphGetOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Channel.ReadBasic.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$tabId = "b6d407d0-aaaa-4aff-a6e2-5dedf2076542"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/tabs/" + $tabId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	ConvertFrom-Json –InputObject $myResult
}
#gavdcodeend 082

#gavdcodebegin 083
Function TeamsPsGraphCreateOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Directory.ReadWrite.All, Group.ReadWrite.All, Channel.Create

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/tabs"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBind = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" + ` 
										"com.microsoft.teamspace.tab.files.sharepoint"
	$myUrl = "https://m365x829450.sharepoint.com/sites/GroupCreatedWithGraph/" + `
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
Function TeamsPsGraphUpdateOneTabInOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$tabId = "0c6fb5d3-f5fa-4cb3-bed8-a9b0251901fc"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
											$channelId + "/tabs/" + $tabId
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphDeleteOneTabFromOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$tabId = "0c6fb5d3-f5fa-4cb3-bed8-a9b0251901fc"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
											$channelId + "/tabs/" + $tabId

	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete
}
#gavdcodeend 085

#gavdcodebegin 086
Function TeamsPsGraphGetAllUsersInOneTeam
{
	# App Registration type:		Delegation
	# App Registration permissions: Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$Url = "https://graph.microsoft.com/v1.0/groups/" + $teamId + "/members"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	#Write-Host $myResult

	$allUsers = ConvertFrom-Json –InputObject $myResult
	foreach($oneUser in $allUsers) {
		$oneUser.value.displayName
	}
}
#gavdcodeend 086

#gavdcodebegin 087
Function TeamsPsGraphSendMessageToOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: ChannelMessage.Send

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/messages"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
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
Function TeamsPsGraphGetAllMessagesChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.Read, Chat.ReadWrite

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
														$channelId + "/messages"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 088

#gavdcodebegin 089
Function TeamsPsGraphSendMessageReplayToOneChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: ChannelMessage.Send, Group.ReadWrite.All

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$messageId = "1631891785238"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
									$channelId + "/messages/" + $messageId + "/replies"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myBody = "{ `
		'body': {`
			'contentType': 'html', `
			'content': '<strong>This is a response to the Channel message</strong>' `
		}}"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Post `
												-Body $myBody -ContentType $myContentType

	Write-Host $myResult
}
#gavdcodeend 089

#gavdcodebegin 090
Function TeamsPsGraphGetAllReplaysToOneMessagesChannel
{
	# App Registration type:		Delegation
	# App Registration permissions: Chat.Read, Chat.ReadWrite

	$teamId = "5bdad80a-b066-4e0d-88eb-8b959b9fc10a"
	$channelId = "19:7d5c55494eeb4ed5a13f17d234aee753@thread.tacv2"
	$messageId = "1631891785238"
	$Url = "https://graph.microsoft.com/v1.0/teams/" + $teamId + "/channels/" + `
									$channelId + "/messages/" + $messageId + "/replies"
	
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url
	
	Write-Host $myResult
}
#gavdcodeend 090

#gavdcodebegin 091
Function TeamsPsCliGetAllTeams
{
	LoginPsCLI
	
	m365 teams team list

	m365 logout
}
#gavdcodeend 091

#gavdcodebegin 092
Function TeamsPsCliGetTeamsByQuery
{
	LoginPsCLI
	
	m365 teams team list --output json --query "[?displayName == 'Sales and Marketing']"

	m365 logout
}
#gavdcodeend 092

#gavdcodebegin 093
Function TeamsPsCliGetOneTeam
{
	LoginPsCLI
	
	m365 teams team get --id "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"

	m365 logout
}
#gavdcodeend 093

#gavdcodebegin 094
Function TeamsPsCliCreateOneTeam
{
	LoginPsCLI
	
	m365 teams team add --name "TeamCreatedWithCliPnP" `
						--description "Team Created With the CLI PnP" `
						--wait

	m365 logout
}
#gavdcodeend 094

#gavdcodebegin 095
Function TeamsPsCliCloneOneTeam
{
	LoginPsCLI
	
	m365 teams team clone --teamId "b3a39202-0e82-4ddf-98d5-69674d066ea6" `
						  --displayName "TeamClonedWithCliPnP" `
						  --description "Team Cloned With the CLI PnP" `
						  --partsToClone "apps,tabs,settings,channels,members" `
						  --visibility "public"

	m365 logout
}
#gavdcodeend 095

#gavdcodebegin 096
Function TeamsPsCliUpdateOneTeam
{
	LoginPsCLI
	
	m365 teams team set --teamId "b3a39202-0e82-4ddf-98d5-69674d066ea6" `
						--description "Team Updated With the CLI PnP"

	m365 logout
}
#gavdcodeend 096

#gavdcodebegin 097
Function TeamsPsCliArchiveOneTeam
{
	LoginPsCLI
	
	m365 teams team archive --teamId "b3a39202-0e82-4ddf-98d5-69674d066ea6" `
						    --shouldSetSpoSiteReadOnlyForMembers

	m365 logout
}
#gavdcodeend 097

#gavdcodebegin 098
Function TeamsPsCliUnarchiveOneTeam
{
	LoginPsCLI
	
	m365 teams team unarchive --teamId "b3a39202-0e82-4ddf-98d5-69674d066ea6"

	m365 logout
}
#gavdcodeend 098

#gavdcodebegin 099
Function TeamsPsCliDeleteOneTeam
{
	LoginPsCLI
	
	m365 teams team remove --teamId "b3a39202-0e82-4ddf-98d5-69674d066ea6" `
						   --confirm

	m365 logout
}
#gavdcodeend 099

#gavdcodebegin 100
Function TeamsPsCliGetAllChannelsOneTeam
{
	LoginPsCLI
	
	m365 teams channel list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"

	m365 logout
}
#gavdcodeend 100

#gavdcodebegin 101
Function TeamsPsCliGetChannelByQuery
{
	LoginPsCLI
	
	m365 teams channel list --output json `
							--teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
							--query "[?displayName == 'Monthly Reports']"

	m365 logout
}
#gavdcodeend 101

#gavdcodebegin 102
Function TeamsPsCliGetOneChannelFromOneTeam
{
	LoginPsCLI
	
	m365 teams channel get --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						   --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2"

	m365 logout
}
#gavdcodeend 102

#gavdcodebegin 103
Function TeamsPsCliCreateOneChannelInOneTeam
{
	LoginPsCLI
	
	m365 teams channel add --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						   --name "ChannelCreatedWithCliPnP" `
						   --description "Channel Created With the CLI PnP"

	m365 logout
}
#gavdcodeend 103

#gavdcodebegin 104
Function TeamsPsCliUpdateOneChannelInOneTeam
{
	LoginPsCLI
	
	m365 teams channel set --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						   --channelName "ChannelCreatedWithCliPnP" `
						   --description "Channel Updated With the CLI PnP"

	m365 logout
}
#gavdcodeend 104

#gavdcodebegin 105
Function TeamsPsCliDeleteOneChannelFromOneTeam
{
	LoginPsCLI
	
	m365 teams channel remove --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						      --channelName "ChannelCreatedWithCliPnP" `
						      --confirm

	m365 logout
}
#gavdcodeend 105

#gavdcodebegin 106
Function TeamsPsCliGetAllTabs
{
	LoginPsCLI
	
	m365 teams tab list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					    --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2"

	m365 logout
}
#gavdcodeend 106

#gavdcodebegin 107
Function TeamsPsCliGetTabByQuery
{
	LoginPsCLI
	
	m365 teams tab list --output json `
						--teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						--channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
						--query "[?displayName == 'Sales Report']"

	m365 logout
}
#gavdcodeend 107

#gavdcodebegin 108
Function TeamsPsCliGetOneTab
{
	LoginPsCLI
	
	m365 teams tab get --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					   --channelName "Monthly Reports" `
					   --tabName "Sales Report"

	m365 logout
}
#gavdcodeend 108

#gavdcodebegin 109
Function TeamsPsCliAddOneTab
{
	LoginPsCLI
	
	m365 teams tab add --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					   --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
					   --appId "e2acbf5d-6a4f-4d35-a760-503dc0faf314" `
					   --appName "Guitaca Site" `
					   --contentUrl "https://guitaca.com"

	m365 logout
}
#gavdcodeend 109

#gavdcodebegin 110
Function TeamsPsCliDeleteOneTab
{
	LoginPsCLI
	
	m365 teams tab remove --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						  --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
					      --tabId "0ef6aadb-ff6c-450f-aa2f-085be8fc1d21" `
						  --confirm

	m365 logout
}
#gavdcodeend 110

#gavdcodebegin 111
Function TeamsPsCliGetAllUsers
{
	LoginPsCLI
	
	m365 teams user list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					     --role "Owner"

	m365 logout
}
#gavdcodeend 111

#gavdcodebegin 112
Function TeamsPsCliGetUserByQuery
{
	LoginPsCLI
	
	m365 teams user list --output json `
						 --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						 --role "Owner" `
						 --query "[?displayName == 'Megan Bowen']"

	m365 logout
}
#gavdcodeend 112

#gavdcodebegin 113
Function TeamsPsCliAddOneUser
{
	LoginPsCLI
	
	m365 teams user add --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					    --userName "user@domain.onmicrosoft.com" `
						--role "Member"

	m365 logout
}
#gavdcodeend 113

#gavdcodebegin 114
Function TeamsPsCliUpdateOneUser
{
	LoginPsCLI
	
	m365 aad o365group user set --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
								--userName "user@domain.onmicrosoft.com" `
								--role "Owner"

	m365 logout
}
#gavdcodeend 114

#gavdcodebegin 115
Function TeamsPsCliDeleteOneUser
{
	LoginPsCLI
	
	m365 aad o365group user remove --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
								   --userName "user@domain.onmicrosoft.com" `
								   --confirm

	m365 logout
}
#gavdcodeend 115

#gavdcodebegin 116
Function TeamsPsCliGetAllApps
{
	LoginPsCLI
	
	m365 teams app list
	Write-Host ("-------")
	m365 teams app list --all
	Write-Host ("-------")
	m365 teams app list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"

	m365 logout
}
#gavdcodeend 116

#gavdcodebegin 117
Function TeamsPsCliGetAppByQuery
{
	LoginPsCLI
	
	m365 teams app list --all `
						--output json `
						--query "[?displayName == 'MailChimp']"

	m365 logout
}
#gavdcodeend 117

#gavdcodebegin 118
Function TeamsPsCliAddOneApp
{
	LoginPsCLI
	
	m365 teams app install --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					       --appId "ffdb7239-3b58-46ba-b108-7f90a6d8799b"

	m365 logout
}
#gavdcodeend 118

#gavdcodebegin 119
Function TeamsPsCliPublishOneApp
{
	LoginPsCLI
	
	m365 teams app publish --filePath "C:\Projects\MyApp.zip"

	m365 logout
}
#gavdcodeend 119

#gavdcodebegin 120
Function TeamsPsCliUpdateOneApp
{
	LoginPsCLI
	
	m365 teams app update --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b" `
						  --filePath "C:\Projects\MyApp.zip"

	m365 logout
}
#gavdcodeend 120

#gavdcodebegin 121
Function TeamsPsCliUninstallOneApp
{
	LoginPsCLI
	
	m365 teams app uninstall --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
					         --appId "ffdb7239-3b58-46ba-b108-7f90a6d8799b" `
							 --confirm

	m365 logout
}
#gavdcodeend 121

#gavdcodebegin 122
Function TeamsPsCliDeleteOneApp
{
	LoginPsCLI
	
	m365 teams app remove --id "ffdb7239-3b58-46ba-b108-7f90a6d8799b0" `
						  --confirm

	m365 logout
}
#gavdcodeend 122

#gavdcodebegin 123
Function TeamsPsCliGetAllMessages
{
	LoginPsCLI
	
	m365 teams message list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						    --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2"

	m365 logout
}
#gavdcodeend 123

#gavdcodebegin 124
Function TeamsPsCliGetMessageByQuery
{
	LoginPsCLI
	
	m365 teams message list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						    --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
							--output json `
							--query "[?id == '1627202598321']"

	m365 logout
}
#gavdcodeend 124

#gavdcodebegin 125
Function TeamsPsCliGetMessage
{
	LoginPsCLI
	
	m365 teams message get --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
						   --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
						   --messageId "1627202598321"

	m365 logout
}
#gavdcodeend 125

#gavdcodebegin 126
Function TeamsPsCliGetMessageReplays
{
	LoginPsCLI
	
	m365 teams message reply list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
								  --channelId "19:c9640d0f69b84319a8f9c260358e0848@thread.tacv2" `
								  --messageId "1627202602559"

	m365 logout
}
#gavdcodeend 126

#gavdcodebegin 127
Function TeamsPsCliGetSettings
{
	LoginPsCLI
	
	m365 teams membersettings list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"
	Write-Host ("-------")
	m365 teams guestsettings list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"
	Write-Host ("-------")
	m365 teams messagingsettings list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"
	Write-Host ("-------")
	m365 teams funsettings list --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90"

	m365 logout
}
#gavdcodeend 127

#gavdcodebegin 128
Function TeamsPsCliSetSettings
{
	LoginPsCLI
	
	m365 teams funsettings set --teamId "c77f29d7-fdaa-4570-9c3c-210e2d20bc90" `
							   --allowGiphy false

	m365 logout
}
#gavdcodeend 128

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\tmPs.values.config"

#------------------------ Using Teams cmdlets

#LoginPsTeams

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

#TeamsPsGetCsTeamTemplateList
#TeamsPsGetCsTeamTemplate
#TeamsPsGetTeamsApp
#TeamsPsGetOneTeamsAppByIdOrName
#TeamsPsNewTeamsApp
#TeamsPsSetTeamsApp
#TeamsPsDeleteTeamsApp

#------------------------ Using Skype For Business cmdlets

#LoginPsTeamsSkype

#TeamsPsGetCallingPolicy
#TeamsPsGetCallParkPolicy
#TeamsPsGetChannelPolicy
#TeamsPsCreateChannelPolicy
#TeamsPsAssignChannelPolicy
#TeamsPsModifyChannelPolicy
#TeamsPsModifyChannelPolicy
#TeamsPsGetTeamsClientConfiguration
#TeamsPsGetGuestMessagingConfiguration
#TeamsPsGetMeetingBroadcastConfiguration
#TeamsPsRemoveGoogleDrive

#------------------------ Using PnP PowerShell for Teams

#LoginPsPnPPowerShellManaged

#TeamsPsPnPGetAllTeams
#TeamsPsPnPGetOneTeam
#TeamsPsPnPNewTeamByName
#TeamsPsPnPNewTeamByGroup
#TeamsPsPnPSetTeam
#TeamsPsPnPSetPictureTeam
#TeamsPsPnPSetArchivedTeam
#TeamsPsPnPRemoveTeam
#TeamsPsPnPGetAllChannelsTeam
#TeamsPsPnPGetOneChannelTeam
#TeamsPsPnPAddOneChannelTeam
#TeamsPsPnPUpdateOneChannelTeam
#TeamsPsPnPSendMessageToOneChannelTeam
#TeamsPsPnPGetMessagesFromOneChannelTeam
#TeamsPsPnPRemoveOneChannelTeam
#TeamsPsPnPGetAllTabsChannelTeam
#TeamsPsPnPGetOneTabChannelTeam
#TeamsPsPnPAddOneTabChannelTeam
#TeamsPsPnPUpdateOneTabChannelTeam
#TeamsPsPnPDeleteOneTabChannelTeam
#TeamsPsPnPGetAllUsersTeam
#TeamsPsPnPGetAllUsersChannelTeam
#TeamsPsPnPGetAllUsersChannelByRoleTeam
#TeamsPsPnPAddOneUserTeam
#TeamsPsPnPDeleteOneUserTeam
#TeamsPsPnPGetAllAppsTeam
#TeamsPsPnPGetOneAppTeam
#TeamsPsPnPAddOneAppTeam
#TeamsPsPnPUpdateOneAppTeam
#TeamsPsPnPDeleteOneAppTeam

#------------------------ Using Microsoft Graph PowerShell for Teams

#$ClientIDDel = $configFile.appsettings.tmClientIdDel
#$TenantName = $configFile.appsettings.tmTenantName
#$UserName = $configFile.appsettings.tmUserName
#$UserPw = $configFile.appsettings.tmUserPw

#TeamsPsGraphGetJoinedTeams
#TeamsPsGraphGetAllTeamsByGroup
##TeamsPsGraphGetOneTeam
#TeamsPsGraphCreateOneTeam
#TeamsPsGraphCreateOneGroup
#TeamsPsGraphCreateOneTeamFromGroup
#TeamsPsGraphUpdateOneTeam
#TeamsPsGraphDeleteOneTeam
#TeamsPsGraphGetAllChannelsInOneTeam
#TeamsPsGraphGetOneChannelInOneTeam
#TeamsPsGraphCreateOneChannel
#TeamsPsGraphUpdateOneChannel
#TeamsPsGraphDeleteOneChannel
#TeamsPsGraphGetAllTabsInOneChannel
#TeamsPsGraphGetOneTabInOneChannel
#TeamsPsGraphCreateOneTabInOneChannel
#TeamsPsGraphUpdateOneTabInOneChannel
#TeamsPsGraphDeleteOneTabFromOneChannel
#TeamsPsGraphGetAllUsersInOneTeam
#TeamsPsGraphSendMessageToOneChannel
#TeamsPsGraphGetAllMessagesChannel
#TeamsPsGraphSendMessageReplayToOneChannel
#TeamsPsGraphGetAllReplaysToOneMessagesChannel

#------------------------ Using Microsoft PnP CLI for Teams

#TeamsPsCliGetAllTeams
#TeamsPsCliGetTeamsByQuery
#TeamsPsCliGetOneTeam
#TeamsPsCliCreateOneTeam
#TeamsPsCliCloneOneTeam
#TeamsPsCliUpdateOneTeam
#TeamsPsCliArchiveOneTeam
#TeamsPsCliUnarchiveOneTeam
#TeamsPsCliDeleteOneTeam
#TeamsPsCliGetAllChannelsOneTeam
#TeamsPsCliGetChannelByQuery
#TeamsPsCliGetOneChannelFromOneTeam
#TeamsPsCliCreateOneChannelInOneTeam
#TeamsPsCliUpdateOneChannelInOneTeam
#TeamsPsCliDeleteOneChannelFromOneTeam
#TeamsPsCliGetAllTabs
#TeamsPsCliGetTabByQuery
#TeamsPsCliGetOneTab
#TeamsPsCliAddOneTab
#TeamsPsCliDeleteOneTab
#TeamsPsCliGetAllUsers
#TeamsPsCliGetUserByQuery
#TeamsPsCliAddOneUser
#TeamsPsCliUpdateOneUser
#TeamsPsCliDeleteOneUser
#TeamsPsCliGetAllApps
#TeamsPsCliGetAppByQuery
#TeamsPsCliAddOneApp
#TeamsPsCliPublishOneApp
#TeamsPsCliUpdateOneApp
#TeamsPsCliUninstallOneApp
#TeamsPsCliDeleteOneApp
#TeamsPsCliGetAllMessages
#TeamsPsCliGetMessageByQuery
#TeamsPsCliGetMessage
#TeamsPsCliGetMessageReplays
#TeamsPsCliGetSettings
#TeamsPsCliSetSettings

Write-Host "Done"  

