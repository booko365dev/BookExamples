##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphSdk_LoginWithSecret
{
    $ClientID = $configFile.appsettings.ClientIdWithSecret
    $ClientSecret = $configFile.appsettings.ClientSecret
    $TenantName = $configFile.appsettings.TenantName

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$ClientSecret -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $ClientID, $securePW

	Connect-MgGraph -TenantId $TenantName `
					-ClientSecretCredential $myCredentials
}

function PsGraphSdk_LoginWithAccPwMSAL
{
	$TenantName = $configFile.appsettings.TenantName
	$ClientID = $configFile.appsettings.ClientIdWithAccPw
	$UserName = $configFile.appsettings.UserName
	$UserPw = $configFile.appsettings.UserPw

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $UserName, $securePW
	$myToken = Get-MsalToken -TenantId $TenantName `
							-ClientId $ClientId `
							-UserCredential $myCredentials
	$myTokenSecure = ConvertTo-SecureString -String $myToken.AccessToken `
							-AsPlainText -Force

	Connect-MgGraph -AccessToken $myTokenSecure
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsTeamsGraphSdk_GetAllMeTeams
{
    # Requires Team.ReadBasic.All

	PsGraphSdk_LoginWithSecret

	$myUser = Get-MgUser -Filter "mail eq '$($configFile.appsettings.UserName)'"

	Get-MgUserJoinedTeam -UserId $myUser.Id

	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
function PsTeamsGraphSdk_GetAllTeamsByGroup
{
    # Requires Requires Group.Read.All, Group.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	Get-MgGroup -Property "id,displayName,resourceProvisioningOption" 

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
function PsTeamsGraphSdk_GetOneTeam
{
    # Requires Requires Group.Read.All, Group.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myGroupId = "ae67c043-...-6521dd8d4bf9"
	#Get-MgGroup -GroupId $myGroupId 
	Get-MgTeam -TeamId $myGroupId

	Disconnect-MgGraph
}
#gavdcodeend 003

#gavdcodebegin 004
function PsTeamsGraphSdk_CreateOneTeam
{
    # Requires Team.Create

	PsGraphSdk_LoginWithSecret

	$myUser = Get-MgUser -Filter "mail eq '$($configFile.appsettings.UserName)'"

	$myNewTeamProps = @{
		"template@odata.bind" = `
					"https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
		displayName = "Team created with Graph PowerShell SDK"
		description = "Team created with the Graph PowerShell SDK"
		members = @(
			@{
				"@odata.type" = "#microsoft.graph.aadUserConversationMember"
				roles = @(
					"owner"
				)
				"user@odata.bind" = `
					"https://graph.microsoft.com/v1.0/users('" + $($myUser.Id) + "')"
			}
		)
	}

	New-MgTeam -BodyParameter $myNewTeamProps

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
function PsTeamsGraphSdk_CreateOneGroup
{
    # Requires Group.ReadWrite.All or Group.Create

	PsGraphSdk_LoginWithSecret

	$myUser = Get-MgUser -Filter "mail eq '$($configFile.appsettings.UserName)'"

	$myNewGroupProps = @{
		displayName = "Group created with Graph PowerShell SDK"
		description = "Group created with the Graph PowerShell SDK"
		groupTypes = @( )
		mailEnabled = $false
		mailNickname = "GraphPsSdk"
		securityEnabled = $true
		"owners@odata.bind" = @(
			"https://graph.microsoft.com/v1.0/users/" + $($myUser.Id)
		)
			"members@odata.bind" = @(
			"https://graph.microsoft.com/v1.0/users/bd6fe5cc-...-2246d8b7b9fb"
		)
	}

	New-MgGroup -BodyParameter $myNewGroupProps

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
function PsTeamsGraphSdk_CreateOneTeamFromGroup
{
    # Requires Team.Create

	PsGraphSdk_LoginWithSecret

	$myNewTeamProps = @{
		"template@odata.bind" = 
			"https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
		"group@odata.bind" = 
			"https://graph.microsoft.com/v1.0/groups('d1f6494d-...-ec022aca291f')"
	}

	New-MgTeam -BodyParameter $myNewTeamProps

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
function PsTeamsGraphSdk_UpdateOneTeam
{
    # Requires TeamSettings.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myTeamId = "f2998a9b-...-38f331e30a26"
	$myTeamProps = @{
		"displayName" = "Team Updated 02" 
	}

	Update-MgTeam -TeamId $myTeamId -BodyParameter $myTeamProps

	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
function PsTeamsGraphSdk_DeleteOneTeam
{
    # Requires Group.ReadWrite.All

	PsGraphSdk_LoginWithSecret

	$myTeamId = "173f9a33-...-87485f6be098"
	Remove-MgGroup -GroupId $myTeamId

	Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
function PsTeamsGraphSdk_GetAllChannelsInOneTeam
{
    # Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	#Get-MgTeamChannel -TeamId $myTeamId
	Get-MgAllTeamChannel -TeamId $myTeamId

	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
function PsTeamsGraphSdk_GetOneChannelInOneTeam
{
    # Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:5e38bea6f01f44f09b13076b3d6f78d2@thread.tacv2"

	Get-MgTeamChannel -TeamId $myTeamId -ChannelId $myChannelId

	Disconnect-MgGraph
}
#gavdcodeend 010

#gavdcodebegin 011
function PsTeamsGraphSdk_CreateOneChannel
{
    # Requires Channel.Create or Channel.CreateGroup

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	$myNewChannelProps = @{
		displayName = "Channel created with Graph PowerShell SDK"
		description = "Channel created with the Graph PowerShell SDK"
		membershipType = "standard"
	}

	New-MgTeamChannel -TeamId $myTeamId -BodyParameter $myNewChannelProps

	Disconnect-MgGraph
}
#gavdcodeend 011

#gavdcodebegin 012
function PsTeamsGraphSdk_UpdateOneChannel
{
    # Requires ChannelSettings.ReadWrite.All

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:4ae4c97c47494659ace4cfaf39ab3572@thread.tacv2"

	$myChannelProps = @{
		displayName = "Channel created with Graph PowerShell SDK Updated"
	}

	Update-MgTeamChannel -TeamId $myTeamId `
						 -ChannelId $myChannelId `
						 -BodyParameter $myChannelProps

	Disconnect-MgGraph
}
#gavdcodeend 012

#gavdcodebegin 013
function PsTeamsGraphSdk_DeleteOneChannel
{
    # Requires Channel.Delete.All

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:4ae4c97c47494659ace4cfaf39ab3572@thread.tacv2"

	Remove-MgTeamChannel -TeamId $myTeamId -ChannelId $myChannelId

	Disconnect-MgGraph
}
#gavdcodeend 013

#gavdcodebegin 014
function PsTeamsGraphSdk_GetAllTabsInOneChannel
{
    # Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

	Get-MgTeamChannelTab -TeamId $myTeamId `
						 -ChannelId $myChannelId `
						 -ExpandProperty "teamsApp" 


	Disconnect-MgGraph
}
#gavdcodeend 014

#gavdcodebegin 015
function PsTeamsGraphSdk_GetOneTabInOneChannel
{
    # Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
    $myTabId = "3ed5b337-c2c9-4d5d-b7b4-84ff09a8fc1c"

	Get-MgTeamChannelTab -TeamId $myTeamId `
						 -ChannelId $myChannelId `
						 -TeamsTabId $myTabId `
						 -ExpandProperty "teamsApp" 

	Disconnect-MgGraph
}
#gavdcodeend 015

#gavdcodebegin 016
function PsTeamsGraphSdk_CreateOneTabInOneChannel
{
    # Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

    $myBind = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" +
                        "com.microsoft.teamspace.tab.files.sharepoint"
    $myUrl = $configFile.appsettings.SiteBaseUrl +
                        "/sites/TeamcreatedwithGraphPowerShellSDK/Shared%20Documents"

	$myNewTabProps = @{
		displayName = "Document Library2"
		configuration = @{
				entityId = ""
				contentUrl = $myUrl
				websiteUrl = $null
				removeUrl = $null
		}
		"teamsApp@odata.bind" = $myBind
	}

	New-MgTeamChannelTab -TeamId $myTeamId `
						 -ChannelId $myChannelId `
						 -BodyParameter $myNewTabProps

	Disconnect-MgGraph
}
#gavdcodeend 016

#gavdcodebegin 017
function PsTeamsGraphSdk_UpdateOneTabInOneChannel
{
    # Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
	$myTabId = "af349ef9-a5ec-4395-bd86-82984e0f619c"

	$myTabProps = @{
		displayName = "My Docs2"
	}

	Update-MgTeamChannelTab -TeamId $myTeamId `
						 -ChannelId $myChannelId `
						 -TeamsTabId $myTabId `
						 -BodyParameter $myTabProps

	Disconnect-MgGraph
}
#gavdcodeend 017

#gavdcodebegin 018
function PsTeamsGraphSdk_DeleteOneTabInOneChannel
{
    # Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
	$myTabId = "af349ef9-a5ec-4395-bd86-82984e0f619c"

	Remove-MgTeamChannelTab -TeamId $myTeamId `
						    -ChannelId $myChannelId `
						    -TeamsTabId $myTabId

	Disconnect-MgGraph
}
#gavdcodeend 018

#gavdcodebegin 019
function PsTeamsGraphSdk_GetAllUsersInOneTeam
{
    # Requires TeamMember.Read.All or TeamMember.Read.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	Get-MgTeamMember -TeamId $myTeamId

	Disconnect-MgGraph
}
#gavdcodeend 019

#gavdcodebegin 020
function PsTeamsGraphSdk_AddOneUserToOneTeam
{
    # Requires TeamMember.ReadWrite.All or TeamMember.ReadWrite.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myUserId = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb";

	$myNewUserProps = @{
		"@odata.type" = "#microsoft.graph.aadUserConversationMember"
		"roles" = @( "owner" )
		"user@odata.bind" = "https://graph.microsoft.com/v1.0/users('" + $myUserId + "')"
	}

	New-MgTeamMember -TeamId $myTeamId -BodyParameter $myNewUserProps

	Disconnect-MgGraph
}
#gavdcodeend 020

#gavdcodebegin 021
function PsTeamsGraphSdk_DeleteOneUserFromOneTeam
{
    # Requires TeamMember.ReadWrite.All or TeamMember.ReadWrite.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myMemberId = "MCMjMSMjYW...tNGE2MC1iOWMxLTIyNDZkOGI3YjlmYg==";

	Remove-MgTeamMember -TeamId $myTeamId -ConversationMemberId $myMemberId

	Disconnect-MgGraph
}
#gavdcodeend 021

#gavdcodebegin 022
function PsTeamsGraphSdk_SendMessageToOneChannel
{
    # Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    # Must use a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

	$myNewMessageProps = @{
		body = @{
			content = "Message from Graph PowerShell SDK"
		}
	}

	New-MgTeamChannelMessage -TeamId $myTeamId `
							 -ChannelId $myChannelId `
							 -BodyParameter $myNewMessageProps

	Disconnect-MgGraph
}
#gavdcodeend 022

#gavdcodebegin 023
function PsTeamsGraphSdk_GetAllMessagesChannel
{
    # Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

	Get-MgTeamChannelMessage -TeamId $myTeamId -ChannelId $myChannelId

	Disconnect-MgGraph
}
#gavdcodeend 023

#gavdcodebegin 024
function PsTeamsGraphSdk_SendMessageToOneChannel
{
    # Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    # Must use a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    $myMessageId = "1731615743977"

	$myNewMessageProps = @{
		body = @{
			content = "Replay from Message from Graph PowerShell SDK"
		}
	}

	New-MgTeamChannelMessageReply -TeamId $myTeamId `
								  -ChannelId $myChannelId `
								  -ChatMessageId $myMessageId `
								  -BodyParameter $myNewMessageProps

	Disconnect-MgGraph
}
#gavdcodeend 024

#gavdcodebegin 025
function PsTeamsGraphSdk_GetAllReplaysToOneMessagesChannel
{
    # Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

	PsGraphSdk_LoginWithSecret

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
    $myMessageId = "1731615743977"

	Get-MgTeamChannelMessageReply -TeamId $myTeamId `
								  -ChannelId $myChannelId `
								  -ChatMessageId $myMessageId

	Disconnect-MgGraph
}
#gavdcodeend 025

#gavdcodebegin 026
function PsTeamsGraphSdk_GetAllMeetings
{
    # Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

    $startMeeting = "2024-01-09T01:00:00"
    $endMeeting = "2024-11-16T23:59:59"
	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"

	Get-MgUserEvent -UserId $userId `
	-Property "subject,body,bodyPreview,organizer,attendees,start,end,location" `
	-Filter "start/dateTime ge '$($startMeeting)' and end/dateTime le '$($endMeeting)'"

	Disconnect-MgGraph
}
#gavdcodeend 026

#gavdcodebegin 027
function PsTeamsGraphSdk_GetOneMeeting
{
    # Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"
    $myMeetingId = "AAMkAGE0ODQ3NTc1LTZkM2ItN...F9SJ2ZDb7Xo-OrAAGb3qfbAAA=";

	Get-MgUserEvent -UserId $userId `
			-EventId $myMeetingId `
			-Property "subject,body,bodyPreview,organizer,attendees,start,end,location"

	Disconnect-MgGraph
}
#gavdcodeend 027

#gavdcodebegin 028
function PsTeamsGraphSdk_CreateOneMeeting
{
    # Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"

	$myNewMeetingProps = @{
				subject = "Test Meeting from Graph PowerShell SDK"
				body = @{
					contentType = "HTML"
					content = "It is a test meeting"
				}
				start = @{
					dateTime = "2024-11-15T16:00:00"
					timeZone = "Pacific Standard Time"
				}
				end = @{
					dateTime = "2024-11-15T17:00:00"
					timeZone = "Pacific Standard Time"
				}
				location = @{
					displayName = "Somewhere"
				}
				attendees = @(
					@{
						emailAddress = @{
							address = "user@domain.com"
							name = "One User Name"
						}
						type = "required"
					}
				)
				allowNewTimeProposals = $true
			}

	New-MgUserEvent -UserId $userId -BodyParameter $myNewMeetingProps

	Disconnect-MgGraph
}
#gavdcodeend 028

#gavdcodebegin 029
function PsTeamsGraphSdk_DeleteOneMeeting
{
    # Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphSdk_LoginWithAccPwMSAL

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"
    $myMeetingId = "AAMkAGE0ODQ3NTc1LTZkM2ItN...F9SJ2ZDb7Xo-OrAAGb3qfbAAA=";

	Remove-MgUserEvent -UserId $userId -EventId $myMeetingId

	Disconnect-MgGraph
}
#gavdcodeend 029


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 029 *** 

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsTeamsGraphSdk_GetAllMeTeams
#PsTeamsGraphSdk_GetAllTeamsByGroup
#PsTeamsGraphSdk_GetOneTeam
#PsTeamsGraphSdk_CreateOneTeam
#PsTeamsGraphSdk_CreateOneGroup
#PsTeamsGraphSdk_CreateOneTeamFromGroup
#PsTeamsGraphSdk_UpdateOneTeam
#PsTeamsGraphSdk_DeleteOneTeam
#PsTeamsGraphSdk_GetAllChannelsInOneTeam
#PsTeamsGraphSdk_GetOneChannelInOneTeam
#PsTeamsGraphSdk_CreateOneChannel
#PsTeamsGraphSdk_UpdateOneChannel
#PsTeamsGraphSdk_DeleteOneChannel
#PsTeamsGraphSdk_GetAllTabsInOneChannel
#PsTeamsGraphSdk_GetOneTabInOneChannel
#PsTeamsGraphSdk_CreateOneTabInOneChannel
#PsTeamsGraphSdk_UpdateOneTabInOneChannel
#PsTeamsGraphSdk_DeleteOneTabInOneChannel
#PsTeamsGraphSdk_GetAllUsersInOneTeam
#PsTeamsGraphSdk_AddOneUserToOneTeam
#PsTeamsGraphSdk_DeleteOneUserFromOneTeam
#PsTeamsGraphSdk_SendMessageToOneChannel
#PsTeamsGraphSdk_GetAllMessagesChannel
#PsTeamsGraphSdk_SendMessageToOneChannel
#PsTeamsGraphSdk_GetAllReplaysToOneMessagesChannel
#PsTeamsGraphSdk_GetAllMeetings
#PsTeamsGraphSdk_GetOneMeeting
#PsTeamsGraphSdk_CreateOneMeeting
#PsTeamsGraphSdk_DeleteOneMeeting

Write-Host "Done"
