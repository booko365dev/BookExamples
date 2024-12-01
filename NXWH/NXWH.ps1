
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphCli_LoginWithCertificate
{
	mgc login --tenant-id $configFile.appsettings.TenantName `
			  --client-id $configFile.appsettings.ClientIdWithCert `
			  --certificate-thumb-print $configFile.appsettings.CertificateThumbprint `
			  --strategy ClientCertificate
}

function PsGraphCli_LoginWithDeviceCode
{
	mgc login --tenant-id $configFile.appsettings.TenantName `
			  --client-id $configFile.appsettings.ClientIdWithAccPw `
			  --strategy DeviceCode
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsTeamsGraphCli_GetAllMeTeams
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphCli_LoginWithCertificate

	$myUser = mgc users list --filter "mail eq '$($configFile.appsettings.UserName)'"
	$myUserId = ($myUser | ConvertFrom-Json).value[0].id

	mgc users joined-teams list --user-id $myUserId

	mgc logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsTeamsGraphCli_GetAllTeamsByGroup
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphCli_LoginWithCertificate

	mgc groups list --select "id,displayName,resourceProvisioningOption"

	mgc logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsTeamsGraphCli_GetOneTeam
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphCli_LoginWithCertificate

    $myGroupId = "ae67c043-...-6521dd8d4bf9"
	#mgc groups get --group-id $myGroupId
	mgc teams get --team-id $myGroupId

	mgc logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsTeamsGraphCli_CreateOneTeam
{
	# Requires Team.Create

	PsGraphCli_LoginWithCertificate

	$myUser = mgc users list --filter "mail eq '$($configFile.appsettings.UserName)'"
	$myUserId = ($myUser | ConvertFrom-Json).value[0].id

	$myNewTeamProps = @{
		"template@odata.bind" = `
					"https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
		displayName = "Team created with Graph CLI"
		description = "Team created with the Graph CLI"
		members = @(
			@{
				"@odata.type" = "#microsoft.graph.aadUserConversationMember"
				roles = @(
					"owner"
				)
				"user@odata.bind" = `
					"https://graph.microsoft.com/v1.0/users('" + $($myUserId) + "')"
			}
		)
	}
	$myNewTeamPropsJson = $myNewTeamProps | ConvertTo-Json -Depth 10
	
	mgc teams create --body $myNewTeamPropsJson

	mgc logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsTeamsGraphCli_CreateOneGroup
{
	# Requires Group.ReadWrite.All or Group.Create

	PsGraphCli_LoginWithCertificate

	$myUser = mgc users list --filter "mail eq '$($configFile.appsettings.UserName)'"
	$myUserId = ($myUser | ConvertFrom-Json).value[0].id

	$myNewGroupProps = @{
		displayName = "Group created with Graph CLI"
		description = "Group created with the Graph CLI"
		groupTypes = @( )
		mailEnabled = $false
		mailNickname = "GraphPsSdk"
		securityEnabled = $true
		"owners@odata.bind" = @(
			"https://graph.microsoft.com/v1.0/users/" + $($myUserId)
		)
			"members@odata.bind" = @(
			"https://graph.microsoft.com/v1.0/users/bd6fe5cc-...-2246d8b7b9fb"
		)
	}
	$myNewGroupPropsJson = $myNewGroupProps | ConvertTo-Json -Depth 10
	
	mgc groups create --body $myNewGroupPropsJson

	mgc logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsTeamsGraphCli_CreateOneTeamFromGroup
{
	# Requires Team.Create

	PsGraphCli_LoginWithCertificate

	$myNewTeamProps = @{
		"template@odata.bind" = 
			"https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
		"group@odata.bind" = 
			"https://graph.microsoft.com/v1.0/groups('d1f6494d-...-ec022aca291f')"
	}
	$myNewTeamPropsJson = $myNewTeamProps | ConvertTo-Json -Depth 10
	
	mgc teams create --body $myNewTeamPropsJson

	mgc logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsTeamsGraphCli_UpdateOneTeam
{
	# Requires TeamSettings.ReadWrite.All

	PsGraphCli_LoginWithCertificate

	$myTeamId = "f2998a9b-491a-4a13-9f40-38f331e30a26"
	$myTeamProps = @{
		"displayName" = "Team Updated 03" 
	}
	$myTeamPropsJson = $myTeamProps | ConvertTo-Json -Depth 10
	
	mgc teams patch --team-id $myTeamId --body $myTeamPropsJson

	mgc logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsTeamsGraphCli_DeleteOneTeam
{
	# Requires Group.ReadWrite.All

	PsGraphCli_LoginWithCertificate

	$myTeamId = "1e4e7c92-...-4f91f4f48c91"
	mgc groups delete --group-id $myTeamId

	mgc logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsTeamsGraphCli_GetAllChannelsInOneTeam
{
	# Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	#mgc teams channels list --team-id $myTeamId
	mgc teams all-channels list --team-id $myTeamId

	mgc logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsTeamsGraphCli_GetOneChannelInOneTeam
{
	# Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:5e38bea6f01f44f09b13076b3d6f78d2@thread.tacv2"

	mgc teams channels get --team-id $myTeamId --channel-id $myChannelId

	mgc logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsTeamsGraphCli_CreateOneChannel
{
	# Requires Channel.Create or Channel.CreateGroup

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	$myNewChannelProps = @{
		displayName = "Channel created with Graph CLI"
		description = "Channel created with the Graph CLI"
		membershipType = "standard"
	}
	$myNewChannelPropsJson = $myNewChannelProps | ConvertTo-Json -Depth 10

	mgc teams channels create --team-id $myTeamId --body $myNewChannelPropsJson

	mgc logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsTeamsGraphCli_UpdateOneChannel
{
	# Requires ChannelSettings.ReadWrite.All

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:cd7b26776db449019dc68b48c364bddb@thread.tacv2";

	$myChannelProps = @{
		displayName = "Channel created with Graph CLI Updated"
	}
	$myChannelPropsJson = $myChannelProps | ConvertTo-Json -Depth 10

	mgc teams channels patch --team-id $myTeamId `
							 --channel-id $myChannelId `
							 --body $myChannelPropsJson

	mgc logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsTeamsGraphCli_DeleteOneChannel
{
	# Requires Channel.Delete.All

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:cd7b26776db449019dc68b48c364bddb@thread.tacv2";

	mgc teams channels delete --team-id $myTeamId --channel-id $myChannelId

	mgc logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsTeamsGraphCli_GetAllTabsInOneChannel
{
	# Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

	mgc teams channels tabs list --team-id $myTeamId `
								 --channel-id $myChannelId `
								 --expand "teamsApp"

	mgc logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsTeamsGraphCli_GetOneTabInOneChannel
{
	# Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    $myTabId = "3ed5b337-c2c9-4d5d-b7b4-84ff09a8fc1c"

	mgc teams channels tabs get --team-id $myTeamId `
								--channel-id $myChannelId `
								--teams-tab-id $myTabId `
								--expand "teamsApp"

	mgc logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsTeamsGraphCli_CreateOneTabInOneChannel
{
	# Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

    $myBind = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" +
                        "com.microsoft.teamspace.tab.files.sharepoint"
    $myUrl = $configFile.appsettings.SiteBaseUrl +
                        "/sites/TeamcreatedwithGraphPowerShellSDK/Shared%20Documents"

	$myNewTabProps = @{
		displayName = "Document Library"
		configuration = @{
				entityId = ""
				contentUrl = $myUrl
				websiteUrl = $null
				removeUrl = $null
		}
		"teamsApp@odata.bind" = $myBind
	}
	$myNewTabPropsJson = $myNewTabProps | ConvertTo-Json -Depth 10

	mgc teams channels tabs create --team-id $myTeamId `
								   --channel-id $myChannelId `
								   --body $myNewTabPropsJson

	mgc logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsTeamsGraphCli_UpdateOneTabInOneChannel
{
	# Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
	$myTabId = "9773cf5c-72af-4c01-a1f9-ef9cb7512aa1"

	$myTabProps = @{
		displayName = "My Docs"
	}
	$myTabPropsJson = $myTabProps | ConvertTo-Json -Depth 10

	mgc teams channels tabs patch --team-id $myTeamId `
								  --channel-id $myChannelId `
								  --teams-tab-id $myTabId `
								  --body $myTabPropsJson

	mgc logout
}
#gavdcodeend 017

#gavdcodebegin 018
function PsTeamsGraphCli_DeleteOneTabInOneChannel
{
	# Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
	$myTabId = "9773cf5c-72af-4c01-a1f9-ef9cb7512aa1"

	mgc teams channels tabs delete --team-id $myTeamId `
								   --channel-id $myChannelId `
								   --teams-tab-id $myTabId

	mgc logout
}
#gavdcodeend 018

#gavdcodebegin 019
function PsTeamsGraphCli_GetAllUsersInOneTeam
{
	# Requires TeamMember.Read.All or TeamMember.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"

	mgc teams members list --team-id $myTeamId

	mgc logout
}
#gavdcodeend 019

#gavdcodebegin 020
function PsTeamsGraphCli_AddOneUserToOneTeam
{
	# Requires TeamMember.ReadWrite.All or TeamMember.ReadWrite.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myUserId = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb";

	$myNewUserProps = @{
		"@odata.type" = "#microsoft.graph.aadUserConversationMember"
		"roles" = @( "owner" )
		"user@odata.bind" = "https://graph.microsoft.com/v1.0/users('" + $myUserId + "')"
	}
	$myNewUserPropsJson = $myNewUserProps | ConvertTo-Json -Depth 10

	mgc teams members create --team-id $myTeamId `
							 --body $myNewUserPropsJson

	mgc logout
}
#gavdcodeend 020

#gavdcodebegin 021
function PsTeamsGraphCli_DeleteOneUserFromOneTeam
{
	# Requires TeamMember.Read.All or TeamMember.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myMemberId = "MCMjMSMjYW...tNGE2MC1iOWMxLTIyNDZkOGI3YjlmYg==";

	mgc teams members delete --team-id $myTeamId `
							 --conversation-member-id $myMemberId

	mgc logout
}
#gavdcodeend 021

#gavdcodebegin 022
function PsTeamsGraphCli_SendMessageToOneChannel
{
	# Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    # Must use a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

	$myNewMessageProps = @{
		body = @{
			content = "Message from Graph CLI"
		}
	}
	$myNewMessagePropsJson = $myNewMessageProps | ConvertTo-Json -Depth 10

	mgc teams channels messages create --team-id $myTeamId `
									   --channel-id $myChannelId `
									   --body $myNewMessagePropsJson

	mgc logout
}
#gavdcodeend 022

#gavdcodebegin 023
function PsTeamsGraphCli_GetAllMessagesChannel
{
	# Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"

	mgc teams channels messages list --team-id $myTeamId `
									 --channel-id $myChannelId

	mgc logout
}
#gavdcodeend 023

#gavdcodebegin 024
function PsTeamsGraphCli_SendMessageReplayToOneChannel
{
	# Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    # Must use a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
    $myMessageId = "1731614789042"

	$myNewMessageProps = @{
		body = @{
			content = "Replay for Message from Graph CLI"
		}
	}
	$myNewMessagePropsJson = $myNewMessageProps | ConvertTo-Json -Depth 10

	mgc teams channels messages replies create --team-id $myTeamId `
											   --channel-id $myChannelId `
											   --chat-message-id $myMessageId `
											   --body $myNewMessagePropsJson

	mgc logout
}
#gavdcodeend 024

#gavdcodebegin 025
function PsTeamsGraphCli_GetAllReplaysToOneMessagesChannel
{
	# Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

	PsGraphCli_LoginWithCertificate

    $myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf"
    $myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2"
    $myMessageId = "1731614789042"

	mgc teams channels messages replies list --team-id $myTeamId `
											 --channel-id $myChannelId `
											 --chat-message-id $myMessageId

	mgc logout
}
#gavdcodeend 025

#gavdcodebegin 026
function PsTeamsGraphCli_GetAllMeetings
{
	# Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

    $startMeeting = "2024-01-09T01:00:00"
    $endMeeting = "2024-11-16T23:59:59"
	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"

	mgc users events list --user-id $userId `
	--select "subject,body,bodyPreview,organizer,attendees,start,end,location" `
	--filter "start/dateTime ge '$($startMeeting)' and end/dateTime le '$($endMeeting)'"

	mgc logout
}
#gavdcodeend 026

#gavdcodebegin 027
function PsTeamsGraphCli_GetOneMeeting
{
	# Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"
    $myMeetingId = "AAMkAGE0ODQ3N...F9SJ2ZDb7Xo-OrAAGb3qfbAAA="

	mgc users events get --user-id $userId `
		--event-id $myMeetingId `
		--select "subject,body,bodyPreview,organizer,attendees,start,end,location"

	mgc logout
}
#gavdcodeend 027

#gavdcodebegin 028
function PsTeamsGraphCli_CreateOneMeeting
{
	# Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"

	$myNewMeetingProps = @{
				subject = "Test Meeting from Graph CLI"
				body = @{
					contentType = "HTML"
					content = "It is a test meeting"
				}
				start = @{
					dateTime = "2024-11-15T14:00:00"
					timeZone = "Pacific Standard Time"
				}
				end = @{
					dateTime = "2024-11-15T15:00:00"
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
	$myNewMeetingPropsJson = $myNewMeetingProps | ConvertTo-Json -Depth 10

	mgc users events create --user-id $userId --body $myNewMeetingPropsJson

	mgc logout
}
#gavdcodeend 028

#gavdcodebegin 029
function PsTeamsGraphCli_DeleteOneMeeting
{
	# Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    # Using a Delegate Authentication Provider
	PsGraphCli_LoginWithDeviceCode

	$userId = "acc28fcb-5261-47f8-960b-715d2f98a431"
    $myMeetingId = "AAMkAGE0ODQ...OrAAAAAAENAAC1vtBLB-F9SJ2ZDb7Xo-OrAAIkvLpdAAA="

	mgc users events delete --user-id $userId --event-id $myMeetingId

	mgc logout
}
#gavdcodeend 029


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 029 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#*** Using the MS Graph CLI
#		ATTENTION: There is a Windows Environment Variable already configured in the computer
#					to redirect the commands to the mgc.exe directory (see instructions in the book)
#PsTeamsGraphCli_GetAllMeTeams
#PsTeamsGraphCli_GetAllTeamsByGroup
#PsTeamsGraphCli_GetOneTeam
#PsTeamsGraphCli_CreateOneTeam
#PsTeamsGraphCli_CreateOneGroup
#PsTeamsGraphCli_CreateOneTeamFromGroup
#PsTeamsGraphCli_UpdateOneTeam
#PsTeamsGraphCli_DeleteOneTeam
#PsTeamsGraphCli_GetAllChannelsInOneTeam
#PsTeamsGraphCli_GetOneChannelInOneTeam
#PsTeamsGraphCli_CreateOneChannel
#PsTeamsGraphCli_UpdateOneChannel
#PsTeamsGraphCli_DeleteOneChannel
#PsTeamsGraphCli_GetAllTabsInOneChannel
#PsTeamsGraphCli_GetOneTabInOneChannel
#PsTeamsGraphCli_CreateOneTabInOneChannel
#PsTeamsGraphCli_UpdateOneTabInOneChannel
#PsTeamsGraphCli_DeleteOneTabInOneChannel
#PsTeamsGraphCli_GetAllUsersInOneTeam
#PsTeamsGraphCli_AddOneUserToOneTeam
#PsTeamsGraphCli_DeleteOneUserFromOneTeam
#PsTeamsGraphCli_SendMessageToOneChannel
#PsTeamsGraphCli_GetAllMessagesChannel
#PsTeamsGraphCli_SendMessageReplayToOneChannel
#PsTeamsGraphCli_GetAllReplaysToOneMessagesChannel
#PsTeamsGraphCli_GetAllMeetings
#PsTeamsGraphCli_GetOneMeeting
#PsTeamsGraphCli_CreateOneMeeting
#PsTeamsGraphCli_DeleteOneMeeting

Write-Host "Done" 

