using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Configuration;
using static Microsoft.Graph.Constants;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------
static GraphServiceClient CsTeamsGraphSdk_LoginWithSecret()
{
    string TenantIdToConn = ConfigurationManager.AppSettings["TenantName"];
    string ClientIdToConn = ConfigurationManager.AppSettings["ClientIdWithSecret"];
    string ClientSecretToConn = ConfigurationManager.AppSettings["ClientSecret"];

    string[] myScopes = ["https://graph.microsoft.com/.default"];

    ClientSecretCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
    };

    ClientSecretCredential secretCredential =
                    new(TenantIdToConn, ClientIdToConn, ClientSecretToConn,
                        clientOptionsCredential);
    GraphServiceClient graphClient = new(secretCredential, myScopes);

    return graphClient;
}

static GraphServiceClient CsTeamsGraphSdk_LoginWithAccPw()
{
    string TenantIdToConn = ConfigurationManager.AppSettings["TenantName"];
    string ClientIdToConn = ConfigurationManager.AppSettings["ClientIdWithAccPw"];
    string UserToConn = ConfigurationManager.AppSettings["UserName"];
    string PasswordToConn = ConfigurationManager.AppSettings["UserPw"];

    string[] myScopes = ["https://graph.microsoft.com/.default"];
    UsernamePasswordCredentialOptions clientOptionsCredential = new()
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
    };

    UsernamePasswordCredential accPwCredential =
                        new(UserToConn, PasswordToConn, TenantIdToConn,
                            ClientIdToConn, clientOptionsCredential);
    GraphServiceClient graphClient = new(accPwCredential, myScopes);
    
    return graphClient;
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsTeamsGraphSdk_GetAllMeTeams()
{
    // Requires Team.ReadBasic.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    TeamCollectionResponse allTeams = myGraphClient
                                .Users[ConfigurationManager.AppSettings["UserName"]]
                                .JoinedTeams.GetAsync().Result;

    foreach (Team oneTeam in allTeams.Value)
    {
        Console.WriteLine(oneTeam.DisplayName + " - " + oneTeam.Id);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsTeamsGraphSdk_GetAllTeamsByGroup()
{
    // Requires Group.Read.All, Group.ReadWrite.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    GroupCollectionResponse allGroups = 
                            myGraphClient.Groups.GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Select = 
                            ["id", "displayName", "resourceProvisioningOption"];
    }).Result;

    foreach (Group oneGroup in allGroups.Value)
    {
        Console.WriteLine(oneGroup.DisplayName + " - " + oneGroup.Id);
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsTeamsGraphSdk_GetOneTeam()
{
    // Requires Group.Read.All, Group.ReadWrite.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myGroupId = "ae67c043-7a8b-4bd6-b82f-6521dd8d4bf9";
    //Group myGroup = myGraphClient.Groups[myGroupId].GetAsync().Result;
    Team myTeam = myGraphClient.Teams[myGroupId].GetAsync().Result;

    //Console.WriteLine(myGroup.DisplayName + " - " + myGroup.Id);
    Console.WriteLine(myTeam.DisplayName + " - " + myTeam.Id);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsTeamsGraphSdk_CreateOneTeam()
{
    // Requires Team.Create

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    UserCollectionResponse myUser = myGraphClient.Users
            .GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Filter = 
                    $"mail eq '" + ConfigurationManager.AppSettings["UserName"] + "'";
            }).Result;
    string myUserId = myUser.Value[0].Id;

    Team myNewTeamProps = new()
    {
        AdditionalData = new Dictionary<string, object>
        {
            {
                "template@odata.bind" ,
                "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
            },
        },
        DisplayName = "Team created with Graph CS SDK",
        Description = "Team created with the Graph CS SDK",
        Members =
        [
            new AadUserConversationMember
            {
                OdataType = "#microsoft.graph.aadUserConversationMember",
                Roles =
                [
                    "owner"
                ],
                AdditionalData = new Dictionary<string, object>
                {
                    {
                        "user@odata.bind" , 
                        "https://graph.microsoft.com/v1.0/users('" + myUserId + "')"
                    },
                },
            },
        ]
    };

    Team myNewTeam = myGraphClient.Teams.PostAsync(myNewTeamProps).Result;
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsTeamsGraphSdk_CreateOneGroup()
{
    // Requires Group.ReadWrite.All or Group.Create

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    UserCollectionResponse myUser = myGraphClient.Users
            .GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Filter = 
                    $"mail eq '" + ConfigurationManager.AppSettings["UserName"] + "'";
            }).Result;
    string myUserId = myUser.Value[0].Id;

    Group myNewGroupProps = new()
    {
        DisplayName = "Group created with Graph CS SDK",
        Description = "Team created with the Graph CS SDK",
        GroupTypes = [],
        MailEnabled = false,
        MailNickname = "GraphCsSdk",
        SecurityEnabled = true,
        AdditionalData = new Dictionary<string, object>
        {
            {
                "owners@odata.bind" , new List<string>
                {
                    "https://graph.microsoft.com/v1.0/users/" + myUserId
                }
            },
            {
                "members@odata.bind" , new List<string>
                {
                    "https://graph.microsoft.com/v1.0/users/bd6fe5cc-...-2246d8b7b9fb"
                }
            }
        }
    };

    Group myGroup = myGraphClient.Groups.PostAsync(myNewGroupProps).Result;

    Console.WriteLine(myGroup.Id);
}
//gavdcodeend 005

//gavdcodebegin 006
static void CsTeamsGraphSdk_CreateOneTeamFromGroup()
{
    // Requires Team.Create

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    Team myNewTeamProps = new()
    {
        AdditionalData = new Dictionary<string, object>
        {
            {
                "template@odata.bind", 
                        "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
            },
            {
                "group@odata.bind",
                    "https://graph.microsoft.com/v1.0/groups('d006b3da-..-4cf037b6f4e3')"
            }
        }
    };
    
    Team myTeam = myGraphClient.Teams.PostAsync(myNewTeamProps).Result;
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsTeamsGraphSdk_UpdateOneTeam()
{
    // Requires TeamSettings.ReadWrite.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "f2998a9b-...-38f331e30a26";
    Team myTeamProps = new()
    {
        DisplayName = "Team Updated 01"
    };

    Team myTeam = myGraphClient.Teams[myTeamId].PatchAsync(myTeamProps).Result;
}
//gavdcodeend 007

//gavdcodebegin 008
static async void CsTeamsGraphSdk_DeleteOneTeam()
{
    // Requires Group.ReadWrite.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "f2998a9b-...-38f331e30a26";
    await myGraphClient.Groups[myTeamId].DeleteAsync();
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsTeamsGraphSdk_GetAllChannelsInOneTeam()
{
    // Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";

    //ChannelCollectionResponse allChannels = myGraphClient
    //                .Teams[myTeamId]
    //                .Channels.GetAsync().Result;
    ChannelCollectionResponse allChannels = myGraphClient
                    .Teams[myTeamId]
                    .AllChannels.GetAsync().Result;

    foreach (Channel oneChannel in allChannels.Value)
    {
        Console.WriteLine(oneChannel.DisplayName + " - " + oneChannel.Id);
    }
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsTeamsGraphSdk_GetOneChannelInOneTeam()
{
    // Requires Channel.ReadBasic.All or ChannelSettings.Read.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:5e38bea6f01f44f09b13076b3d6f78d2@thread.tacv2";

    Channel oneChannel = myGraphClient
                    .Teams[myTeamId]
                    .Channels[myChannelId].GetAsync().Result;

    Console.WriteLine(oneChannel.DisplayName + " - " + oneChannel.Id);
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsTeamsGraphSdk_CreateOneChannel()
{
    // Requires Channel.Create or Channel.CreateGroup

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";

    Channel myNewChannelProps = new()
    {
        DisplayName = "Channel created with Graph CS SDK",
        Description = "Channel created with the Graph CS SDK",
        MembershipType = ChannelMembershipType.Standard
    };

    Channel myChannel = myGraphClient.Teams[myTeamId].Channels
                                    .PostAsync(myNewChannelProps).Result;

    Console.WriteLine(myChannel.DisplayName + " - " + myChannel.Id);
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsTeamsGraphSdk_UpdateOneChannel()
{
    // Requires ChannelSettings.ReadWrite.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    Channel myChannelProps = new()
    {
        DisplayName = "Channel created with Graph CS SDK Updated"
    };

    myGraphClient.Teams[myTeamId].Channels[myChannelId].PatchAsync(myChannelProps);
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsTeamsGraphSdk_DeleteOneChannel()
{
    // Requires Channel.Delete.All

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    myGraphClient.Teams[myTeamId].Channels[myChannelId].DeleteAsync();
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsTeamsGraphSdk_GetAllTabsInOneChannel()
{
    // Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    TeamsTabCollectionResponse allTabs = myGraphClient.Teams[myTeamId]
                    .Channels[myChannelId].Tabs.GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Expand = ["teamsApp"];
    }).Result;

    foreach (TeamsTab oneTab in allTabs.Value)
    {
        Console.WriteLine(oneTab.DisplayName + " - " + oneTab.Id);
    }
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsTeamsGraphSdk_GetOneTabInOneChannel()
{
    // Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    string myTabId = "3ed5b337-c2c9-4d5d-b7b4-84ff09a8fc1c";

    TeamsTab myTab = myGraphClient.Teams[myTeamId]
             .Channels[myChannelId].Tabs[myTabId].GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Expand = ["teamsApp"];
    }).Result;

    Console.WriteLine(myTab.DisplayName + " - " + myTab.Id);
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsTeamsGraphSdk_CreateOneTabInOneChannel()
{
    // Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    string myBind = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/" +
                        "com.microsoft.teamspace.tab.files.sharepoint";
    string myUrl = ConfigurationManager.AppSettings["SiteBaseUrl"] +
                        "/sites/TeamcreatedwithGraphPowerShellSDK/Shared%20Documents";

    TeamsTab myNewTabProps = new()
    {
        DisplayName = "Document Library",
        Configuration = new TeamsTabConfiguration
        {
            EntityId = "",
            ContentUrl = myUrl,
            WebsiteUrl = null,
            RemoveUrl = null
        },
        AdditionalData = new Dictionary<string, object>
        {
            {
                "teamsApp@odata.bind" , myBind
            }
        }
    };

    TeamsTab myTab = myGraphClient.Teams[myTeamId].Channels[myChannelId]
                .Tabs.PostAsync(myNewTabProps).Result;

    Console.WriteLine(myTab.DisplayName + " - " + myTab.Id);
}
//gavdcodeend 016

//gavdcodebegin 017
static void CsTeamsGraphSdk_UpdateOneTabInOneChannel()
{
    // Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    string myTabId = "16c04584-d3f0-4cde-9370-4cf2d0a1ef0c";

    TeamsTab myNewTabProps = new()
    {
        DisplayName = "My Docs"
    };

    TeamsTab myTab = myGraphClient.Teams[myTeamId].Channels[myChannelId]
                .Tabs[myTabId].PatchAsync(myNewTabProps).Result;

    Console.WriteLine(myTab.DisplayName + " - " + myTab.Id);
}
//gavdcodeend 017

//gavdcodebegin 018
static void CsTeamsGraphSdk_DeleteOneTabFromOneChannel()
{
    // Requires Directory.ReadWrite.All, Group.Read.All, TeamsTab.ReadWriteForTeam

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    string myTabId = "16c04584-d3f0-4cde-9370-4cf2d0a1ef0c";

    myGraphClient.Teams[myTeamId].Channels[myChannelId].Tabs[myTabId].DeleteAsync();
}
//gavdcodeend 018

//gavdcodebegin 019
static void CsTeamsGraphSdk_GetAllUsersInOneTeam()
{
    // Requires TeamMember.Read.All or TeamMember.Read.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";

    ConversationMemberCollectionResponse myMembers = myGraphClient
                            .Teams[myTeamId].Members.GetAsync().Result;

    foreach (ConversationMember oneMember in myMembers.Value)
    {
        Console.WriteLine(oneMember.DisplayName + " - " + oneMember.Id);
    }
}
//gavdcodeend 019

//gavdcodebegin 020
static void CsTeamsGraphSdk_AddOneUserToOneTeam()
{
    // Requires TeamMember.ReadWrite.All or TeamMember.ReadWrite.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myUserId = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb";

    AadUserConversationMember myNewUserProps = new ()
    {
        OdataType = "#microsoft.graph.aadUserConversationMember",
        Roles =
        [
            "owner",
        ],
        AdditionalData = new Dictionary<string, object>
        {
            {
                "user@odata.bind" , "https://graph.microsoft.com/v1.0/users" + 
                                                            "('" + myUserId + "')"
            }
        }
    };

    ConversationMember result = myGraphClient.Teams[myTeamId].Members.
                                                    PostAsync(myNewUserProps).Result;
}
//gavdcodeend 020

//gavdcodebegin 021
static void CsTeamsGraphSdk_DeleteOneUserFromOneTeam()
{
    // Requires TeamMember.ReadWrite.All or TeamMember.ReadWrite.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myMemberId = "MCMjMSMjYW...tNGE2MC1iOWMxLTIyNDZkOGI3YjlmYg==";

    myGraphClient.Teams[myTeamId].Members[myMemberId].DeleteAsync();
}
//gavdcodeend 021

//gavdcodebegin 022
static void CsTeamsGraphSdk_SendMessageToOneChannel()
{
    // Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    // Must use a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    ChatMessage myNewMessageProps = new()
    {
        Body = new ItemBody
        {
            Content = "Message from Graph CS SDK"
        }
    };

    ChatMessage myMessage = myGraphClient.Teams[myTeamId]
                .Channels[myChannelId].Messages.PostAsync(myNewMessageProps).Result;
}
//gavdcodeend 022

//gavdcodebegin 023
static void CsTeamsGraphSdk_GetAllMessagesChannel()
{
    // Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";

    ChatMessageCollectionResponse allMessages = myGraphClient.Teams[myTeamId]
                .Channels[myChannelId].Messages.GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Top = 10;
    }).Result;

    foreach (ChatMessage oneMessage in allMessages.Value)
    {
        Console.WriteLine(oneMessage.Body.Content + " - " + oneMessage.Id);
    }
}
//gavdcodeend 023

//gavdcodebegin 024
static void CsTeamsGraphSdk_SendMessageReplayToOneChannel()
{
    // Requires ChannelMessage.Send (it works only with Delegate Authentication Provider)

    // Must use a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    string myMessageId = "1731533254015";

    ChatMessage myNewMessageProps = new()
    {
        Body = new ItemBody
        {
            Content = "Replay for Message from Graph CS SDK"
        }
    };

    ChatMessage myMessage = myGraphClient.Teams[myTeamId]
                .Channels[myChannelId].Messages[myMessageId].Replies
                .PostAsync(myNewMessageProps).Result;
}
//gavdcodeend 024

//gavdcodebegin 025
static void CsTeamsGraphSdk_GetAllReplaysToOneMessagesChannel()
{
    // Requires ChannelMessage.Read.All or ChannelMessage.Read.Group

    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithSecret();

    string myTeamId = "3fdc7296-67ee-4c73-b275-b1abd0d585bf";
    string myChannelId = "19:7c88dbaeec484330b930fa35d8bc1e88@thread.tacv2";
    string myMessageId = "1731533254015";

    ChatMessageCollectionResponse allMessages = myGraphClient.Teams[myTeamId]
                .Channels[myChannelId].Messages[myMessageId].Replies
                .GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Top = 10;
                }).Result;

    foreach (ChatMessage oneMessage in allMessages.Value)
    {
        Console.WriteLine(oneMessage.Body.Content + " - " + oneMessage.Id);
    }
}
//gavdcodeend 025

//gavdcodebegin 026
static void CsTeamsGraphSdk_GetAllMeetings()
{
    // Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    // Using a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    string startMeeting = "2024-01-09T01:00:00";
    string endMeeting = "2024-11-16T23:59:59";

    // For Application registration use "/users/userId/" instead of "/me/"
    EventCollectionResponse allMeetings = myGraphClient.Me.Events
                                        .GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Filter = 
                "start/dateTime ge '" + startMeeting + 
                "' and end/dateTime le '" + endMeeting + "'";
        requestConfiguration.QueryParameters.Select = 
                ["subject", "body", "bodyPreview", "organizer", "attendees", 
                 "start", "end", "location"];
        requestConfiguration.Headers.Add
                ("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
    }).Result;


    foreach (Event oneMeeting in allMeetings.Value)
    {
        Console.WriteLine(oneMeeting.Start.DateTime + " - " + 
                          oneMeeting.Subject + " - " + 
                          oneMeeting.Id);
    }
}
//gavdcodeend 026

//gavdcodebegin 027
static void CsTeamsGraphSdk_GetOneMeeting()
{
    // Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    // Using a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    string myMeetingId = "AAMkAGE0ODQ3N...F9SJ2ZDb7Xo-OrAAGb3qfbAAA=";

    // For Application registration use "/users/userId/" instead of "/me/"
    Event myMeeting = 
        myGraphClient.Me.Events[myMeetingId].GetAsync((requestConfiguration) =>
    {
        requestConfiguration.QueryParameters.Select = 
                ["subject", "body", "bodyPreview", "organizer", "attendees", 
                 "start", "end", "location", "hideAttendees"];
        requestConfiguration.Headers.Add
                ("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
    }).Result;

    Console.WriteLine(myMeeting.Start.DateTime + " - " +
                      myMeeting.Subject + " - " +
                      myMeeting.Id);
}
//gavdcodeend 027

//gavdcodebegin 028
static void CsTeamsGraphSdk_CreateOneMeeting()
{
    // Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    // Using a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    Event requestBody = new ()
    {
        Subject = "Test Meeting from Graph CS SDK",
        Body = new ItemBody
        {
            ContentType = BodyType.Html,
            Content = "It is a test meeting",
        },
        Start = new DateTimeTimeZone
        {
            DateTime = "2024-11-15T12:00:00",
            TimeZone = "Pacific Standard Time",
        },
        End = new DateTimeTimeZone
        {
            DateTime = "2024-11-15T13:00:00",
            TimeZone = "Pacific Standard Time",
        },
        Location = new Location
        {
            DisplayName = "Somewhere",
        },
        Attendees =
    [
        new() {
            EmailAddress = new EmailAddress
            {
                Address = "user@domain.com",
                Name = "One User Name",
            },
            Type = AttendeeType.Required,
        },
    ],
        AllowNewTimeProposals = true
    };

    Event myMeeting = myGraphClient.Me.Events.PostAsync(requestBody, 
        (requestConfiguration) =>
    {
        requestConfiguration.Headers.Add
                    ("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
    }).Result;
    // For Application registration use "/users/userId/" instead of "/me/"

    Console.WriteLine(myMeeting.Start.DateTime + " - " +
                      myMeeting.Subject + " - " +
                      myMeeting.Id);
}
//gavdcodeend 028

//gavdcodebegin 029
static void CsTeamsGraphSdk_DeleteOneMeeting()
{
    // Requires Calendars.ReadBasic, Calendars.Read, Calendars.ReadWrite

    // Using a Delegate Authentication Provider
    GraphServiceClient myGraphClient = CsTeamsGraphSdk_LoginWithAccPw();

    string myMeetingId = "AAMkAGE0ODQ3N...F9SJ2ZDb7Xo-OrAAGb3qfbAAA=";

    myGraphClient.Me.Events[myMeetingId].DeleteAsync();
}
//gavdcodeend 029


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

// *** Latest Source Code Index: 029 ***

//CsTeamsGraphSdk_GetAllMeTeams();
//CsTeamsGraphSdk_GetAllTeamsByGroup();
//CsTeamsGraphSdk_GetOneTeam();
//CsTeamsGraphSdk_CreateOneTeam();
//CsTeamsGraphSdk_CreateOneGroup();
//CsTeamsGraphSdk_CreateOneTeamFromGroup();
//CsTeamsGraphSdk_UpdateOneTeam();
//CsTeamsGraphSdk_DeleteOneTeam();
//CsTeamsGraphSdk_GetAllChannelsInOneTeam();
//CsTeamsGraphSdk_GetOneChannelInOneTeam();
//CsTeamsGraphSdk_CreateOneChannel();
//CsTeamsGraphSdk_UpdateOneChannel();
//CsTeamsGraphSdk_DeleteOneChannel();
//CsTeamsGraphSdk_GetAllTabsInOneChannel();
//CsTeamsGraphSdk_GetOneTabInOneChannel();
//CsTeamsGraphSdk_CreateOneTabInOneChannel();
//CsTeamsGraphSdk_UpdateOneTabInOneChannel();
//CsTeamsGraphSdk_DeleteOneTabFromOneChannel();
//CsTeamsGraphSdk_GetAllUsersInOneTeam();
//CsTeamsGraphSdk_AddOneUserToOneTeam();
//CsTeamsGraphSdk_DeleteOneUserFromOneTeam();
//CsTeamsGraphSdk_SendMessageToOneChannel();
//CsTeamsGraphSdk_GetAllMessagesChannel();
//CsTeamsGraphSdk_SendMessageReplayToOneChannel();
//CsTeamsGraphSdk_GetAllReplaysToOneMessagesChannel();
//CsTeamsGraphSdk_GetAllMeetings();
//CsTeamsGraphSdk_GetOneMeeting();
//CsTeamsGraphSdk_CreateOneMeeting();
//CsTeamsGraphSdk_DeleteOneMeeting();

Console.WriteLine("Done");


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used
