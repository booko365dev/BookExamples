##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function ManagedIdentyPsRest_SetClaims
{
	#-------------------------------------
	# Connect to Azure AD using User Account and Password

	$AzureUser = "user@domain.onmicrosoft.com"
	$AzurePW = "password"
	$AzureTenantId = "ade56059-xxxx-xxxx-xxxx-e4772a8168ca"

	$secPW = ConvertTo-SecureString -String $AzurePW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $AzureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $AzureTenantId

	#-------------------------------------
	# Assign one Permission

	$appId = "08e5988a-xxxx-xxxx-xxxx-e17ba60b1651"  #--> Managed Identity identifier

	# Connect with Graph using the Principal
	$AccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
	Connect-Graph -AccessToken $AccessToken.Token
	$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '$($appId)'"

	# Get permissions for the Principal (not necessarily to set them, 
	#	but just to know which are already set)
	$myPermsDel = Get-MgServicePrincipalOauth2PermissionGrant `
					-ServicePrincipalId $GraphApp.Id # for Delegated permissions
	[Array]$myPermsApp = Get-MgServicePrincipalAppRoleAssignment `
					-ServicePrincipalId $GraphApp.Id # for Application permissions

	$GraphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph
	# Get the Resource for Graph
	$myResource = Get-MgServicePrincipal -Filter "appId eq '$($GraphAppId)'" 

	# To connect with the MicrosoftTeams module instead of MSGraph, use the next two lines
	# $TeamsAppId = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239" # The ID MicrosoftTeams module
	# Get the Resource for MicrosoftTeams
	# $myResource = Get-MgServicePrincipal -Filter "appId eq '$($TeamsAppId)'"

	# Get the AppRole "Sites.FullControl.All" for the Resource
	$myAppRole = $myResource.AppRoles | Where-Object {$_.Value -eq 'Sites.FullControl.All'}
	#$myAppRole = $myResource.AppRoles | Where-Object {$_.Value -eq 'Sites.ReadWrite.All'}
	#$myAppRole = $myResource.AppRoles | Where-Object {$_.Value -eq 'Sites.Read.All'}

	$AppRoleAssignment = @{
		"PrincipalId" = $GraphApp.Id
		"ResourceId" = $myResource.Id
		"AppRoleId" =  $myAppRole.Id}

	# Assign the AppRole to the Principal --> Type is "Application"
	New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $GraphApp.Id `
											-BodyParameter $AppRoleAssignment

	Disconnect-Graph
	Disconnect-AzAccount
}
#gavdcodeend 001

#gavdcodebegin 002
Function ManagedIdentyPsRest_RemoveClaims
{
	#-------------------------------------
	# Connect to Azure AD using User Account and Password

	$AzureUser = "user@domain.onmicrosoft.com"
	$AzurePW = "password"
	$AzureTenantId = "ade56059-xxxx-xxxx-xxxx-e4772a8168ca"

	$secPW = ConvertTo-SecureString -String $AzurePW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $AzureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $AzureTenantId

	#-------------------------------------
	# Remove one Permission

	[Array]$myPermissions = Get-MgServicePrincipalAppRoleAssignment `
													-ServicePrincipalId $GraphApp.Id

	$GraphAppIdDel = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph
	$myResourceDel = Get-MgServicePrincipal -Filter "appId eq '$($GraphAppIdDel)'" 
	$myAppRoleDel = $myResourceDel.AppRoles | Where-Object `
												{$_.Value -eq 'Sites.FullControl.All'}
	# Find the AppRoleAssignment for the AppRole "Sites.FullControl.All" in the 
	#		array of permissions
	$myAssignment = $myPermissions | Where-Object {$_.AppRoleId -eq $myAppRoleDel.Id}

	# Remove the AppRoleAssignment from the Principal
	Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $GraphApp.Id `
											   -AppRoleAssignmentId $myAssignment.Id

	#-------------------------------------

	Disconnect-Graph
	Disconnect-AzAccount
}
#gavdcodeend 002

#gavdcodebegin 003
Function ManagedIdentyPsPnp_SetClaims
{
	#-------------------------------------
	# Connect to Azure AD using User Account and Password

	$AzureUser = "user@domain.onmicrosoft.com"
	$AzurePW = "password"
	$AzureTenantId = "ade56059-xxxx-xxxx-xxxx-e4772a8168ca"

	$secPW = ConvertTo-SecureString -String $AzurePW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $AzureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $AzureTenantId

	#-------------------------------------
	# Connect to PnP using Url and Token
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
	Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/Test_Site" `
												-AccessToken $myAccessToken.Token

	# Get Permissions for the Principal
	#Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole
	#Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole

	#-------------------------------------
	# Assign one Permission
	# Use the "ObjectId" of the Principal, not its "ApplicationId"
	Add-PnPAzureADServicePrincipalAppRole `
									-Principal "79087a2a-xxxx-xxxx-xxxx-4066eb05be51" `
									-AppRole "Sites.FullControl.All" `
									-BuiltInType MicrosoftGraph

	# PnP - Use the "ObjectId" of the Principal, not its "ApplicationId"
	# Add-PnPAzureADServicePrincipalAppRole `
	#								-Principal "af64fcae-xxxx-xxxx-xxxx-8cacbb649ae5" `
	#                               -AppRole "Sites.FullControl.All" `
	#                               -BuiltInType SharePointOnline

	# Disconnect-Graph
	Disconnect-PnPOnline
	Disconnect-AzAccount
}
#gavdcodeend 003

#gavdcodebegin 004
Function ManagedIdentyPsPnp_RemoveClaims
{
	#-------------------------------------
	# Connect to Azure AD using User Account and Password

	$AzureUser = "user@domain.onmicrosoft.com"
	$AzurePW = "password"
	$AzureTenantId = "ade56059-xxxx-xxxx-xxxx-e4772a8168ca"

	$secPW = ConvertTo-SecureString -String $AzurePW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $AzureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $AzureTenantId

	#-------------------------------------
	# Connect to PnP using Url and Token
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
	Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/Test_Site" `
												-AccessToken $myAccessToken.Token

	# Get Permissions for the Principal
	#Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole
	#Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole

	#-------------------------------------
	# Removes one Permission - It removes the "AppRoleName" for Graph and SharePoint

	Remove-PnPAzureADServicePrincipalAssignedAppRole `
									-Principal "683da884-xxxx-xxxx-xxxx-bc7f26027b8d" `
	                                -AppRoleName "Sites.FullControl.All"
	Remove-PnPAzureADServicePrincipalAssignedAppRole `
									-Principal "683da884-xxxx-xxxx-xxxx-bc7f26027b8d" `
	                                -AppRoleName "Group.Read.All"

	#-------------------------------------

	# Disconnect-Graph
	Disconnect-PnPOnline
	Disconnect-AzAccount
}
#gavdcodeend 004

#gavdcodebegin 005
Function ManagedIdentyPsRest_Example
{
	Import-Module -Name "Microsoft.Graph.Authentication"
	Import-Module -Name "Microsoft.Graph.Sites"

	# Connect to Azure to get the security token
	Connect-AzAccount -Identity
	# Get the Azure token to be used to login in Graph
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

	# Do I have a token?
	Write-Output "-- myToken: " $myAccessToken.Token

	# Connect to the Graph SDK with the acquired access token
	Connect-Graph -AccessToken $myAccessToken.Token

	$myHeaders = @{
		'Authorization' = "Bearer $($myAccessToken.Token)"
		'Content-Type'  = 'application/json'
	}
	$myUri = "https://graph.microsoft.com/v1.0/sites"

	$myResult = Invoke-RestMethod -Uri $myUri -Headers $myHeaders -Method Get
	Write-Output "- RestOutput: " $myResult.value

	#$myObject = ConvertFrom-Json –InputObject $myResult.value 
	#Write-Output "- JSON: " $myObject.value.subject
}
#gavdcodeend 005

#gavdcodebegin 006
Function ManagedIdentyPsGraphSdk_Example
{
	Import-Module -Name "Microsoft.Graph.Authentication"
	Import-Module -Name "Microsoft.Graph.Sites"

	# Connect to Azure to get the security token
	Connect-AzAccount -Identity
	# Get the Azure token to be used to login in Graph
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

	# Do I have a token?
	Write-Output "-- myToken: " $myAccessToken.Token

	# Connect to the Graph SDK with the acquired access token (use Connect-Graph 
	#	or Connect-MgGraph)
	Connect-MgGraph -AccessToken $myAccessToken.Token

	# Do something with Graph
	$mySite = Get-MgSite -SiteId "cc41b7f0-xxxx-xxxx-xxxx-5afb83a210a8"
	Write-Output "-- GraphOutput: " $mySite.DisplayName " - " $mySite.Id

	# Disconnect from the Graph SDK
	$disco = Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function ManagedIdentyPsPnp_Example
{
	Import-Module -Name "PnP.PowerShell"

	# Connect to PnP using the Managed Identity
	Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/MySite" -ManagedIdentity

	# Do I have a token?
	#Get-PnPAccessToken # OK - Token to consume the Microsoft Graph API
	#Get-PnPAccessToken -Decoded  # OK - Gives no Claims for SharePoint

	# Do something with PnP
	$mySite = Get-PnPSite -Includes Id,RootWeb.Title
	Write-Output "- PnPOutput: " $mySite.RootWeb.Title " - " $mySite.Id " - " $mySite.Url

	# Disconnect from PnP
	$disc = Disconnect-PnPOnline

	Write-Output "Done"
}
#gavdcodeend 007

#gavdcodebegin 008
Function ManagedIdentyPsCli_Example
{
	# For Non *-Cs cmdlets - the Microsoft Graph API permissions needed are: 
	#   Organization.Read.All, User.Read.All, Group.ReadWrite.All, 
	#   AppCatalog.ReadWrite.All, TeamSettings.ReadWrite.All, 
	#   Channel.Delete.All, ChannelSettings.ReadWrite.All, 
	#   ChannelMember.ReadWrite.All

	Import-Module -Name "MicrosoftTeams"

	# Connect to MicrosoftTeams
	Connect-MicrosoftTeams -Identity
	# -Identity parameter: Login using managed service identity in the current 
	#       environment. This is currently not supported for *-Cs cmdlets.

	# Do something with MicrosoftTeams
	$myTeams = Get-Team -User "user@domain.onmicrosoft.com"
	Write-Output "-- TeamsOutput: " $myTeams

	# Disconnect from Microsoft Teams
	$disc = Disconnect-MicrosoftTeams
}
#gavdcodeend 008

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 008 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> The scripts must be run from an Azure Automation Runbook or Azure Function
#		The examples cannot be run from an external script. Managed Identities work
#		only if the code is running from an script inside an Azure Service

Write-Host "Done"  

