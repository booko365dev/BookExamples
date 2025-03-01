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
function PsGraphPsSdkManagedIdenty_SetClaims
{
	# Requires the Az PowerShell module 
    param (
        [string] $azureUser,
        [string] $azureUserPw,
        [string] $azureTenantId,
        [string] $appId,  #--> Managed Identity identifier
        [string] $appRole
    )

	# Connect to Azure AD using User Account and Password
	$secPW = ConvertTo-SecureString -String $azureUserPW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $azureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $azureTenantId

	# Connect to MS Graph using the Principal
	$accessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $secureToken = ConvertTo-SecureString $accessToken.Token -AsPlainText -Force
	Connect-Graph -AccessToken $secureToken

	$myPrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($appId)'"

	# Get permissions for the Principal (not necessarily to set them, 
	#	but just to know which are already set)
	#[Array]$myPermsApp = Get-MgServicePrincipalOauth2PermissionGrant `
	#				-ServicePrincipalId $myPrincipal.Id   # for Delegated permissions
	#[Array]$myPermsApp = Get-MgServicePrincipalAppRoleAssignment `
	#				-ServicePrincipalId $myPrincipal.Id   # for Application permissions

	# Get the Resource for the Graph API
	$graphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph API
	$myResource = Get-MgServicePrincipal -Filter "appId eq '$($graphAppId)'" 

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-MgServicePrincipal -Filter "appId eq '$($SharePointAppId)'"

	# Get the AppRole for the Resource. Use for example 'Sites.FullControl.All'
    $myAppRole = $myResource.AppRoles | Where-Object {$_.Value -eq $appRole}

	$appRoleAssignment = @{
		"PrincipalId" = $myPrincipal.Id
		"ResourceId" = $myResource.Id
		"AppRoleId" =  $myAppRole.Id}

	# Assign the AppRole to the Principal --> Type is "Application"
	New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $myPrincipal.Id `
											-BodyParameter $appRoleAssignment

	Disconnect-Graph
	Disconnect-AzAccount
}
#gavdcodeend 001

#gavdcodebegin 002
function PsGraphPsSdkManagedIdenty_RemoveClaims
{
	# Requires the Az PowerShell module 
    param (
        [string] $azureUser,
        [string] $azureUserPw,
        [string] $azureTenantId,
        [string] $appId,  #--> Managed Identity identifier
        [string] $appRole
    )

	# Connect to Azure AD using User Account and Password
	$secPW = ConvertTo-SecureString -String $azureUserPW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $azureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $azureTenantId

	# Connect to MS Graph using the Principal
	$accessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
    $secureToken = ConvertTo-SecureString $accessToken.Token -AsPlainText -Force
	Connect-Graph -AccessToken $secureToken
	
	$myPrincipal = Get-MgServicePrincipal -Filter "AppId eq '$($appId)'"

	#[Array]$myPermsApp = Get-MgServicePrincipalOauth2PermissionGrant `
	#				-ServicePrincipalId $myPrincipal.Id   # for Delegated permissions
	[Array]$myPermsApp = Get-MgServicePrincipalAppRoleAssignment `
					-ServicePrincipalId $myPrincipal.Id   # for Application permissions

	# Get the Resource for the Graph API
	$graphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph API
	$myResource = Get-MgServicePrincipal -Filter "appId eq '$($graphAppId)'" 

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-MgServicePrincipal -Filter "appId eq '$($SharePointAppId)'" 

	# Find the AppRoleAssignment for the AppRole "Sites.FullControl.All" in the 
	#		array of permissions
	$myAppRole = $myResource.AppRoles | Where-Object {$_.Value -eq $appRole}
	$myAssignment = $myPermsApp | Where-Object {$_.AppRoleId -eq $myAppRole.Id}

	# Remove the AppRoleAssignment from the Principal
	Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $myPrincipal.Id `
											   -AppRoleAssignmentId $myAssignment.Id

	Disconnect-Graph
	Disconnect-AzAccount
}
#gavdcodeend 002

#gavdcodebegin 003
function PsPnPManagedIdenty_SetClaims
{
	# Requires the Az PowerShell module 
    param (
        [string] $azureUser,
        [string] $azureUserPw,
        [string] $azureTenantId,
        [string] $appId,  #--> Managed Identity identifier
        [string] $appRole
    )

	# Connect to Azure AD using User Account and Password
	$secPW = ConvertTo-SecureString -String $azureUserPW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $azureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $azureTenantId

	# Connect to PnP using Url and Token
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
	Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/Test_Site" `
												-AccessToken $myAccessToken.Token

	## Get Permissions for the Principal for the Microsoft Graph API
	#Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole
	## Get Permissions for the Principal for the SharePoin API
	#Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole

	## Assign one Permission for the Microsoft Graph API
	# Use the "ObjectId" of the Principal, not its "ApplicationId"
	Add-PnPAzureADServicePrincipalAppRole `
									-Principal $appId `
									-AppRole $appRole `
									-BuiltInType MicrosoftGraph

	## Assign one Permission for the SharePoint API
	# PnP - Use the "ObjectId" of the Principal, not its "ApplicationId"
	# Add-PnPAzureADServicePrincipalAppRole `
	#								-Principal $appId `
	#                               -AppRole $appRole `
	#                               -BuiltInType SharePointOnline

	Disconnect-PnPOnline
	Disconnect-AzAccount
}
#gavdcodeend 003

#gavdcodebegin 004
function PsPnPManagedIdenty_RemoveClaims
{
	# Requires the Az PowerShell module 
    param (
        [string] $azureUser,
        [string] $azureUserPw,
        [string] $azureTenantId,
        [string] $appId,  #--> Managed Identity identifier
        [string] $appRole
    )

	# Connect to Azure AD using User Account and Password
	$secPW = ConvertTo-SecureString -String $azureUserPW -AsPlainText -Force
	$myCredential = New-Object -TypeName "System.Management.Automation.PSCredential" `
							   -ArgumentList $azureUser,$secPW
	Connect-AzAccount -Credential $myCredential -Tenant $azureTenantId

	# Connect to PnP using Url and Token
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"
	Connect-PnPOnline -Url "https://domain.sharepoint.com/sites/Test_Site" `
												-AccessToken $myAccessToken.Token

	## Get Permissions for the Principal for the Microsoft Graph API
	#Get-PnPAzureADServicePrincipal -BuiltInType MicrosoftGraph | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole
	## Get Permissions for the Principal for the SharePoint API
	#Get-PnPAzureADServicePrincipal -BuiltInType SharePointOnline | `
	#									Get-PnPAzureADServicePrincipalAvailableAppRole

	# Removes one Permission - It removes the "AppRoleName" for MS Graph and SharePoint
	Remove-PnPAzureADServicePrincipalAssignedAppRole `
									-Principal $appId `
	                                -AppRoleName $appRole

	Disconnect-PnPOnline
	Disconnect-AzAccount
}
#gavdcodeend 004

#gavdcodebegin 010
function PsAzManagedIdenty_SetClaimsById
{
    param (
        [string] $appId,
        [string] $appRole
    )

    # Connect to Azure
    Connect-AzAccount
    # Select-AzSubscription -SubscriptionId $subscriptionId  # Only if needed

    # Get the Managed Identity Service Principal
    $myPrincipal = Get-AzADServicePrincipal -ApplicationId $appId

	# Get the Resource for the MS Graph API
	$graphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph API
	$myResource = Get-AzADServicePrincipal -ApplicationId $graphAppId 

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-AzADServicePrincipal -ApplicationId $SharePointAppId

	# Get the AppRole for the Resource. Use for example 'Sites.FullControl.All'
    $myAppRole = $myResource.AppRole | Where-Object {$_.Value -eq $appRole}

    # Assign the AppRole to the Managed Identity
    New-AzADServicePrincipalAppRoleAssignment -ServicePrincipalId $myPrincipal.Id `
                                              -ResourceId $myResource.Id `
                                              -AppRoleId $myAppRole.Id

	Disconnect-AzAccount
}
#gavdcodeend 010

#gavdcodebegin 011
function PsAzManagedIdenty_SetClaimsByName {
    param(
        [string] $managedIdentityName,
        [string] $resourceGroupName,
        [string] $appRole
    )

    # Connect to Azure
    Connect-AzAccount
    #Select-AzSubscription -SubscriptionId $subscriptionId  # Only if needed

    # Get the Managed Identity Service Principal
    $myPrincipal = Get-AzUserAssignedIdentity -ResourceGroupName $resourceGroupName `
											  -Name $managedIdentityName

    # Get the resource for the Microsoft Graph API
    $graphAppId = "00000003-0000-0000-c000-000000000000"  # The Identifier for Graph API
    $myResource = Get-AzADServicePrincipal -ApplicationId $graphAppId

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-AzADServicePrincipal -ApplicationId $SharePointAppId

	# Get the AppRole for the Resource. Use for example 'Sites.FullControl.All'
    $myAppRole = $myResource.AppRole | Where-Object {$_.Value -eq $appRole}

    # Assign the App Role to the Managed Identity
    New-AzADServicePrincipalAppRoleAssignment `
										-ServicePrincipalId $myPrincipal.PrincipalId `
                                        -ResourceId $myResource.Id `
                                        -AppRoleId $myAppRole.Id

	Disconnect-AzAccount
}
#gavdcodeend 011

#gavdcodebegin 012
function PsAzManagedIdenty_RemoveClaimsById
{
    param (
        [string] $appId,
        [string] $appRole
    )

    # Connect to Azure
    Connect-AzAccount
    # Select-AzSubscription -SubscriptionId $subscriptionId  # Only if needed

    # Get the Managed Identity Service Principal
    $myPrincipal = Get-AzADServicePrincipal -ApplicationId $appId

	##[Array]$myPermsApp = Get-AzADServicePrincipalOauth2PermissionGrant `
	##				-ServicePrincipalId $myPrincipal.Id   # for Del permissions
	[Array]$myPermsApp = Get-AzADServicePrincipalAppRoleAssignment `
					-ServicePrincipalId $myPrincipal.Id   # for App permissions

	# Get the Resource for the MS Graph API
	$graphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph API
	$myResource = Get-AzADServicePrincipal -ApplicationId $graphAppId 

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-AzADServicePrincipal -ApplicationId $SharePointAppId

	# Find the AppRoleAssignment for the AppRole "Sites.FullControl.All" in the 
	#		array of permissions
	$myAppRole = $myResource.AppRole | Where-Object {$_.Value -eq $appRole}
	$myAssignment = $myPermsApp | Where-Object {$_.AppRoleId -eq $myAppRole.Id}

    # Remove the AppRoleAssigment from the Managed Identity Principal
    Remove-AzADServicePrincipalAppRoleAssignment `
										-ServicePrincipalId $myPrincipal.Id `
                                        -AppRoleAssignmentId  $myAssignment.Id

	Disconnect-AzAccount
}
#gavdcodeend 012

#gavdcodebegin 013
function PsAzManagedIdenty_RemoveClaimsByName
{
    param (
        [string] $managedIdentityName,
        [string] $resourceGroupName,
        [string] $appRole
    )

    # Connect to Azure
    Connect-AzAccount
    # Select-AzSubscription -SubscriptionId $subscriptionId  # Only if needed

    # Get the Managed Identity Service Principal
    $myPrincipal = Get-AzUserAssignedIdentity -ResourceGroupName $resourceGroupName `
											  -Name $managedIdentityName

	##[Array]$myPermsApp = Get-AzADServicePrincipalOauth2PermissionGrant `
	##				-ServicePrincipalId $myPrincipal.PrincipalId   # for Del permissions
	[Array]$myPermsApp = Get-AzADServicePrincipalAppRoleAssignment `
					-ServicePrincipalId $myPrincipal.PrincipalId   # for App permissions

	# Get the Resource for the MS Graph API
	$graphAppId = "00000003-0000-0000-c000-000000000000" # The Identifier for Graph API
	$myResource = Get-AzADServicePrincipal -ApplicationId $graphAppId 

	# To use the Office 365 SharePoint Online module instead of MSGraph:
	# Get the Resource for Office 365 SharePoint Online
	#$SharePointAppId = "00000003-0000-0ff1-ce00-000000000000" # The ID for SP
	#$myResource = Get-AzADServicePrincipal -ApplicationId $SharePointAppId

	# Find the AppRoleAssignment for the AppRole "Sites.FullControl.All" in the 
	#		array of permissions
	$myAppRole = $myResource.AppRole | Where-Object {$_.Value -eq $appRole}
	$myAssigment = $myPermsApp | Where-Object {$_.AppRoleId -eq $myAppRole.Id}

    # Remove the AppRoleAssigment from the Managed Identity Principal
    Remove-AzADServicePrincipalAppRoleAssignment `
										-ServicePrincipalId $myPrincipal.PrincipalId `
                                        -AppRoleAssignmentId  $myAssigment.Id

	Disconnect-AzAccount
}
#gavdcodeend 013

#gavdcodebegin 009
function PsAzManagedIdenty_CreateManagedIdentity_UserAssigned
{
	# Requires the Az PowerShell module 
    param (
        [string] $resourceGroupName,
        [string] $location,
        [string] $identityName
    )

    # Connect to Azure
    Connect-AzAccount

    # Select the Azure subscription to use
    #Select-AzSubscription -SubscriptionId $subscriptionId # Only if needed

    # Create a Resource Group (if it doesn't exist)
    # Check if the resource group exists
    $resourceGroup = Get-AzResourceGroup -Name $resourceGroupName `
										 -ErrorAction SilentlyContinue

    if (-not $resourceGroup) {
        Write-Host "Resource group '$resourceGroupName' does not exist. Creating..."
        $resourceGroup = New-AzResourceGroup -Name $resourceGroupName `
											 -Location $location
        Write-Host "Resource group '$resourceGroupName' created"
    } else {
        Write-Host "Resource group '$resourceGroupName' already exists"
    }

    # Create User-Assigned Managed Identity
    $identity = New-AzUserAssignedIdentity -Name $identityName `
										   -ResourceGroupName $resourceGroupName `
										   -Location $location

    Write-Host "User-assigned managed identity '$identityName' created"

    # Output Identity Details
    $identity | Format-List
}
#gavdcodeend 009

#gavdcodebegin 014
function PsAzManagedIdenty_DeleteManagedIdentity_UserAssigned
{
	# Requires the Az PowerShell module 
    param (
        [string] $resourceGroupName,
        [string] $identityName
    )

    # Connect to Azure
    Connect-AzAccount

    # Select the Azure subscription to use
    #Select-AzSubscription -SubscriptionId $subscriptionId  # Only if necessary

    $done = Remove-AzUserAssignedIdentity -Name $identityName `
										  -ResourceGroupName $resourceGroupName

	if ($done) {
		Write-Host "User-assigned managed identity '$identityName' removed"
	}
}
#gavdcodeend 014

#gavdcodebegin 015
function PsAzManagedIdenty_GetManagedIdentities_AllUserAssigned
{
    # Connect to Azure
    Connect-AzAccount

    # Select the Azure subscription to use
    #Select-AzSubscription -SubscriptionId $subscriptionId  # Only if necessary

	Get-AzUserAssignedIdentity
}
#gavdcodeend 015

#gavdcodebegin 016
function PsAzManagedIdenty_GetManagedIdentities_OneUserAssignedByGroup
{
	# Requires the Az PowerShell module 
    param (
        [string] $resourceGroupName
    )

    # Connect to Azure
    Connect-AzAccount

    # Select the Azure subscription to use
    #Select-AzSubscription -SubscriptionId $subscriptionId  # Only if necessary

	Get-AzUserAssignedIdentity -ResourceGroupName $resourceGroupName
}
#gavdcodeend 016

#gavdcodebegin 017
function PsAzManagedIdenty_GetManagedIdentities_OneUserAssignedByGroupAndName
{
	# Requires the Az PowerShell module 
    param (
        [string] $resourceGroupName,
        [string] $identityName
    )

    # Connect to Azure
    Connect-AzAccount

    # Select the Azure subscription to use
    #Select-AzSubscription -SubscriptionId $subscriptionId  # Only if necessary

	Get-AzUserAssignedIdentity -ResourceGroupName $resourceGroupName `
							   -Name $identityName
}
#gavdcodeend 017

#----EXAMPLES--------------------------------------------------------
##==> The next scripts must be run from an Azure Automation Runbook or Azure function
#		The examples cannot be run from an external script. Managed Identities work
#		only if the code is running from an script inside an Azure Service

#gavdcodebegin 019
function PsGraphRestApi_ManagedIdentyExample
{
	# Define the resource URI and OAuth2 endpoint
	$resource = "https://graph.microsoft.com"
	$endpoint = "http://169.254.169.254/metadata/identity/oauth2/token"
	$apiVersion = "2018-02-01"
	$clientId = "93a29f3d-xxxx-xxxx-xxxx-1df61b36c9d9"

	# Build the request URI
	$uri = "$($endpoint)?api-version=$($apiVersion)&resource=`
			$($resource)&client_id=$($clientId)"

	# Set headers
	$headers = @{
		Metadata = "true"
	}

	# Obtain the access token
	$tokenResponse = Invoke-RestMethod -Method Get -Uri $uri `
									   -Headers $headers -ErrorAction Stop
	$accessToken = $tokenResponse.access_token

	if (!$accessToken) {
		Write-Error "Failed to obtain access token."
		return
	}

	# Specify the SharePoint site URL
	$siteUrl = "https://[TenantName].sharepoint.com/sites/[SiteName]"

	# URL-encode the site URL
	$encodedSiteUrl = [System.Uri]::EscapeDataString($siteUrl)

	# Build the site endpoint
	$siteEndpoint = "https://graph.microsoft.com/v1.0/sites/$($encodedSiteUrl)"

	# Set the Authorization header
	$headers = @{
		Authorization = "Bearer $accessToken"
	}

	# Get the site information
	try {
		$siteResponse = Invoke-RestMethod -Method Get -Uri $siteEndpoint 
										  -Headers $headers -ErrorAction Stop
	} catch {
		Write-Error "Error retrieving site information: $_"
		return
	}

	$siteId = $siteResponse.id

	if (!$siteId) {
		Write-Error "Failed to retrieve site ID."
		return
	}

	# Build the lists endpoint
	$listsEndpoint = "https://graph.microsoft.com/v1.0/sites/$siteId/lists"

	# Get the lists
	try {
		$listsResponse = Invoke-RestMethod -Method Get -Uri $listsEndpoint 
										   -Headers $headers -ErrorAction Stop
	} catch {
		Write-Error "Error retrieving lists: $_"
		return
	}

	# Output the lists
	foreach ($list in $listsResponse.value) {
		Write-Output "List Title: $($list.displayName)"
		Write-Output "List ID: $($list.id)"
		Write-Output "---------------------------------"
	}
}
#gavdcodeend 019

#gavdcodebegin 005
function PsGraphRestApi_ManagedIdentyExample_AzToken
{
	Import-Module -Name "Microsoft.Graph.Authentication"
	Import-Module -Name "Microsoft.Graph.Sites"

	# Connect to Azure (System Assigned Identity) to get the security token
	Connect-AzAccount -Identity
	## Connect to Azure (User Assigned Identity) to get the security token
	#Connect-AzAccount -Identity -AccountId "93a29f3d-xxxx-xxxx-xxxx-1df61b36c9d9"

	# Get the Azure token to be used to login in Graph
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

	# Do I have a token?
	Write-Output "-- myToken: " $myAccessToken.Token

	# Connect to Graph with the acquired access token
	[SecureString]$secureToken = ConvertTo-SecureString -String `
								 $myAccessToken.Token -AsPlainText -Force
	Connect-MgGraph -AccessToken $secureToken

	$myHeaders = @{
		'Authorization' = "Bearer $($myAccessToken.Token)"
		'Content-Type'  = 'application/json'
	}
	$myUri = "https://graph.microsoft.com/v1.0/sites"

	$myResult = Invoke-RestMethod -Uri $myUri -Headers $myHeaders -Method Get
	Write-Output "-- RestOutput: " $myResult.value

	#$myObject = ConvertFrom-Json –InputObject $myResult.value 
	#Write-Output "- JSON: " $myObject.value.subject
}
#gavdcodeend 005

#gavdcodebegin 006
function PsGraphPsSdk_ManagedIdentyExample_AzToken
{
	Import-Module -Name "Microsoft.Graph.Authentication"
	Import-Module -Name "Microsoft.Graph.Sites"

	# Connect to Azure (System Assigned Identity) to get the security token
	Connect-AzAccount -Identity
	## Connect to Azure (User Assigned Identity) to get the security token
	#Connect-AzAccount -Identity -AccountId "93a29f3d-xxxx-xxxx-xxxx-1df61b36c9d9"

	# Get the Azure token to be used to login in Graph
	$myAccessToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"

	# Do I have a token?
	Write-Output "-- myToken: " $myAccessToken.Token

	# Connect to Graph with the acquired access token
	[SecureString]$secureToken = ConvertTo-SecureString -String `
								 $myAccessToken.Token -AsPlainText -Force
	Connect-MgGraph -AccessToken $secureToken

	# Do something with Graph
	$mySite = Get-MgSite -SiteId "cc41b7f0-xxxx-xxxx-xxxx-5afb83a210a8"
	Write-Output "-- GraphOutput: " $mySite.DisplayName " - " $mySite.Id

	# Disconnect from the Graph SDK
	$disconn = Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 018
function PsGraphPsSdk_ManagedIdentyExample
{
	Import-Module -Name "Microsoft.Graph.Authentication"
	Import-Module -Name "Microsoft.Graph.Sites"

	# Connect to MS Graph (System Assigned Identity)
	Connect-MgGraph -Identity
	## Connect to MS Graph (User Assigned Identity)
	#Connect-MgGraph -Identity -ClientId "93a29f3d-xxxx-xxxx-xxxx-1df61b36c9d9"

	# Do something with Graph
	$mySite = Get-MgSite -SiteId "cc41b7f0-xxxx-xxxx-xxxx-5afb83a210a8"
	Write-Output "-- GraphOutput: " $mySite.DisplayName " - " $mySite.Id

	# Disconnect from the Graph SDK
	$disconn = Disconnect-MgGraph
}
#gavdcodeend 018

#gavdcodebegin 007
function PsPnP_ManagedIdentyExample
{
	Import-Module -Name "PnP.PowerShell"

	# Connect to PnP using a System Assigned Managed Identity
	Connect-PnPOnline -Url "https://[domain].sharepoint.com/sites/[SiteName]" `
					  -ManagedIdentity
	## Connect to PnP using a User Assigned Managed Identity
	#Connect-PnPOnline -Url "https://[domain].sharepoint.com/sites/[SiteName]" `
	#	-ManagedIdentity `
	#	-UserAssignedManagedIdentityObjectId "93a29f3d-xxxx-xxxx-xxxx-1df61b36c9d9"

	# Do I have a token?
	#Get-PnPAccessToken # Token to consume the Microsoft Graph API (if necessary)

	# Do something with PnP
	$mySite = Get-PnPSite -Includes Id,RootWeb.Title
	Write-Output "-- PnPOutput: " $mySite.RootWeb.Title " - " $mySite.Id " - " $mySite.Url

	# Disconnect from PnP
	$disc = Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 008
function PsTeamsPs_ManagedIdentyExample # Deprecated
{
	# Deprecated by Microsoft
	# There is no more direct support for Managed Identities in the Teams PowerShell module
	
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

# Example CSharp for Functions (in other repo) (Graph REST API and Graph CSharp SDK)

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 019 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

#PsAzManagedIdenty_CreateManagedIdentity_UserAssigned -resourceGroupName "Chapter06" `
#             -location "westeurope" -identityName "ManagedIdentityUserAssigned_01"

#PsAzManagedIdenty_DeleteManagedIdentity_UserAssigned -resourceGroupName "Chapter06" `
#									 -identityName "ManagedIdentityUserAssigned_01"

#PsGraphPsSdkManagedIdenty_SetClaims -azureUser $configFile.appsettings.UserName `
#									-azureUserPw $configFile.appsettings.UserPw `
#									-azureTenantId "ade56059-89c0-4594-90c3-e4772a8168ca" `
#									-appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#									-appRole "Sites.FullControl.All"

#PsGraphPsSdkManagedIdenty_RemoveClaims -azureUser $configFile.appsettings.UserName `
#									-azureUserPw $configFile.appsettings.UserPw `
#									-azureTenantId "ade56059-89c0-4594-90c3-e4772a8168ca" `
#									-appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#									-appRole "Sites.FullControl.All"

#PsPnPManagedIdenty_SetClaims -azureUser $configFile.appsettings.UserName `
#									-azureUserPw $configFile.appsettings.UserPw `
#									-azureTenantId "ade56059-89c0-4594-90c3-e4772a8168ca" `
#									-appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#									-appRole "Sites.FullControl.All"

#PsPnPManagedIdenty_RemoveClaims -azureUser $configFile.appsettings.UserName `
#									-azureUserPw $configFile.appsettings.UserPw `
#									-azureTenantId "ade56059-89c0-4594-90c3-e4772a8168ca" `
#									-appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#									-appRole "Sites.FullControl.All"

#PsAzManagedIdenty_SetClaimsById -appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#								-appRole "Sites.FullControl.All"

#PsAzManagedIdenty_SetClaimsByName -managedIdentityName "ManagedIdentityUserAssigned_01" `
#								  -resourceGroupName "Chapter06" `
#								  -appRole "Sites.FullControl.All"

#PsAzManagedIdenty_RemoveClaimsById -appId "ff852e59-f3a9-4445-b27f-4d15b2659fc1" `
#								  -appRole "Sites.FullControl.All"

#PsAzManagedIdenty_RemoveClaimsByName -managedIdentityName "ManagedIdentityUserAssigned_01" `
#								  -resourceGroupName "Chapter06" `
#								  -appRole "Sites.FullControl.All"

#PsAzManagedIdenty_GetManagedIdentities_AllUserAssigned
#PsAzManagedIdenty_GetManagedIdentities_OneUserAssignedByGroup `
#								-resourceGroupName "Chapter06"
#PsAzManagedIdenty_GetManagedIdentities_OneUserAssignedByGroupAndName `
#								-resourceGroupName "Chapter06" `
#								-identityName "ManagedIdentityUserAssigned_01"

Write-Host "Done"  

