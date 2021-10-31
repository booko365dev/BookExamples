 
#gavdcodebegin 01
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
#gavdcodeend 01 

#gavdcodebegin 07
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
#gavdcodeend 07 

#*** Using Classic PowerShell cmdlets ---------------------------------------------------

#gavdcodebegin 02
Function GrPsGetTeam()
{
	$Url = "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										  -ClientSecret $ClientSecretApp `
										  -TenantName $TenantName

	<#
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	#>
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 02 

#gavdcodebegin 03
Function GrPsCreateChannel()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										  -ClientSecret $ClientSecretApp `
										  -TenantName $TenantName

	<#
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	#>
    
	$myBody = "{ 'displayName':'Graph Channel 25', `
                 'description':'Channel created with Graph' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
											-ContentType $myContentType -Method Post

	Write-Host $myResult
}
#gavdcodeend 03 

#gavdcodebegin 04
Function GrPsGetChannel()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" +
							"/channels/19:a43782b050c547da86a05649bf00083e@thread.tacv2"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										  -ClientSecret $ClientSecretApp `
										  -TenantName $TenantName

	<#
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	#>
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url

	Write-Host $myResult
}
#gavdcodeend 04 

#gavdcodebegin 05
Function GrPsUpdateChannel()
{
	$Url = 
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" +
							"/channels/19:a43782b050c547da86a05649bf00083e@thread.tacv2"

	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										  -ClientSecret $ClientSecretApp `
										  -TenantName $TenantName

	<#
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	#>
    
	$myBody = "{ 'description':'Channel Description Updated' }"
	$myContentType = "application/json"
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)";
				   'IF-MATCH' = '*' }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Body $myBody `
											-ContentType $myContentType -Method Patch

	Write-Host $myResult
}
#gavdcodeend 05

#gavdcodebegin 06
Function GrPsDeleteChannel()
{
	$Url = 
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:a43782b050c547da86a05649bf00083e@thread.tacv2"
	
	$myOAuth = Get-AzureTokenApplication -ClientID $ClientIDApp `
										  -ClientSecret $ClientSecretApp `
										  -TenantName $TenantName

	<#
	$myOAuth = Get-AzureTokenDelegation -ClientID $ClientIDDel `
										-TenantName $TenantName `
										-UserName $UserName `
										-UserPw $UserPw
	#>
	
	$myHeader = @{ 'Authorization' = "$($myOAuth.token_type) $($myOAuth.access_token)" }
	$myResult = Invoke-WebRequest -Headers $myHeader -Uri $Url -Method Delete

	Write-Host $myResult
}
#gavdcodeend 06

#*** Using Classic PowerShell cmdlets ---------------------------------------------------

#gavdcodebegin 08
Function GrPsLoginGraphInteractive()
{
	Connect-Graph
}
#gavdcodeend 08

#gavdcodebegin 09
Function GrPsLoginGraphAssignRights()
{
	Connect-Graph -Scopes "User.Read","User.ReadWrite.All","Mail.ReadWrite",`
				"Directory.Read.All","Chat.ReadWrite", "People.Read", `
				"Group.Read.All", "Tasks.ReadWrite", `
				"Sites.Manage.All"
}
#gavdcodeend 09

#gavdcodebegin 10
Function GrPsLoginGraphApplication()
{
	Connect-Graph -TenantId "e82079f9-5663-48aa-a0c7-c9b7b4db2cf5"
}
#gavdcodeend 10

#gavdcodebegin 11
Function GrPsGetGroupsSelect()
{
	Get-MgGroup | Select-Object id, DisplayName, GroupTypes
}
#gavdcodeend 11

#*** Using PowerShell-MicrosoftGraphAPI module ------------------------------------------

#gavdcodebegin 12
Function GrPsGetToken()
{
	Import-Module .\MicrosoftGraph.psm1

	$myCredential = New-Object System.Management.Automation.PSCredential(`
			$ClientIDApp,(ConvertTo-SecureString $ClientSecretApp -AsPlainText -Force))
	$myToken = Get-MSGraphAuthToken -Credential $myCredential -TenantID $TenantID

	Return $myToken
}
#gavdcodeend 12

#gavdcodebegin 13
Function GrPsGetTeamWithModule()
{
	$myToken = GrPsGetToken
	Invoke-MSGraphQuery `
		-URI "https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" `
		-Token $myToken
}
#gavdcodeend 13

#gavdcodebegin 14
Function GrPsCreateChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'displayName':'Graph Channel 40', `
				 'description':'Channel created with Graph' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Post
}
#gavdcodeend 14

#gavdcodebegin 15
Function GrPsUpdateChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Patch
}
#gavdcodeend 15

#gavdcodebegin 16
Function GrPsDeleteChannelWithModule()
{
	$Url = `
		"https://graph.microsoft.com/v1.0/teams/5b409eec-a4ae-4f04-a354-0434c444265d" + 
							"/channels/19:bb17af0c3a894262809c5412606f09f3@thread.tacv2"
	
	$myToken = GrPsGetToken
	$myBody = "{ 'description':'Channel Description Updated' }"
	Invoke-MSGraphQuery `
		-URI $Url `
		-Body $myBody `
		-Token $myToken `
		-Meth Delete
}
#gavdcodeend 16

#*** Using Graph PowerShell PnP ---------------------------------------------------------------

#gavdcodebegin 17
Function GrPsLoginGraphPnP()
{
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -DeviceLogin -LaunchBrowser

	#Disconnect-PnPOnline
}
#gavdcodeend 17

#gavdcodebegin 18
Function GrPsLoginGraphPnPExample()
{
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -DeviceLogin -LaunchBrowser
	Get-PnPTeamsUser -Team "Design"

	#Disconnect-PnPOnline
}
#gavdcodeend 18

#gavdcodebegin 19
Function GrPsLoginGraphPnPInteractive()
{
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -Interactive
}
#gavdcodeend 19

#gavdcodebegin 20
Function GrPsLoginGraphPnPWithCredentials()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
		$configFile.appsettings.UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential `
		-argumentlist $configFile.appsettings.UserName, $securePW

	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -Credentials $myCredentials
}
#gavdcodeend 20

#gavdcodebegin 21
Function GrPsLoginGraphPnPGetToken()
{
	Connect-PnPOnline -Url $configFile.appsettings.TenantUrl -DeviceLogin -LaunchBrowser
	Get-PnPGraphAccessToken -Decoded

	#Disconnect-PnPOnline
}
#gavdcodeend 21

#----------------------------------------------------------------------------------------

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$TenantID = $configFile.appsettings.TenantId
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#*** Using Classic PowerShell cmdlets
#GrPsGetTeam
#GrPsCreateChannel
#GrPsGetChannel
#GrPsUpdateChannel
#GrPsDeleteChannel

#*** Using Microsoft Graph PowerShell cmdlets
#GrPsLoginGraphInteractive
#GrPsLoginGraphAssignRights
#GrPsLoginGraphApplication
#GrPsGetGroupsSelect

#*** Using PowerShell-MicrosoftGraphAPI module
#GrPsGetToken
#GrPsGetTeamWithModule
#GrPsCreateChannelWithModule
#GrPsUpdateChannelWithModule
#GrPsDeleteChannelWithModule

#*** Using Graph PowerShell PnP
#GrPsLoginGraphPnP
#GrPsLoginGraphPnPExample
#GrPsLoginGraphPnPInteractive
#GrPsLoginGraphPnPWithCredentials
#GrPsLoginGraphPnPGetToken

Write-Host "Done" 
