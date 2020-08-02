 
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

#----------------------------------------------------------------------------------------

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

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\grPs.values.config"

$ClientIDApp = $configFile.appsettings.ClientIdApp
$ClientSecretApp = $configFile.appsettings.ClientSecretApp
$ClientIDDel = $configFile.appsettings.ClientIdDel
$TenantName = $configFile.appsettings.TenantName
$UserName = $configFile.appsettings.UserName
$UserPw = $configFile.appsettings.UserPw

#GrPsGetTeam
#GrPsCreateChannel
#GrPsGetChannel
#GrPsUpdateChannel
#GrPsDeleteChannel

Write-Host "Done" 
