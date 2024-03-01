
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function LoginPsCsom  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.SiteCollUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}
#gavdcodeend 001

#gavdcodebegin 002
Function LoginPsPSO  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}
#gavdcodeend 002

#gavdcodebegin 003
Function LoginPsPnP #*** LEGACY CODE *** 
{
    # ATTENTION: Using the deprecated PnP-PowerShell module
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#gavdcodeend 003

#gavdcodebegin 016
Function LoginPsPnPPowerShell
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#gavdcodeend 016

#gavdcodebegin 018
Function LoginPsPnPPowerShellCertificate
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}
#gavdcodeend 018

#gavdcodebegin 020
Function LoginPsPnPPowerShellCertificateBase64
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}
#gavdcodeend 020

#gavdcodebegin 021
Function LoginPsPnPPowerShellInteractive
{
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}
#gavdcodeend 021

#gavdcodebegin 014
Function PsCli_LoginWithAccPw
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}
#gavdcodeend 014

#gavdcodebegin 023
Function PsCli_LoginWithSecret
{
	m365 login --authType secret `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithSecret `
			   --secret $configFile.appsettings.ClientSecret
}
#gavdcodeend 023

#gavdcodebegin 025
Function PsCli_LoginWithCertificate
{
	m365 login --authType certificate `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithCert `
			   --certificateFile $configFile.appsettings.CertificateFilePath `
			   --password $configFile.appsettings.CertificateFilePw
}
#gavdcodeend 025


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 004
Function Invoke-RestSPO #*** LEGACY CODE ***  
{
	# ATTENTION: Legacy - Using the deprecated PnP-PowerShell module
	Param (
		[Parameter(Mandatory=$True)]
		[String]$Url,
 
		[Parameter(Mandatory=$False)]
		[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = `
								[Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password,
 
		[Parameter(Mandatory=$False)]
		[String]$Metadata,

		[Parameter(Mandatory=$False)]
		[System.Byte[]]$Body,
 
		[Parameter(Mandatory=$False)]
		[String]$RequestDigest,
 
		[Parameter(Mandatory=$False)]
		[String]$ETag,
 
		[Parameter(Mandatory=$False)]
		[String]$XHTTPMethod,

		[Parameter(Mandatory=$False)]
		[System.String]$Accept = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[String]$ContentType = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[Boolean]$BinaryStringResponseBody = $False
	)
 
	if([string]::IsNullOrEmpty($Password)) {
		$SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
	}
	else {
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
	}
 
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials(`
															$UserName, $SecurePassword)
	$request = [System.Net.WebRequest]::Create($Url)
	$request.Credentials = $credentials
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	$request.ContentType = $ContentType
	$request.Accept = $Accept
	$request.Method=$Method
 
	if($RequestDigest) { 
		$request.Headers.Add("X-RequestDigest", $RequestDigest)
	}
	if($ETag) { 
		$request.Headers.Add("If-Match", $ETag)
	}
	if($XHTTPMethod) { 
		$request.Headers.Add("X-HTTP-Method", $XHTTPMethod)
	}
	if($Metadata -or $Body) {
		if($Metadata) {     
			$Body = [byte[]][char[]]$Metadata
		}      
		$request.ContentLength = $Body.Length 
		$stream = $request.GetRequestStream()
		$stream.Write($Body, 0, $Body.Length)
	}
	else {
		$request.ContentLength = 0
	}

	$response = $request.GetResponse()
	try {
		if($BinaryStringResponseBody -eq $False) {    
			$streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
			try {
				$data=$streamReader.ReadToEnd()
				$results = $data | ConvertFrom-Json
				$results.d 
			}
			finally {
				$streamReader.Dispose()
			}
		}
		else {
			$dataStream = New-Object System.IO.MemoryStream
			try {
			Stream-CopyTo -Source $response.GetResponseStream() -Destination $dataStream
			$dataStream.ToArray()
			}
			finally {
				$dataStream.Dispose()
			} 
		}
	}
	finally {
		$response.Dispose()
	}
}
#gavdcodeend 004
 
#gavdcodebegin 005
Function Get-SPOContextInfo #*** LEGACY CODE *** 
{
	# ATTENTION: Legacy - Using the deprecated PnP-PowerShell module
	Param(
		[Parameter(Mandatory=$True)]
		[String]$WebUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password
	)
   
	$Url = $WebUrl + "/_api/contextinfo"
	Invoke-RestSPO $Url Post $UserName $Password
}

Function Stream-CopyTo(
	[System.IO.Stream]$Source, [System.IO.Stream]$Destination) #*** LEGACY CODE *** 
{
	# ATTENTION: Legacy - Using the deprecated PnP-PowerShell module
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}
#gavdcodeend 005

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 006
Function PsRestGetExample #*** LEGACY CODE *** 
{
	# ATTENTION: Legacy - Using the deprecated PnP-PowerShell module
	$webUrl = $configFile.appsettings.SiteCollUrl
	$userName = $configFile.appsettings.UserName
	$password = $configFile.appsettings.UserPw

	$endpointUrl = $WebUrl + "/_api/web/Created"
	$myResult = Invoke-RestSPO -Url $endpointUrl `
							   -Method Get `
							   -UserName $userName `
							   -Password $password
	$myResult
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsRestPostExample #*** LEGACY CODE *** 
{
	# ATTENTION: Legacy - Using the deprecated PnP-PowerShell module
	$webUrl = $configFile.appsettings.SiteCollUrl
	$userName = $configFile.appsettings.UserName
	$password = $configFile.appsettings.UserPw

	$endpointUrl = $WebUrl + "/_api/web/lists"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.List' }; 
				Title = 'NewListRest'; 
				BaseTemplate = 100; 
				Description = 'Test NewListRest'; 
	            AllowContentTypes = $true;
	            ContentTypesEnabled = $true
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl `
									  -UserName $userName `
									  -Password $password
	Invoke-RestSPO -Url $endpointUrl `
				   -Method Post `
				   -UserName $userName `
				   -Password $password `
				   -Metadata $myPayload `
				   -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsCsomExample  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$spCtx = LoginPsCsom
	
	$rootWeb = $spCtx.Web
	$spCtx.Load($rootWeb)
	$spCtx.ExecuteQuery()
	$spCtx.Dispose()
	
	Write-Host $rootWeb.Created  
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsPsoExample  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	LoginPsPSO
	Test-SPOSite $configFile.appsettings.SiteCollUrl
	Disconnect-SPOService
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsPnpExample    #*** LEGACY CODE ***
{ 
    # ATTENTION: Using the deprecated PnP-PowerShell module
	LoginPsPnP
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 017
Function PsPnpPowerShellExample
{
	LoginPsPnPPowerShell
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 017

#gavdcodebegin 019
Function PsPnpPowerShellCertificateExample
{
	LoginPsPnPPowerShellCertificate
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 022
Function PsPnpPowerShellInteractiveExample
{
	LoginPsPnPPowerShellInteractive
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 022

#gavdcodebegin 011
Function PsPnpRestGetWebExample
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken
	
	$endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	Write-Host $dataObject.d.Title
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsPnpRestGetItemsExample
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken
	
	$endpointUrl = $configFile.appsettings.SiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/items" + 
						"?`$filter=startswith(Title,'ItemOne')"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content.Replace("ID", "_ID") | ConvertFrom-Json
	Write-Host $dataObject.d.results.Title
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsPnpRestPostExample
{
	LoginPsPnPPowerShellWithAccPwDefault
	$myOAuth = Get-PnPAppAuthAccessToken
	
	$endpointUrl = $configFile.appsettings.SiteCollUrl + 
						"/_api/web/lists/GetByTitle('TestList')/items"
	$myPayload = 
		"{
		  '__metadata': { 'type': 'SP.ListItem' },
		  'Title': 'ItemTwo'
		}"
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }

	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 013

#gavdcodebegin 015
Function PsCli_ExampleWithAccPw
{
	PsCli_LoginWithAccPw
	
	m365 spo site list --type TeamSite

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 024
Function PsCli_ExampleWithSecret
{
	PsCli_LoginWithSecret
	
	m365 teams team list

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 026
Function PsCli_ExampleWithCertificate
{
	PsCli_LoginWithCertificate
	
	m365 spo tenant settings list

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
Function PsCli_GetRoleAssigmentsById
{
	PsCli_LoginWithAccPw
	
	m365 entra approleassignment list --appId "377568de-bf5d-4ea6-b4d7-b8e6dac488c9"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsCli_GetRoleAssigmentsByName
{
	PsCli_LoginWithAccPw
	
	m365 entra approleassignment list --appDisplayName "TestAppReg"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
Function PsCli_AddRoleAssigments
{
	PsCli_LoginWithAccPw
	
	m365 entra approleassignment add --appId "377568de-bf5d-4ea6-b4d7-b8e6dac488c9" `
									 --resource "Microsoft Graph" `
									 --scopes "Sites.FullControl.All"

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsCli_DeleteRoleAssigments
{
	PsCli_LoginWithAccPw
	
	m365 entra approleassignment remove --appId "377568de-bf5d-4ea6-b4d7-b8e6dac488c9" `
									    --resource "Microsoft Graph" `
									    --scopes "Sites.FullControl.All" `
										--force

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsCli_GetAccessTokenGraph
{
	PsCli_LoginWithAccPw
	
	m365 util accesstoken get --resource graph

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
Function PsCli_GetAccessTokenSharePoint
{
	PsCli_LoginWithAccPw
	
	m365 spo set --url "https://[domain].sharepoint.com"
	m365 util accesstoken get --resource sharepoint

	m365 logout
}
#gavdcodeend 032


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 032 ***

#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

##==> CSOM
#PsCsomExample

##==> PSO
#PsPsoExample

##==> PnP-PowerShell (Legacy)
# ATTENTION: Using the deprecated PnP-PowerShell module
#PsPnpExample

##==> PnP PowerShell (Replacement of the legacy PnP-PowerShell)
#PsPnpPowerShellExample
#PsPnpPowerShellCertificateExample
#PsPnpPowerShellInteractiveExample

#==> REST PowerShell cmdlets (Legacy)
# ATTENTION: Using the deprecated PnP-PowerShell module
#PsRestGetExample       ## Simple GET request without body
#PsRestPostExample      ## Full POST query with data in the body

#==> REST PnP PowerShell cmdlets
#PsPnpRestGetWebExample
#PsPnpRestGetItemsExample
#PsPnpRestPostExample

##==> CLI
#PsCli_ExampleWithAccPw
#PsCli_ExampleWithSecret
#PsCli_ExampleWithCertificate
#PsCli_GetRoleAssigmentsById
#PsCli_GetRoleAssigmentsByName
#PsCli_AddRoleAssigments
#PsCli_DeleteRoleAssigments
#PsCli_GetAccessTokenGraph
#PsCli_GetAccessTokenSharePoint

Write-Host "Done"
