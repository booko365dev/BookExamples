
#gavdcodebegin 01
Function LoginPsCsom()  #*** USE POWERSHELL 5.x, NOT 7.x ***
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
#gavdcodeend 01

#gavdcodebegin 02
Function LoginPsPSO()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}
#gavdcodeend 02

#gavdcodebegin 03
Function LoginPsPnP() #*** LEGACY CODE *** 
{
    # ATTENTION: Using the deprecated PnP-PowerShell module
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#gavdcodeend 03

#gavdcodebegin 16
Function LoginPsPnPPowerShell()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#gavdcodeend 16

#gavdcodebegin 18
Function LoginPsPnPPowerShellCertificate()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}
#gavdcodeend 18

#gavdcodebegin 20
Function LoginPsPnPPowerShellCertificateBase64()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}
#gavdcodeend 20

#gavdcodebegin 21
Function LoginPsPnPPowerShellInteractive()
{
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}
#gavdcodeend 21

#gavdcodebegin 14
Function LoginPsCLIWithAccPw()
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}
#gavdcodeend 14

#gavdcodebegin 23
Function LoginPsCLIWithSecret()
{
	m365 login --authType secret `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithSecret `
			   --secret $configFile.appsettings.ClientSecret
}
#gavdcodeend 23

#gavdcodebegin 25
Function LoginPsCLIWithCertificate()
{
	m365 login --authType certificate `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithCert `
			   --certificateFile $configFile.appsettings.CertificateFilePath `
			   --password $configFile.appsettings.CertificateFilePw
}
#gavdcodeend 25

#gavdcodebegin 04
Function Invoke-RestSPO() #*** LEGACY CODE ***  
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
#gavdcodeend 04
 
#gavdcodebegin 05
Function Get-SPOContextInfo() #*** LEGACY CODE *** 
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
#gavdcodeend 05

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

#gavdcodebegin 06
Function PsRestGetExample() #*** LEGACY CODE *** 
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
#gavdcodeend 06

#gavdcodebegin 07
Function PsRestPostExample() #*** LEGACY CODE *** 
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
#gavdcodeend 07

#gavdcodebegin 08
Function PsCsomExample()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	$spCtx = LoginPsCsom
	
	$rootWeb = $spCtx.Web
	$spCtx.Load($rootWeb)
	$spCtx.ExecuteQuery()
	$spCtx.Dispose()
	
	Write-Host $rootWeb.Created  
}
#gavdcodeend 08

#gavdcodebegin 09
Function PsPsoExample()  #*** USE POWERSHELL 5.x, NOT 7.x ***
{
	LoginPsPSO
	Test-SPOSite $configFile.appsettings.SiteCollUrl
	Disconnect-SPOService
}
#gavdcodeend 09

#gavdcodebegin 10
Function PsPnpExample()    #*** LEGACY CODE ***
{ 
    # ATTENTION: Using the deprecated PnP-PowerShell module
	LoginPsPnP
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 10

#gavdcodebegin 17
Function PsPnpPowerShellExample()
{
	LoginPsPnPPowerShell
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 17

#gavdcodebegin 19
Function PsPnpPowerShellCertificateExample()
{
	LoginPsPnPPowerShellCertificate
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 19

#gavdcodebegin 22
Function PsPnpPowerShellInteractiveExample()
{
	LoginPsPnPPowerShellInteractive
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 22

#gavdcodebegin 11
Function PsPnpRestGetWebExample()
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
#gavdcodeend 11

#gavdcodebegin 12
Function PsPnpRestGetItemsExample()
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
#gavdcodeend 12

#gavdcodebegin 13
Function PsPnpRestPostExample()
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
#gavdcodeend 13

#gavdcodebegin 15
Function PsCliExampleWithAccPw()
{
	LoginPsCLIWithAccPw
	
	m365 spo site list --type TeamSite

	m365 logout
}
#gavdcodeend 15

#gavdcodebegin 24
Function PsCliExampleWithSecret()
{
	LoginPsCLIWithSecret
	
	m365 teams team list

	m365 logout
}
#gavdcodeend 24

#gavdcodebegin 26
Function PsCliExampleWithCertificate()
{
	LoginPsCLIWithCertificate
	
	m365 spo tenant settings list

	m365 logout
}
#gavdcodeend 26

#----------------------------------------------------------------------------------------

# Running the Functions
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
#PsCliExampleWithAccPw
#PsCliExampleWithSecret
#PsCliExampleWithCertificate

Write-Host ""  
Write-Host "Done"
Write-Host ""
