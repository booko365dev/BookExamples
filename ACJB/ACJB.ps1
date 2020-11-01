
#gavdcodebegin 01
Function LoginPsCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}
#gavdcodeend 01

#gavdcodebegin 02
Function LoginPsPSO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.spAdminUrl -Credential $myCredentials
}
#gavdcodeend 02

#gavdcodebegin 03
Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}
#gavdcodeend 03

#gavdcodebegin 04
Function Invoke-RestSPO() {
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
Function Get-SPOContextInfo(){
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

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) {
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}
#gavdcodeend 05

#----------------------------------------------------------------------------------------

#gavdcodebegin 06
Function PsRestGetExample(){
	$webUrl = $configFile.appsettings.spUrl
	$userName = $configFile.appsettings.spUserName
	$password = $configFile.appsettings.spUserPw

	$endpointUrl = $WebUrl + "/_api/web/Created"
	$myResult = Invoke-RestSPO -Url $endpointUrl `
							   -Method Get `
							   -UserName $userName `
							   -Password $password
	$myResult
}
#gavdcodeend 06

#gavdcodebegin 07
Function PsRestPostExample(){
	$webUrl = $configFile.appsettings.spUrl
	$userName = $configFile.appsettings.spUserName
	$password = $configFile.appsettings.spUserPw

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
Function PsCsomExample(){
	$spCtx = LoginPsCsom
	
	$rootWeb = $spCtx.Web
	$spCtx.Load($rootWeb)
	$spCtx.ExecuteQuery()
	$spCtx.Dispose()
	
	Write-Host $rootWeb.Created  
}
#gavdcodeend 08

#gavdcodebegin 09
Function PsPsoExample(){
	LoginPsPSO
	Test-SPOSite $configFile.appsettings.spUrl
	Disconnect-SPOService
}
#gavdcodeend 09

#gavdcodebegin 10
Function PsPnpExample(){
	LoginPsPnP
	
	$myWeb = Get-PnPWeb
	
	Write-Host $myWeb.Title
	
	#Disconnect-PnPOnline
}
#gavdcodeend 10

#gavdcodebegin 11
Function PsPnpRestGetExample(){
	LoginPsPnP
	
	$myWeb = Invoke-PnPSPRestMethod -Url /_api/web
	
	Write-Host $myWeb.Title
}
#gavdcodeend 11

#gavdcodebegin 12
Function PsPnpRestPostExample01(){
	LoginPsPnP
	
	$myBody = "{'Title':'Test01'}"
	Invoke-PnPSPRestMethod -Method Post `
						   -Url "/_api/web/lists/GetByTitle('TestList')/items" `
						   -Content $myBody
}
#gavdcodeend 12

#gavdcodebegin 13
Function PsPnpRestPostExample02(){
	LoginPsPnP
	
	$myBody = "{ '__metadata': { 'type': 'SP.ListItem' }, 'Title': 'Test02'}"
	Invoke-PnPSPRestMethod -Method Post `
						   -Url "/_api/web/lists/GetByTitle('TestList')/items" `
						   -Content $myBody `
						   -ContentType "application/json;odata=verbose"
}
#gavdcodeend 13

#----------------------------------------------------------------------------------------

# Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

##==> CSOM
#PsCsomExample

##==> PSO
#PsPsoExample

##==> PnP PowerShell
#PsPnpExample

#==> REST PowerShell cmdlets
#PsRestGetExample       ## Simple GET request without body
#PsRestPostExample      ## Full POST query with data in the body

#==> REST PnP PowerShell cmdlets
#PsPnpRestGetExample
#PsPnpRestPostExample01
#PsPnpRestPostExample02

Write-Host ""  

