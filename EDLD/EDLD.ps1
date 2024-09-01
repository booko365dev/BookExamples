
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function Invoke-RestSPO 
{  #*** LEGACY CODE ***
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
				$results = $data.ToString().Replace("ID", "_ID") | ConvertFrom-Json
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
 
Function Get-SPOContextInfo
{  #*** LEGACY CODE ***
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

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) 
{
	  #*** LEGACY CODE ***
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

Function PsSpPnpPowerShell_LoginWithAccPw
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

# Using Basic Authentication, this is Legacy code, and cannot be used anymore

#gavdcodebegin 001
Function PsSpSpRestApi_CreateOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemRestPs'
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 001 

#gavdcodebegin 002
Function PsSpSpRestApi_UploadOneDocument  #*** LEGACY CODE ***
{
	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $WebUrl + "/_api/web/GetFolderByServerRelativeUrl('" + $WebUrlRel + 
				"/TestLibrary')/Files/add(url='" + $FileInfo.Name + "',overwrite=true)"
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)
	$contextInfo = Get-SPOContextInfo  $WebUrl $UserName $Password
	$data = Invoke-RestSPO -Url $endpointUrl -Method Post -UserName $userName -Password `
						$password -Body $FileContent -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsSpSpRestApi_DownloadOneDocument  #*** LEGACY CODE ***
{
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $WebUrl + "/_api/web/GetFileByServerRelativeUrl('" + $WebUrlRel + 
				"/TestLibrary/TestDocument01.docx')/`$value"
	$fileContent = Invoke-RestSPO -Url $endpointUrl -Method Get -UserName $UserName `
									-Password $Password -BinaryStringResponseBody $True
	$fileName = [System.IO.Path]::GetFileName("TestDwload.docx")
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsSpSpRestApi_ReadAllListItems  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsSpSpRestApi_ReadOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items(26)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsSpSpRestApi_ReadAllLibraryDocs  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsSpSpRestApi_ReadOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items(25)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsSpSpRestApi_UpdateOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(26)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemCsRest_Updated'
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsSpSpRestApi_UpdateOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestLibrary')/items(25)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'TestDocument01_Updated.docx'
			   } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsSpSpRestApi_DeleteOneListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(26)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsSpSpRestApi_DeleteOneLibraryDoc  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestLibrary')/items(25)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsSpSpRestApi_BreakSecurityInheritanceListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/" +
						"items(27)/breakroleinheritance(copyRoleAssignments=false," +
						"clearSubscopes=true)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsSpSpRestApi_ResetSecurityInheritanceListItem  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/" +
						"items(27)/resetroleinheritance"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsSpSpRestApi_AddUserToSecurityRoleInListItem  #*** LEGACY CODE ***
{
	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $WebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
							-Password $password -RequestDigest `
							$contextInfo.GetContextWebInformation.FormDigestValue `
							-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsSpSpRestApi_UpdateUserSecurityRoleInListItem  #*** LEGACY CODE ***
{
	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Find the RoleDefinitions
	$endpointUrl = $WebUrl + "/_api/web/roledefinitions?$select=Id&" +
	                                "$filter=startswith(Name,'Full Control')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $roleId = $data.results[0].Id
	$data | ConvertTo-Json

    # Add the User in the RoleDefinion to the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + 
			$roleId + ")"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "MERGE"
	$data | ConvertTo-Json
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsSpSpRestApi_DeleteUserFromSecurityRoleInListItem  #*** LEGACY CODE ***
{
    # Find the User
	$endpointUrl = $WebUrl + "/_api/web/siteusers?$select=Id&" +
	                                "$filter=startswith(Title,'MOD')"
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName `
																-Password $password
    $userId = $data.results[0].Id
	$data | ConvertTo-Json

    # Remove the User from the List
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(27)/" +
					     "roleassignments/getbyprincipalid(principalid=" + $userId + ")";
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName `
						-Password $password -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue `
						-ETag "*" -XHTTPMethod "DELETE"
	$data | ConvertTo-Json
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsSpSpRestApi_CreateOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/Folders"
    $myPayload = @{
				__metadata = @{ 'type' = 'SP.Folder' };
				ServerRelativeUrl = $myServerRelativeUrl
			} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
    
	$data | ConvertTo-Json
}
#gavdcodeend 017 

#gavdcodebegin 018
Function PsSpSpRestApi_ReadAllFolders  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                            $myServerRelativeUrl + "')/ListItemAllFields"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsSpSpRestApi_RenameOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                            $myServerRelativeUrl + "')/ListItemAllFields"
    $myPayload = @{
        __metadata = @{ 'type' = 'SP.Data.TestDocumentsItem' };
        Title = 'RestFolderPSRenamed';
        FileLeafRef = 'RestFolderPSRenamend'
    } | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue -ETag "*" `
						-XHTTPMethod "MERGE"

	$data | ConvertTo-Json
}
#gavdcodeend 019 

#gavdcodebegin 020
Function PsSpSpRestApi_DeleteOneFolder  #*** LEGACY CODE ***
{
    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + 
                                                $myServerRelativeUrl + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 020 

#gavdcodebegin 021
Function PsSpSpRestApi_CreateOneAttachment  #*** LEGACY CODE ***
{
	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\Test.csv")
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

    $endpointUrl = $webUrl + "/_api/lists/GetByTitle('TestList')" + `
                        "/items(3)/AttachmentFiles/add(FileName='" + $FileInfo.Name + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Body $FileContent -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue
    
	$data | ConvertTo-Json
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsSpSpRestApi_ReadAllAttachments  #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/lists/GetByTitle('TestList')" + `
                                                        "/items(3)/AttachmentFiles"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsSpSpRestApi_DownloadOneAttachmentByFileName  #*** LEGACY CODE ***
{
	$myFileName = "Test.csv"

	$endpointUrl = $WebUrl + "/_api/lists/GetByTitle('TestList')" +
                                      "/items(3)/AttachmentFiles('" + $myFileName + "')" +
                                      "/$value"
	$fileContent = Invoke-RestSPO -Url $endpointUrl -Method Get -UserName $UserName `
									-Password $Password -BinaryStringResponseBody $True
	$fileName = [System.IO.Path]::GetFileName($myFileName)
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsSpSpRestApi_DeleteOneAttachmentByFileName  #*** LEGACY CODE ***
{
	$myFileName = "Test.csv"

	$endpointUrl = $WebUrl + "/_api/lists/GetByTitle('TestList')" + `
                                      "/items(3)/AttachmentFiles('" + $myFileName + "')"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 024

#----------------------------------------------------------------------------------------

# Using Azure AD Authentication through Connect-PnPOnline and an Account/PW App Registration

#gavdcodebegin 101
Function PsSpSpRestApi_CreateOneListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/web/lists/getbytitle('TestList')/items"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemRestPs'
			   } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 101 

#gavdcodebegin 102
Function PsSpSpRestApi_UploadOneDocumentAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\TestDocument01.docx")
	$WebUri = [System.Uri]$configFile.appsettings.SiteCollUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
				"/_api/web/GetFolderByServerRelativeUrl('" + $WebUrlRel + `
				"/TestLibrary')/Files/add(url='" + $FileInfo.Name + "',overwrite=true)"
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 102

#gavdcodebegin 103
Function PsSpSpRestApi_DownloadOneDocumentAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$WebUri = [System.Uri]$configFile.appsettings.SiteCollUrl
	$WebUrlRel = $WebUri.AbsolutePath

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
				"/_api/web/GetFileByServerRelativeUrl('" + $WebUrlRel + `
				"/TestLibrary/TestDocument01.docx')/`$value"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$fileContent = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$fileName = [System.IO.Path]::GetFileName("TestDowload.docx")
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 103

#gavdcodebegin 104
Function PsSpSpRestApi_ReadAllListItemsAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestList')/items?" + `
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 104

#gavdcodebegin 105
Function PsSpSpRestApi_ReadOneListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestList')/items(2)?" + `
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 105

#gavdcodebegin 106
Function PsSpSpRestApi_ReadAllLibraryDocsAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestLibrary')/items?" + `
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 106

#gavdcodebegin 107
Function PsSpSpRestApi_ReadOneLibraryDocAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/web/lists/getbytitle('TestLibrary')/items(1)?" + 
						"$select=Title,Id"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 107

#gavdcodebegin 108
Function PsSpSpRestApi_UpdateOneListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/items(2)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'NewListItemCsRest_Updated'
			   } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 108

#gavdcodebegin 109
Function PsSpSpRestApi_UpdateOneLibraryDocAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestLibrary')/items(1)"
	$myPayload = @{ 
				__metadata = @{ 'type' = 'SP.ListItem' }; 
				Title = 'TestDocument01_Updated.docx'
			   } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 109

#gavdcodebegin 110
Function PsSpSpRestApi_DeleteOneListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/items(2)"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 110

#gavdcodebegin 111
Function PsSpSpRestApi_DeleteOneLibraryDocAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestLibrary')/items(1)"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 111

#gavdcodebegin 112
Function PsSpSpRestApi_BreakSecurityInheritanceListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/" + `
						"items(3)/breakroleinheritance(copyRoleAssignments=false," +
						"clearSubscopes=true)"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 112

#gavdcodebegin 113
Function PsSpSpRestApi_ResetSecurityInheritanceListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/" + `
						"items(3)/resetroleinheritance"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 113

#gavdcodebegin 114
Function PsSpSpRestApi_AddUserToSecurityRoleInListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/siteusers?$select=Id&" + `
	                        "$filter=startswith(Title,'Ade')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Find the RoleDefinitions
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/roledefinitions?$select=Id&" + `
	                        "$filter=startswith(Name,'Full Control')"
	
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Add the User in the RoleDefinion to the Item
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
			"/_api/lists/getbytitle('TestList')/items(3)/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 114

#gavdcodebegin 115
Function PsSpSpRestApi_UpdateUserSecurityRoleInListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	# Inheritance MUST be broken
    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?$select=Id&" + `
	                                "$filter=startswith(Title,'Ade')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Find the RoleDefinitions
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/roledefinitions?$select=Id&" + `
	                                "$filter=startswith(Name,'Edit')"
	
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$roleId = $dataObject.d.results[0].Id
	Write-Host "RoleId - " $roleId

    # Update the User in the RoleDefinion to the Item
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
			"/_api/lists/getbytitle('TestList')/items(3)/" + `
			"roleassignments/addroleassignment(principalid=" + $userId + ",roledefid=" + `
			$roleId + ")"
	
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 115

#gavdcodebegin 116
Function PsSpSpRestApi_DeleteUserFromSecurityRoleInListItemAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    # Find the User
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/siteusers?$select=Id&" + `
	                                "$filter=startswith(Title,'MOD')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$dataObject = $data.content | ConvertFrom-Json
	$userId = $dataObject.d.results[0].Id
	Write-Host "UserId - " $userId

    # Remove the User from the List
	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
						"/_api/lists/getbytitle('TestList')/items(3)/" + `
					    "roleassignments/getbyprincipalid(principalid=" + $userId + ")";

	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 116

#gavdcodebegin 117
Function PsSpSpRestApi_CreateOneFolderAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + "/_api/web/Folders"
    $myPayload = @{
				__metadata = @{ 'type' = 'SP.Folder' };
				ServerRelativeUrl = $myServerRelativeUrl
			} | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 117 

#gavdcodebegin 118
Function PsSpSpRestApi_ReadAllFoldersAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/GetFolderByServerRelativeUrl('" + `
                            $myServerRelativeUrl + "')/ListItemAllFields"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 118

#gavdcodebegin 119
Function PsSpSpRestApi_RenameOneFolderAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
							"/_api/web/GetFolderByServerRelativeUrl('" + `
                            $myServerRelativeUrl + "')/ListItemAllFields"
    $myPayload = @{
        __metadata = @{ 'type' = 'SP.Data.TestLibraryItem' };
        Title = 'RestFolderPSRenamed';
        FileLeafRef = 'RestFolderPSRenamed'
    } | ConvertTo-Json
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Merge `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 119 

#gavdcodebegin 120
Function PsSpSpRestApi_DeleteOneFolderAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myServerRelativeUrl = "/sites/[SiteName]/[LibraryName]/RestFolderPS"

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/web/GetFolderByServerRelativeUrl('" + `
                                    $myServerRelativeUrl + "')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 120 

#gavdcodebegin 121
Function PsSpSpRestApi_CreateOneAttachmentAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$FileInfo = New-Object System.IO.FileInfo("C:\Temporary\Test.csv")
	$FileContent = [System.IO.File]::ReadAllBytes($FileInfo.FullName)

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
					  "/_api/lists/GetByTitle('TestList')" + `
                      "/items(3)/AttachmentFiles/add(FileName='" + $FileInfo.Name + "')"
	$myPayload = $FileContent

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 121

#gavdcodebegin 122
Function PsSpSpRestApi_ReadAllAttachmentsAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $endpointUrl = $configFile.appsettings.SiteCollUrl + `
								"/_api/lists/GetByTitle('TestList')" + `
                                "/items(3)/AttachmentFiles"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 122

#gavdcodebegin 123
Function PsSpSpRestApi_DownloadOneAttachmentByFileNameAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$myFileName = "Test.csv"

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
									"/_api/lists/GetByTitle('TestList')" + `
                                    "/items(3)/AttachmentFiles('" + $myFileName + "')" + `
                                    "/$value"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$fileContent = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	$fileName = [System.IO.Path]::GetFileName($myFileName)
	$downloadFilePath = [System.IO.Path]::Combine("C:\Temporary", $fileName)
	[System.IO.File]::WriteAllBytes($downloadFilePath, $fileContent)
}
#gavdcodeend 123

#gavdcodebegin 124
Function PsSpSpRestApi_DeleteOneAttachmentByFileNameAD
{
	PsSpPnpPowerShell_LoginWithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$myFileName = "Test.csv"

	$endpointUrl = $configFile.appsettings.SiteCollUrl + `
										"/_api/lists/GetByTitle('TestList')" + `
										"/items(3)/AttachmentFiles('" + $myFileName + "')"
	
	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose'; `
				   'If-Match' = '*' }
	$data = Invoke-WebRequest -Method Delete `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 124

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 124 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

## Using LEGACY Classic Authentication (Do not use this code anymore)
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#PsSpSpRestApi_CreateOneListItem
#PsSpSpRestApi_UploadOneDocument
#PsSpSpRestApi_DownloadOneDocument
#PsSpSpRestApi_ReadAllListItems
#PsSpSpRestApi_ReadOneListItem
#PsSpSpRestApi_ReadAllLibraryDocs
#PsSpSpRestApi_ReadOneLibraryDoc
#PsSpSpRestApi_UpdateOneListItem
#PsSpSpRestApi_UpdateOneLibraryDoc
#PsSpSpRestApi_DeleteOneListItem
#PsSpSpRestApi_DeleteOneLibraryDoc
#PsSpSpRestApi_BreakSecurityInheritanceListItem
#PsSpSpRestApi_ResetSecurityInheritanceListItem
#PsSpSpRestApi_AddUserToSecurityRoleInListItem
#PsSpSpRestApi_UpdateUserSecurityRoleInListItem
#PsSpSpRestApi_DeleteUserFromSecurityRoleInListItem
#PsSpSpRestApi_CreateOneFolder
#PsSpSpRestApi_ReadAllFolders
#PsSpSpRestApi_RenameOneFolder
#PsSpSpRestApi_DeleteOneFolder
#PsSpSpRestApi_CreateOneAttachment
#PsSpSpRestApi_ReadAllAttachments
#PsSpSpRestApi_DownloadOneAttachmentByFileName
#PsSpSpRestApi_DeleteOneAttachmentByFileName

## Using Azure AD Authentication
#PsSpSpRestApi_CreateOneListItemAD
#PsSpSpRestApi_UploadOneDocumentAD
#PsSpSpRestApi_DownloadOneDocumentAD
#PsSpSpRestApi_ReadAllListItemsAD
#PsSpSpRestApi_ReadOneListItemAD
#PsSpSpRestApi_ReadAllLibraryDocsAD
#PsSpSpRestApi_ReadOneLibraryDocAD
#PsSpSpRestApi_UpdateOneListItemAD
#PsSpSpRestApi_UpdateOneLibraryDocAD
#PsSpSpRestApi_DeleteOneListItemAD
#PsSpSpRestApi_DeleteOneLibraryDocAD
#PsSpSpRestApi_BreakSecurityInheritanceListItemAD
#PsSpSpRestApi_ResetSecurityInheritanceListItemAD
#PsSpSpRestApi_AddUserToSecurityRoleInListItemAD
#PsSpSpRestApi_UpdateUserSecurityRoleInListItemAD
#PsSpSpRestApi_DeleteUserFromSecurityRoleInListItemAD
#PsSpSpRestApi_CreateOneFolderAD
#PsSpSpRestApi_ReadAllFoldersAD
#PsSpSpRestApi_RenameOneFolderAD
#PsSpSpRestApi_DeleteOneFolderAD
#PsSpSpRestApi_CreateOneAttachmentAD
#PsSpSpRestApi_ReadAllAttachmentsAD
#PsSpSpRestApi_DownloadOneAttachmentByFileNameAD
#PsSpSpRestApi_DeleteOneAttachmentByFileNameAD

Write-Host "Done" 
