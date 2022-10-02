Function Invoke-RestSPO() {  #*** LEGACY CODE ***
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
 
Function Get-SPOContextInfo(){  #*** LEGACY CODE ***
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
	  #*** LEGACY CODE ***
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

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

# Using Basic Authentication, this is Legacy code, and cannot be used anymore

#gavdcodebegin 01
Function SpPsRestCreateOneListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 01 

#gavdcodebegin 02
Function SpPsRestUploadOneDocument()  #*** LEGACY CODE ***
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
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsRestDownloadOneDocument()  #*** LEGACY CODE ***
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
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsRestReadAllListItems()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsRestReadOneListItem()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestList')/items(26)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsRestReadAllLibraryDocs()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsRestReadOneLibraryDoc()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/web/lists/getbytitle('TestLibrary')/items(25)?" + 
																	"$select=Title,Id"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsRestUpdateOneListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsRestUpdateOneLibraryDoc()  #*** LEGACY CODE ***
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
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsRestDeleteOneListItem()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/items(26)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsRestDeleteOneLibraryDoc()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestLibrary')/items(25)"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue `
		-ETag "*" -XHTTPMethod "DELETE"

	$data | ConvertTo-Json
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsRestBreakSecurityInheritanceListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsRestResetSecurityInheritanceListItem()  #*** LEGACY CODE ***
{
	$endpointUrl = $WebUrl + "/_api/lists/getbytitle('TestList')/" +
						"items(27)/resetroleinheritance"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
		$password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue

	$data | ConvertTo-Json
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsRestAddUserToSecurityRoleInListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsRestUpdateUserSecurityRoleInListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsRestDeleteUserFromSecurityRoleInListItem()  #*** LEGACY CODE ***
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
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsRestCreateOneFolder()  #*** LEGACY CODE ***
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
#gavdcodeend 17 

#gavdcodebegin 18
Function SpPsRestReadAllFolders()  #*** LEGACY CODE ***
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
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsRestRenameOneFolder()  #*** LEGACY CODE ***
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
#gavdcodeend 19 

#gavdcodebegin 20
Function SpPsRestDeleteOneFolder()  #*** LEGACY CODE ***
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
#gavdcodeend 20 

#gavdcodebegin 21
Function SpPsRestCreateOneAttachment()  #*** LEGACY CODE ***
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
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsRestReadAllAttachments()  #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/lists/GetByTitle('TestList')" + `
                                                        "/items(3)/AttachmentFiles"
	$contextInfo = Get-SPOContextInfo -WebUrl $WebUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
		  $password -RequestDigest $contextInfo.GetContextWebInformation.FormDigestValue 

	$data | ConvertTo-Json
}
#gavdcodeend 22

#gavdcodebegin 23
Function SpPsRestDownloadOneAttachmentByFileName()  #*** LEGACY CODE ***
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
#gavdcodeend 23

#gavdcodebegin 24
Function SpPsRestDeleteOneAttachmentByFileName()  #*** LEGACY CODE ***
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
#gavdcodeend 24

#----------------------------------------------------------------------------------------

# Using Azure AD Authentication through Connect-PnPOnline and an Account/PW App Registration

#gavdcodebegin 101
Function SpPsRestCreateOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestUploadOneDocumentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDownloadOneDocumentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadAllListItemsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadAllLibraryDocsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestUpdateOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestUpdateOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDeleteOneListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDeleteOneLibraryDocAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestBreakSecurityInheritanceListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestResetSecurityInheritanceListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestAddUserToSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestUpdateUserSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDeleteUserFromSecurityRoleInListItemAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestCreateOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadAllFoldersAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestRenameOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDeleteOneFolderAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestCreateOneAttachmentAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestReadAllAttachmentsAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDownloadOneAttachmentByFileNameAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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
Function SpPsRestDeleteOneAttachmentByFileNameAD
{
	LoginPsPnPPowerShellWithAccPwDefault
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

#----------------------------------------------------------------------------------------

## Running the Functions
[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

## Using Legacy Classic Authentication
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRestCreateOneListItem
#SpPsRestUploadOneDocument
#SpPsRestDownloadOneDocument
#SpPsRestReadAllListItems
#SpPsRestReadOneListItem
#SpPsRestReadAllLibraryDocs
#SpPsRestReadOneLibraryDoc
#SpPsRestUpdateOneListItem
#SpPsRestUpdateOneLibraryDoc
#SpPsRestDeleteOneListItem
#SpPsRestDeleteOneLibraryDoc
#SpPsRestBreakSecurityInheritanceListItem
#SpPsRestResetSecurityInheritanceListItem
#SpPsRestAddUserToSecurityRoleInListItem
#SpPsRestUpdateUserSecurityRoleInListItem
#SpPsRestDeleteUserFromSecurityRoleInListItem
#SpPsRestCreateOneFolder
#SpPsRestReadAllFolders
#SpPsRestRenameOneFolder
#SpPsRestDeleteOneFolder
#SpPsRestCreateOneAttachment
#SpPsRestReadAllAttachments
#SpPsRestDownloadOneAttachmentByFileName
#SpPsRestDeleteOneAttachmentByFileName

## Using Azure AD Authentication
#SpPsRestCreateOneListItemAD
#SpPsRestUploadOneDocumentAD
#SpPsRestDownloadOneDocumentAD
#SpPsRestReadAllListItemsAD
#SpPsRestReadOneListItemAD
#SpPsRestReadAllLibraryDocsAD
#SpPsRestReadOneLibraryDocAD
#SpPsRestUpdateOneListItemAD
#SpPsRestUpdateOneLibraryDocAD
#SpPsRestDeleteOneListItemAD
#SpPsRestDeleteOneLibraryDocAD
#SpPsRestBreakSecurityInheritanceListItemAD
#SpPsRestResetSecurityInheritanceListItemAD
#SpPsRestAddUserToSecurityRoleInListItemAD
#SpPsRestUpdateUserSecurityRoleInListItemAD
#SpPsRestDeleteUserFromSecurityRoleInListItemAD
#SpPsRestCreateOneFolderAD
#SpPsRestReadAllFoldersAD
#SpPsRestRenameOneFolderAD
#SpPsRestDeleteOneFolderAD
#SpPsRestCreateOneAttachmentAD
#SpPsRestReadAllAttachmentsAD
#SpPsRestDownloadOneAttachmentByFileNameAD
#SpPsRestDeleteOneAttachmentByFileNameAD

Write-Host "Done" 
