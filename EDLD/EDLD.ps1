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

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsRestCreateOneListItem()
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
Function SpPsRestUploadOneDocument()
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
Function SpPsRestDownloadOneDocument()
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
Function SpPsRestReadAllListItems()
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
Function SpPsRestReadOneListItem()
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
Function SpPsRestReadAllLibraryDocs()
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
Function SpPsRestReadOneLibraryDoc()
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
Function SpPsRestUpdateOneListItem()
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
Function SpPsRestUpdateOneLibraryDoc()
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
Function SpPsRestDeleteOneListItem()
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
Function SpPsRestDeleteOneLibraryDoc()
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
Function SpPsRestBreakSecurityInheritanceListItem()
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
Function SpPsRestResetSecurityInheritanceListItem()
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
Function SpPsRestAddUserToSecurityRoleInListItem()
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
Function SpPsRestUpdateUserSecurityRoleInListItem()
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
Function SpPsRestDeleteUserFromSecurityRoleInListItem()
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
Function SpPsRestCreateOneFolder()
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
Function SpPsRestReadAllFolders()
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
Function SpPsRestRenameOneFolder()
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
Function SpPsRestDeleteOneFolder()
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
Function SpPsRestCreateOneAttachment()
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
Function SpPsRestReadAllAttachments()
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
Function SpPsRestDownloadOneAttachmentByFileName()
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
Function SpPsRestDeleteOneAttachmentByFileName()
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

## Running the Functions
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$webUrl = $configFile.appsettings.spUrl
$userName = $configFile.appsettings.spUserName
$password = $configFile.appsettings.spUserPw

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

Write-Host "Done" 
