Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

Function SpPsPnpCreateOneListItem()
{
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
}

Function SpPsPnpUploadOneDocument()
{
	$fileFullPath = "C:\Temporary\TestDocument01.docx"
	Add-PnPFile -Path $fileFullPath -Folder "TestLibrary"
}

Function SpPsPnpUploadSeveralDocuments()
{
	$filesPath = "C:\Temporary\"
	$myFiles = Get-ChildItem -Path $filesPath -Force -Recurse
 
	ForEach ($oneFile in $myFiles) {
		Add-PnPFile -Path "$($oneFile.Directory)\$($oneFile.Name)" -Folder "TestLibrary" `
											-Values @{"Title" = $($oneFile.Name)}
	}
}

Function SpPsPnpDownloadOneDocument()
{
	Get-PnPFile -Url  "/TestLibrary/TestDocument01.docx" `
				-Path "C:\Temporary\" `
				-FileName "TestDocument01_Dnld.docx" `
				-AsFile
}

Function SpPsPnpReadAllListItems()
{
	Get-PnPListItem -List "TestList"
}

Function SpPsPnpReadOneListItem()
{
	Get-PnPListItem -List "TestList" -Id 44
}

Function SpPsPnpFindOneLibraryDocument()
{
	Find-PnPFile -List "TestLibrary" -Match *.docx
}

Function SpPsPnpCopyOneLibraryDocument()
{
	Copy-PnPFile -SourceUrl "TestLibrary/TestDocument01.docx" `
						-TargetUrl "OtherTestLibrary/TestDocument01.docx"
}

Function SpPsPnpMoveOneLibraryDocument()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Move-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetUrl ($WebUrlRel + "/OtherTestLibrary/TestDocument01.docx")
}

Function SpPsPnpUpdateOneListItem()
{
	Set-PnPListItem -List "TestList" -Identity 44 `
			-Values @{"Title" = "NewListItemPsPnpUpdated"}
}

Function SpPsPnpRenameOneLibraryDocument()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Rename-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetFileName "TestDocument01_Renamed.docx"
}

Function SpPsPnpDeleteOneListItem()
{
	Remove-PnPListItem -List "TestList" -Identity "44" -Force -Recycle
}

Function SpPsPnpDeleteToRecycleOneListItem()
{
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "45" -Force
}

Function SpPsPnpDeleteOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Remove-PnPFile -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx") -Recycle
}

Function SpPsPnpResetVersionOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Reset-PnPFileVersion -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}

Function SpPsPnpCheckOutOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedOut -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}

Function SpPsPnpCheckInOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedIn -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
			-CheckinType MinorCheckin -Comment "Changed by PowerShell"
}

Function SpPsPnpAddUserToSecurityRole()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.spUserName -AddRole 'Read'
}

Function SpPsPnpRemoveUserFromSecurityRole()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.spUserName -RemoveRole 'Read'
}

Function SpPsPnpResetSecurityInheritance()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 -InheritPermissions
}

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsPnP

#SpPsPnpCreateOneListItem
#SpPsPnpUploadOneDocument
#SpPsPnpUploadSeveralDocuments
#SpPsPnpDownloadOneDocument
#SpPsPnpReadAllListItems
#SpPsPnpReadOneListItem
#SpPsPnpFindOneLibraryDocument
#SpPsPnpCopyOneLibraryDocument
#SpPsPnpMoveOneLibraryDocument
#SpPsPnpUpdateOneListItem
#SpPsPnpRenameOneLibraryDocument
#SpPsPnpDeleteOneListItem
#SpPsPnpDeleteToRecycleOneListItem
#SpPsPnpDeleteOneLibraryDoc
#SpPsPnpResetVersionOneLibraryDoc
#SpPsPnpCheckOutOneLibraryDoc
#SpPsPnpCheckInOneLibraryDoc
#SpPsPnpAddUserToSecurityRole
#SpPsPnpRemoveUserFromSecurityRole
#SpPsPnpResetSecurityInheritance

Write-Host "Done"
