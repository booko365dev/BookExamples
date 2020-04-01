Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsPnpCreateOneListItem()
{
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnpUploadOneDocument()
{
	$fileFullPath = "C:\Temporary\TestDocument01.docx"
	Add-PnPFile -Path $fileFullPath -Folder "TestLibrary"
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnpUploadSeveralDocuments()
{
	$filesPath = "C:\Temporary\"
	$myFiles = Get-ChildItem -Path $filesPath -Force -Recurse
 
	ForEach ($oneFile in $myFiles) {
		Add-PnPFile -Path "$($oneFile.Directory)\$($oneFile.Name)" -Folder "TestLibrary" `
											-Values @{"Title" = $($oneFile.Name)}
	}
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnpDownloadOneDocument()
{
	Get-PnPFile -Url  "/TestLibrary/TestDocument01.docx" `
				-Path "C:\Temporary\" `
				-FileName "TestDocument01_Dnld.docx" `
				-AsFile
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnpReadAllListItems()
{
	Get-PnPListItem -List "TestList"
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnpReadOneListItem()
{
	Get-PnPListItem -List "TestList" -Id 44
}
#gavdcodeend 06

#gavdcodebegin 10
Function SpPsPnpFindOneLibraryDocument()
{
	Find-PnPFile -List "TestLibrary" -Match *.docx
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsPnpCopyOneLibraryDocument()
{
	Copy-PnPFile -SourceUrl "TestLibrary/TestDocument01.docx" `
						-TargetUrl "OtherTestLibrary/TestDocument01.docx"
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsPnpMoveOneLibraryDocument()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Move-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetUrl ($WebUrlRel + "/OtherTestLibrary/TestDocument01.docx")
}
#gavdcodeend 12

#gavdcodebegin 07
Function SpPsPnpUpdateOneListItem()
{
	Set-PnPListItem -List "TestList" -Identity 44 `
			-Values @{"Title" = "NewListItemPsPnpUpdated"}
}
#gavdcodeend 07

#gavdcodebegin 13
Function SpPsPnpRenameOneLibraryDocument()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Rename-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetFileName "TestDocument01_Renamed.docx"
}
#gavdcodeend 13

#gavdcodebegin 08
Function SpPsPnpDeleteOneListItem()
{
	Remove-PnPListItem -List "TestList" -Identity "44" -Force -Recycle
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnpDeleteToRecycleOneListItem()
{
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "45" -Force
}
#gavdcodeend 09

#gavdcodebegin 14
Function SpPsPnpDeleteOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Remove-PnPFile -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx") -Recycle
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsPnpResetVersionOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Reset-PnPFileVersion -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsPnpCheckOutOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedOut -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsPnpCheckInOneLibraryDoc()
{
	$webUrl = $configFile.appsettings.spUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedIn -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
			-CheckinType MinorCheckin -Comment "Changed by PowerShell"
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsPnpAddUserToSecurityRole()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.spUserName -AddRole 'Read'
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsPnpRemoveUserFromSecurityRole()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.spUserName -RemoveRole 'Read'
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsPnpResetSecurityInheritance()
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 -InheritPermissions
}
#gavdcodeend 20

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