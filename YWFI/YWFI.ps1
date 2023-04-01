Function LoginPsPnP
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function SpPsPnp_CreateOneListItem  #*** LEGACY CODE ***
{
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsPnp_UploadOneDocument  #*** LEGACY CODE ***
{
	$fileFullPath = "C:\Temporary\TestDocument01.docx"
	Add-PnPFile -Path $fileFullPath -Folder "TestLibrary"
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsPnp_UploadSeveralDocuments  #*** LEGACY CODE ***
{
	$filesPath = "C:\Temporary\"
	$myFiles = Get-ChildItem -Path $filesPath -Force -Recurse
 
	ForEach ($oneFile in $myFiles) {
		Add-PnPFile -Path "$($oneFile.Directory)\$($oneFile.Name)" -Folder "TestLibrary" `
											-Values @{"Title" = $($oneFile.Name)}
	}
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsPnp_DownloadOneDocument  #*** LEGACY CODE ***
{
	Get-PnPFile -Url  "/TestLibrary/TestDocument01.docx" `
				-Path "C:\Temporary\" `
				-FileName "TestDocument01_Dnld.docx" `
				-AsFile
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsPnp_ReadAllListItems  #*** LEGACY CODE ***
{
	Get-PnPListItem -List "TestList"
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsPnp_ReadOneListItem  #*** LEGACY CODE ***
{
	Get-PnPListItem -List "TestList" -Id 44
}
#gavdcodeend 006

#gavdcodebegin 010
Function SpPsPnp_FindOneLibraryDocument  #*** LEGACY CODE ***
{
	Find-PnPFile -List "TestLibrary" -Match *.docx
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpPsPnp_CopyOneLibraryDocument  #*** LEGACY CODE ***
{
	Copy-PnPFile -SourceUrl "TestLibrary/TestDocument01.docx" `
						-TargetUrl "OtherTestLibrary/TestDocument01.docx"
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpPsPnp_MoveOneLibraryDocument  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Move-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetUrl ($WebUrlRel + "/OtherTestLibrary/TestDocument01.docx")
}
#gavdcodeend 012

#gavdcodebegin 007
Function SpPsPnp_UpdateOneListItem  #*** LEGACY CODE ***
{
	Set-PnPListItem -List "TestList" -Identity 44 `
			-Values @{"Title" = "NewListItemPsPnpUpdated"}
}
#gavdcodeend 007

#gavdcodebegin 013
Function SpPsPnp_RenameOneLibraryDocument  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Rename-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetFileName "TestDocument01_Renamed.docx"
}
#gavdcodeend 013

#gavdcodebegin 008
Function SpPsPnp_DeleteOneListItem  #*** LEGACY CODE ***
{
	Remove-PnPListItem -List "TestList" -Identity "44" -Force -Recycle
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsPnp_DeleteToRecycleOneListItem  #*** LEGACY CODE ***
{
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "45" -Force
}
#gavdcodeend 009

#gavdcodebegin 014
Function SpPsPnp_DeleteOneLibraryDoc  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Remove-PnPFile -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx") -Recycle
}
#gavdcodeend 014

#gavdcodebegin 015
Function SpPsPnp_ResetVersionOneLibraryDoc  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Reset-PnPFileVersion -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}
#gavdcodeend 015

#gavdcodebegin 016
Function SpPsPnp_CheckOutOneLibraryDoc  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedOut -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx")
}
#gavdcodeend 016

#gavdcodebegin 017
Function SpPsPnp_CheckInOneLibraryDoc  #*** LEGACY CODE ***
{
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedIn -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
			-CheckinType MinorCheckin -Comment "Changed by PowerShell"
}
#gavdcodeend 017

#gavdcodebegin 018
Function SpPsPnp_AddUserToSecurityRole  #*** LEGACY CODE ***
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.UserName -AddRole 'Read'
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpPsPnp_RemoveUserFromSecurityRole  #*** LEGACY CODE ***
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 `
				-User $configFile.appsettings.UserName -RemoveRole 'Read'
}
#gavdcodeend 019

#gavdcodebegin 020
Function SpPsPnp_ResetSecurityInheritance  #*** LEGACY CODE ***
{
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 35 -InheritPermissions
}
#gavdcodeend 020

#gavdcodebegin 021
Function SpPsPnp_AddFolder  #*** LEGACY CODE ***
{
	Add-PnPFolder -Name "PnPPowerShellFolder" -Folder "TestDocuments"
}
#gavdcodeend 021

#gavdcodebegin 022
Function SpPsPnp_ResolveFolder  #*** LEGACY CODE ***
{
	Resolve-PnPFolder -SiteRelativePath "TestDocuments/PnPPowerShellFolderResolve"
}
#gavdcodeend 022

#gavdcodebegin 023
Function SpPsPnp_GetFolder  #*** LEGACY CODE ***
{
	Get-PnPFolder -Url "TestDocuments/PnPPowerShellFolder"
}
#gavdcodeend 023

#gavdcodebegin 024
Function SpPsPnp_GetFolderItem  #*** LEGACY CODE ***
{
	Get-PnPFolderItem -FolderSiteRelativeUrl  "TestDocuments/PnPPowerShellFolder"
}
#gavdcodeend 024

#gavdcodebegin 025
Function SpPsPnp_RenameFolder  #*** LEGACY CODE ***
{
	Rename-PnPFolder -Folder "TestDocuments/PnPPowerShellFolder" `
					 -TargetFolderName "PnPPowerShellFolderRenamed"
}
#gavdcodeend 025

#gavdcodebegin 026
Function SpPsPnp_MoveFolder  #*** LEGACY CODE ***
{
	Move-PnPFolder -Folder "TestDocuments/PnPPowerShellFolder" `
				   -TargetFolder "Shared Documents"
}
#gavdcodeend 026

#gavdcodebegin 027
Function SpPsPnp_RemoveFolder  #*** LEGACY CODE ***
{
	Remove-PnPFolder -Name "PnPPowerShellFolder" `
				     -Folder "TestDocuments" `
					 -Recycle
}
#gavdcodeend 027

#gavdcodebegin 028
Function SpPsPnp_AddRightsFolder  #*** LEGACY CODE ***
{
	Set-PnPFolderPermission -List "TestDocuments" `
							-Identity "TestDocuments\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-AddRole "Contribute"
}
#gavdcodeend 028

#gavdcodebegin 029
Function SpPsPnp_RemoveRightsFolder  #*** LEGACY CODE ***
{
	Set-PnPFolderPermission -List "TestDocuments" `
							-Identity "TestDocuments\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-RemoveRole "Contribute"
}
#gavdcodeend 029

#gavdcodebegin 030
Function SpPsPnp_ReadAllAttachments  #*** LEGACY CODE ***
{
	$myListitem = Get-PnPListItem -List "TestList" -Id 3
	$myAttachments = Get-PnPProperty -ClientObject $myListitem -Property "AttachmentFiles"
	foreach ($oneAttachment in $myAttachments) {
		Write-Host "File Name - " $oneAttachment.ServerRelativeUrl
	}
}
#gavdcodeend 030

#gavdcodebegin 031
Function SpPsPnp_DownloadAllAttachments  #*** LEGACY CODE ***
{
	$myListitem = Get-PnPListItem -List "TestList" -Id 3
	$myAttachments = Get-PnPProperty -ClientObject $myListitem -Property "AttachmentFiles"
	$myFilesPath = "C:\Temporary\"
	foreach ($oneAttachment in $myAttachments) {
		Write-Host "File Name - " $oneAttachment.FileName
		Get-PnPFile -Url $oneAttachment.ServerRelativeUrl `
					-FileName $oneAttachement.FileName `
					-Path $myFilesPath `
					-AsFile
	}
}
#gavdcodeend 031

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsPnP

#SpPsPnp_CreateOneListItem
#SpPsPnp_UploadOneDocument
#SpPsPnp_UploadSeveralDocuments
#SpPsPnp_DownloadOneDocument
#SpPsPnp_ReadAllListItems
#SpPsPnp_ReadOneListItem
#SpPsPnp_FindOneLibraryDocument
#SpPsPnp_CopyOneLibraryDocument
#SpPsPnp_MoveOneLibraryDocument
#SpPsPnp_UpdateOneListItem
#SpPsPnp_RenameOneLibraryDocument
#SpPsPnp_DeleteOneListItem
#SpPsPnp_DeleteToRecycleOneListItem
#SpPsPnp_DeleteOneLibraryDoc
#SpPsPnp_ResetVersionOneLibraryDoc
#SpPsPnp_CheckOutOneLibraryDoc
#SpPsPnp_CheckInOneLibraryDoc
#SpPsPnp_AddUserToSecurityRole
#SpPsPnp_RemoveUserFromSecurityRole
#SpPsPnp_ResetSecurityInheritance
#SpPsPnp_AddFolder
#SpPsPnp_ResolveFolder
#SpPsPnp_GetFolder
#SpPsPnp_GetFolderItem
#SpPsPnp_RenameFolder
#SpPsPnp_MoveFolder
#SpPsPnp_RemoveFolder
#SpPsPnp_AddRightsFolder
#SpPsPnp_RemoveRightsFolder
#SpPsPnp_ReadAllAttachments
#SpPsPnp_DownloadAllAttachments

Write-Host "Done"