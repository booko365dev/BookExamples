
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnPPowerShellWithAccPw($FullSiteUrl)
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.UserName, $securePW
		Connect-PnPOnline -Url $FullSiteUrl -Credentials $myCredentials
	}
}

Function LoginPsPnPPowerShellWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function LoginPsPnPPowerShellWithCertificate
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}

Function LoginPsPnPPowerShellWithCertificateBase64
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function SpPsPnpPowerShell_CreateOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
	Disconnect-PnPOnline
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsPnpPowerShell_UploadOneDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$fileFullPath = "C:\Temporary\TestDocument01.docx"
	Add-PnPFile -Path $fileFullPath -Folder "TestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsPnpPowerShell_UploadSeveralDocuments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$filesPath = "C:\Temporary\"
	$myFiles = Get-ChildItem -Path $filesPath -Force -Recurse
 
	ForEach ($oneFile in $myFiles) {
		Add-PnPFile -Path "$($oneFile.Directory)\$($oneFile.Name)" -Folder "TestLibrary" `
											-Values @{"Title" = $($oneFile.Name)}
	}
	Disconnect-PnPOnline
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsPnpPowerShell_DownloadOneDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPFile -Url  "/TestLibrary/TestDocument01.docx" `
				-Path "C:\Temporary\" `
				-FileName "TestDocument01_Downnload.docx" `
				-AsFile
	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsPnpPowerShell_ReadAllListItems
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPListItem -List "TestList"
	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsPnpPowerShell_ReadOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPListItem -List "TestList" -Id 4
	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 010
function SpPsPnpPowerShell_FindOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Find-PnPFile -List "TestLibrary" -Match *.docx
	Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsPnpPowerShell_CopyOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Copy-PnPFile -SourceUrl "TestLibrary/TestDocument01.docx" `
						-TargetUrl "OtherTestLibrary/TestDocument01.docx"
	Disconnect-PnPOnline
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsPnpPowerShell_MoveOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Move-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetUrl ($WebUrlRel + "/OtherTestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 012

#gavdcodebegin 007
function SpPsPnpPowerShell_UpdateOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListItem -List "TestList" -Identity 4 `
			-Values @{"Title" = "NewListItemPsPnpUpdated"}
	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 013
function SpPsPnpPowerShell_RenameOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Rename-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetFileName "TestDocument01_Renamed.docx"
	Disconnect-PnPOnline
}
#gavdcodeend 013

#gavdcodebegin 008
function SpPsPnpPowerShell_DeleteOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Remove-PnPListItem -List "TestList" -Identity "4" -Force -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsPnpPowerShell_DeleteToRecycleOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "5" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 009

#gavdcodebegin 014
function SpPsPnpPowerShell_DeleteOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Remove-PnPFile -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx") -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsPnpPowerShell_ResetVersionOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Reset-PnPFileVersion -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 015

#gavdcodebegin 016
function SpPsPnpPowerShell_CheckOutOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedOut -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsPnpPowerShell_CheckInOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedIn -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
			-CheckinType MinorCheckin -Comment "Changed by PowerShell"
	Disconnect-PnPOnline
}
#gavdcodeend 017

#gavdcodebegin 018
function SpPsPnpPowerShell_AddUserToSecurityRole
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 `
				-User $configFile.appsettings.UserName -AddRole 'Read'
	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsPnpPowerShell_RemoveUserFromSecurityRole
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 `
				-User $configFile.appsettings.UserName -RemoveRole 'Read'
	Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsPnpPowerShell_ResetSecurityInheritance
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 -InheritPermissions
	Disconnect-PnPOnline
}
#gavdcodeend 020

#gavdcodebegin 021
function SpPsPnpPowerShell_AddFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Add-PnPFolder -Name "PnPPowerShellFolder" -Folder "TestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 021

#gavdcodebegin 022
function SpPsPnpPowerShell_ResolveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Resolve-PnPFolder -SiteRelativePath "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 022

#gavdcodebegin 023
function SpPsPnpPowerShell_GetFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 023

#gavdcodebegin 024
function SpPsPnpPowerShell_GetFolderItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 024

#gavdcodebegin 025
function SpPsPnpPowerShell_RenameFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Rename-PnPFolder -Folder "TestLibrary/PnPPowerShellFolder" `
					 -TargetFolderName "PnPPowerShellFolderRenamed"
	Disconnect-PnPOnline
}
#gavdcodeend 025

#gavdcodebegin 026
function SpPsPnpPowerShell_MoveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Move-PnPFolder -Folder "TestLibrary/PnPPowerShellFolder" `
				   -TargetFolder "OtherTestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 026

#gavdcodebegin 027
function SpPsPnpPowerShell_RemoveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Remove-PnPFolder -Name "PnPPowerShellFolder" `
				     -Folder "TestLibrary" `
					 -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 027

#gavdcodebegin 028
function SpPsPnpPowerShell_AddRightsFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPFolderPermission -List "TestLibrary" `
							-Identity "TestLibrary\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-AddRole "Contribute"
	Disconnect-PnPOnline
}
#gavdcodeend 028

#gavdcodebegin 029
function SpPsPnpPowerShell_RemoveRightsFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPFolderPermission -List "TestLibrary" `
							-Identity "TestLibrary\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-RemoveRole "Contribute"
	Disconnect-PnPOnline
}
#gavdcodeend 029

#gavdcodebegin 030
function SpPsPnpPowerShell_ReadAllAttachments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$myListitem = Get-PnPListItem -List "TestList" -Id 3
	$myAttachments = Get-PnPProperty -ClientObject $myListitem -Property "AttachmentFiles"
	foreach ($oneAttachment in $myAttachments) {
		Write-Host "File Name - " $oneAttachment.ServerRelativeUrl
	}
	Disconnect-PnPOnline
}
#gavdcodeend 030

#gavdcodebegin 031
function SpPsPnpPowerShell_DownloadAllAttachments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
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
	Disconnect-PnPOnline
}
#gavdcodeend 031


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 31 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#SpPsPnpPowerShell_CreateOneListItem
#SpPsPnpPowerShell_UploadOneDocument
#SpPsPnpPowerShell_UploadSeveralDocuments
#SpPsPnpPowerShell_DownloadOneDocument
#SpPsPnpPowerShell_ReadAllListItems
#SpPsPnpPowerShell_ReadOneListItem
#SpPsPnpPowerShell_FindOneLibraryDocument
#SpPsPnpPowerShell_CopyOneLibraryDocument
#SpPsPnpPowerShell_MoveOneLibraryDocument
#SpPsPnpPowerShell_UpdateOneListItem
#SpPsPnpPowerShell_RenameOneLibraryDocument
#SpPsPnpPowerShell_DeleteOneListItem
#SpPsPnpPowerShell_DeleteToRecycleOneListItem
#SpPsPnpPowerShell_DeleteOneLibraryDoc
#SpPsPnpPowerShell_ResetVersionOneLibraryDoc
#SpPsPnpPowerShell_CheckOutOneLibraryDoc
#SpPsPnpPowerShell_CheckInOneLibraryDoc
#SpPsPnpPowerShell_AddUserToSecurityRole
#SpPsPnpPowerShell_RemoveUserFromSecurityRole
#SpPsPnpPowerShell_ResetSecurityInheritance
#SpPsPnpPowerShell_AddFolder
#SpPsPnpPowerShell_ResolveFolder
#SpPsPnpPowerShell_GetFolder
#SpPsPnpPowerShell_GetFolderItem
#SpPsPnpPowerShell_RenameFolder
#SpPsPnpPowerShell_MoveFolder
#SpPsPnpPowerShell_RemoveFolder
#SpPsPnpPowerShell_AddRightsFolder
#SpPsPnpPowerShell_RemoveRightsFolder
#SpPsPnpPowerShell_ReadAllAttachments
#SpPsPnpPowerShell_DownloadAllAttachments

Write-Host "Done" 
