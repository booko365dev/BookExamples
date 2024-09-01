
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpPnpPowerShell_LoginWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function PsSpPnpPowerShell_LoginWithAccPw($FullSiteUrl)
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

Function PsSpPnpPowerShell_LoginWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function PsSpPnpPowerShell_LoginWithCertificate
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

Function PsSpPnpPowerShell_LoginWithCertificateBase64
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
function PsSpPnpPowerShell_CreateOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
	Disconnect-PnPOnline
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpPnpPowerShell_UploadOneDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$fileFullPath = "C:\Temporary\TestDocument01.docx"
	Add-PnPFile -Path $fileFullPath -Folder "TestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpPnpPowerShell_UploadSeveralDocuments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
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
function PsSpPnpPowerShell_DownloadOneDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPFile -Url  "/TestLibrary/TestDocument01.docx" `
				-Path "C:\Temporary\" `
				-FileName "TestDocument01_Downnload.docx" `
				-AsFile
	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpPnpPowerShell_ReadAllListItems
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPListItem -List "TestList"
	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpPnpPowerShell_ReadOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPListItem -List "TestList" -Id 4
	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 010
function PsSpPnpPowerShell_FindOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Find-PnPFile -List "TestLibrary" -Match *.docx
	Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpPnpPowerShell_CopyOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Copy-PnPFile -SourceUrl "TestLibrary/TestDocument01.docx" `
						-TargetUrl "OtherTestLibrary/TestDocument01.docx"
	Disconnect-PnPOnline
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpPnpPowerShell_MoveOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Move-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetUrl ($WebUrlRel + "/OtherTestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 012

#gavdcodebegin 007
function PsSpPnpPowerShell_UpdateOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListItem -List "TestList" -Identity 4 `
			-Values @{"Title" = "NewListItemPsPnpUpdated"}
	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 013
function PsSpPnpPowerShell_RenameOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Rename-PnPFile -ServerRelativeUrl ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
				-TargetFileName "TestDocument01_Renamed.docx"
	Disconnect-PnPOnline
}
#gavdcodeend 013

#gavdcodebegin 008
function PsSpPnpPowerShell_DeleteOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Remove-PnPListItem -List "TestList" -Identity "4" -Force -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpPnpPowerShell_DeleteToRecycleOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "5" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 009

#gavdcodebegin 014
function PsSpPnpPowerShell_DeleteOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Remove-PnPFile -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx") -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpPnpPowerShell_ResetVersionOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Reset-PnPFileVersion -ServerRelativeUrl `
							($WebUrlRel + "/TestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpPnpPowerShell_CheckOutOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedOut -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx")
	Disconnect-PnPOnline
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpPnpPowerShell_CheckInOneLibraryDoc
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$webUrl = $configFile.appsettings.SiteCollUrl
	$WebUri = [System.Uri]$WebUrl
	$WebUrlRel = $WebUri.AbsolutePath
	
	Set-PnPFileCheckedIn -Url ($WebUrlRel + "/TestLibrary/TestDocument01.docx") `
			-CheckinType MinorCheckin -Comment "Changed by PowerShell"
	Disconnect-PnPOnline
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpPnpPowerShell_AddUserToSecurityRole
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 `
				-User $configFile.appsettings.UserName -AddRole 'Read'
	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpPnpPowerShell_RemoveUserFromSecurityRole
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 `
				-User $configFile.appsettings.UserName -RemoveRole 'Read'
	Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpPnpPowerShell_ResetSecurityInheritance
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 -InheritPermissions
	Disconnect-PnPOnline
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpPnpPowerShell_AddFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Add-PnPFolder -Name "PnPPowerShellFolder" -Folder "TestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpPnpPowerShell_ResolveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Resolve-PnPFolder -SiteRelativePath "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 022

#gavdcodebegin 023
function PsSpPnpPowerShell_GetFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 023

#gavdcodebegin 024
function PsSpPnpPowerShell_GetFolderItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 024

#gavdcodebegin 025
function PsSpPnpPowerShell_RenameFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Rename-PnPFolder -Folder "TestLibrary/PnPPowerShellFolder" `
					 -TargetFolderName "PnPPowerShellFolderRenamed"
	Disconnect-PnPOnline
}
#gavdcodeend 025

#gavdcodebegin 026
function PsSpPnpPowerShell_MoveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Move-PnPFolder -Folder "TestLibrary/PnPPowerShellFolder" `
				   -TargetFolder "OtherTestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 026

#gavdcodebegin 027
function PsSpPnpPowerShell_RemoveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Remove-PnPFolder -Name "PnPPowerShellFolder" `
				     -Folder "TestLibrary" `
					 -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 027

#gavdcodebegin 028
function PsSpPnpPowerShell_AddRightsFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPFolderPermission -List "TestLibrary" `
							-Identity "TestLibrary\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-AddRole "Contribute"
	Disconnect-PnPOnline
}
#gavdcodeend 028

#gavdcodebegin 029
function PsSpPnpPowerShell_RemoveRightsFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPFolderPermission -List "TestLibrary" `
							-Identity "TestLibrary\PnPPowerShellFolder" `
							-User "user@domain.OnMicrosoft.com" `
							-RemoveRole "Contribute"
	Disconnect-PnPOnline
}
#gavdcodeend 029

#gavdcodebegin 030
function PsSpPnpPowerShell_ReadAllAttachments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$myListitem = Get-PnPListItem -List "TestList" -Id 3
	$myAttachments = Get-PnPProperty -ClientObject $myListitem -Property "AttachmentFiles"
	foreach ($oneAttachment in $myAttachments) {
		Write-Host "File Name - " $oneAttachment.ServerRelativeUrl
	}
	Disconnect-PnPOnline
}
#gavdcodeend 030

#gavdcodebegin 031
function PsSpPnpPowerShell_DownloadAllAttachments
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
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

#PsSpPnpPowerShell_CreateOneListItem
#PsSpPnpPowerShell_UploadOneDocument
#PsSpPnpPowerShell_UploadSeveralDocuments
#PsSpPnpPowerShell_DownloadOneDocument
#PsSpPnpPowerShell_ReadAllListItems
#PsSpPnpPowerShell_ReadOneListItem
#PsSpPnpPowerShell_FindOneLibraryDocument
#PsSpPnpPowerShell_CopyOneLibraryDocument
#PsSpPnpPowerShell_MoveOneLibraryDocument
#PsSpPnpPowerShell_UpdateOneListItem
#PsSpPnpPowerShell_RenameOneLibraryDocument
#PsSpPnpPowerShell_DeleteOneListItem
#PsSpPnpPowerShell_DeleteToRecycleOneListItem
#PsSpPnpPowerShell_DeleteOneLibraryDoc
#PsSpPnpPowerShell_ResetVersionOneLibraryDoc
#PsSpPnpPowerShell_CheckOutOneLibraryDoc
#PsSpPnpPowerShell_CheckInOneLibraryDoc
#PsSpPnpPowerShell_AddUserToSecurityRole
#PsSpPnpPowerShell_RemoveUserFromSecurityRole
#PsSpPnpPowerShell_ResetSecurityInheritance
#PsSpPnpPowerShell_AddFolder
#PsSpPnpPowerShell_ResolveFolder
#PsSpPnpPowerShell_GetFolder
#PsSpPnpPowerShell_GetFolderItem
#PsSpPnpPowerShell_RenameFolder
#PsSpPnpPowerShell_MoveFolder
#PsSpPnpPowerShell_RemoveFolder
#PsSpPnpPowerShell_AddRightsFolder
#PsSpPnpPowerShell_RemoveRightsFolder
#PsSpPnpPowerShell_ReadAllAttachments
#PsSpPnpPowerShell_DownloadAllAttachments

Write-Host "Done" 
