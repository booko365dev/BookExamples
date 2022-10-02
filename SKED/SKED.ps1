
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

#gavdcodebegin 01
function SpPsPnpPowerShell_CreateOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Add-PnPListItem -List "TestList" -Values @{"Title" = "NewListItemPsPnp"}
	Disconnect-PnPOnline
}
#gavdcodeend 01

#gavdcodebegin 02
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
#gavdcodeend 02

#gavdcodebegin 03
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
#gavdcodeend 03

#gavdcodebegin 04
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
#gavdcodeend 04

#gavdcodebegin 05
function SpPsPnpPowerShell_ReadAllListItems
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPListItem -List "TestList"
	Disconnect-PnPOnline
}
#gavdcodeend 05

#gavdcodebegin 06
function SpPsPnpPowerShell_ReadOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPListItem -List "TestList" -Id 4
	Disconnect-PnPOnline
}
#gavdcodeend 06

#gavdcodebegin 10
function SpPsPnpPowerShell_FindOneLibraryDocument
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Find-PnPFile -List "TestLibrary" -Match *.docx
	Disconnect-PnPOnline
}
#gavdcodeend 10

#gavdcodebegin 11
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
#gavdcodeend 11

#gavdcodebegin 12
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
#gavdcodeend 12

#gavdcodebegin 07
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
#gavdcodeend 07

#gavdcodebegin 13
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
#gavdcodeend 13

#gavdcodebegin 08
function SpPsPnpPowerShell_DeleteOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Remove-PnPListItem -List "TestList" -Identity "4" -Force -Recycle
	Disconnect-PnPOnline
}
#gavdcodeend 08

#gavdcodebegin 09
function SpPsPnpPowerShell_DeleteToRecycleOneListItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Move-PnPListItemToRecycleBin -List "TestList" -Identity "5" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 09

#gavdcodebegin 14
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
#gavdcodeend 14

#gavdcodebegin 15
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
#gavdcodeend 15

#gavdcodebegin 16
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
#gavdcodeend 16

#gavdcodebegin 17
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
#gavdcodeend 17

#gavdcodebegin 18
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
#gavdcodeend 18

#gavdcodebegin 19
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
#gavdcodeend 19

#gavdcodebegin 20
function SpPsPnpPowerShell_ResetSecurityInheritance
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListItemPermission -List 'TestLibrary' -Identity 4 -InheritPermissions
	Disconnect-PnPOnline
}
#gavdcodeend 20

#gavdcodebegin 21
function SpPsPnpPowerShell_AddFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Add-PnPFolder -Name "PnPPowerShellFolder" -Folder "TestLibrary"
	Disconnect-PnPOnline
}
#gavdcodeend 21

#gavdcodebegin 22
function SpPsPnpPowerShell_ResolveFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Resolve-PnPFolder -SiteRelativePath "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 22

#gavdcodebegin 23
function SpPsPnpPowerShell_GetFolder
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 23

#gavdcodebegin 24
function SpPsPnpPowerShell_GetFolderItem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPFolder -Url "TestLibrary/PnPPowerShellFolder"
	Disconnect-PnPOnline
}
#gavdcodeend 24

#gavdcodebegin 25
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
#gavdcodeend 25

#gavdcodebegin 26
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
#gavdcodeend 26

#gavdcodebegin 27
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
#gavdcodeend 27

#gavdcodebegin 28
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
#gavdcodeend 28

#gavdcodebegin 29
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
#gavdcodeend 29

#gavdcodebegin 30
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
#gavdcodeend 30

#gavdcodebegin 31
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
#gavdcodeend 31


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

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
