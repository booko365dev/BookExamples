##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphCli_LoginWithSecret
{
	$env:AZURE_TENANT_ID = "ade56059-89c0-4594-90c3-e4772a8168ca"
	$env:AZURE_CLIENT_ID = "3d16c7bc-a14e-454e-81cd-da571cf2a8e3"
	$env:AZURE_CLIENT_SECRET = "IIX8Q~M5q-FCotlzO7GA4bLAQBPtF9dHwr.CcbsN"
	
	mgc login --strategy Environment
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsGraphCli_GetAllSpEmbeddedContainers
{
	PsGraphCli_LoginWithSecret

	$myContainerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4"

	mgc storage file-storage containers list `
						--filter "containerTypeId eq $($myContainerTypeId)"

	mgc logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsGraphCli_GetOneSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9Gonxt4GLB1...qiYjT8Qtn88"

	mgc storage file-storage containers get `
						--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsGraphCli_GetDriveSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9Gonxt4GLB1...qiYjT8Qtn88"

	mgc storage file-storage containers drive get `
						--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsGraphCli_ActivateSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!A7pQX90BaEWOOSirXHOIrs_2HzkZmN9Gonxt4GLB1...qiYjT8Qtn88"

	mgc storage file-storage containers activate post `
						--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsGraphCli_CreateSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerTypeId = "86d347fa-234c-4e53-973e-a1d0100d6fe4"

	$myBody = '{
				  "displayName": "Test Storage Container CLI", 
				  "description": "It is a CLI Test Storage Container", 
				  "containerTypeId": "' + $myContainerTypeId + '" 
			   }'

	mgc storage file-storage containers create --body $myBody

	mgc logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsGraphCli_UpdateSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	$myBody = '{
				  "displayName": "Test Storage Container CLI Updated", 
				  "description": "It is a CLI Test Storage Container Updated" 
			   }'

	mgc storage file-storage containers patch `
					--file-storage-container-id $myContainerId `
					--body $myBody

	mgc logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsGraphCli_DeleteSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	mgc storage file-storage containers delete `
					--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsGraphCli_DeletePermanentlySpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	mgc storage file-storage containers permanent-delete `
					--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsGraphCli_GetAllPermissionsSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	mgc storage file-storage containers permissions list `
					--file-storage-container-id $myContainerId

	mgc logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsGraphCli_AddPermissionSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	$myBody = '{
				  "roles": ["reader"],
				  "grantedToV2": {
					"user": {
					  "userPrincipalName": "user@domain.onmicrosoft.com"
					}
				  } 
			   }'

	mgc storage file-storage containers permissions create `
					--file-storage-container-id $myContainerId `
					--body $myBody

	mgc logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsGraphCli_UpdatePermissionSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	$myPermissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZWxldkBndWl0YWNhZG...zb2Z0LmNvbQ"

	$myBody = '{
				  "roles": ["owner"]
			   }'

	mgc storage file-storage containers permissions patch `
					--file-storage-container-id $myContainerId `
					--permission-id $myPermissionId `
					--body $myBody

	mgc logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsGraphCli_DeletePermissionSpEmbeddedContainer
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	$myPermissionId = "X2k6MCMuZnxtZW1iZXJzaGlwfGFkZWxldkBndWl0YWNhZG...zb2Z0LmNvbQ"

	mgc storage file-storage containers permissions delete `
					--file-storage-container-id $myContainerId `
					--permission-id $myPermissionId

	mgc logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsGraphCli_UploadOneFileToLibrary   # +++> TODO
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"

	$filePath = "C:\Temporary\TestWordFile.docx"
	$fileName = "TestWordFile.docx"

	$binaryStream = [System.IO.File]::ReadAllBytes($filePath)
	
	mgc drives items content put --drive-id $myContainerId `
								 --input-file $binaryStream `
								 --name $fileName

	mgc logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsGraphCli_GetAllFilesInFolderLibrary
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	
	$myRoot = mgc drives root get --drive-id $myContainerId | ConvertFrom-Json

	mgc drives items children list --drive-id $myContainerId `
								   --drive-item-id $myRoot.id

	mgc logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsGraphCli_DownloadFileFromSpEmbeddedContainer   # +++> TODO
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	$myFileId = "017IUVQBN6Y2GOVW7725BZO354PWSELRRZ"

    mgc drives items content get --drive-id $myContainerId `
                                 --drive-item-id $myFileId `
								 --output-file "C:\Temporary\MyDownloadedFile.docx"

	mgc logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsGraphCli_GetAllMetadataItemSpEmbeddedContainer   # +++> TODO
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	$myFileId = "017IUVQBL7CYLYXDN2HBE3QE5ELS74DS7N"

    mgc drives items get --drive-id $myDriveId `
						 --drive-item-id $myFileId

	mgc logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsGraphCli_DeleteFileFromSpEmbeddedContainer   # +++> TODO
{
	PsGraphCli_LoginWithSecret

	$myContainerId = "b!FAN0qpPlN0OAsMD7pGXdsc_2HzkZmN9Gonxt4...6RqiYjT8Qtn88"
	$myFileId = "017IUVQBL7CYLYXDN2HBE3QE5ELS74DS7N"

    mgc drives items delete --drive-id $myDriveId `
						    --drive-item-id $myFileId

	mgc logout
}
#gavdcodeend 017


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 017 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#		ATTENTION: There is a Windows Environment Variable already configured in the computer
#					to redirect the commands to the mgc.exe directory (see instructions in the book)
#PsGraphCli_GetAllSpEmbeddedContainers
#PsGraphCli_GetOneSpEmbeddedContainer
#PsGraphCli_GetDriveSpEmbeddedContainer
#PsGraphCli_ActivateSpEmbeddedContainer
#PsGraphCli_CreateSpEmbeddedContainer
#PsGraphCli_UpdateSpEmbeddedContainer
#PsGraphCli_DeleteSpEmbeddedContainer
#PsGraphCli_DeletePermanentlySpEmbeddedContainer
#PsGraphCli_GetAllPermissionsSpEmbeddedContainer
#PsGraphCli_AddPermissionSpEmbeddedContainer
#PsGraphCli_UpdatePermissionSpEmbeddedContainer
#PsGraphCli_DeletePermissionSpEmbeddedContainer
#PsGraphCli_UploadOneFileToLibrary
#PsGraphCli_GetAllFilesInFolderLibrary
#PsGraphCli_DownloadFileFromSpEmbeddedContainer
#PsGraphCli_GetAllMetadataItemSpEmbeddedContainer
#PsGraphCli_DeleteFileFromSpEmbeddedContainer

Write-Host "Done" 
