
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpGraphCli_LoginWithCertificate
{
	mgc login --tenant-id $configFile.appsettings.TenantName `
			  --client-id $configFile.appsettings.ClientIdWithCert `
			  --certificate-thumb-print $configFile.appsettings.CertificateThumbprint `
			  --strategy ClientCertificate
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpGraphCli_GetAllItemsInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"

	mgc sites lists items list --site-id $mySiteId `
							   --list-id $myListId

	mgc logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpGraphCli_GetOneListItemInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myItemId = "11"

	mgc sites lists items get --site-id $mySiteId `
							  --list-id $myListId `
							  --list-item-id $myItemId

	mgc logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpGraphCli_CreateOneListItemInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myBody = @{	fields = @{
						Title = "PsSpGraphCli_Item"
					}
				} | ConvertTo-Json -Depth 5

	mgc sites lists items create --site-id $mySiteId `
								 --list-id $myListId `
								 --body $myBody

	mgc logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpGraphCli_UpdateOneListItemInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myItemId = "12"
	$myBody = @{	Title = "List Item updated"
				} | ConvertTo-Json -Depth 5

	mgc sites lists items patch --site-id $mySiteId `
								--list-id $myListId `
								--list-item-id $myItemId `
								--body $myBody

	mgc logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpGraphCli_DeleteOneListItemFromList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myItemId = "12"

	mgc sites lists items delete --site-id $mySiteId `
								 --list-id $myListId `
								 --list-item-id $myItemId

	mgc logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpGraphCli_GetDrivesInSite
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"

	$allDrives = mgc sites drives list --site-id $mySiteId | ConvertFrom-Json

	foreach ($oneDrive in $allDrives.value)
	{
		Write-Host "Drive name: " $oneDrive.name " - Drive identifier:" $oneDrive.id
	}

	mgc logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpGraphCli_GetAllFilesInLibrary
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"

    $myLibraryRootId = mgc drives root get --drive-id $myDriveId | ConvertFrom-Json
	write-host "Root ID of Library - " $myLibraryRootId.id
    $myDocs = mgc drives items children list --drive-id $myDriveId `
											 --drive-item-id $myLibraryRootId.id `
											  | ConvertFrom-Json

	foreach ($oneDoc in $myDocs.value)
	{
		Write-Host "Doc name: " $oneDoc.name " - Doc identifier:" $oneDoc.id
	}

	mgc logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpGraphCli_GetAllFilesInFolderLibrary
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFolderId = "01IAJF3RFCB4MZLKVESFH2OPTWWU7NJHOX"

    $myDocs = mgc drives items children list --drive-id $myDriveId `
											 --drive-item-id $myFolderId `
											  | ConvertFrom-Json

	foreach ($oneDoc in $myDocs.value)
	{
		Write-Host "Doc name: " $oneDoc.name " - Doc identifier:" $oneDoc.id
	}

	mgc logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpGraphCli_GetOneFileMetadata
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    $fileMetadata = mgc drives items get --drive-id $myDriveId `
										 --drive-item-id $myFileId
									  	  | ConvertFrom-Json

    Write-Host "Name - $($fileMetadata.Name)"
    Write-Host "ID - $($fileMetadata.Id)"
    Write-Host "Size - $($fileMetadata.Size) bytes"
    Write-Host "Created DateTime - $($fileMetadata.CreatedDateTime)"
    Write-Host "Last Modified DateTime - $($fileMetadata.LastModifiedDateTime)"

	mgc logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpGraphCli_UploadFileToLibrary   # +++> TODO
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    $uploadResult = mgc drives items upload --drive-id $myDriveId `
                                            --item-path $uploadPath `
                                            --file $localFilePath | ConvertFrom-Json

    if ($uploadResult) {
        Write-Host "File uploaded successfully. Item ID: $($uploadResult.id)"
    } else {
        Write-Host "Failed to upload the file."
    }

    mgc logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpGraphCli_DownloadFileFromLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    mgc drives items content get --drive-id $myDriveId `
                                 --drive-item-id $myFileId `
								 --output-file "C:\Temporary\MyDownloadedFile.docx"

    mgc logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpGraphCli_UpdateOneFileMetadata
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"
	$myBody = @{	name = "List Item updated"
				} | ConvertTo-Json -Depth 5

    mgc drives items patch --drive-id $myDriveId `
                           --drive-item-id $myFileId `
						   --body $myBody

    mgc logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpGraphCli_DeleteOneFile
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    mgc drives items delete --drive-id $myDriveId `
                            --drive-item-id $myFileId

    mgc logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpGraphCli_CopyFile
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"
	$myNewDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm"

    $myNewLibraryRootId = mgc drives root get --drive-id $myNewDriveId `
											   | ConvertFrom-Json

	$myBody = @{	parentReference = @{
						driveId = $($myNewDriveId)
						id = $($myNewLibraryRootId.id)
					}
					name = "TestDocument(copy).docx"
				}
	$bodyJson = $myBody | ConvertTo-Json

    mgc drives items copy post --drive-id $myDriveId `
							   --drive-item-id $myFileId `
							   --body $bodyJson

    mgc logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpGraphCli_MoveFile
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"
	$myNewDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm"

    $myNewLibraryRootId = mgc drives root get --drive-id $myNewDriveId `
											   | ConvertFrom-Json

	$myBody = @{	parentReference = @{
						driveId = $($myNewDriveId)
						id = $($myNewLibraryRootId.id)
					}
					name = "TestDocument(moved).docx"
				}
	$bodyJson = $myBody | ConvertTo-Json

    mgc drives items patch --drive-id $myDriveId `
						   --drive-item-id $myFileId `
						   --body $bodyJson

    mgc logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpGraphCli_CreateFolderInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"

    $myLibraryRootId = mgc drives root get --drive-id $myDriveId `
											   | ConvertFrom-Json

	$myBody = @{	name = "NewFolderInLibraryFromGraphCli"
					folder = @{
					}
					"@microsoft.graph.conflictBehavior" = "rename"
				}
	$bodyJson = $myBody | ConvertTo-Json

    mgc drives items children create --drive-id $myDriveId `
									 --drive-item-id $myLibraryRootId `
									 --body $bodyJson

    mgc logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpGraphCli_CheckOutFileInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    mgc drives items checkout post --drive-id $myDriveId `
								   --drive-item-id $myFileId

    mgc logout
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpGraphCli_CheckInFileInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

	$myBody = @{ comment = "Checked in by the CLI" }
	$bodyJson = $myBody | ConvertTo-Json

    mgc drives items checkin post --drive-id $myDriveId `
								  --drive-item-id $myFileId `
								  --body $bodyJson

    mgc logout
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpGraphCli_GetPermissionsFileInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    mgc drives items permissions list --drive-id $myDriveId `
									  --drive-item-id $myFileId

    mgc logout
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpGraphCli_CreatePermissionFileInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

	$myBody = @{	recipients = @(
						@{
							email = "user@domain.onmicrosoft.com"
						}
					)
					message = "This is a file with permissions"
					requireSignIn = $true
					sendInvitation = $true
					roles = @(
						"write"
					)
					password = "password123"
					expirationDateTime = "2024-12-31T23:59:00.000Z"
				}
	$bodyJson = $myBody | ConvertTo-Json

    mgc drives items invite post --drive-id $myDriveId `
								 --drive-item-id $myFileId `
								 --body $bodyJson

    mgc logout
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpGraphCli_DeletePermissionFileInLibrary
{
    # Requires Delegated rights: Files.ReadWrite.All

    PsSpGraphCli_LoginWithCertificate

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"
	$myPermissionId = "aTowIy5mfG1lbWJlcn2QGd1aXRhY2FkZXYub25taWNyb3NvZnQuY29t"

    mgc drives items permissions delete --drive-id $myDriveId `
									    --drive-item-id $myFileId `
										--permission-id $myPermissionId

    mgc logout
}
#gavdcodeend 021


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 021 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#*** Using the MS Graph CLI
#		ATTENTION: There is a Windows Environment Variable already configured in the computer
#					to redirect the commands to the mgc.exe directory (see instructions in the book)
#PsSpGraphCli_GetAllItemsInList
#PsSpGraphCli_GetOneListItemInList
#PsSpGraphCli_CreateOneListItemInList
#PsSpGraphCli_UpdateOneListItemInList
#PsSpGraphCli_DeleteOneListItemFromList
#PsSpGraphCli_GetDrivesInSite
#PsSpGraphCli_GetAllFilesInLibrary
#PsSpGraphCli_GetAllFilesInFolderLibrary
#PsSpGraphCli_GetOneFileMetadata
#PsSpGraphCli_UploadFileToLibrary
#PsSpGraphCli_DownloadFileFromLibrary
#PsSpGraphCli_UpdateOneFileMetadata
#PsSpGraphCli_DeleteOneFile
#PsSpGraphCli_CopyFile
#PsSpGraphCli_MoveFile
#PsSpGraphCli_CreateFolderInLibrary
#PsSpGraphCli_CheckOutFileInLibrary
#PsSpGraphCli_CheckInFileInLibrary
#PsSpGraphCli_GetPermissionsFileInLibrary
#PsSpGraphCli_CreatePermissionFileInLibrary
#PsSpGraphCli_DeletePermissionFileInLibrary

Write-Host "Done" 

