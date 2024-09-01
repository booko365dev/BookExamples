##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpGraphSdk_LoginWithInteraction
{
	Connect-Graph
}

Function PsSpGraphSdk_LoginWithAccPwMsal
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $UserName, $securePW

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 
	$myTokenSecure = ConvertTo-SecureString -String $myToken.AccessToken `
											-AsPlainText -Force

	Connect-Graph -AccessToken $myTokenSecure
}

Function PsSpGraphSdk_LoginWithSecretMsal
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$secureSecret = ConvertTo-SecureString -String `
								$ClientSecret -AsPlainText -Force

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -ClientSecret ($secureSecret)
	$myTokenSecure = ConvertTo-SecureString -String $myToken.AccessToken `
											-AsPlainText -Force

	Connect-Graph -AccessToken $myTokenSecure
}

function PsSpGraphSdk_LoginWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$ClientSecret -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $ClientID, $securePW

	Connect-MgGraph -TenantId $TenantName `
					-ClientSecretCredential $myCredentials
}

Function PsSpGraphSdk_LoginWithCertificate
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$CertificateThumbprint
	)

	Connect-MgGraph -TenantId $TenantName `
					-ClientId $ClientId `
					-CertificateThumbprint $CertificateThumbprint
}

Function PsSpGraphSdk_LoginWithCertificateFile
{
	[SecureString]$secureCertPw = ConvertTo-SecureString -String `
							$configFile.appSettings.CertificateFilePw -AsPlainText -Force

	$myCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(`
							  $configFile.appSettings.CertificateFilePath, $secureCertPw)
	
	Connect-MgGraph -TenantId $configFile.appsettings.TenantName `
					-ClientId $configFile.appsettings.ClientIdWithCert `
					-Certificate $myCert 
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Other routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpGraphSdk_LoginSetVersion
{
	#Select-MgProfile -Name "beta"
	#Select-MgProfile -Name "v1.0"
}

Function PsSpGraphSdk_AssignRights
{
	Connect-Graph -Scopes "Directory.AccessAsUser.All, Directory.ReadWrite.All"
	Get-MgUser
	Disconnect-MgGraph
}

Function PsSpGraphSdk_CheckAvailableRights
{
	Find-MgGraphPermission "user" -PermissionType Application
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function PsSpGraphSdk_GetAllItemsInList
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"

	Get-MgSiteListItem -SiteId $mySiteId `
					   -ListId $myListId
	# Use the SiteId, not the WebId or the RootWeb.ID 

	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
Function PsSpGraphSdk_GetOneItemInList
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myListItemId = "10"

	Get-MgSiteListItem -SiteId $mySiteId `
					   -ListId $myListId `
					   -ListItemId $myListItemId
	# Use the SiteId, not the WebId or the RootWeb.ID 

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsSpGraphSdk_CreateOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw
	
	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"

	New-MgSiteListItem -SiteId $mySiteId `
					   -ListId $myListId `
					   -Fields @{Title = "New_SpPsGraphSdkItem"}

	Disconnect-MgGraph
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsSpGraphSdk_UpdateOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw
	
	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myListItemId = "11"

	Update-MgSiteListItem -SiteId $mySiteId `
					      -ListId $myListId `
						  -ListItemId $myListItemId `
						  -Fields @{Title = "Update_SpPsGraphSdkItem"}

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsSpGraphSdk_DeleteOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "c6d81938-b786-4af4-b2dd-a2132787f1d9"
	$myListItemId = "10"

	Remove-MgSiteListItem -SiteId $mySiteId `
					      -ListId $myListId `
						  -ListItemId $myListItemId

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsSpGraphSdk_GetDriveOfLibraryByLibraryId
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw
	
	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "b331af17-edc3-4058-909e-e6fa74abe946"

	$myDrive = Get-MgSiteListDrive -SiteId $mySiteId `
								   -ListId $myListId

	Write-Host "Drive ID for Library - " $myDrive.Id

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsSpGraphSdk_GetAllFilesInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

    $myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"

    $myLibraryRootId = Get-MgDriveRoot -DriveId $myDriveId
	write-host "Root ID of Library - " $myLibraryRootId.Id
    $myDocs = Get-MgDriveItemChild -DriveId $myDriveId -DriveItemId $myLibraryRootId.Id

    foreach ($oneDoc in $myDocs) {
        Write-Host "Name: $($oneDoc.Name), ID: $($oneDoc.Id)"
    }

    Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsSpGraphSdk_GetAllFilesInFolderLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

    $myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFolderId = "01IAJF3RFCB4MZLKVESFH2OPTWWU7NJHOX"

    $myDocs = Get-MgDriveItemChild -DriveId $myDriveId -DriveItemId $myFolderId

    foreach ($oneDoc in $myDocs) {
        Write-Host "Name: $($oneDoc.Name), ID: $($oneDoc.Id)"
    }

    Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsSpGraphSdk_GetOneFileMetadata
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

    $myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    $fileMetadata = Get-MgDriveItem -DriveId $myDriveId -DriveItemId $myFileId

    Write-Host "Name - $($fileMetadata.Name)"
    Write-Host "ID - $($fileMetadata.Id)"
    Write-Host "Size - $($fileMetadata.Size) bytes"
    Write-Host "Created DateTime - $($fileMetadata.CreatedDateTime)"
    Write-Host "Last Modified DateTime - $($fileMetadata.LastModifiedDateTime)"

    Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsSpGraphSdk_UploadOneFileToLibrary   # +++> TODO
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

    $myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
    $myFolderId = "root"
	$myFilePath = "C:\Path\To\Your\File.txt"

    $myFileName = [System.IO.Path]::GetFileName($myFilePath)
    ##$uploadPath = if ($myFolderId -eq "root") { "root:/$myFileName:" } else { "items/$myFolderId:/$fileName:" }

    ##Add-MgDriveItemContent -DriveId $myDriveId  ===> No such command available at the moment (2024-09)
 
	Disconnect-MgGraph
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsSpGraphSdk_DownloadOneFileFromLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

    $myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

	Get-MgDriveItemContent -DriveId $myDriveId `
						   -DriveItemId  $myFileId `
						   -OutFile "C:\Temporary\MyDownloadedFile.docx"

	Disconnect-MgGraph
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsSpGraphSdk_UpdateOneFileMetadata
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

	$metadataToUpdate = @{
		"name" = "UpdatedTestDocument.docx";
		"description" = "Updated description for the file";
	}
    $jsonMetadata = $metadataToUpdate | ConvertTo-Json

    Update-MgDriveItem -DriveId $myDriveId `
                       -DriveItemId $myFileId `
                       -BodyParameter $jsonMetadata

    Disconnect-MgGraph
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsSpGraphSdk_DeleteOneFile
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RFN3G62VR3YQ5BKRCXVYRUHBHFW"

    Remove-MgDriveItem -DriveId $myDriveId `
					   -DriveItemId $myFileId `
					   -Confirm:$false

    Disconnect-MgGraph
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsSpGraphSdk_CopyFile
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RDCCOXYUSRL3NDYSAMQFGPNB26O"
	$myNewDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm"

    $myNewLibraryRootId = Get-MgDriveRoot -DriveId $myNewDriveId

	$myBody = @{	parentReference = @{
						driveId = $($myNewDriveId)
						id = $($myNewLibraryRootId.Id)
					}
					name = "TestDocument(copy).jpg"
				}
	$bodyJson = $myBody | ConvertTo-Json

	Copy-MgDriveItem -DriveId $myDriveId `
					 -DriveItemId $myFileId `
					 -BodyParameter $bodyJson

    Disconnect-MgGraph
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsSpGraphSdk_MoveFile
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RDCCOXYUSRL3NDYSAMQFGPNB26O"
	$myNewDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQoj6VQi19WQoEABpsrQuCm"

    $myNewLibraryRootId = Get-MgDriveRoot -DriveId $myNewDriveId

	$myBody = @{	parentReference = @{
						driveId = $($myNewDriveId)
						id = $($myNewLibraryRootId.Id)
					}
					name = "TestDocument(moved).jpg"
				}
	$bodyJson = $myBody | ConvertTo-Json

    Update-MgDriveItem -DriveId $myDriveId `
                       -DriveItemId $myFileId `
                       -BodyParameter $bodyJson

    Disconnect-MgGraph
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsSpGraphSdk_CreateFolderInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"

    $myLibraryRootId = Get-MgDriveRoot -DriveId $myDriveId

	$myBody = @{	name = "NewFolderInLibraryFromGraphSdk"
					folder = @{
					}
					"@microsoft.graph.conflictBehavior" = "rename"
				}
	$bodyJson = $myBody | ConvertTo-Json

    New-MgDriveItemChild -DriveId $myDriveId `
						 -DriveItemId $myLibraryRootId `
						 -BodyParameter $bodyJson

    Disconnect-MgGraph
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsSpGraphSdk_CheckOutFileInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    Invoke-MgCheckoutDriveItem -DriveId $myDriveId `
							   -DriveItemId $myFileId

    Disconnect-MgGraph
}
#gavdcodeend 017

#gavdcodebegin 018
Function PsSpGraphSdk_CheckInFileInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    Invoke-MgCheckinDriveItem -DriveId $myDriveId `
							  -DriveItemId $myFileId

    Disconnect-MgGraph
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsSpGraphSdk_GetPermissionsFileInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"

    Get-MgDriveItemPermission -DriveId $myDriveId `
							  -DriveItemId $myFileId

    Disconnect-MgGraph
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsSpGraphSdk_CreatePermissionFileInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

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

    Invoke-MgInviteDriveItem -DriveId $myDriveId `
							 -DriveItemId $myFileId `
							 -BodyParameter $bodyJson

    Disconnect-MgGraph
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsSpGraphSdk_DeletePermissionFileInLibrary
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

    PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							    -ClientID $configFile.appsettings.ClientIdWithAccPw `
							    -UserName $configFile.appsettings.UserName `
							    -UserPw $configFile.appsettings.UserPw

	$myDriveId = "b!WhHukVuKrUmWJ5na4EOUq74PGWyAYoJInwMa9_X4aGQXrzGzw-1YQJCe5vp0q-lG"
	$myFileId = "01IAJF3RCKV2M3NS37VNFJD35UORISHQS7"
	$myPermissionId = "aTowIy5mfG1lbWJlcnNoaXB8YWRlbGVXRhY2FkZXYub25taWNyb3NvZnQuY29t"

    Remove-MgDriveItemPermission -DriveId $myDriveId `
								 -DriveItemId $myFileId `
								 -PermissionId $myPermissionId

    Disconnect-MgGraph
}
#gavdcodeend 021


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 021 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsSpGraphSdk_GetAllItemsInList
#PsSpGraphSdk_GetOneItemInList
#PsSpGraphSdk_CreateOneListItem
#PsSpGraphSdk_UpdateOneListItem
#PsSpGraphSdk_DeleteOneListItem
#PsSpGraphSdk_GetDriveOfLibraryByLibraryId
#PsSpGraphSdk_GetAllFilesInLibrary
#PsSpGraphSdk_GetAllFilesInFolderLibrary
#PsSpGraphSdk_GetOneFileMetadata
#PsSpGraphSdk_UploadOneFileToLibrary
#PsSpGraphSdk_DownloadOneFileFromLibrary
#PsSpGraphSdk_UpdateOneFileMetadata
#PsSpGraphSdk_CopyFile
#PsSpGraphSdk_MoveFile
#PsSpGraphSdk_DeleteOneFile
#PsSpGraphSdk_CreateFolderInLibrary
#PsSpGraphSdk_CheckOutFileInLibrary
#PsSpGraphSdk_CheckInFileInLibrary
#PsSpGraphSdk_GetPermissionsFileInLibrary
#PsSpGraphSdk_CreatePermissionFileInLibrary
#PsSpGraphSdk_DeletePermissionFileInLibrary

Write-Host "Done" 

