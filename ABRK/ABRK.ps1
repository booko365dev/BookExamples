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

Function PsSpGraphSdk_LoginAssignRights
{
	Connect-Graph -Scopes "Directory.AccessAsUser.All, Directory.ReadWrite.All"
	Get-MgUser
	Disconnect-MgGraph
}

Function PsSpGraphSdk_LoginCheckAvailableRights
{
	Find-MgGraphPermission "user" -PermissionType Application
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function PsSpGraphSdk_GetAllLists
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteList -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab"

	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
Function PsSpGraphSdk_GetOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteList -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
				   -ListId "73de12ba-40e6-426c-890f-952cc9b3c74c"

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsSpGraphSdk_CreateOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	New-MgSiteList -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
				   -Description "New List" `
				   -DisplayName "NewListGraphSdk"

	Disconnect-MgGraph
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsSpGraphSdk_UpdateOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Update-MgSiteList -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
					  -ListId "0b423122-1a90-4451-bc92-5d25433c6962" `
					  -Description "List Updated"

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsSpGraphSdk_DeleteOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteList -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
					  -ListId "0b423122-1a90-4451-bc92-5d25433c6962"

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsSpGraphSdk_GetAllFieldsList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListColumn -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
						 -ListId "0b423122-1a90-4451-bc92-5d25433c6962"

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsSpGraphSdk_GetOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListColumn -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
						 -ListId "0b423122-1a90-4451-bc92-5d25433c6962" `
						 -ColumnDefinitionId "fa564e0f-0c70-4ab9-b863-0177e6ddd247"

	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsSpGraphSdk_CreateOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	$myField = @{ AllowMultipleLines="true"; TextType="plain" }
	New-MgSiteListColumn -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
						 -ListId "0b423122-1a90-4451-bc92-5d25433c6962" `
						 -DisplayName "MyTextField" `
						 -Name "MyTextField" `
						 -Text $myField

	Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsSpGraphSdk_UpdateOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Update-MgSiteListColumn -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
							-ListId "0b423122-1a90-4451-bc92-5d25433c6962" `
							-ColumnDefinitionId "fb5f4740-bf6f-4753-b98c-5c583bb4fd1e" `
							-Description "Field Description Updated"

	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsSpGraphSdk_DeleteOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteListColumn -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
							-ListId "0b423122-1a90-4451-bc92-5d25433c6962" `
						    -ColumnDefinitionId "fb5f4740-bf6f-4753-b98c-5c583bb4fd1e"

	Disconnect-MgGraph
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsSpGraphSdk_GetAllContentTypesSite
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteContentType -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab"

	Disconnect-MgGraph
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsSpGraphSdk_GetAllContentTypesList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListContentType -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
							  -ListId "0b423122-1a90-4451-bc92-5d25433c6962"

	Disconnect-MgGraph
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsSpGraphSdk_GetOneContentTypeSite
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteContentType -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
						  -ContentTypeId "0x01010B"

	Disconnect-MgGraph
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsSpGraphSdk_CreateOneContentTypeSite
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	$ContentTypeParams = @{
		name = "docSet"
		description = "My custom ContentType"
		base = @{
			name = "My ContentType"
			id = "0x010101"
		}
		group = "Document Content Types"
	}

	New-MgSiteContentType -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
						  -BodyParameter $ContentTypeParams

	Disconnect-MgGraph
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsSpGraphSdk_DeleteOneContentTypeSite
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	PsSpGraphSdk_LoginWithAccPwMsal -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteContentType -SiteId "91ee115a-8a5b-49ad-9627-99dae04394ab" `
							 -ContentTypeId "0x01010100044D900EDE741843A113CA8148553442"

	Disconnect-MgGraph
}
#gavdcodeend 015


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 015 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsSpGraphSdk_GetAllLists
#PsSpGraphSdk_GetOneList
#PsSpGraphSdk_CreateOneList
#PsSpGraphSdk_UpdateOneList
#PsSpGraphSdk_DeleteOneList
#PsSpGraphSdk_GetAllFieldsList
#PsSpGraphSdk_GetOneFieldList
#PsSpGraphSdk_CreateOneFieldList
#PsSpGraphSdk_UpdateOneFieldList
#PsSpGraphSdk_DeleteOneFieldList
#PsSpGraphSdk_GetAllContentTypesSite
#PsSpGraphSdk_GetAllContentTypesList
#PsSpGraphSdk_GetOneContentTypeSite
#PsSpGraphSdk_DeleteOneContentTypeSite

Write-Host "Done" 

