##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function GrPsLoginGraphSDKWithInteraction
{
	Connect-Graph
}

Function GrPsLoginGraphSDKWithAccPw
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

	Connect-Graph -AccessToken $myToken.AccessToken
}

Function GrPsLoginGraphSDKWithSecret
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

	Connect-Graph -AccessToken $myToken.AccessToken
}

Function GrPsLoginGraphSDKWithCertificate
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

Function GrPsLoginGraphSDKWithCertificateFile
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

Function GrPsLoginGraphSDKSetVersion
{
	#Select-MgProfile -Name "beta"
	#Select-MgProfile -Name "v1.0"
}

Function GrPsLoginGraphSDKAssignRights
{
	Connect-Graph -Scopes "Directory.AccessAsUser.All, Directory.ReadWrite.All"
	Get-MgUser
	Disconnect-MgGraph
}

Function GrPsLoginGraphSDKCheckAvailableRights
{
	Find-MgGraphPermission "user" -PermissionType Application
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function SpPsGraphSdk_GetAllLists
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteList -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e"

	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsGraphSdk_GetOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteList -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
				   -ListId "cb3841d2-6561-452c-bbaf-08338bfa0029"

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsGraphSdk_CreateOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	New-MgSiteList -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
				   -Description "New List" `
				   -DisplayName "NewListGraphSdk"

	Disconnect-MgGraph
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsGraphSdk_UpdateOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Update-MgSiteList -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
					  -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b" `
					  -Description "List Updated"

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsGraphSdk_DeleteOneList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteList -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
					  -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b"

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpPsGraphSdk_GetAllFieldsList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListColumn -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
						 -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b"

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpPsGraphSdk_GetOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListColumn -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
						 -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b" `
						 -ColumnDefinitionId "bc91a437-52e7-49e1-8c4e-4698904b2b6d"

	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
Function SpPsGraphSdk_CreateOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	$myField = @{ AllowMultipleLines="true"; TextType="plain" }
	New-MgSiteListColumn -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
						 -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b" `
						 -DisplayName "MyTextField" `
						 -Name "MyTextField" `
						 -Text $myField

	Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpPsGraphSdk_UpdateOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Update-MgSiteListColumn -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
							-ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b" `
							-ColumnDefinitionId "0f8ccb8b-ecb7-437c-a2e8-fbab08a6b544" `
							-Description "Field Description Updated"

	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpPsGraphSdk_DeleteOneFieldList
{
	# Requires Delegated rights: 
	#						Sites.Read.All, Sites.ReadWrite.All, Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteListColumn -SiteId "870ae987-120f-45ed-aa6e-b4a6b7bc226e" `
						    -ListId "f4c8f7ea-62ec-4e64-9aaf-a66c5e12134b" `
						    -ColumnDefinitionId "0f8ccb8b-ecb7-437c-a2e8-fbab08a6b544"

	Disconnect-MgGraph
}
#gavdcodeend 010


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#SpPsGraphSdk_GetAllLists
#SpPsGraphSdk_GetOneList
#SpPsGraphSdk_CreateOneList
#SpPsGraphSdk_UpdateOneList
#SpPsGraphSdk_DeleteOneList
#SpPsGraphSdk_GetAllFieldsList
#SpPsGraphSdk_GetOneFieldList
#SpPsGraphSdk_CreateOneFieldList
#SpPsGraphSdk_UpdateOneFieldList
#SpPsGraphSdk_DeleteOneFieldList

Write-Host "Done" 

