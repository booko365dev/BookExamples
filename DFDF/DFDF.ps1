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
Function SpPsGraphSdk_GetAllListsItems
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListItem -SiteId "5e83b6c6-f7da-4215-acfa-bc93a95bc356" `
					   -ListId "25820c38-3423-49bf-8c38-f774337af29c"
	# Use the SiteId, not the WebId or the RootWeb.ID 

	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
Function SpPsGraphSdk_GetOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgSiteListItem -SiteId "5e83b6c6-f7da-4215-acfa-bc93a95bc356" `
					   -ListId "25820c38-3423-49bf-8c38-f774337af29c" `
					   -ListItemId "8"
	# Use the SiteId, not the WebId or the RootWeb.ID 

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 003
Function SpPsGraphSdk_CreateOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	New-MgSiteListItem -SiteId "5e83b6c6-f7da-4215-acfa-bc93a95bc356" `
					   -ListId "25820c38-3423-49bf-8c38-f774337af29c" `
					   -Fields @{Title = "New_SpPsGraphSdkItem"}

	Disconnect-MgGraph
}
#gavdcodeend 003

#gavdcodebegin 004
Function SpPsGraphSdk_UpdateOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Update-MgSiteListItem -SiteId "5e83b6c6-f7da-4215-acfa-bc93a95bc356" `
						  -ListId "25820c38-3423-49bf-8c38-f774337af29c" `
						  -ListItemId "9" `
						  -Fields @{Title = "Update_SpPsGraphSdkItem"}

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpPsGraphSdk_DeleteOneListItem
{
	# Requires Delegated rights: 
	#						Sites.Read.All/Sites.ReadWrite.All/Sites.FullControl.All

	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Remove-MgSiteListItem -SiteId "5e83b6c6-f7da-4215-acfa-bc93a95bc356" `
						  -ListId "25820c38-3423-49bf-8c38-f774337af29c" `
						  -ListItemId "8"

	Disconnect-MgGraph
}
#gavdcodeend 005

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 05 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#SpPsGraphSdk_GetAllListsItems
#SpPsGraphSdk_GetOneListItem
#SpPsGraphSdk_CreateOneListItem
#SpPsGraphSdk_UpdateOneListItem
#SpPsGraphSdk_DeleteOneListItem

Write-Host "Done" 

