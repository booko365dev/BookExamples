##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsGraphSdk_LoginWithAccPwMSAL
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

	Connect-MgGraph -AccessToken $myTokenSecure
}

function PsGraphSdk_LoginWithSecret
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

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Other routines ***---------------------------
##---------------------------------------------------------------------------------------


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpGraphSdk_GetAllSiteCollections
{
	# App Registration type: Graph 
	# App Registration permissions: Sites.Read.All, Sites.ReadWrite.All
	
	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	Get-MgSite
    
	Disconnect-MgGraph
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpGraphSdk_GetOneSiteCollection
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	Get-MgSite -Search "Test*"
	Get-MgSite -SiteId "e7cd70c0-6e48-48b4-9380-caab1d1e8433"

	Disconnect-MgGraph
}
#gavdcodeend 002

#gavdcodebegin 004
function PsSpGraphSdk_GetFollowedSiteCollections
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All
	# Works only for Delegated permissions

	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	Get-MgUserFollowedSite -UserId "acc28fcb-5261-47f8-960b-715d2f98a431"

	Disconnect-MgGraph
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpGraphSdk_FollowSiteCollections
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	$sitesToFollow = @{
		value = @(
			@{
				id = "domain.sharepoint.com,"`
					"e7cd70c0-6e48-48b4-9380-caab1d1e8433,"`
					"7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
			}
		)
	}

	Add-MgUserFollowedSite -UserId "acc28fcb-5261-47f8-960b-715d2f98a431" `
							-BodyParameter $sitesToFollow

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpGraphSdk_UnFollowSiteCollections
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	$sitesToUnFollow = @{
		value = @(
			@{
				id = "domain.sharepoint.com,"1
					"e7cd70c0-6e48-48b4-9380-caab1d1e8433,"`
					"7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
			}
		)
	}

	Remove-MgUserFollowedSite -UserId "acc28fcb-5261-47f8-960b-715d2f98a431" `
							  -BodyParameter $sitesToUnFollow

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 003
function PsSpGraphSdk_GetWebs
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsGraphSdk_LoginWithSecret -TenantName $configFile.appsettings.TenantName `
								-ClientID $configFile.appsettings.ClientIdWithSecret `
								-ClientSecret $configFile.appsettings.ClientSecret

	Get-MgSubSite -SiteId "7f80b1d6-885a-4630-91c1-57b45c9159cb"

	Disconnect-MgGraph
}
#gavdcodeend 003

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 006 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsSpGraphSdk_GetAllSiteCollections
#PsSpGraphSdk_GetOneSiteCollection
#PsSpGraphSdk_GetFollowedSiteCollections
#PsSpGraphSdk_FollowSiteCollections
#PsSpGraphSdk_UnFollowSiteCollections
#PsSpGraphSdk_GetWebs

Write-Host "Done" 

