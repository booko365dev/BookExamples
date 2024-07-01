
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
function PsSpGraphCli_GetAllSiteCollections
{
	# Requires Application rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	mgc sites list

	mgc logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpGraphCli_GetOneSiteCollection
{
	# Requires Application rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"

	mgc sites get --site-id $SiteId

	mgc logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpGraphCli_GetFollowedSiteCollections
{
	# Requires Delegated rights: Sites.Read.All
	# Works only for Delegated permissions

	PsSpGraphCli_LoginWithCertificate

	$UserId = "acc28fcb-5261-47f8-960b-715d2f98a431"

	mgc users followed-sites list --user-id $UserId

	mgc logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpGraphCli_FollowSiteCollections
{
	# Requires Application rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$UserId = "acc28fcb-5261-47f8-960b-715d2f98a431"
	$myBody = @{	value = @(
						@{
							id = "domain.sharepoint.com,
									e7cd70c0-6e48-48b4-9380-caab1d1e8433,
									7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
						}
					)
				} | ConvertTo-Json -Depth 5

	mgc users followed-sites add post --user-id $UserId `
									  --body $myBody

	mgc logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpGraphCli_UnFollowSiteCollections
{
	# Requires Application rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$UserId = "acc28fcb-5261-47f8-960b-715d2f98a431"
	$myBody = @{	value = @(
						@{
							id = "domain.sharepoint.com,
									e7cd70c0-6e48-48b4-9380-caab1d1e8433,
									7dc381ab-fa9a-41a7-a98b-7dafa684eb1c"
						}
					)
				} | ConvertTo-Json -Depth 5

	mgc users followed-sites remove post --user-id $UserId `
									     --body $myBody

	mgc logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpGraphCli_GetOneSiteCollectionAllPermissions
{
	# Requires Application rights: Sites.FullControl.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"

	mgc sites permissions list --site-id $SiteId

	mgc logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpGraphCli_GetOneSiteCollectionOnePermission
{
	# Requires Application rights: Sites.FullControl.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$PermissionId = "aTowaS50fG1zLnNwL...C00NTk0LTkwYzMtZTQ3NzJhODE2OGNh"

	mgc sites permissions get --site-id $SiteId `
							  --permission-id $PermissionId

	mgc logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpGraphCli_CreatePermissionInSiteCollection
{
	# Requires Application rights: Sites.FullControl.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myBody = @{     roles = @("write")
					 grantedToIdentities = @(
						@{
							application = @{
								id = "bd6fe5cc-462a-4a60-b9c1-2246d8b7b9fb"
								displayName = "Rights From GraphCLI"
							}
						}
					)
				} | ConvertTo-Json -Depth 5

	mgc sites permissions create --site-id $SiteId `
								 --body $myBody

	mgc logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpGraphCli_UpdatePermissionInSiteCollection
{
	# Requires Application rights: Sites.FullControl.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$PermissionId = "aTowaS50fG1zLnNwL...C00NTk0LTkwYzMtZTQ3NzJhODE2OGNh"
	$myBody = @{     roles = @("read")	} | ConvertTo-Json -Depth 5

	mgc sites permissions patch --site-id $SiteId `
								--permission-id $PermissionId `
								--body $myBody

	mgc logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpGraphCli_DeletePermissionInSiteCollection
{
	# Requires Application rights: Sites.FullControl.All

	PsSpGraphCli_LoginWithCertificate

	$SiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$PermissionId = "aTowaS50fG1zLnNwL...C00NTk0LTkwYzMtZTQ3NzJhODE2OGNh"

	mgc sites permissions delete --site-id $SiteId `
								 --permission-id $PermissionId

	mgc logout
}
#gavdcodeend 010

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 010 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#*** Using the MS Graph CLI
#		ATTENTION: There is a Windows Environment Variable already configured in the computer
#					to redirect the commands to the mgc.exe directory (see instructions in the book)
#PsSpGraphCli_GetAllSiteCollections
#PsSpGraphCli_GetOneSiteCollection
#PsSpGraphCli_GetFollowedSiteCollections
#PsSpGraphCli_FollowSiteCollections
#PsSpGraphCli_UnFollowSiteCollections
#PsSpGraphCli_GetOneSiteCollectionAllPermissions
#PsSpGraphCli_GetOneSiteCollectionOnePermission
#PsSpGraphCli_CreatePermissionInSiteCollection
#PsSpGraphCli_UpdatePermissionInSiteCollection
#PsSpGraphCli_DeletePermissionInSiteCollection

Write-Host "Done" 

