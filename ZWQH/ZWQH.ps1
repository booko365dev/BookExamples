
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsCLI()
{
 m365 login --authType certificate `
		    --tenant $cnfTenantName `
		    --appId $cnfClientIdWithCert `
		    --certificateFile $cnfCertificateFilePath `
		    --password $cnfCertificateFilePw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function SpPsCliM365_GetTenantProperties
{
	m365 spo tenant settings list
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsCliM365_UpdateTenantProperties
{
	m365 spo tenant settings set --OneDriveForGuestsEnabled false
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsCliM365_GetRecycleBin
{
	m365 spo tenant recyclebinitem list
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsCliM365_GetRecycleBinQuery
{
	m365 spo tenant recyclebinitem list --query "[?Status == 'Recycled']"
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsCliM365_RestoreRecycleBin
{
	m365 spo tenant recyclebinitem restore `
					--url "https://domain.sharepoint.com/sites/MySite" `
					--wait
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsCliM365_DeleteFromRecycleBin
{
	m365 spo tenant recyclebinitem remove `
					--url "https://domain.sharepoint.com/sites/MySite" `
					--confirm
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsCliM365_GetAllTenantThemes
{
	m365 spo theme list
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsCliM365_GetOneTenantTheme
{
	m365 spo theme get --name "myTheme"
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsCliM365_CreateTenantTheme
{
	m365 spo theme set --name "myTheme01" --theme "C:\Temporary\myThemeColors.json"
}
#gavdcodeend 009

#gavdcodebegin 010
function SpPsCliM365_ApplyTenantTheme
{
	m365 spo theme apply --name "myTheme" `
						 --webUrl "https://domain.sharepoint.com/sites/MySite"
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsCliM365_DeleteTenantTheme
{
	m365 spo theme remove --name Contoso-Blue --confirm
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsCliM365_GetAppCatalog
{
	m365 spo tenant appcatalogurl get
}
#gavdcodeend 012

#gavdcodebegin 013
function SpPsCliM365_CreateAppCatalog
{
	m365 spo tenant appcatalog add `
				--url https://domain.sharepoint.com/sites/AppCatalog `
				--owner user@domain.onmicrosoft.com `
				--timeZone 4 `
				--wait
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsCliM365_GetCdn
{
	m365 spo cdn get --type Public
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsCliM365_SetCdn
{
	m365 spo cdn set --type Public --enabled true
}
#gavdcodeend 015

#gavdcodebegin 016
function SpPsCliM365_GetCdnOrigens
{
	m365 spo cdn origin list --type Public
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsCliM365_CreateCdnOrigen
{
	m365 spo cdn origin add --type Public --origin "*/sites/USSales/myCDN"
}
#gavdcodeend 017

#gavdcodebegin 018
function SpPsCliM365_DeleteCdnOrigen
{
	m365 spo cdn origin remove --type Public --origin "*/sites/USSales/myCDN"
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsCliM365_GetCdnPolicy
{
	m365 spo cdn policy list --type Private
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsCliM365_SetCdnPolicy
{
	m365 spo cdn policy set `
			--type Public `
			--policy IncludeFileExtensions `
			--value "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON"
}
#gavdcodeend 020


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 020 ***

#region ConfigValuesCS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value     # https://domain.onmicrosoft.com
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value   # https://domain.sharepoint.com
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value  # https://domain-admin.sharepoint.com
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value   # https://domain.sharepoint.com/sites/TestSite
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

#------- Using the CLI for Microsoft 365 --------
# Connect to Office 365
$spCtx = LoginPsCLI

#SpPsCliM365_GetTenantProperties
#SpPsCliM365_UpdateTenantProperties
#SpPsCliM365_GetRecycleBin
#SpPsCliM365_GetRecycleBinQuery
#SpPsCliM365_RestoreRecycleBin
#SpPsCliM365_DeleteFromRecycleBin
#SpPsCliM365_GetAllTenantThemes
#SpPsCliM365_GetOneTenantTheme
#SpPsCliM365_CreateTenantTheme
#SpPsCliM365_ApplyTenantTheme
#SpPsCliM365_DeleteTenantTheme
#SpPsCliM365_GetAppCatalog
#SpPsCliM365_CreateAppCatalog
#SpPsCliM365_GetCdn
#SpPsCliM365_SetCdn
#SpPsCliM365_GetCdnOrigens
#SpPsCliM365_CreateCdnOrigen
#SpPsCliM365_DeleteCdnOrigen
#SpPsCliM365_GetCdnPolicy
#SpPsCliM365_SetCdnPolicy

# Disconnect from Office 365
m365 logout

Write-Host "Done" 
