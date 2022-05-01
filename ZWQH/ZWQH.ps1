
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsCLI()
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 01
function PsCliSharePoint_GetTenantProperties
{
	m365 spo tenant settings list
}
#gavdcodeend 01

#gavdcodebegin 02
function PsCliSharePoint_UpdateTenantProperties
{
	m365 spo tenant settings set --OneDriveForGuestsEnabled false
}
#gavdcodeend 02

#gavdcodebegin 03
function PsCliSharePoint_GetRecycleBin
{
	m365 spo tenant recyclebinitem list
}
#gavdcodeend 03

#gavdcodebegin 04
function PsCliSharePoint_GetRecycleBinQuery
{
	m365 spo tenant recyclebinitem list --query "[?Status == 'Recycled']"
}
#gavdcodeend 04

#gavdcodebegin 05
function PsCliSharePoint_RestoreRecycleBin
{
	m365 spo tenant recyclebinitem restore `
					--url "https://domain.sharepoint.com/sites/MySite" `
					--wait
}
#gavdcodeend 05

#gavdcodebegin 06
function PsCliSharePoint_DeleteFromRecycleBin
{
	m365 spo tenant recyclebinitem remove `
					--url "https://domain.sharepoint.com/sites/MySite" `
					--confirm
}
#gavdcodeend 06

#gavdcodebegin 07
function PsCliSharePoint_GetAllTenantThemes
{
	m365 spo theme list
}
#gavdcodeend 07

#gavdcodebegin 08
function PsCliSharePoint_GetOneTenantTheme
{
	m365 spo theme get --name "myTheme"
}
#gavdcodeend 08

#gavdcodebegin 09
function PsCliSharePoint_CreateTenantTheme
{
	m365 spo theme set --name "myTheme01" --theme "C:\Temporary\myThemeColors.json"
}
#gavdcodeend 09

#gavdcodebegin 10
function PsCliSharePoint_ApplyTenantTheme
{
	m365 spo theme apply --name "myTheme" `
						 --webUrl "https://domain.sharepoint.com/sites/MySite"
}
#gavdcodeend 10

#gavdcodebegin 11
function PsCliSharePoint_DeleteTenantTheme
{
	m365 spo theme remove --name Contoso-Blue --confirm
}
#gavdcodeend 11

#gavdcodebegin 12
function PsCliSharePoint_GetAppCatalog
{
	m365 spo tenant appcatalogurl get
}
#gavdcodeend 12

#gavdcodebegin 13
function PsCliSharePoint_CreateAppCatalog
{
	m365 spo tenant appcatalog add `
				--url https://domain.sharepoint.com/sites/AppCatalog `
				--owner user@domain.onmicrosoft.com `
				--timeZone 4 `
				--wait
}
#gavdcodeend 13

#gavdcodebegin 14
function PsCliSharePoint_GetCdn
{
	m365 spo cdn get --type Public
}
#gavdcodeend 14

#gavdcodebegin 15
function PsCliSharePoint_SetCdn
{
	m365 spo cdn set --type Public --enabled true
}
#gavdcodeend 15

#gavdcodebegin 16
function PsCliSharePoint_GetCdnOrigens
{
	m365 spo cdn origin list --type Public
}
#gavdcodeend 16

#gavdcodebegin 17
function PsCliSharePoint_CreateCdnOrigen
{
	m365 spo cdn origin add --type Public --origin "*/sites/USSales/myCDN"
}
#gavdcodeend 17

#gavdcodebegin 18
function PsCliSharePoint_DeleteCdnOrigen
{
	m365 spo cdn origin remove --type Public --origin "*/sites/USSales/myCDN"
}
#gavdcodeend 18

#gavdcodebegin 19
function PsCliSharePoint_GetCdnPolicy
{
	m365 spo cdn policy list --type Private
}
#gavdcodeend 19

#gavdcodebegin 20
function PsCliSharePoint_SetCdnPolicy
{
	m365 spo cdn policy set `
			--type Public `
			--policy IncludeFileExtensions `
			--value "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON"
}
#gavdcodeend 20


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
# Connect to Office 365
$spCtx = LoginPsCLI

#PsCliSharePoint_GetTenantProperties
#PsCliSharePoint_UpdateTenantProperties
#PsCliSharePoint_GetRecycleBin
#PsCliSharePoint_GetRecycleBinQuery
#PsCliSharePoint_RestoreRecycleBin
#PsCliSharePoint_DeleteFromRecycleBin
#PsCliSharePoint_GetAllTenantThemes
#PsCliSharePoint_GetOneTenantTheme
#PsCliSharePoint_CreateTenantTheme
#PsCliSharePoint_ApplyTenantTheme
#PsCliSharePoint_DeleteTenantTheme
#PsCliSharePoint_GetAppCatalog
#PsCliSharePoint_CreateAppCatalog
#PsCliSharePoint_GetCdn
#PsCliSharePoint_SetCdn
#PsCliSharePoint_GetCdnOrigens
#PsCliSharePoint_CreateCdnOrigen
#PsCliSharePoint_DeleteCdnOrigen
#PsCliSharePoint_GetCdnPolicy
#PsCliSharePoint_SetCdnPolicy

# Disconnect from Office 365
m365 logout

Write-Host "Done" 
