
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
function PsCliSharePoint_GetModernSiteCollections
{
	$spCtx = LoginPsCLI
	
	m365 spo site list
	#m365 spo site list --type "TeamSite"
	#m365 spo site list --deleted

	m365 logout
}
#gavdcodeend 01

#gavdcodebegin 02
function PsCliSharePoint_GetOneModernSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site list --type TeamSite --filter "Url -like 'contoso'"
	#m365 spo site list --type TeamSite `
	#			--filter "Url -eq 'https://domain.sharepoint.com/sites/Contoso'"

	m365 logout
}
#gavdcodeend 02

#gavdcodebegin 03
function PsCliSharePoint_GetClassicSiteCollections
{
	$spCtx = LoginPsCLI
	
	m365 spo site classic list
	#m365 spo site classic list --webTemplate "STS#3" #"APPCATALOG#0"

	m365 logout
}
#gavdcodeend 03

#gavdcodebegin 04
function PsCliSharePoint_GetPropertiesOneSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site get --url $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 04

#gavdcodebegin 05
function PsCliSharePoint_CreateSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site add --type "CommunicationSite" `
					  --siteDesign "Showcase" `
					  --url ($configFile.appsettings.SiteBaseUrl + "/sites/NewComFromCLI") `
					  --title "NewCommunicationSiteFromCLI" `
					  --description "New Communication Site" `
					  --owners "user@domain.OnMicrosoft.com" `
					  --shareByEmailEnabled `
					  --allowFileSharingForGuestUsers

	m365 spo site add --type "TeamSite" `
					  --alias "NewTeamSite" `
					  --title "NewTeamSite" `
					  --description "New Team Site" `
					  --isPublic `
					  --lcid 3082 `
					  --owners "user@domain.OnMicrosoft.com"

	m365 spo site add --type "ClassicSite" `
					  --url ($configFile.appsettings.SiteBaseUrl + "/sites/NewClaFromCLI") `
					  --title "NewClassicTeamSite" `
					  --description "New Classic Team Site" `
					  --owners "user@domain.OnMicrosoft.com" `
					  --timeZone 4 `
					  --lcid 1043 `
					  --storageQuota 1 `
					  --storageQuotaWarningLevel 1 `
					  --resourceQuota  1 `
					  --resourceQuotaWarningLevel 1 `
					  --webTemplate "STS#0"

	m365 logout
}
#gavdcodeend 05

#gavdcodebegin 06
function PsCliSharePoint_RenameSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site rename `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI") `
			--newSiteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
			--newSiteTitle "Updated Site Coll" `
			--wait flag

	m365 logout
}
#gavdcodeend 06

#gavdcodebegin 07
function PsCliSharePoint_UpdateSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site set `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--owners "user@domain.OnMicrosoft.com"

	m365 logout
}
#gavdcodeend 07

#gavdcodebegin 08
function PsCliSharePoint_DeleteSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site remove `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--skipRecycleBin `
				--confirm

	m365 logout
}
#gavdcodeend 08

#gavdcodebegin 09
function PsCliSharePoint_GetRecyclebonSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 09

#gavdcodebegin 10
function PsCliSharePoint_GetRecyclebinQuerySiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			 --query "[?LeafName == 'Document.docx']"

	m365 logout
}
#gavdcodeend 10

#gavdcodebegin 11
function PsCliSharePoint_GetRecyclebinTypeSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--type "files"  #"folders" "listItems"

	m365 logout
}
#gavdcodeend 11

#gavdcodebegin 12
function PsCliSharePoint_GetRecyclebinRestoreSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem restore `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--ids "4f219584-c959-400e-911a-d9ec5b460ea1"

	m365 logout
}
#gavdcodeend 12

#gavdcodebegin 13
function PsCliSharePoint_SetChromeSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site chrome set `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewCommunicationSite") `
		--headerLayout "Extended" --logoAlignment "Right" --disableFooter "true"

	m365 logout
}
#gavdcodeend 13

#gavdcodebegin 14
function PsCliSharePoint_GetWebs
{
	$spCtx = LoginPsCLI
	
	m365 spo web list `
			--webUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 14

#gavdcodebegin 15
function PsCliSharePoint_GetSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web get `
	--webUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/Subweb01") 
	--withGroups

	m365 logout
}
#gavdcodeend 15

#gavdcodebegin 16
function PsCliSharePoint_CreateSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web add `
	--parentWebUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
	--webUrl "MySubweb" `
	--title "My Sub-Site" `
	--description "This is my sub-site" `
	--webTemplate "STS#0" `
	--locale "1033" `
	--inheritNavigation `
	--breakInheritance

	m365 logout
}
#gavdcodeend 16

#gavdcodebegin 17
function PsCliSharePoint_DeleteSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web remove `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/MySubweb") `
		--confirm

	m365 logout
}
#gavdcodeend 17

#gavdcodebegin 18
function PsCliSharePoint_LanguagesSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web installedlanguage list `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 18

#gavdcodebegin 19
function PsCliSharePoint_ReindexSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web reindex  `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 19


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#PsCliSharePoint_GetModernSiteCollections
#PsCliSharePoint_GetOneModernSiteCollection
#PsCliSharePoint_GetClassicSiteCollections
#PsCliSharePoint_GetPropertiesOneSiteCollection
#PsCliSharePoint_CreateSiteCollection
#PsCliSharePoint_RenameSiteCollection
#PsCliSharePoint_UpdateSiteCollection
#PsCliSharePoint_DeleteSiteCollection
#PsCliSharePoint_GetRecyclebonSiteCollection
#PsCliSharePoint_GetRecyclebinQuerySiteCollection
#PsCliSharePoint_GetRecyclebinTypeSiteCollection
#PsCliSharePoint_GetRecyclebinRestoreSiteCollection
#PsCliSharePoint_SetChromeSiteCollection
#PsCliSharePoint_GetWebs
#PsCliSharePoint_GetSubWeb
#PsCliSharePoint_CreateSubWeb
#PsCliSharePoint_DeleteSubWeb
#PsCliSharePoint_LanguagesSubWeb
#PsCliSharePoint_ReindexSubWeb

Write-Host "Done" 
