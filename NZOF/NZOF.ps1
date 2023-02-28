
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

#gavdcodebegin 001
function SpPsCliM365_GetModernSiteCollections
{
	$spCtx = LoginPsCLI
	
	m365 spo site list
	#m365 spo site list --type "TeamSite"
	#m365 spo site list --deleted

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsCliM365_GetOneModernSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site list --type TeamSite --filter "Url -like 'contoso'"
	#m365 spo site list --type TeamSite `
	#			--filter "Url -eq 'https://domain.sharepoint.com/sites/Contoso'"

	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsCliM365_GetClassicSiteCollections
{
	$spCtx = LoginPsCLI
	
	m365 spo site classic list
	#m365 spo site classic list --webTemplate "STS#3" #"APPCATALOG#0"

	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsCliM365_GetPropertiesOneSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site get --url $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsCliM365_CreateSiteCollection
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
#gavdcodeend 005

#gavdcodebegin 006
function SpPsCliM365_RenameSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site rename `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI") `
			--newSiteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
			--newSiteTitle "Updated Site Coll" `
			--wait flag

	m365 logout
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsCliM365_UpdateSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site set `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--owners "user@domain.OnMicrosoft.com"

	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsCliM365_DeleteSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site remove `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--skipRecycleBin `
				--confirm

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsCliM365_GetRecyclebonSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 009

#gavdcodebegin 010
function SpPsCliM365_GetRecyclebinQuerySiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			 --query "[?LeafName == 'Document.docx']"

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsCliM365_GetRecyclebinTypeSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--type "files"  #"folders" "listItems"

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsCliM365_GetRecyclebinRestoreSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site recyclebinitem restore `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--ids "4f219584-c959-400e-911a-d9ec5b460ea1"

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function SpPsCliM365_SetChromeSiteCollection
{
	$spCtx = LoginPsCLI
	
	m365 spo site chrome set `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewCommunicationSite") `
		--headerLayout "Extended" --logoAlignment "Right" --disableFooter "true"

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsCliM365_GetWebs
{
	$spCtx = LoginPsCLI
	
	m365 spo web list `
			--webUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsCliM365_GetSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web get `
	--webUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/Subweb01") 
	--withGroups

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function SpPsCliM365_CreateSubWeb
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
#gavdcodeend 016

#gavdcodebegin 017
function SpPsCliM365_DeleteSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web remove `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/MySubweb") `
		--confirm

	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
function SpPsCliM365_LanguagesSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web installedlanguage list `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsCliM365_ReindexSubWeb
{
	$spCtx = LoginPsCLI
	
	m365 spo web reindex  `
		--webUrl ($configFile.appsettings.SiteBaseUrl + `
													"/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 019


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#SpPsCliM365_GetModernSiteCollections
#SpPsCliM365_GetOneModernSiteCollection
#SpPsCliM365_GetClassicSiteCollections
#SpPsCliM365_GetPropertiesOneSiteCollection
#SpPsCliM365_CreateSiteCollection
#SpPsCliM365_RenameSiteCollection
#SpPsCliM365_UpdateSiteCollection
#SpPsCliM365_DeleteSiteCollection
#SpPsCliM365_GetRecyclebonSiteCollection
#SpPsCliM365_GetRecyclebinQuerySiteCollection
#SpPsCliM365_GetRecyclebinTypeSiteCollection
#SpPsCliM365_GetRecyclebinRestoreSiteCollection
#SpPsCliM365_SetChromeSiteCollection
#SpPsCliM365_GetWebs
#SpPsCliM365_GetSubWeb
#SpPsCliM365_CreateSubWeb
#SpPsCliM365_DeleteSubWeb
#SpPsCliM365_LanguagesSubWeb
#SpPsCliM365_ReindexSubWeb

Write-Host "Done" 
