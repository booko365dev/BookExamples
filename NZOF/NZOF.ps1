
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpCliM365_LoginWithAccPw
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCliM365_GetModernSiteCollections
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site list
	#m365 spo site list --type "TeamSite"
	#m365 spo site list --deleted

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpCliM365_GetOneModernSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site list --type TeamSite --filter "Url -like 'contoso'"
	#m365 spo site list --type TeamSite `
	#			--filter "Url -eq 'https://domain.sharepoint.com/sites/Contoso'"

	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpCliM365_GetClassicSiteCollections
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site classic list
	#m365 spo site classic list --webTemplate "STS#3" #"APPCATALOG#0"

	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpCliM365_GetPropertiesOneSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site get --url $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpCliM365_CreateSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
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
function PsSpCliM365_RenameSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site rename `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI") `
			--newSiteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
			--newSiteTitle "Updated Site Coll" `
			--wait flag

	m365 logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpCliM365_UpdateSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site set `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--owners "user@domain.OnMicrosoft.com"

	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpCliM365_DeleteSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site remove `
				--url ($configFile.appsettings.SiteBaseUrl + "/sites/UpdSiteFromCLI") `
				--skipRecycleBin `
				--confirm

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpCliM365_GetRecyclebonSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpCliM365_GetRecyclebinQuerySiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			 --query "[?LeafName == 'Document.docx']"

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpCliM365_GetRecyclebinTypeSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site recyclebinitem list `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--type "files"  #"folders" "listItems"

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpCliM365_GetRecyclebinRestoreSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site recyclebinitem restore `
			--siteUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
			--ids "4f219584-c959-400e-911a-d9ec5b460ea1"

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpCliM365_SetChromeSiteCollection
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo site chrome set `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewCommunicationSite") `
		--headerLayout "Extended" --logoAlignment "Right" --disableFooter "true"

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpCliM365_GetWebs
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web list `
			--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01")

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpCliM365_GetSubWeb
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web get `
	--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/Subweb01") `
	--withGroups

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpCliM365_CreateSubWeb
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web add `
	--parentWebUrl ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01") `
	--url "MySubweb" `
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
function PsSpCliM365_DeleteSubWeb
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web remove `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/MySubweb") `
		--confirm

	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpCliM365_LanguagesSubWeb
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web installedlanguage list `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpCliM365_ReindexSubWeb
{
	$spCtx = PsSpCliM365_LoginWithAccPw
	
	m365 spo web reindex  `
		--url ($configFile.appsettings.SiteBaseUrl + "/sites/NewSiteFromCLI01/Subsite01") 

	m365 logout
}
#gavdcodeend 019


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 019 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#PsSpCliM365_GetModernSiteCollections
#PsSpCliM365_GetOneModernSiteCollection
#-->PsSpCliM365_GetClassicSiteCollections
#PsSpCliM365_GetPropertiesOneSiteCollection
#PsSpCliM365_CreateSiteCollection
#PsSpCliM365_RenameSiteCollection
#PsSpCliM365_UpdateSiteCollection
#PsSpCliM365_DeleteSiteCollection
#PsSpCliM365_GetRecyclebonSiteCollection
#PsSpCliM365_GetRecyclebinQuerySiteCollection
#PsSpCliM365_GetRecyclebinTypeSiteCollection
#PsSpCliM365_GetRecyclebinRestoreSiteCollection
#PsSpCliM365_SetChromeSiteCollection
#PsSpCliM365_GetWebs
#PsSpCliM365_GetSubWeb
#PsSpCliM365_CreateSubWeb
#PsSpCliM365_DeleteSubWeb
#PsSpCliM365_LanguagesSubWeb
#PsSpCliM365_ReindexSubWeb

Write-Host "Done" 
