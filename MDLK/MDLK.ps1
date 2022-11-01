
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsCLI
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#------- Term Store --------
#gavdcodebegin 01
function SpPsCliM365_CreateTermGroup
{
	$spCtx = LoginPsCLI

	m365 spo term group add --name "PsCliTermGroup" `
							--description "Group description"


	m365 logout
}
#gavdcodeend 01

#gavdcodebegin 02
function SpPsCliM365_FindAllTermGroups
{
	$spCtx = LoginPsCLI

	m365 spo term group list
	
	m365 logout
}
#gavdcodeend 02

#gavdcodebegin 03
function SpPsCliM365_FindOneTermGroup
{
	$spCtx = LoginPsCLI

	m365 spo term group get --name "PsCliTermGroup"

	$myTermGroup = m365 spo term group get --id "ba048db1-5337-4795-b662-262af3789994"
	$myTermGroupObj = $myTermGroup | ConvertFrom-Json
	Write-Host $myTermGroupObj.Name
	
	m365 logout
}
#gavdcodeend 03

#gavdcodebegin 04
function SpPsCliM365_CreateTermSet
{
	$spCtx = LoginPsCLI

	m365 spo term set add --name "PsCliTermSet" `
						  --termGroupName "PsCliTermGroup" `
						  --description "Set description"

	
	m365 logout
}
#gavdcodeend 04

#gavdcodebegin 05
function SpPsCliM365_FindAllTermSets
{
	$spCtx = LoginPsCLI

	m365 spo term set list --termGroupName "PsCliTermGroup"
	
	m365 logout
}
#gavdcodeend 05

#gavdcodebegin 06
function SpPsCliM365_FindOneTermSet
{
	$spCtx = LoginPsCLI

	m365 spo term set get --termGroupName "PsCliTermGroup" `
						  --name "PsCliTermSet"

	$myTermSet = m365 spo term set get --termGroupId "ba048db1-5337-4795-b662-262af3789994" `
									   --id "777bcaac-309e-4075-94c8-c0ec582a8d16"
	$myTermSetObj = $myTermSet | ConvertFrom-Json
	Write-Host $myTermSetObj.Name
	
	m365 logout
}
#gavdcodeend 06

#gavdcodebegin 07
function SpPsCliM365_CreateTerm
{
	$spCtx = LoginPsCLI

	m365 spo term add --name "PsCliTerm" `
					  --termGroupName "PsCliTermGroup" `
					  --termSetName "PsCliTermSet" `
					  --description "Term description"
	
	m365 logout
}
#gavdcodeend 07

#gavdcodebegin 08
function SpPsCliM365_FindAllTerms
{
	$spCtx = LoginPsCLI

	m365 spo term list --termGroupName "PsCliTermGroup" `
					   --termSetName "PsCliTermSet" `
	
	m365 logout
}
#gavdcodeend 08

#gavdcodebegin 09
function SpPsCliM365_FindOneTermSet
{
	$spCtx = LoginPsCLI

	m365 spo term get --termGroupName "PsCliTermGroup" `
					  --termSetName "PsCliTermSet" `
					  --name "PsCliTerm"

	$myTerm01 = m365 spo term get --termGroupId "ba048db1-5337-4795-b662-262af3789994" `
								  --termSetId "777bcaac-309e-4075-94c8-c0ec582a8d16" `
								  --id "f4da2cc4-0acf-4268-84b8-36b0473abf3c"
	$myTerm01Obj = $myTerm01 | ConvertFrom-Json
	Write-Host $myTerm01Obj.Name

	$myTerm02 = m365 spo term get --id "f4da2cc4-0acf-4268-84b8-36b0473abf3c"
	$myTerm02Obj = $myTerm02 | ConvertFrom-Json
	Write-Host $myTerm02Obj.Name
	
	m365 logout
}
#gavdcodeend 09

#------- Search --------
#gavdcodebegin 10
function SpPsCliM365_Search
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo search --queryText "teams" `
								 --trimDuplicates 
								#--rowLimit 50 --allResults
	$myResultsObj = $myResults | ConvertFrom-Json
	foreach($oneResult in $myResultsObj) {
		Write-Host $oneResult.Title " - " $oneResult.OriginalPath
	}
	
	m365 logout
}
#gavdcodeend 10

#gavdcodebegin 11
function SpPsCliM365_SearchProperties
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo search --queryText "teams" `
								 --trimDuplicates `
								 --selectProperties "Title,ModifiedBy"
	Write-Host $myResults
	
	m365 logout
}
#gavdcodeend 11

#gavdcodebegin 12
function SpPsCliM365_SearchSorting
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo search --queryText "teams" `
								 --selectProperties "Title,ModifiedBy" `
								 --rowLimit 100 `
								 --sortList "ModifiedBy:ascending"
	Write-Host $myResults
	
	m365 logout
}
#gavdcodeend 12

#------- User Profile --------
#gavdcodebegin 13
function SpPsCliM365_GetUserProfile
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo userprofile get --userName $configFile.appsettings.UserName
	Write-Host $myResults
	
	m365 logout
}
#gavdcodeend 13

#gavdcodebegin 14
function SpPsCliM365_UpdateSingleUserProfile
{
	$spCtx = LoginPsCLI

	m365 spo userprofile set --userName $configFile.appsettings.UserName `
							 --propertyName "AboutMe" `
							 --propertyValue "Modified with the CLI"
	
	m365 logout
}
#gavdcodeend 14

#gavdcodebegin 15
function SpPsCliM365_UpdateMultipleUserProfile
{
	$spCtx = LoginPsCLI

	m365 spo userprofile set --userName $configFile.appsettings.UserName `
							 --propertyName "SPS-Skills" `
							 --propertyValue "SharePoint, Development"
	
	m365 logout
}
#gavdcodeend 15

#------- Modern Pages --------
#gavdcodebegin 16
function SpPsCliM365_FindAllPages
{
	$spCtx = LoginPsCLI

	m365 spo page list --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 16

#gavdcodebegin 17
function SpPsCliM365_FindOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page get --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPage.aspx" `
					  --metadataOnly
	
	m365 logout
}
#gavdcodeend 17

#gavdcodebegin 18
function SpPsCliM365_CreateOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page add --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPageCLI.aspx" `
					  --title "My Modern Page CLI" `
					  --description "This is a Modern Page" `
					  --layoutType Article `
					  --commentsEnabled
	
	m365 logout
}
#gavdcodeend 18

#gavdcodebegin 19
function SpPsCliM365_UpdateOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page set --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPageCLI.aspx" `
					  --publishMessage "Page Published 01" `
					  --publish
	
	m365 logout
}
#gavdcodeend 19

#gavdcodebegin 20
function SpPsCliM365_CopyOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page copy --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceName "MyModernPageCLI.aspx" `
					   --targetUrl "MyCLI.aspx" `
					   --overwrite
	
	m365 logout
}
#gavdcodeend 20

#gavdcodebegin 21
function SpPsCliM365_DeleteOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page remove --webUrl $configFile.appsettings.SiteCollUrl `
					     --name "MyModernPageCLI.aspx" `
					     --confirm
	
	m365 logout
}
#gavdcodeend 21

#gavdcodebegin 22
function SpPsCliM365_FindAllSectionsInPage
{
	$spCtx = LoginPsCLI

	m365 spo page section list --webUrl $configFile.appsettings.SiteCollUrl `
							   --name "MyModernPageCLI.aspx"
	
	m365 logout
}
#gavdcodeend 22

#gavdcodebegin 23
function SpPsCliM365_FindOneSectionInPage
{
	$spCtx = LoginPsCLI

	m365 spo page section get --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --section 1
	
	m365 logout
}
#gavdcodeend 23

#gavdcodebegin 24
function SpPsCliM365_CreateOneSectionInPage
{
	$spCtx = LoginPsCLI

	m365 spo page section add --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --sectionTemplate TwoColumn `
							  --order 3 --debug

	#m365 spo page set --webUrl $configFile.appsettings.SiteCollUrl `
	#				  --name "MyModernPageCLI.aspx" `
	#				  --publish
	
	m365 logout
}
#gavdcodeend 24

#gavdcodebegin 25
function SpPsCliM365_FindAllColumnsInSectionPage
{
	$spCtx = LoginPsCLI

	m365 spo page column list --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --section 1
	
	m365 logout
}
#gavdcodeend 25

#gavdcodebegin 26
function SpPsCliM365_FindOneColumnInSectionPage
{
	$spCtx = LoginPsCLI

	m365 spo page column get --webUrl $configFile.appsettings.SiteCollUrl `
							 --name "MyModernPageCLI.aspx" `
							 --section 1 `
							 --column 1
	
	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
function SpPsCliM365_FindAllControlsInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control list --webUrl $configFile.appsettings.SiteCollUrl `
							   --name "MyModernPageCLI.aspx"
	
	m365 logout
}
#gavdcodeend 27

#gavdcodebegin 28
function SpPsCliM365_FindOneControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control get --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --id "0a3acab9-07bb-4cfd-bac9-b640b126e6a8"
	
	m365 logout
}
#gavdcodeend 28

#gavdcodebegin 29
function SpPsCliM365_CreateTextControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page text add --webUrl $configFile.appsettings.SiteCollUrl `
						   --pageName "MyModernPageCLI.aspx" `
						   --section 1 `
	                       --column 1 `
	                       --text "Text in my control"
	
	m365 logout
}
#gavdcodeend 29

#gavdcodebegin 30
function SpPsCliM365_CreateWebPartControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page clientsidewebpart add --webUrl $configFile.appsettings.SiteCollUrl `
										--pageName "MyModernPageCLI.aspx" `
										--standardWebPart "LinkPreview" `
										--section 1 `
										--column 1 `
										--order 2 `
										--webPartProperties '"{""url"":""https://a.com""}"'
	
	m365 logout
}
#gavdcodeend 30

#gavdcodebegin 31
function SpPsCliM365_ModifyControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control set --webUrl $configFile.appsettings.SiteCollUrl `
						      --name "MyModernPageCLI.aspx" `
						      --id "67b2104a-0271-4961-983a-b3fc7f556075" `
	                          --webPartProperties '"{""url"":""https://b.com""}"'
	
	m365 logout
}
#gavdcodeend 31

#gavdcodebegin 32
function SpPsCliM365_ModifyHeaderPage
{
	$spCtx = LoginPsCLI

	m365 spo page header set --webUrl $configFile.appsettings.SiteCollUrl `
						      --pageName "MyModernPageCLI.aspx" `
						      --textAlignment Center `
	                          --showPublishDate
	
	m365 spo page set --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPageCLI.aspx" `
					  --publish

	m365 logout
}
#gavdcodeend 32

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Term Store --------
#SpPsCliM365_CreateTermGroup
#SpPsCliM365_FindAllTermGroups
#SpPsCliM365_FindOneTermGroup
#SpPsCliM365_CreateTermSet
#SpPsCliM365_FindAllTermSets
#SpPsCliM365_FindOneTermSet
#SpPsCliM365_CreateTerm
#SpPsCliM365_FindAllTerms
#SpPsCliM365_FindOneTermSet

#------- Search --------
#SpPsCliM365_Search
#SpPsCliM365_SearchProperties
#SpPsCliM365_SearchSorting

#------- User Profile --------
#SpPsCliM365_GetUserProfile
#SpPsCliM365_UpdateSingleUserProfile
#SpPsCliM365_UpdateMultipleUserProfile

#------- Modern Pages --------
#SpPsCliM365_FindAllPages
#SpPsCliM365_FindOnePage
#SpPsCliM365_CreateOnePage
#SpPsCliM365_UpdateOnePage
#SpPsCliM365_CopyOnePage
#SpPsCliM365_DeleteOnePage
#SpPsCliM365_FindAllSectionsInPage
#SpPsCliM365_FindOneSectionInPage
#SpPsCliM365_CreateOneSectionInPage
#SpPsCliM365_FindAllColumnsInSectionPage
#SpPsCliM365_FindOneColumnInSectionPage
#SpPsCliM365_FindAllControlsInPage
#SpPsCliM365_FindOneControlInPage
#SpPsCliM365_CreateTextControlInPage
#SpPsCliM365_CreateWebPartControlInPage
#SpPsCliM365_ModifyControlInPage
#SpPsCliM365_ModifyHeaderPage

Write-Host "Done" 
