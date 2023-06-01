
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
#gavdcodebegin 001
function SpPsCliM365_CreateTermGroup
{
	$spCtx = LoginPsCLI

	m365 spo term group add --name "PsCliTermGroup" `
							--description "Group description"


	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsCliM365_FindAllTermGroups
{
	$spCtx = LoginPsCLI

	m365 spo term group list
	
	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsCliM365_FindOneTermGroup
{
	$spCtx = LoginPsCLI

	m365 spo term group get --name "PsCliTermGroup"

	$myTermGroup = m365 spo term group get --id "ba048db1-5337-4795-b662-262af3789994"
	$myTermGroupObj = $myTermGroup | ConvertFrom-Json
	Write-Host $myTermGroupObj.Name
	
	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsCliM365_CreateTermSet
{
	$spCtx = LoginPsCLI

	m365 spo term set add --name "PsCliTermSet" `
						  --termGroupName "PsCliTermGroup" `
						  --description "Set description"

	
	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsCliM365_FindAllTermSets
{
	$spCtx = LoginPsCLI

	m365 spo term set list --termGroupName "PsCliTermGroup"
	
	m365 logout
}
#gavdcodeend 005

#gavdcodebegin 006
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
#gavdcodeend 006

#gavdcodebegin 007
function SpPsCliM365_CreateTerm
{
	$spCtx = LoginPsCLI

	m365 spo term add --name "PsCliTerm" `
					  --termGroupName "PsCliTermGroup" `
					  --termSetName "PsCliTermSet" `
					  --description "Term description"
	
	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsCliM365_FindAllTerms
{
	$spCtx = LoginPsCLI

	m365 spo term list --termGroupName "PsCliTermGroup" `
					   --termSetName "PsCliTermSet" `
	
	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
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
#gavdcodeend 009

#------- Search --------
#gavdcodebegin 010
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
#gavdcodeend 010

#gavdcodebegin 011
function SpPsCliM365_SearchProperties
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo search --queryText "teams" `
								 --trimDuplicates `
								 --selectProperties "Title,ModifiedBy"
	Write-Host $myResults
	
	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
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
#gavdcodeend 012

#------- User Profile --------
#gavdcodebegin 013
function SpPsCliM365_GetUserProfile
{
	$spCtx = LoginPsCLI

	$myResults = m365 spo userprofile get --userName $configFile.appsettings.UserName
	Write-Host $myResults
	
	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsCliM365_UpdateSingleUserProfile
{
	$spCtx = LoginPsCLI

	m365 spo userprofile set --userName $configFile.appsettings.UserName `
							 --propertyName "AboutMe" `
							 --propertyValue "Modified with the CLI"
	
	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsCliM365_UpdateMultipleUserProfile
{
	$spCtx = LoginPsCLI

	m365 spo userprofile set --userName $configFile.appsettings.UserName `
							 --propertyName "SPS-Skills" `
							 --propertyValue "SharePoint, Development"
	
	m365 logout
}
#gavdcodeend 015

#------- Modern Pages --------
#gavdcodebegin 016
function SpPsCliM365_FindAllPages
{
	$spCtx = LoginPsCLI

	m365 spo page list --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsCliM365_FindOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page get --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPage.aspx" `
					  --metadataOnly
	
	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
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
#gavdcodeend 018

#gavdcodebegin 019
function SpPsCliM365_UpdateOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page set --webUrl $configFile.appsettings.SiteCollUrl `
					  --name "MyModernPageCLI.aspx" `
					  --publishMessage "Page Published 01" `
					  --publish
	
	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsCliM365_CopyOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page copy --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceName "MyModernPageCLI.aspx" `
					   --targetUrl "MyCLI.aspx" `
					   --overwrite
	
	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
function SpPsCliM365_DeleteOnePage
{
	$spCtx = LoginPsCLI

	m365 spo page remove --webUrl $configFile.appsettings.SiteCollUrl `
					     --name "MyModernPageCLI.aspx" `
					     --confirm
	
	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 022
function SpPsCliM365_FindAllSectionsInPage
{
	$spCtx = LoginPsCLI

	m365 spo page section list --webUrl $configFile.appsettings.SiteCollUrl `
							   --name "MyModernPageCLI.aspx"
	
	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
function SpPsCliM365_FindOneSectionInPage
{
	$spCtx = LoginPsCLI

	m365 spo page section get --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --section 1
	
	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
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
#gavdcodeend 024

#gavdcodebegin 025
function SpPsCliM365_FindAllColumnsInSectionPage
{
	$spCtx = LoginPsCLI

	m365 spo page column list --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --section 1
	
	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function SpPsCliM365_FindOneColumnInSectionPage
{
	$spCtx = LoginPsCLI

	m365 spo page column get --webUrl $configFile.appsettings.SiteCollUrl `
							 --name "MyModernPageCLI.aspx" `
							 --section 1 `
							 --column 1
	
	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function SpPsCliM365_FindAllControlsInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control list --webUrl $configFile.appsettings.SiteCollUrl `
							   --name "MyModernPageCLI.aspx"
	
	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function SpPsCliM365_FindOneControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control get --webUrl $configFile.appsettings.SiteCollUrl `
							  --name "MyModernPageCLI.aspx" `
							  --id "0a3acab9-07bb-4cfd-bac9-b640b126e6a8"
	
	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
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
#gavdcodeend 029

#gavdcodebegin 030
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
#gavdcodeend 030

#gavdcodebegin 031
function SpPsCliM365_ModifyControlInPage
{
	$spCtx = LoginPsCLI

	m365 spo page control set --webUrl $configFile.appsettings.SiteCollUrl `
						      --name "MyModernPageCLI.aspx" `
						      --id "67b2104a-0271-4961-983a-b3fc7f556075" `
	                          --webPartProperties '"{""url"":""https://b.com""}"'
	
	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
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
#gavdcodeend 032

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
