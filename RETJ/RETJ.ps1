
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

#gavdcodebegin 001
function SpPsCliM365_CreateListWithParameters
{
	$spCtx = LoginPsCLI

	m365 spo list add --title "NewListPsCli" `
					  --baseTemplate GenericList `
					  --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsCliM365_CreateListWithSchema
{
	$spCtx = LoginPsCLI

	$listSchema = "<List xmlns:ows='Microsoft SharePoint' `
						Title='NewListPsCliWithSchema' `
						Url='Lists/NewListPsCliWithSchema' `
						BaseType='0' `
						xmlns='http://schemas.microsoft.com/sharepoint/'> `
					<MetaData><ContentTypes></ContentTypes>`
							  <Fields>[fields definition]</Fields>`
							  <Views>[vies definition]</Views>`
							  <Forms>[forms definition]</Forms>`
					</MetaData></List>"

	m365 spo list add --title "NewListPsCliWithSchema" `
					  --baseTemplate GenericList `
					  --webUrl $configFile.appsettings.SiteCollUrl `
					  --schemaXml $listSchema
	
	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsCliM365_GetAllLists
{
	$spCtx = LoginPsCLI

	m365 spo list list --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsCliM365_GetOneList
{
	$spCtx = LoginPsCLI

	m365 spo list get --id "cb3841d2-6561-452c-bbaf-08338bfa0029" `
					  --webUrl $configFile.appsettings.SiteCollUrl
#	m365 spo list get --title "NewListPsCli" `
#					  --webUrl $configFile.appsettings.SiteCollUrl
#	m365 spo list get --title "NewListPsCli" `
#					  --webUrl $configFile.appsettings.SiteCollUrl `
#					  --properties "Title,Id"
	
	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsCliM365_UpdateOneList
{
	$spCtx = LoginPsCLI

	m365 spo list set --id "bf830e20-8288-45c8-901f-68e14ac61075" `
					  --webUrl $configFile.appsettings.SiteCollUrl `
					  --description "List updated"
	
	m365 logout
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsCliM365_DeleteOneList
{
	$spCtx = LoginPsCLI

	m365 spo list remove --id "bf830e20-8288-45c8-901f-68e14ac61075" `
						 --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsCliM365_GetAllFieldsInList
{
	$spCtx = LoginPsCLI

	m365 spo field list --listTitle "NewListPsCli" `
					    --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsCliM365_GetOneFieldInList
{
	$spCtx = LoginPsCLI

	m365 spo field get --listTitle "NewListPsCli" `
					   --fieldTitle "Title" `
					   --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsCliM365_CreateOneFieldInList
{
	$spCtx = LoginPsCLI

	$fieldXml = `
			"<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	m365 spo field add --listTitle "NewListPsCli" `
					   --xml $fieldXml `
					   --webUrl $configFile.appsettings.SiteCollUrl `
					   --options AddToAllContentTypes

	m365 logout
}
#gavdcodeend 009

#gavdcodebegin 010
function SpPsCliM365_UpdateOneFieldInList
{
	$spCtx = LoginPsCLI

	m365 spo field set --listTitle "NewListPsCli" `
					   --name "MyMultilineField" `
					   --Description "New Field Description" `
					   --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsCliM365_DeleteOneFieldFromList
{
	$spCtx = LoginPsCLI

	m365 spo field remove --listTitle "NewListPsCli" `
						  --fieldTitle "MyMultilineField" `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --confirm

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsCliM365_BreakInheritanceList
{
	$spCtx = LoginPsCLI

	m365 spo list roleinheritance break --listTitle "NewListPsCli" `
										--webUrl $configFile.appsettings.SiteCollUrl `
										--clearExistingPermissions

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function SpPsCliM365_RestoreInheritanceList
{
	$spCtx = LoginPsCLI

	m365 spo list roleinheritance reset --listTitle "NewListPsCli" `
										--webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsCliM365_GetAllRoledefinition
{
	$spCtx = LoginPsCLI

	m365 spo roledefinition list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 014


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#SpPsCliM365_CreateListWithParameters
#SpPsCliM365_CreateListWithSchema
#SpPsCliM365_GetAllLists
#SpPsCliM365_GetOneList
#SpPsCliM365_UpdateOneList
#SpPsCliM365_DeleteOneList
#SpPsCliM365_GetAllFieldsInList
#SpPsCliM365_GetOneFieldInList
#SpPsCliM365_CreateOneFieldInList
#SpPsCliM365_UpdateOneFieldInList
#SpPsCliM365_DeleteOneFieldFromList
#SpPsCliM365_BreakInheritanceList
#SpPsCliM365_RestoreInheritanceList
#SpPsCliM365_GetAllRoledefinition

Write-Host "Done" 
