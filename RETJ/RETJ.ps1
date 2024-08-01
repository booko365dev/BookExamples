
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpCliM365_LoginWithAccPw
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCliM365_CreateListWithParameters
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list add --title "NewListPsCli" `
					  --baseTemplate GenericList `
					  --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpCliM365_CreateListWithSchema
{
	$spCtx = PsSpCliM365_LoginWithAccPw

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
function PsSpCliM365_GetAllLists
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list list --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpCliM365_GetOneList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list get --id "97ee9e5c-a332-468e-bfcd-252b1cee3b5c" `
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
function PsSpCliM365_UpdateOneList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list set --id "97ee9e5c-a332-468e-bfcd-252b1cee3b5c" `
					  --webUrl $configFile.appsettings.SiteCollUrl `
					  --description "List updated"
	
	m365 logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpCliM365_DeleteOneList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list remove --title "NewListPsCli" `
						 --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpCliM365_GetAllFieldsInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo field list --listTitle "NewListPsCli" `
					    --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpCliM365_GetOneFieldInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo field get --listTitle "NewListPsCli" `
					   --fieldTitle "Title" `
					   --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpCliM365_CreateOneFieldInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

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
function PsSpCliM365_UpdateOneFieldInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo field set --listTitle "NewListPsCli" `
					   --title "MyMultilineField" `
					   --Description "New Field Description" `
					   --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpCliM365_DeleteOneFieldFromList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo field remove --listTitle "NewListPsCli" `
						  --title "MyMultilineField" `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --force

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpCliM365_BreakInheritanceList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list roleinheritance break --listTitle "NewListPsCli" `
										--webUrl $configFile.appsettings.SiteCollUrl `
										--clearExistingPermissions

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpCliM365_RestoreInheritanceList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list roleinheritance reset --listTitle "NewListPsCli" `
										--webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpCliM365_GetAllRoledefinition
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo roledefinition list --webUrl $configFile.appsettings.SiteCollUrl

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpCliM365_GetOneRoledefinition
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo roledefinition get --webUrl $configFile.appsettings.SiteCollUrl `
								--id "1073741924"

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpCliM365_AddRoledefinition
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo roledefinition add --webUrl $configFile.appsettings.SiteCollUrl `
								--name "NewRoleDefinitionFromM365" `
								--description "Is is my description" `
					--rights "ViewListItems,AddListItems,EditListItems,DeleteListItems"

	m365 logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpCliM365_RemoveRoledefinition
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo roledefinition remove --webUrl $configFile.appsettings.SiteCollUrl `
								--id "1073741924" `
								--force

	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpCliM365_AssignRoledefinitionToListAndUser
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list roleassignment add --webUrl $configFile.appsettings.SiteCollUrl `
									 --listTitle "NewListPsCli" `
									 --upn $configFile.appsettings.UserName `
									 --roleDefinitionId "1073741924"

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpCliM365_RemoveRoledefinitionFromListForUser
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list roleassignment remove --webUrl $configFile.appsettings.SiteCollUrl `
										--listTitle "NewListPsCli" `
										--upn $configFile.appsettings.UserName

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpCliM365_GetAllContentTypesList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list contenttype list --webUrl $configFile.appsettings.SiteCollUrl `
								   --listTitle "NewListPsCli"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpCliM365_AddOneContentTypeToList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list contenttype add --webUrl $configFile.appsettings.SiteCollUrl `
								  --listTitle "NewListPsCli" `
								  --id "0x010101"

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpCliM365_SetDefaultContentTypeInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list contenttype default set --webUrl $configFile.appsettings.SiteCollUrl `
										  --listTitle "NewListPsCli" `
										  --contentTypeId "0x010101"

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
function PsSpCliM365_DeleteContentTypeFromList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list contenttype remove --webUrl $configFile.appsettings.SiteCollUrl `
								     --listTitle "NewListPsCli" `
								     --id "0x010101"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
function PsSpCliM365_GetAllViewsInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view list --webUrl $configFile.appsettings.SiteCollUrl `
								     --listTitle "NewListPsCli"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 025
function PsSpCliM365_GetOneViewInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view get --webUrl $configFile.appsettings.SiteCollUrl `
						   --listTitle "NewListPsCli" `
						   --title "All Items"

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function PsSpCliM365_AddOneViewToList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view add --webUrl $configFile.appsettings.SiteCollUrl `
						   --listTitle "NewListPsCli" `
						   --title "My View" `
						   --fields "Title,MyMultilineField" `
						   --paged --default

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function PsSpCliM365_UpdateOneViewInList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view set --webUrl $configFile.appsettings.SiteCollUrl `
						   --listTitle "NewListPsCli" `
						   --title "My View" `
						   --Title "My View Updated"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function PsSpCliM365_AddOneFieldToViewForList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view field add --webUrl $configFile.appsettings.SiteCollUrl `
								 --listTitle "NewListPsCli" `
								 --viewTitle "My View Updated" `
								 --title "My Text Field"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
function PsSpCliM365_UpdateOneFieldForViewForList
{
	$spCtx = PsSpCliM365_LoginWithAccPw


	m365 spo list view field set --webUrl $configFile.appsettings.SiteCollUrl `
								 --listTitle "NewListPsCli" `
								 --viewTitle "My View Updated" `
								 --title "My Text Field" `
								 --position 1

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
function PsSpCliM365_DeleteOneFieldFromViewForList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view field remove --webUrl $configFile.appsettings.SiteCollUrl `
									--listTitle "NewListPsCli" `
									--viewTitle "My View Updated" `
									--title "My Text Field"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
function PsSpCliM365_DeleteOneViewFromList
{
	$spCtx = PsSpCliM365_LoginWithAccPw

	m365 spo list view remove --webUrl $configFile.appsettings.SiteCollUrl `
									--listTitle "NewListPsCli" `
									--title "My View Updated"

	m365 logout
}
#gavdcodeend 031


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 031 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#PsSpCliM365_CreateListWithParameters
#PsSpCliM365_CreateListWithSchema
#PsSpCliM365_GetAllLists
#PsSpCliM365_GetOneList
#PsSpCliM365_UpdateOneList
#PsSpCliM365_DeleteOneList
#PsSpCliM365_GetAllFieldsInList
#PsSpCliM365_GetOneFieldInList
#PsSpCliM365_CreateOneFieldInList
#PsSpCliM365_UpdateOneFieldInList
#PsSpCliM365_DeleteOneFieldFromList
#PsSpCliM365_BreakInheritanceList
#PsSpCliM365_RestoreInheritanceList
#PsSpCliM365_GetAllRoledefinition
#PsSpCliM365_GetOneRoledefinition
#PsSpCliM365_AddRoledefinition
#PsSpCliM365_AssignRoledefinitionToListAndUser
#PsSpCliM365_RemoveRoledefinitionFromListForUser
#PsSpCliM365_GetAllContentTypesList
#PsSpCliM365_AddOneContentTypeToList
#PsSpCliM365_SetDefaultContentTypeInList
#PsSpCliM365_DeleteContentTypeFromList
#PsSpCliM365_GetAllViewsInList
#PsSpCliM365_GetOneViewInList
#PsSpCliM365_AddOneViewToList
#PsSpCliM365_UpdateOneViewInList
#PsSpCliM365_AddOneFieldToViewForList
#PsSpCliM365_UpdateOneFieldForViewForList
#PsSpCliM365_DeleteOneFieldFromViewForList
#PsSpCliM365_DeleteOneViewFromList

Write-Host "Done" 
