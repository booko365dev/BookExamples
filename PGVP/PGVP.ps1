
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
function SpPsCliM365_CreateListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem add --contentType "Item" `
						  --listTitle "TestList" `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --Title "ItemFromSpPsCli"

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsCliM365_GetAllListItemsInList
{
	$spCtx = LoginPsCLI

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsCliM365_GetAllListItemsInListSelectFields
{
	$spCtx = LoginPsCLI

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --fields "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsCliM365_GetOneListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem get --listTitle "TestList" `
						  --id 6 `
						  --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsCliM365_GetOneListItemSelectFields
{
	$spCtx = LoginPsCLI

	m365 spo listitem get --listTitle "TestList" `
						  --id 6 `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --properties "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsCliM365_GetOneListItemByCAML
{
	$spCtx = LoginPsCLI
	
	$myCAML = "<View>`
				<Query>`
					<Where>`
						<Eq>`
							<FieldRef Name='Title' />`
							<Value Type='Text'>ItemFromSpPsCli</Value>`
						</Eq>`
					</Where>`
				</Query>`
			</View>"

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --camlQuery $myCAML
	
	m365 logout
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsCliM365_GetOneListItemByFilter
{
	$spCtx = LoginPsCLI

	$myFilter = "Title eq 'ItemFromSpPsCli'"

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --filter $myFilter
	
	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsCliM365_UpdateOneListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem set --contentType "Item" `
						  --listTitle "TestList" `
						  --id 6 `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --Title "ItemFromSpPsCliUpdated"

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsCliM365_DeleteOneListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem remove --listTitle "TestList" `
							 --id 6 `
							 --webUrl $configFile.appsettings.SiteCollUrl `
							 --confirm

	m365 logout
}
#gavdcodeend 009

#gavdcodebegin 010
function SpPsCliM365_UploadOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx"

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsCliM365_UploadOneDocumentByRelativePath
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "/Sites/Test_Guitaca/TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx"

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsCliM365_UploadOneDocumentAndChangeFieldValue
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx" `
					  --Title "DocumentFromSpPsCli"

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function SpPsCliM365_GetAllDocumentsInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo file list --webUrl $configFile.appsettings.SiteCollUrl `
					   --folder "TestLibrary"

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsCliM365_GetOneDocumentProperties
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0"

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsCliM365_GetOneDocumentPropertiesAsListItem
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "69e3efeb-45e3-4d81-8aba-7a3dc35d648f" `
					  --asListItem

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function SpPsCliM365_GetOneDocumentPropertiesAsString
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0" `
					  --asString

	m365 logout
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsCliM365_DownloadOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0" `
					  --asFile `
					  --path "C:\Temporary\TestDocument02.docx"

	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
function SpPsCliM365_UpdateOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file rename --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/TestDocument01.docx" `
						 --targetFileName "TestDocument01Updated.docx" `
						 --force

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsCliM365_CopyOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file copy --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument01.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsCliM365_MoveOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file move --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument01.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
function SpPsCliM365_DeleteOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file remove --webUrl $configFile.appsettings.SiteCollUrl `
						 --url "/TestLibrary/TestDocument01.docx"

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 022
function SpPsCliM365_CreateFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder add --webUrl $configFile.appsettings.SiteCollUrl `
						--parentFolderUrl "/TestLibrary" `
						--name 'NewFolderCli'

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
function SpPsCliM365_GetFoldersInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder list --webUrl $configFile.appsettings.SiteCollUrl `
						 --parentFolderUrl "/TestLibrary"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
function SpPsCliM365_GetOneFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder get --webUrl $configFile.appsettings.SiteCollUrl `
						--folderUrl "/TestLibrary"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 025
function SpPsCliM365_RenameOneFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder rename --webUrl $configFile.appsettings.SiteCollUrl `
						   --folderUrl "/TestLibrary/NewFolderCli" `
						   --name "NewFolderCliUpdated"

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function SpPsCliM365_CopyOneFolderToOtherLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder copy --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function SpPsCliM365_MoveOneFolderToOtherLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder move --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function SpPsCliM365_DeleteOneFolderFromLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder remove --webUrl $configFile.appsettings.SiteCollUrl `
						   --folderUrl "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
function SpPsCliM365_GetAttachementsInItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem attachment list --webUrl $configFile.appsettings.SiteCollUrl `
									  --listTitle "TestList" `
									  --itemId 8

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
function SpPsCliM365_BreakInheritanceItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem roleinheritance break --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 8

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
function SpPsCliM365_RestoreInheritanceItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem roleinheritance reset --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 8

	m365 logout
}
#gavdcodeend 031


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 31 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#SpPsCliM365_CreateListItem
#SpPsCliM365_GetAllListItemsInList
#SpPsCliM365_GetAllListItemsInListSelectFields
#SpPsCliM365_GetOneListItem
#SpPsCliM365_GetOneListItemSelectFields
#SpPsCliM365_GetOneListItemByCAML
#SpPsCliM365_GetOneListItemByFilter
#SpPsCliM365_UpdateOneListItem
#SpPsCliM365_DeleteOneListItem
#SpPsCliM365_UploadOneDocument
#SpPsCliM365_UploadOneDocumentByRelativePath
#SpPsCliM365_UploadOneDocumentAndChangeFieldValue
#SpPsCliM365_GetAllDocumentsInLibrary
#SpPsCliM365_GetOneDocumentProperties
#SpPsCliM365_GetOneDocumentPropertiesAsListItem
#SpPsCliM365_GetOneDocumentPropertiesAsString
#SpPsCliM365_DownloadOneDocument
#SpPsCliM365_UpdateOneDocument
#SpPsCliM365_CopyOneDocument
#SpPsCliM365_MoveOneDocument
#SpPsCliM365_DeleteOneDocument
#SpPsCliM365_CreateFolderInLibrary
#SpPsCliM365_GetFoldersInLibrary
#SpPsCliM365_GetOneFolderInLibrary
#SpPsCliM365_RenameOneFolderInLibrary
#SpPsCliM365_CopyOneFolderToOtherLibrary
#SpPsCliM365_MoveOneFolderToOtherLibrary
#SpPsCliM365_DeleteOneFolderFromLibrary
#SpPsCliM365_GetAttachementsInItem
#SpPsCliM365_BreakInheritanceItem
#SpPsCliM365_RestoreInheritanceItem

Write-Host "Done" 
