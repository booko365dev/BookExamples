
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

#gavdcodebegin 01
function SpPsCliM365_CreateListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem add --contentType "Item" `
						  --listTitle "TestList" `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --Title "ItemFromSpPsCli"

	m365 logout
}
#gavdcodeend 01

#gavdcodebegin 02
function SpPsCliM365_GetAllListItemsInList
{
	$spCtx = LoginPsCLI

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 02

#gavdcodebegin 03
function SpPsCliM365_GetAllListItemsInListSelectFields
{
	$spCtx = LoginPsCLI

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --fields "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 03

#gavdcodebegin 04
function SpPsCliM365_GetOneListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem get --listTitle "TestList" `
						  --id 6 `
						  --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 04

#gavdcodebegin 05
function SpPsCliM365_GetOneListItemSelectFields
{
	$spCtx = LoginPsCLI

	m365 spo listitem get --listTitle "TestList" `
						  --id 6 `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --properties "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 05

#gavdcodebegin 06
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
#gavdcodeend 06

#gavdcodebegin 07
function SpPsCliM365_GetOneListItemByFilter
{
	$spCtx = LoginPsCLI

	$myFilter = "Title eq 'ItemFromSpPsCli'"

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --filter $myFilter
	
	m365 logout
}
#gavdcodeend 07

#gavdcodebegin 08
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
#gavdcodeend 08

#gavdcodebegin 09
function SpPsCliM365_DeleteOneListItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem remove --listTitle "TestList" `
							 --id 6 `
							 --webUrl $configFile.appsettings.SiteCollUrl `
							 --confirm

	m365 logout
}
#gavdcodeend 09

#gavdcodebegin 10
function SpPsCliM365_UploadOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx"

	m365 logout
}
#gavdcodeend 10

#gavdcodebegin 11
function SpPsCliM365_UploadOneDocumentByRelativePath
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "/Sites/Test_Guitaca/TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx"

	m365 logout
}
#gavdcodeend 11

#gavdcodebegin 12
function SpPsCliM365_UploadOneDocumentAndChangeFieldValue
{
	$spCtx = LoginPsCLI

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument01.docx" `
					  --Title "DocumentFromSpPsCli"

	m365 logout
}
#gavdcodeend 12

#gavdcodebegin 13
function SpPsCliM365_GetAllDocumentsInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo file list --webUrl $configFile.appsettings.SiteCollUrl `
					   --folder "TestLibrary"

	m365 logout
}
#gavdcodeend 13

#gavdcodebegin 14
function SpPsCliM365_GetOneDocumentProperties
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0"

	m365 logout
}
#gavdcodeend 14

#gavdcodebegin 15
function SpPsCliM365_GetOneDocumentPropertiesAsListItem
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "69e3efeb-45e3-4d81-8aba-7a3dc35d648f" `
					  --asListItem

	m365 logout
}
#gavdcodeend 15

#gavdcodebegin 16
function SpPsCliM365_GetOneDocumentPropertiesAsString
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0" `
					  --asString

	m365 logout
}
#gavdcodeend 16

#gavdcodebegin 17
function SpPsCliM365_DownloadOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "ad0beee5-d5d0-4ba3-86ff-fa0a0e8f06e0" `
					  --asFile `
					  --path "C:\Temporary\TestDocument02.docx"

	m365 logout
}
#gavdcodeend 17

#gavdcodebegin 18
function SpPsCliM365_UpdateOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file rename --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/TestDocument01.docx" `
						 --targetFileName "TestDocument01Updated.docx" `
						 --force

	m365 logout
}
#gavdcodeend 18

#gavdcodebegin 19
function SpPsCliM365_CopyOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file copy --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument01.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 19

#gavdcodebegin 20
function SpPsCliM365_MoveOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file move --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument01.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 20

#gavdcodebegin 21
function SpPsCliM365_DeleteOneDocument
{
	$spCtx = LoginPsCLI

	m365 spo file remove --webUrl $configFile.appsettings.SiteCollUrl `
						 --url "/TestLibrary/TestDocument01.docx"

	m365 logout
}
#gavdcodeend 21

#gavdcodebegin 22
function SpPsCliM365_CreateFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder add --webUrl $configFile.appsettings.SiteCollUrl `
						--parentFolderUrl "/TestLibrary" `
						--name 'NewFolderCli'

	m365 logout
}
#gavdcodeend 22

#gavdcodebegin 23
function SpPsCliM365_GetFoldersInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder list --webUrl $configFile.appsettings.SiteCollUrl `
						 --parentFolderUrl "/TestLibrary"

	m365 logout
}
#gavdcodeend 23

#gavdcodebegin 24
function SpPsCliM365_GetOneFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder get --webUrl $configFile.appsettings.SiteCollUrl `
						--folderUrl "/TestLibrary"

	m365 logout
}
#gavdcodeend 24

#gavdcodebegin 25
function SpPsCliM365_RenameOneFolderInLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder rename --webUrl $configFile.appsettings.SiteCollUrl `
						   --folderUrl "/TestLibrary/NewFolderCli" `
						   --name "NewFolderCliUpdated"

	m365 logout
}
#gavdcodeend 25

#gavdcodebegin 26
function SpPsCliM365_CopyOneFolderToOtherLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder copy --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 26

#gavdcodebegin 27
function SpPsCliM365_MoveOneFolderToOtherLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder move --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 27

#gavdcodebegin 28
function SpPsCliM365_DeleteOneFolderFromLibrary
{
	$spCtx = LoginPsCLI

	m365 spo folder remove --webUrl $configFile.appsettings.SiteCollUrl `
						   --folderUrl "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 28

#gavdcodebegin 29
function SpPsCliM365_GetAttachementsInItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem attachment list --webUrl $configFile.appsettings.SiteCollUrl `
									  --listTitle "TestList" `
									  --itemId 8

	m365 logout
}
#gavdcodeend 29

#gavdcodebegin 30
function SpPsCliM365_BreakInheritanceItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem roleinheritance break --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 8

	m365 logout
}
#gavdcodeend 30

#gavdcodebegin 31
function SpPsCliM365_RestoreInheritanceItem
{
	$spCtx = LoginPsCLI

	m365 spo listitem roleinheritance reset --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 8

	m365 logout
}
#gavdcodeend 31


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

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
