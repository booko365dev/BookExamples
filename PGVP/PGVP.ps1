
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpCliM365_Login
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCliM365_CreateListItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem add --contentType "Item" `
						  --listTitle "TestList" `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --Title "ItemFromSpPsCli"

	m365 logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpCliM365_GetAllListItemsInList
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpCliM365_GetAllListItemsInListSelectFields
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --fields "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpCliM365_GetOneListItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem get --listTitle "TestList" `
						  --id 9 `
						  --webUrl $configFile.appsettings.SiteCollUrl
	
	m365 logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpCliM365_GetOneListItemSelectFields
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem get --listTitle "TestList" `
						  --id 9 `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --properties "ID,Title,Modified"
	
	m365 logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpCliM365_GetOneListItemByCAML
{
	$spCtx = PsSpCliM365_Login
	
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
function PsSpCliM365_GetOneListItemByFilter
{
	$spCtx = PsSpCliM365_Login

	$myFilter = "Title eq 'ItemFromSpPsCli'"

	m365 spo listitem list --listTitle "TestList" `
						   --webUrl $configFile.appsettings.SiteCollUrl `
						   --filter $myFilter
	
	m365 logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpCliM365_UpdateOneListItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem set --contentType "Item" `
						  --listTitle "TestList" `
						  --id 9 `
						  --webUrl $configFile.appsettings.SiteCollUrl `
						  --Title "ItemFromSpPsCliUpdated"

	m365 logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpCliM365_DeleteOneListItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem remove --listTitle "TestList" `
							 --id 9 `
							 --webUrl $configFile.appsettings.SiteCollUrl `
							 --force

	m365 logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpCliM365_UploadOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument.docx"

	m365 logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpCliM365_UploadOneDocumentByRelativePath
{
	$spCtx = PsSpCliM365_Login

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "/Sites/Test_Guitaca/TestLibrary" `
					  --path "C:\Temporary\TestDocument.docx"

	m365 logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpCliM365_UploadOneDocumentAndChangeFieldValue
{
	$spCtx = PsSpCliM365_Login

	m365 spo file add --webUrl $configFile.appsettings.SiteCollUrl `
					  --folder "TestLibrary" `
					  --path "C:\Temporary\TestDocument.docx" `
					  --Title "DocumentFromSpPsCli"

	m365 logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpCliM365_GetAllDocumentsInLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo file list --webUrl $configFile.appsettings.SiteCollUrl `
					   --folderUrl "TestLibrary"

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpCliM365_GetOneDocumentProperties
{
	$spCtx = PsSpCliM365_Login

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "c0b79c1b-c642-4b64-bbbb-ff89ed4365bb"

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpCliM365_GetOneDocumentPropertiesAsListItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "c0b79c1b-c642-4b64-bbbb-ff89ed4365bb" `
					  --asListItem

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsSpCliM365_GetOneDocumentPropertiesAsString
{
	$spCtx = PsSpCliM365_Login

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "c0b79c1b-c642-4b64-bbbb-ff89ed4365bb" `
					  --asString

	m365 logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsSpCliM365_DownloadOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file get --webUrl $configFile.appsettings.SiteCollUrl `
					  --id "c0b79c1b-c642-4b64-bbbb-ff89ed4365bb" `
					  --asFile `
					  --path "C:\Temporary\TestDocument01.docx"

	m365 logout
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpCliM365_UpdateOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file rename --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/TestDocument.docx" `
						 --targetFileName "TestDocument01Updated.docx" `
						 --force

	m365 logout
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpCliM365_CopyOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file copy --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpCliM365_MoveOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file move --webUrl $configFile.appsettings.SiteCollUrl `
					   --sourceUrl "/TestLibrary/TestDocument.docx" `
					   --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpCliM365_DeleteOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file remove --webUrl $configFile.appsettings.SiteCollUrl `
						 --url "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 036
function PsSpCliM365_CheckoutOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file checkout --webUrl $configFile.appsettings.SiteCollUrl `
						   --url "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 036

#gavdcodebegin 037
function PsSpCliM365_CheckoutUndoOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file checkout undo --webUrl $configFile.appsettings.SiteCollUrl `
								--fileUrl "/TestLibrary/TestDocument.docx" `
								--force

	m365 logout
}
#gavdcodeend 037

#gavdcodebegin 038
function PsSpCliM365_CheckinOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file checkin --webUrl $configFile.appsettings.SiteCollUrl `
						  --url "/TestLibrary/TestDocument.docx" `
						  --type "Major" `
						  --comment "Checkin from CLI"

	m365 logout
}
#gavdcodeend 038

#gavdcodebegin 022
function PsSpCliM365_CreateFolderInLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder add --webUrl $configFile.appsettings.SiteCollUrl `
						--parentFolderUrl "/TestLibrary" `
						--name "NewFolderCli"

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
function PsSpCliM365_GetFoldersInLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder list --webUrl $configFile.appsettings.SiteCollUrl `
						 --parentFolderUrl "/TestLibrary"

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
function PsSpCliM365_GetOneFolderInLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder get --webUrl $configFile.appsettings.SiteCollUrl `
						--url "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 025
function PsSpCliM365_RenameOneFolderInLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder set --webUrl $configFile.appsettings.SiteCollUrl `
						--url "/TestLibrary/NewFolderCli" `
						--name "NewFolderCliUpdated"

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function PsSpCliM365_CopyOneFolderToOtherLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder copy --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function PsSpCliM365_MoveOneFolderToOtherLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder move --webUrl $configFile.appsettings.SiteCollUrl `
						 --sourceUrl "/TestLibrary/NewFolderCli" `
						 --targetUrl "/sites/Test_Guitaca/OtherTestLibrary/"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function PsSpCliM365_DeleteOneFolderFromLibrary
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder remove --webUrl $configFile.appsettings.SiteCollUrl `
						   --url "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 032
function PsSpCliM365_AddAttachementsToItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem attachment add --webUrl $configFile.appsettings.SiteCollUrl `
									 --listTitle "TestList" `
									 --listItemId  9 `
									 --filePath "C:\Temporary\TestDocument.docx"

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 029
function PsSpCliM365_GetAllAttachementsInItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem attachment list --webUrl $configFile.appsettings.SiteCollUrl `
									  --listTitle "TestList" `
									  --listItemId  9

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 033
function PsSpCliM365_GetOneAttachementsInItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem attachment get --webUrl $configFile.appsettings.SiteCollUrl `
									  --listTitle "TestList" `
									  --listItemId  9 `
									  --fileName "TestDocument.docx"

	m365 logout
}
#gavdcodeend 033

#gavdcodebegin 034
function PsSpCliM365_UpdateAttachementsInItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem attachment set --webUrl $configFile.appsettings.SiteCollUrl `
									 --listTitle "TestList" `
									 --listItemId  9 `
									 --filePath "C:\Temporary\TestDocument.docx" `
									 --fileName "TestDocument.docx"

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 035
function PsSpCliM365_DeleteAttachementsInItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem attachment remove --webUrl $configFile.appsettings.SiteCollUrl `
										--listTitle "TestList" `
										--listItemId  9 `
										--fileName "TestDocument.docx"

	m365 logout
}
#gavdcodeend 035

#gavdcodebegin 039
function PsSpCliM365_GetAllSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink list --webUrl $configFile.appsettings.SiteCollUrl `
								   --fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 039

#gavdcodebegin 040
function PsSpCliM365_GetOneSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink get --webUrl $configFile.appsettings.SiteCollUrl `
								  --fileUrl "/TestLibrary/TestDocument.docx" `
								  --id "503e3411-478f-4196-8650-84d293613237"

	m365 logout
}
#gavdcodeend 040

#gavdcodebegin 041
function PsSpCliM365_AddSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink add --webUrl $configFile.appsettings.SiteCollUrl `
								  --fileUrl "/TestLibrary/TestDocument.docx" `
								  --type "view" `
								  --scope "organization"

	m365 logout
}
#gavdcodeend 041

#gavdcodebegin 042
function PsSpCliM365_UpdateSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink set --webUrl $configFile.appsettings.SiteCollUrl `
								  --fileUrl "/TestLibrary/TestDocument.docx" `
								  --id "503e3411-478f-4196-8650-84d293613237" `
								  --expirationDateTime "2024-01-09T16:57:00.000Z"

	m365 logout
}
#gavdcodeend 042

#gavdcodebegin 043
function PsSpCliM365_DeleteSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink remove --webUrl $configFile.appsettings.SiteCollUrl `
								     --fileUrl "/TestLibrary/TestDocument.docx" `
								     --id "503e3411-478f-4196-8650-84d293613237" `
								     --force

	m365 logout
}
#gavdcodeend 043

#gavdcodebegin 044
function PsSpCliM365_ClearSharesLinkOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinglink clear --webUrl $configFile.appsettings.SiteCollUrl `
								     --fileUrl "/TestLibrary/TestDocument.docx" `
								     --scope "organization" `
								     --force

	m365 logout
}
#gavdcodeend 044

#gavdcodebegin 045
function PsSpCliM365_GetAllSharingOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file sharinginfo get --webUrl $configFile.appsettings.SiteCollUrl `
								  --fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 045

#gavdcodebegin 046
function PsSpCliM365_GetAllVersionsOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file version list --webUrl $configFile.appsettings.SiteCollUrl `
							   --fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 046

#gavdcodebegin 047
function PsSpCliM365_GetOneVersionOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file version get --webUrl $configFile.appsettings.SiteCollUrl `
							  --fileUrl "/TestLibrary/TestDocument.docx" `
							  --label "1.0"

	m365 logout
}
#gavdcodeend 047

#gavdcodebegin 048
function PsSpCliM365_RemoveOneVersionOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file version remove --webUrl $configFile.appsettings.SiteCollUrl `
								 --fileUrl "/TestLibrary/TestDocument.docx" `
								 --label "1.0"

	m365 logout
}
#gavdcodeend 048

#gavdcodebegin 049
function PsSpCliM365_RestoreOneVersionOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file version restore --webUrl $configFile.appsettings.SiteCollUrl `
								  --fileUrl "/TestLibrary/TestDocument.docx" `
								  --label "3.0"

	m365 logout
}
#gavdcodeend 049

#gavdcodebegin 050
function PsSpCliM365_ClearOneVersionOneDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file version clear --webUrl $configFile.appsettings.SiteCollUrl `
								--fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 050

#gavdcodebegin 030
function PsSpCliM365_BreakInheritanceItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem roleinheritance break --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 9

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
function PsSpCliM365_RestoreInheritanceItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem roleinheritance reset --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 9

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 051
function PsSpCliM365_AssignRoleItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem roleassignment add --webUrl $configFile.appsettings.SiteCollUrl `
										 --listTitle "TestList" `
										 --listItemId 9 `
										 --upn "user@domain.onmicrosoft.com" `
										 --roleDefinitionName "Full Control"

	m365 logout
}
#gavdcodeend 051

#gavdcodebegin 052
function PsSpCliM365_RemoveRoleItem
{
	$spCtx = PsSpCliM365_Login

	m365 spo listitem roleassignment remove --webUrl $configFile.appsettings.SiteCollUrl `
											--listTitle "TestList" `
											--listItemId 9 `
											--upn "user@domain.onmicrosoft.com"

	m365 logout
}
#gavdcodeend 052

#gavdcodebegin 053
function PsSpCliM365_BreakInheritanceDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file roleinheritance break --webUrl $configFile.appsettings.SiteCollUrl `
										--fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 053

#gavdcodebegin 054
function PsSpCliM365_RestoreInheritanceDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file roleinheritance reset --webUrl $configFile.appsettings.SiteCollUrl `
										--fileUrl "/TestLibrary/TestDocument.docx"

	m365 logout
}
#gavdcodeend 054

#gavdcodebegin 055
function PsSpCliM365_AssignRoleDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file roleassignment add --webUrl $configFile.appsettings.SiteCollUrl `
									 --fileUrl "/TestLibrary/TestDocument.docx" `
									 --upn "user@domain.onmicrosoft.com" `
									 --roleDefinitionName "Full Control"

	m365 logout
}
#gavdcodeend 055

#gavdcodebegin 056
function PsSpCliM365_RemoveRoleDocument
{
	$spCtx = PsSpCliM365_Login

	m365 spo file roleassignment remove --webUrl $configFile.appsettings.SiteCollUrl `
										--fileUrl "/TestLibrary/TestDocument.docx" `
										--upn "user@domain.onmicrosoft.com"

	m365 logout
}
#gavdcodeend 056

#gavdcodebegin 057
function PsSpCliM365_BreakInheritanceFolder
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder roleinheritance break --webUrl $configFile.appsettings.SiteCollUrl `
										  --folderUrl "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 057

#gavdcodebegin 058
function PsSpCliM365_RestoreInheritanceFolder
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder roleinheritance reset --webUrl $configFile.appsettings.SiteCollUrl `
										  --folderUrl "/TestLibrary/NewFolderCli"

	m365 logout
}
#gavdcodeend 058

#gavdcodebegin 059
function PsSpCliM365_AssignRoleFolder
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder roleassignment add --webUrl $configFile.appsettings.SiteCollUrl `
									   --folderUrl "/TestLibrary/NewFolderCli" `
									   --upn "user@domain.onmicrosoft.com" `
									   --roleDefinitionName "Full Control"

	m365 logout
}
#gavdcodeend 059

#gavdcodebegin 060
function PsSpCliM365_RemoveRoleFolder
{
	$spCtx = PsSpCliM365_Login

	m365 spo folder roleassignment remove --webUrl $configFile.appsettings.SiteCollUrl `
										  --folderUrl "/TestLibrary/NewFolderCli" `
										  --upn "user@domain.onmicrosoft.com"

	m365 logout
}
#gavdcodeend 060


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 060 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Using the CLI for Microsoft 365 --------
#PsSpCliM365_CreateListItem
#PsSpCliM365_GetAllListItemsInList
#PsSpCliM365_GetAllListItemsInListSelectFields
#PsSpCliM365_GetOneListItem
#PsSpCliM365_GetOneListItemSelectFields
#PsSpCliM365_GetOneListItemByCAML
#PsSpCliM365_GetOneListItemByFilter
#PsSpCliM365_UpdateOneListItem
#PsSpCliM365_DeleteOneListItem
#PsSpCliM365_UploadOneDocument
#PsSpCliM365_UploadOneDocumentByRelativePath
#PsSpCliM365_UploadOneDocumentAndChangeFieldValue
#PsSpCliM365_GetAllDocumentsInLibrary
#PsSpCliM365_GetOneDocumentProperties
#PsSpCliM365_GetOneDocumentPropertiesAsListItem
#PsSpCliM365_GetOneDocumentPropertiesAsString
#PsSpCliM365_DownloadOneDocument
#PsSpCliM365_UpdateOneDocument
#PsSpCliM365_CopyOneDocument
#PsSpCliM365_MoveOneDocument
#PsSpCliM365_DeleteOneDocument
#PsSpCliM365_CheckoutOneDocument
#PsSpCliM365_CheckoutUndoOneDocument
#PsSpCliM365_CheckinOneDocument
#PsSpCliM365_CreateFolderInLibrary
#PsSpCliM365_GetFoldersInLibrary
#PsSpCliM365_GetOneFolderInLibrary
#PsSpCliM365_RenameOneFolderInLibrary
#PsSpCliM365_CopyOneFolderToOtherLibrary
#PsSpCliM365_MoveOneFolderToOtherLibrary
#PsSpCliM365_DeleteOneFolderFromLibrary
#PsSpCliM365_AddAttachementsToItem
#PsSpCliM365_GetAllAttachementsInItem
#PsSpCliM365_GetOneAttachementsInItem
#PsSpCliM365_UpdateAttachementsInItem
#PsSpCliM365_DeleteAttachementsInItem
#PsSpCliM365_GetAllSharesLinkOneDocument
#PsSpCliM365_GetOneSharesLinkOneDocument
#PsSpCliM365_AddSharesLinkOneDocument
#PsSpCliM365_UpdateSharesLinkOneDocument
#PsSpCliM365_DeleteSharesLinkOneDocument
#PsSpCliM365_ClearSharesLinkOneDocument
#PsSpCliM365_GetAllSharingOneDocument
#PsSpCliM365_GetAllVersionsOneDocument
#PsSpCliM365_GetOneVersionOneDocument
#PsSpCliM365_RemoveOneVersionOneDocument
#PsSpCliM365_RestoreOneVersionOneDocument
#PsSpCliM365_ClearOneVersionOneDocument
#PsSpCliM365_BreakInheritanceItem
#PsSpCliM365_RestoreInheritanceItem
#PsSpCliM365_AssignRoleItem
#PsSpCliM365_RemoveRoleItem
#PsSpCliM365_BreakInheritanceDocument
#PsSpCliM365_RestoreInheritanceDocument
#PsSpCliM365_AssignRoleDocument
#PsSpCliM365_RemoveRoleDocument
#PsSpCliM365_BreakInheritanceFolder
#PsSpCliM365_RestoreInheritanceFolder
#PsSpCliM365_AssignRoleFolder
#PsSpCliM365_RemoveRoleFolder

Write-Host "Done" 
