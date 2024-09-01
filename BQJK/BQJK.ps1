
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpCsom_Login
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.SiteCollUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function PsSpCsom_CreateOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItemCreationInfo = `
				New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	$newListItem = $myList.AddItem($myListItemCreationInfo)
	$newListItem["Title"] = "NewListItemPsCsom"
	$newListItem.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 001

#gavdcodebegin 002
Function PsSpCsom_CreateMultipleItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	for($intCounter = 0; $intCounter -lt 4; $intCounter++) {
		$myListItemCreationInfo = `
				New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
		$newListItem = $myList.AddItem($myListItemCreationInfo)
		$newListItem["Title"] = $intCounter.ToString() + "-NewListItemPsCsom"
		$newListItem.Update()
	}

	$spCtx.ExecuteQuery()
}
#gavdcodeend 002

#gavdcodebegin 003
Function PsSpCsom_UploadOneDocument($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$filePath = "C:\Temporary\"
    $fileName = "TestDocument01.docx"
	$fileFullPath = $filePath + $fileName

	$myFileInfo = New-Object System.IO.FileInfo($fileName)
    $spCtx.Load($myList.RootFolder)
    $spCtx.ExecuteQuery()

    $fileUrl = $myList.RootFolder.ServerRelativeUrl + "/" + $myFileInfo.Name
	$fileMode = [System.IO.FileMode]::Open
	$myFileStream = New-Object System.IO.FileStream $fileFullPath, $fileMode
	[Microsoft.SharePoint.Client.File]::SaveBinaryDirect($spCtx, $fileUrl, `
															$myFileStream, $true)
}
#gavdcodeend 003

#gavdcodebegin 004
Function PsSpCsom_UploadOneDocumentFileCrInfo($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$filePath = "C:\Temporary\"
    $fileName = "TestDocument01.docx"
	$fileFullPath = $filePath + $fileName

    $spCtx.Load($myList.RootFolder)
    $spCtx.ExecuteQuery()

	$fileMode = [System.IO.FileMode]::Open
	$myFileStream = New-Object System.IO.FileStream $fileFullPath, $fileMode

	$myFileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$myFileCreationInfo.Overwrite = $true
	$myFileCreationInfo.ContentStream = $myFileStream
	$myFileCreationInfo.Url = $fileName

	$newFile = $myList.RootFolder.Files.Add($myFileCreationInfo)
	$spCtx.Load($newFile)
	$spCtx.ExecuteQuery()
}
#gavdcodeend 004

#gavdcodebegin 005
Function PsSpCsom_DownloadOneDocument($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$filePath = "C:\Temporary\"

	$listItemId = 27
	$myListItem = $myList.GetItemById($listItemId)
	$spCtx.Load($myListItem)
	$spCtx.Load($myListItem.File)
	$spCtx.ExecuteQuery()

	$fileRef = $myListItem.File.ServerRelativeUrl
	if ($spCtx.HasPendingRequest) { $spCtx.ExecuteQuery() }
	$myFileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($spCtx, $fileRef)
	$fileName = $filePath + $myListItem.File.Name
	$myFileStream = [System.IO.File]::Create($fileName)
	$myFileInfo.Stream.CopyTo($myFileStream)
	$myFileStream.Close()
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsSpCsom_ReadAllListItems($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $allItems = $myList.GetItems(`
					[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    foreach ($oneItem in $allItems) {
        Write-Host ($oneItem["Title"] + " - " + $oneItem.Id)
    }
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsSpCsom_ReadAllLibraryDocs($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

    $allItems = $myList.GetItems(`
					[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    foreach ($oneItem in $allItems) {
        Write-Host ($oneItem["FileLeafRef"] + " - " + $oneItem.Id)
    }
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsSpCsom_ReadOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$filterField = 28
    $rowLimit = 10
    $myViewXml = "
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='ID' />
                        <Value Type='Number'>" + $filterField + "</Value>
                    </Eq>
                </Where>
            </Query>
            <ViewFields>
                <FieldRef Name='Title' />
            </ViewFields>
            <RowLimit>" + $rowLimit + "</RowLimit>
        </View>"

    $myCamlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	$myCamlQuery.ViewXml = $myViewXml

    $allItems = $myList.GetItems($myCamlQuery)
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    foreach ($oneItem in $allItems) {
        Write-Host ($oneItem["Title"] + " - " + $oneItem.Id)
    }
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsSpCsom_ReadOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$filterField = 27
    $rowLimit = 10
    $myViewXml = "
        <View>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='ID' />
                        <Value Type='Number'>" + $filterField + "</Value>
                    </Eq>
                </Where>
            </Query>
            <ViewFields>
                <FieldRef Name='FileLeafRef' />
            </ViewFields>
            <RowLimit>" + $rowLimit + "</RowLimit>
        </View>"

    $myCamlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	$myCamlQuery.ViewXml = $myViewXml

    $allItems = $myList.GetItems($myCamlQuery)
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    Write-Host ($allItems[0]["FileLeafRef"] + " - " + $allItems[0].Id)
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsSpCsom_UpdateOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItem = $myList.GetItemById(28)
	$myListItem["Title"] = "NewListItemPsCsomUpdated"
	$myListItem.Update()
	$spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    Write-Host ("Item Title - " + $myListItem["Title"])
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsSpCsom_UpdateOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItem = $myList.GetItemById(27)
	$myListItem["Title"] = "LibraryDocPsCsomUpdated.docx"
	$myListItem.Update()
	$spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    Write-Host ("Item Title - " + $myListItem["Title"])
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsSpCsom_DeleteOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItem = $myList.GetItemById(28)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 012

#gavdcodebegin 013
Function PsSpCsom_DeleteAllListItems($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItems = $myList.GetItems(`
					[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
	$spCtx.Load($myListItems)
    $spCtx.ExecuteQuery()

	foreach ($oneItem in $myListItems) {
        $oneItemToDelete = $myList.GetItemById($oneItem.Id)
        $oneItemToDelete.DeleteObject()
    }

	$spCtx.ExecuteQuery()
}
#gavdcodeend 013

#gavdcodebegin 014
Function PsSpCsom_DeleteOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItem = $myList.GetItemById(27)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 014

#gavdcodebegin 015
Function PsSpCsom_DeleteAllLibraryDocs($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItems = $myList.GetItems(`
					[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
	$spCtx.Load($myListItems)
    $spCtx.ExecuteQuery()

	foreach ($oneItem in $myListItems) {
        $oneItemToDelete = $myList.GetItemById($oneItem.Id)
        $oneItemToDelete.DeleteObject()
    }

	$spCtx.ExecuteQuery()
}
#gavdcodeend 015

#gavdcodebegin 016
Function PsSpCsom_BreakSecurityInheritanceListItem($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $myListItem = $myList.GetItemById(43)
    $spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    $myListItem.BreakRoleInheritance($false, $true)
    $myListItem.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 016

#gavdcodebegin 017
Function PsSpCsom_ResetSecurityInheritanceListItem($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $myListItem = $myList.GetItemById(43)
    $spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    $myListItem.ResetRoleInheritance()
    $myListItem.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 017

 #gavdcodebegin 018
Function PsSpCsom_AddUserToSecurityRoleInListItem($spCtx)
{
	$myWeb = $spCtx.Web
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
	$myListItem = $myList.GetItemById(43)

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $roleDefinition = `
        New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(`
									[Microsoft.SharePoint.Client.RoleType]::Reader))

    $myListItem.RoleAssignments.Add($myUser, $roleDefinition)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsSpCsom_UpdateUserSecurityRoleInListItem($spCtx)
{
	$myWeb = $spCtx.Web
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
	$myListItem = $myList.GetItemById(43)

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $roleDefinition = `
        New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(`
							[Microsoft.SharePoint.Client.RoleType]::Administrator))

    $myRoleAssignment = $myListItem.RoleAssignments.GetByPrincipal($myUser)
    $myRoleAssignment.ImportRoleDefinitionBindings($roleDefinition)

    $myRoleAssignment.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsSpCsom_DeleteUserFromSecurityRoleInListItem($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $myListItem = $myList.GetItemById(43)

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $myListItem.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsSpCsom_CreateFolderInLibrary($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $myWeb.Lists.GetByTitle("TestDocuments")

    $myFolder01 = $myList.RootFolder.Folders.Add("FirstLevelFolderPS")
    $myFolder01.Update()
    $mySubFolder = $myFolder01.Folders.Add("SecondLevelFolderPS")
    $mySubFolder.Update()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsSpCsom_CreateFolderWithInfo($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $myWeb.Lists.GetByTitle("TestList")

    $infoFolder = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $infoFolder.UnderlyingObjectType = `
                          [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder
    $infoFolder.LeafName = "FolderWithInfoPS"
    $newItem = $myList.AddItem($infoFolder)
    $newItem["Title"] = "FolderWithInfoPS"
    $newItem.Update()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsSpCsom_AddItemInFolder($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $myWeb.Lists.GetByTitle("TestList")

    $myListItemCreationInfo =
                  New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $myListItemCreationInfo.FolderUrl = $spCtx.Url + "/lists/TestList/FolderWithInfoPS"
    $newListItem = $myList.AddItem($myListItemCreationInfo)
    $newListItem["Title"] = "NewListItemInFolderPsCsom"
    $newListItem.Update()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsSpCsom_UploadOneDocumentInFolder($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestDocuments")

    $filePath = "C:\Temporary\"
    $fileName = "TestDocument01.docx"
	$fileFullPath = $filePath + $fileName
    
    $spCtx.Load($myList.RootFolder)
    $spCtx.ExecuteQuery()
    
	$fileMode = [System.IO.FileMode]::Open
	$myFileStream = New-Object System.IO.FileStream $fileFullPath, $fileMode

	$myFileCreationInfo = 
                    New-Object Microsoft.SharePoint.Client.FileCreationInformation
	$myFileCreationInfo.Overwrite = $true
	$myFileCreationInfo.ContentStream = $myFileStream
	$myFileCreationInfo.Url = 
                    $spCtx.Url + "/TestDocuments/FirstLevelFolderPS/" + $fileName

	$newFile = $myList.RootFolder.Files.Add($myFileCreationInfo)
    $spCtx.Load($newFile)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 024

#gavdcodebegin 025
Function PsSpCsom_ReadAllFolders($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $allFolders = $myList.GetItems(`
                    [Microsoft.SharePoint.Client.CamlQuery]::CreateAllFoldersQuery())
    $spCtx.Load($allFolders)
    $spCtx.ExecuteQuery()

    foreach ($oneFolder in $allFolders) {
        Write-Host($oneFolder["FileLeafRef"] + " - " + $oneFolder["ServerUrl"])
    }
}
#gavdcodeend 025

#gavdcodebegin 026
Function PsSpCsom_ReadAllItemsInFolder($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $myQuery = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $myQuery.FolderServerRelativeUrl = 
                            "/sites/[SiteNm]/Lists/TestList/FolderWithInfoPS"
    $allItems = $myList.GetItems($myQuery)
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    foreach ($oneItem in $allItems) {
        Write-Host($oneItem["Title"] + " - " + $oneItem.Id);
    }
}
#gavdcodeend 026

#gavdcodebegin 027
Function PsSpCsom_DeleteOneFolder($spCtx)
{
    $folderRelativeUrl = "/sites/[SiteName]/Lists/TestList/FolderWithInfoPS"
    $myFolder = $spCtx.Web.GetFolderByServerRelativeUrl($folderRelativeUrl)

    $myFolder.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 027

#gavdcodebegin 028
Function PsSpCsom_CreateOneAttachment($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $listItemId = 3
    $myListItem = $myList.GetItemById($listItemId)

    $myFilePath = "C:\Temporary\Test.csv"
    $myFileName = "Test.csv"
    $myAttachmentInfo = 
             New-Object Microsoft.SharePoint.Client.AttachmentCreationInformation
    $myAttachmentInfo.FileName = $myFileName
    
    $fileMode = [System.IO.FileMode]::Open
    $myFileStream = New-Object System.IO.FileStream $myFilePath, $FileMode
    $myAttachmentInfo.ContentStream = $myFileStream
    $myAttachment = $myListItem.AttachmentFiles.Add($myAttachmentInfo)
    $spCtx.Load($myAttachment)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 028

#gavdcodebegin 029
Function PsSpCsom_ReadAllAttachments($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $listItemId = 3
    $myListItem = $myList.GetItemById($listItemId)

    $allAttachments = $myListItem.AttachmentFiles
    $spCtx.Load($allAttachments)
    $spCtx.ExecuteQuery()

    foreach ($oneAttachment in $allAttachments) {
        Write-Host "File Name - " $oneAttachment.FileName
    }
}
#gavdcodeend 029

#gavdcodebegin 030
Function PsSpCsom_DownloadAllAttachments($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $listItemId = 3
    $myListItem = $myList.GetItemById($listItemId)

    $allAttachments = $myListItem.AttachmentFiles
    $spCtx.Load($allAttachments)
    $spCtx.ExecuteQuery()

    $myFilesPath = "C:\Temporary\"
    foreach ($oneAttachment in $allAttachments) {
        Write-Host "File Name - " $oneAttachment.FileName
        $myFileInfo = [Microsoft.SharePoint.Client.File]::
                         OpenBinaryDirect($spCtx, $oneAttachment.ServerRelativeUrl)
        $spCtx.ExecuteQuery()

        $myFileStream = 
                    [System.IO.File]::Create($myFilesPath + $oneAttachment.FileName)
        $myFileInfo.Stream.CopyTo($myFileStream)
        $myFileStream.Close()
    }
}
#gavdcodeend 030

#gavdcodebegin 031
Function PsSpCsom_DeleteAllAttachments($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $listItemId = 3
    $myListItem = $myList.GetItemById($listItemId)

    $allAttachments = $myListItem.AttachmentFiles
    $spCtx.Load($allAttachments)
    $spCtx.ExecuteQuery()

    foreach ($oneAttachment in $allAttachments) {
        $oneAttachment.DeleteObject()
    }

    $spCtx.ExecuteQuery()
}
#gavdcodeend 031

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 031 ***

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = PsSpCsom_Login

#PsSpCsom_CreateOneListItem $spCtx
#PsSpCsom_CreateMultipleItem $spCtx
#PsSpCsom_UploadOneDocument $spCtx
#PsSpCsom_UploadOneDocumentFileCrInfo $spCtx
#PsSpCsom_DownloadOneDocument $spCtx
#PsSpCsom_ReadAllListItems $spCtx
#PsSpCsom_ReadAllLibraryDocs $spCtx
#PsSpCsom_ReadOneListItem $spCtx
#PsSpCsom_ReadOneLibraryDoc $spCtx
#PsSpCsom_UpdateOneListItem $spCtx
#PsSpCsom_UpdateOneLibraryDoc $spCtx
#PsSpCsom_DeleteOneListItem $spCtx
#PsSpCsom_DeleteAllListItems $spCtx
#PsSpCsom_DeleteOneLibraryDoc $spCtx
#PsSpCsom_DeleteAllLibraryDocs $spCtx
#PsSpCsom_BreakSecurityInheritanceListItem $spCtx
#PsSpCsom_ResetSecurityInheritanceListItem $spCtx
#PsSpCsom_AddUserToSecurityRoleInListItem $spCtx
#PsSpCsom_UpdateUserSecurityRoleInListItem $spCtx
#PsSpCsom_DeleteUserFromSecurityRoleInListItem $spCtx
#PsSpCsom_CreateFolderInLibrary $spCtx
#PsSpCsom_CreateFolderWithInfo $spCtx
#PsSpCsom_AddItemInFolder $spCtx
#PsSpCsom_UploadOneDocumentInFolder $spCtx
#PsSpCsom_ReadAllFolders $spCtx
#PsSpCsom_ReadAllItemsInFolder $spCtx
#PsSpCsom_DeleteOneFolder $spCtx
#PsSpCsom_CreateOneAttachment $spCtx
#PsSpCsom_ReadAllAttachments $spCtx
#PsSpCsom_DownloadAllAttachments $spCtx
#PsSpCsom_DeleteAllAttachments $spCtx

Write-Host "Done"