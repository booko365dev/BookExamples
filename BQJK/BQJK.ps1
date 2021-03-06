﻿Function LoginPsCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsCsomCreateOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItemCreationInfo = `
					New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	$newListItem = $myList.AddItem($myListItemCreationInfo)
	$newListItem["Title"] = "NewListItemPsCsom"
	$newListItem.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsCsomCreateMultipleItem($spCtx)
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
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsCsomUploadOneDocument($spCtx)
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
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsCsomUploadOneDocumentFileCrInfo($spCtx)
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
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsCsomDownloadOneDocument($spCtx)
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
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsCsomReadAllListItems($spCtx)
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
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsCsomReadAllLibraryDocs($spCtx)
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
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsCsomReadOneListItem($spCtx)
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
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsCsomReadOneLibraryDoc($spCtx)
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
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsCsomUpdateOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItem = $myList.GetItemById(28)
	$myListItem["Title"] = "NewListItemPsCsomUpdated"
	$myListItem.Update()
	$spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    Write-Host ("Item Title - " + $myListItem["Title"])
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsCsomUpdateOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItem = $myList.GetItemById(27)
	$myListItem["Title"] = "LibraryDocPsCsomUpdated.docx"
	$myListItem.Update()
	$spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    Write-Host ("Item Title - " + $myListItem["Title"])
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsCsomDeleteOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItem = $myList.GetItemById(28)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsCsomDeleteAllListItems($spCtx)
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
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsCsomDeleteOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItem = $myList.GetItemById(27)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsCsomDeleteAllLibraryDocs($spCtx)
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
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsCsomBreakSecurityInheritanceListItem($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $myListItem = $myList.GetItemById(43)
    $spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    $myListItem.BreakRoleInheritance($false, $true)
    $myListItem.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsCsomResetSecurityInheritanceListItem($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")

    $myListItem = $myList.GetItemById(43)
    $spCtx.Load($myListItem)
    $spCtx.ExecuteQuery()

    $myListItem.ResetRoleInheritance()
    $myListItem.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 17

 #gavdcodebegin 18
Function SpPsCsomAddUserToSecurityRoleInListItem($spCtx)
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
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsCsomUpdateUserSecurityRoleInListItem($spCtx)
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
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsCsomDeleteUserFromSecurityRoleInListItem($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $myListItem = $myList.GetItemById(43)

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $myListItem.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsCsomCreateFolderInLibrary($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $myWeb.Lists.GetByTitle("TestDocuments")

    $myFolder01 = $myList.RootFolder.Folders.Add("FirstLevelFolderPS")
    $myFolder01.Update()
    $mySubFolder = $myFolder01.Folders.Add("SecondLevelFolderPS")
    $mySubFolder.Update()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsCsomCreateFolderWithInfo($spCtx)
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
#gavdcodeend 22

#gavdcodebegin 23
Function SpPsCsomAddItemInFolder($spCtx)
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
#gavdcodeend 23

#gavdcodebegin 24
Function SpPsCsomUploadOneDocumentInFolder($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestDocuments")

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
	$myFileCreationInfo.Url = $spCtx.Url + "/TestDocuments/FirstLevelFolderPS/" + $fileName

	$newFile = $myList.RootFolder.Files.Add($myFileCreationInfo)
    $spCtx.Load($newFile)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 24

#gavdcodebegin 25
Function SpPsCsomReadAllFolders($spCtx)
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
#gavdcodeend 25

#gavdcodebegin 26
Function SpPsCsomReadAllItemsInFolder($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $myQuery = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $myQuery.FolderServerRelativeUrl = "/sites/[SiteName]/Lists/TestList/FolderWithInfoPS"
    $allItems = $myList.GetItems($myQuery)
    $spCtx.Load($allItems)
    $spCtx.ExecuteQuery()

    foreach ($oneItem in $allItems) {
        Write-Host($oneItem["Title"] + " - " + $oneItem.Id);
    }
}
#gavdcodeend 26

#gavdcodebegin 27
Function SpPsCsomDeleteOneFolder($spCtx)
{
    $folderRelativeUrl = "/sites/[SiteName]/Lists/TestList/FolderWithInfoPS"
    $myFolder = $spCtx.Web.GetFolderByServerRelativeUrl($folderRelativeUrl)

    $myFolder.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 27

#gavdcodebegin 28
Function SpPsCsomCreateOneAttachment($spCtx)
{
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $listItemId = 3
    $myListItem = $myList.GetItemById($listItemId)

    $myFilePath = "C:\Temporary\Test.csv"
    $myFileName = "Test.csv"
    $myAttachmentInfo = New-Object Microsoft.SharePoint.Client.AttachmentCreationInformation
    $myAttachmentInfo.FileName = $myFileName
    
    $fileMode = [System.IO.FileMode]::Open
    $myFileStream = New-Object System.IO.FileStream $myFilePath, $FileMode
    $myAttachmentInfo.ContentStream = $myFileStream
    $myAttachment = $myListItem.AttachmentFiles.Add($myAttachmentInfo)
    $spCtx.Load($myAttachment)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 28

#gavdcodebegin 29
Function SpPsCsomReadAllAttachments($spCtx)
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
#gavdcodeend 29

#gavdcodebegin 30
Function SpPsCsomDownloadAllAttachments($spCtx)
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
        $myFileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($spCtx, $oneAttachment.ServerRelativeUrl)
        $spCtx.ExecuteQuery()

        $myFileStream = [System.IO.File]::Create($myFilesPath + $oneAttachment.FileName)
        $myFileInfo.Stream.CopyTo($myFileStream)
        $myFileStream.Close()
    }
}
#gavdcodeend 30

#gavdcodebegin 31
Function SpPsCsomDeleteAllAttachments($spCtx)
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
#gavdcodeend 31

#-----------------------------------------------------------------------------------------


Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsCsom

#SpPsCsomCreateOneListItem $spCtx
#SpPsCsomCreateMultipleItem $spCtx
#SpPsCsomUploadOneDocument $spCtx
#SpPsCsomUploadOneDocumentFileCrInfo $spCtx
#SpPsCsomDownloadOneDocument $spCtx
#SpPsCsomReadAllListItems $spCtx
#SpPsCsomReadAllLibraryDocs $spCtx
#SpPsCsomReadOneListItem $spCtx
#SpPsCsomReadOneLibraryDoc $spCtx
#SpPsCsomUpdateOneListItem $spCtx
#SpPsCsomUpdateOneLibraryDoc $spCtx
#SpPsCsomDeleteOneListItem $spCtx
#SpPsCsomDeleteAllListItems $spCtx
#SpPsCsomDeleteOneLibraryDoc $spCtx
#SpPsCsomDeleteAllLibraryDocs $spCtx
#SpPsCsomBreakSecurityInheritanceListItem $spCtx
#SpPsCsomResetSecurityInheritanceListItem $spCtx
#SpPsCsomAddUserToSecurityRoleInListItem $spCtx
#SpPsCsomUpdateUserSecurityRoleInListItem $spCtx
#SpPsCsomDeleteUserFromSecurityRoleInListItem $spCtx
#SpPsCsomCreateFolderInLibrary $spCtx
#SpPsCsomCreateFolderWithInfo $spCtx
#SpPsCsomAddItemInFolder $spCtx
#SpPsCsomUploadOneDocumentInFolder $spCtx
#SpPsCsomReadAllFolders $spCtx
#SpPsCsomReadAllItemsInFolder $spCtx
#SpPsCsomDeleteOneFolder $spCtx
#SpPsCsomCreateOneAttachment $spCtx
#SpPsCsomReadAllAttachments $spCtx
#SpPsCsomDownloadAllAttachments $spCtx
#SpPsCsomDeleteAllAttachments $spCtx

Write-Host "Done"