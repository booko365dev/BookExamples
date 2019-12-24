Function LoginPsCsom()
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

Function SpPsCsomDeleteOneListItem($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestList")

	$myListItem = $myList.GetItemById(28)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}

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

Function SpPsCsomDeleteOneLibraryDoc($spCtx)
{
	$myList = $spCtx.Web.Lists.GetByTitle("TestLibrary")

	$myListItem = $myList.GetItemById(27)
	$myListItem.DeleteObject()
    $spCtx.ExecuteQuery()
}

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

Function SpPsCsomDeleteUserFromSecurityRoleInListItem($spCtx)
{
    $myWeb = $spCtx.Web
    $myList = $spCtx.Web.Lists.GetByTitle("TestList")
    $myListItem = $myList.GetItemById(43)

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $myListItem.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}

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

Write-Host "Done"
