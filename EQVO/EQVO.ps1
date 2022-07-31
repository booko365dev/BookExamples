
Function LoginPsCsom()
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
#-----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsCsom_CreateOneList($spCtx)
{
	$myWeb = $spCtx.Web

	$myListCreationInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
	$myListCreationInfo.Title = "NewListPsCsom"
	$myListCreationInfo.Description = "New List created using PowerShell CSOM";
	$myListCreationInfo.TemplateType = `
			[Microsoft.SharePoint.Client.ListTemplateType]'GenericList'

	$newList = $myWeb.Lists.Add($myListCreationInfo)
	$newList.OnQuickLaunch = $true
	$newList.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsCsom_ReadAllList($spCtx)
{
	$myWeb = $spCtx.Web
	$allLists = $myWeb.Lists
	$spCtx.Load($allLists)
	$spCtx.ExecuteQuery()

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title + " - " + $oneList.Id
	}
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsCsom_ReadOneList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$spCtx.Load($myList)
	$spCtx.ExecuteQuery();

	Write-Host "List description -" $myList.Description
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsCsom_UpdateOneList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myList.Description = "New List Description"
	$myList.Update()
	$spCtx.Load($myList)
	$spCtx.ExecuteQuery()

	Write-Host "List description -" $myList.Description
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsCsom_DeleteOneList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myList.DeleteObject()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsCsom_AddOneFieldToList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$fieldXml = "<Field DisplayName='MyMultilineField' Type='Note' />"
	$myField = $myList.Fields.AddFieldAsXml($fieldXml, `
						  $true, `
						  [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue)
	$spCtx.ExecuteQuery()
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsCsom_ReadAllFieldsFromList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$allFields = $myList.Fields
	$spCtx.Load($allFields)
	$spCtx.ExecuteQuery()

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsCsom_ReadOneFieldFromList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myField = $myList.Fields.GetByTitle("MyMultilineField")
	$spCtx.Load($myField)
	$spCtx.ExecuteQuery()

	Write-Host $myField.Id "-" $myField.TypeAsString
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsCsom_UpdateOneFieldInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myField = $myList.Fields.GetByTitle("MyMultilineField")

	$myField.Description = "New Field Description"
	$myField.Hidden = $false

	$myField.Update()
	$spCtx.Load($myField)
	$spCtx.ExecuteQuery()

	Write-Host $myField.Description
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsCsom_DeleteOneFieldFromList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myField = $myList.Fields.GetByTitle("MyMultilineField")
	$myField.DeleteObject()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 10

#gavdcodebegin 12
Function SpPsCsom_RetrieveProperties() 
{
	param(
   [Microsoft.SharePoint.Client.ClientObject]$Object = 
		$(throw "Please provide a Client Object"), [string]$PropertyName
	)
	
   $myCtx = $Object.Context
   $myLoad = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load") 
   $myType = $Object.GetType()
   $myCltLoad = $myLoad.MakeGenericMethod($myType) 

   $myParam = [System.Linq.Expressions.Expression]::Parameter(($myType), $myType.Name)
   $myExpr = [System.Linq.Expressions.Expression]::Lambda(
			[System.Linq.Expressions.Expression]::Convert(
			[System.Linq.Expressions.Expression]::PropertyOrField($myParam, $PropertyName),
			[System.Object]
			),
			$($myParam)
   )
   $myExprArray = [System.Array]::CreateInstance($myExpr.GetType(), 1)
   $myExprArray.SetValue($myExpr, 0)
   $myCltLoad.Invoke($myCtx, @($Object, $myExprArray))
}
#gavdcodeend 12

#gavdcodebegin 11
Function SpPsCsom_BreakSecurityInheritanceList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$spCtx.Load($myList)
	$spCtx.ExecuteQuery()

	RetrieveProperties -Object $myList -PropertyName "HasUniqueRoleAssignments"
	$spCtx.ExecuteQuery()

	if ($myList.HasUniqueRoleAssignments -eq $false)
	{
		$myList.BreakRoleInheritance($false, $true)
	}
	$myList.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 11

#gavdcodebegin 13
Function SpPsCsom_ResetSecurityInheritanceList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$spCtx.Load($myList)
	$spCtx.ExecuteQuery();

	RetrieveProperties -Object $myList -PropertyName "HasUniqueRoleAssignments"
	$spCtx.ExecuteQuery()

	if ($myList.HasUniqueRoleAssignments -eq $true)
	{
		$myList.ResetRoleInheritance()
	}
	$myList.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsCsomAddUserToSecurityRoleInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")

	$myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
	$roleDefinition = 
		New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
	$roleDefinition.Add($myWeb.RoleDefinitions.GetByType(`
										[Microsoft.SharePoint.Client.RoleType]::Reader))
	$myRoleAssignment = $myList.RoleAssignments.Add($myUser, $roleDefinition)

	$spCtx.ExecuteQuery()
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsCsom_UpdateUserSecurityRoleInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")

	$myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
	$roleDefinition =
		New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
	$roleDefinition.Add($myWeb.RoleDefinitions.GetByType(`
								[Microsoft.SharePoint.Client.RoleType]::Administrator))

	$myRoleAssignment = $myList.RoleAssignments.GetByPrincipal($myUser)
	$myRoleAssignment.ImportRoleDefinitionBindings($roleDefinition)

	$myRoleAssignment.Update()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsCsom_DeleteUserFromSecurityRoleInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")

	$myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
	$myList.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

	$spCtx.ExecuteQuery()
	$spCtx.Dispose()
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsCsom_ColumnIndex($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListCsCsom")
	$myField = $myList.Fields.GetByTitle("My Text Col")

	$myField.Indexed = $true

	$myField.Update()
	$spCtx.Load($myField)
	$spCtx.ExecuteQuery()

	Write-Host $myField.Description
}
#gavdcodeend 17

#-----------------------------------------------------------------------------------------


Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsCsom

#SpPsCsom_CreateOneList $spCtx
#SpPsCsom_ReadAllList $spCtx
#SpPsCsom_ReadOneList $spCtx
#SpPsCsom_UpdateOneList $spCtx
#SpPsCsom_DeleteOneList $spCtx
#SpPsCsom_AddOneFieldToList $spCtx
#SpPsCsom_ReadAllFieldsFromList $spCtx
#SpPsCsom_ReadOneFieldFromList $spCtx
#SpPsCsom_UpdateOneFieldInList $spCtx
#SpPsCsom_DeleteOneFieldFromList $spCtx
#SpPsCsom_BreakSecurityInheritanceList $spCtx
#SpPsCsom_ResetSecurityInheritanceList $spCtx
#SpPsCsom_AddUserToSecurityRoleInList $spCtx
#SpPsCsom_UpdateUserSecurityRoleInList $spCtx
#SpPsCsom_DeleteUserFromSecurityRoleInList $spCtx
#SpPsCsom_ColumnIndex $spCtx

Write-Host "Done"