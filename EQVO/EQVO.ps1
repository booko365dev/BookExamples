
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
#-----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsCsomCreateOneList($spCtx)
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
Function SpPsCsomReadAllList($spCtx)
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
Function SpPsCsomReadOneList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$spCtx.Load($myList)
	$spCtx.ExecuteQuery();

	Write-Host "List description -" $myList.Description
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsCsomUpdateOneList($spCtx)
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
Function SpPsCsomDeleteOneList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myList.DeleteObject()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsCsomAddOneFieldToList($spCtx)
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
Function SpPsCsomReadAllFieldsFromList($spCtx)
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
Function SpPsCsomReadOneFieldFromList($spCtx)
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
Function SpPsCsomUpdateOneFieldInList($spCtx)
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
Function SpPsCsomDeleteOneFieldFromList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")
	$myField = $myList.Fields.GetByTitle("MyMultilineField")
	$myField.DeleteObject()
	$spCtx.ExecuteQuery()
}
#gavdcodeend 10

#gavdcodebegin 12
Function RetrieveProperties() 
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
Function SpPsCsomBreakSecurityInheritanceList($spCtx)
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
Function SpPsCsomResetSecurityInheritanceList($spCtx)
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

	$myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
	$roleDefinition = 
		New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
	$roleDefinition.Add($myWeb.RoleDefinitions.GetByType(`
										[Microsoft.SharePoint.Client.RoleType]::Reader))
	$myRoleAssignment = $myList.RoleAssignments.Add($myUser, $roleDefinition)

	$spCtx.ExecuteQuery()
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsCsomUpdateUserSecurityRoleInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")

	$myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
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
Function SpPsCsomDeleteUserFromSecurityRoleInList($spCtx)
{
	$myWeb = $spCtx.Web
	$myList = $myWeb.Lists.GetByTitle("NewListPsCsom")

	$myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
	$myList.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

	$spCtx.ExecuteQuery()
	$spCtx.Dispose()
}
#gavdcodeend 16

#-----------------------------------------------------------------------------------------


Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Temporary\spPs.values.config"

$spCtx = LoginPsCsom

#SpPsCsomCreateOneList $spCtx
#SpPsCsomReadAllList $spCtx
#SpPsCsomReadOneList $spCtx
#SpPsCsomUpdateOneList $spCtx
#SpPsCsomDeleteOneList $spCtx
#SpPsCsomAddOneFieldToList $spCtx
#SpPsCsomReadAllFieldsFromList $spCtx
#SpPsCsomReadOneFieldFromList $spCtx
#SpPsCsomUpdateOneFieldInList $spCtx
#SpPsCsomDeleteOneFieldFromList $spCtx
#SpPsCsomBreakSecurityInheritanceList $spCtx
#SpPsCsomResetSecurityInheritanceList $spCtx
#SpPsCsomAddUserToSecurityRoleInList $spCtx
#SpPsCsomUpdateUserSecurityRoleInList $spCtx
#SpPsCsomDeleteUserFromSecurityRoleInList $spCtx

Write-Host "Done"