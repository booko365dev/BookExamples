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

Function LoginCsom($WebFullUrl)
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext($WebFullUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

Function LoginAdminCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spAdminUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

Function SpPsCsomCreateOneSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myUser = $configFile.appsettings.spUserName
    $mySiteCreationProps = New-Object `
				Microsoft.Online.SharePoint.TenantAdministration.SiteCreationProperties
    $mySiteCreationProps.Url = $configFile.appsettings.spBaseUrl + 
                                        "/sites/NewSiteCollectionModernPsCsom"
    $mySiteCreationProps.Title = "NewSiteCollectionModernPsCsom"
    $mySiteCreationProps.Owner = $configFile.appsettings.spUserName
    $mySiteCreationProps.Template = "STS#3"
    $mySiteCreationProps.StorageMaximumLevel = 100
    $mySiteCreationProps.UserCodeMaximumLevel = 50

    $myOps = $myTenant.CreateSite($mySiteCreationProps)
    $spAdminCtx.Load($myOps)
    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomFindWebTemplates($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTemplates = $myTenant.GetSPOTenantWebTemplates(1033, 0)
    
	$spAdminCtx.Load($myTemplates)
    $spAdminCtx.ExecuteQuery()

    foreach ($oneTemplate in $myTemplates)
    {
        Write-Host ($oneTemplate.Name + " - " + $oneTemplate.Title)
    }
}

Function SpPsCsomReadAllSiteCollections($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetSiteProperties(0, $true)
    
	$spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    foreach ($oneSiteColl in $myProps)
    {
        Write-Host ($oneSiteColl.Title + " - " + $oneSiteColl.Url)
    }
}

Function SpPsCsomRemoveSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveSite(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}

Function SpPsCsomRestoreSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RestoreDeletedSite(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}

Function SpPsCsomRemoveDeletedSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveDeletedSite(
        $onfigFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}

Function SpPsCsomCreateGroupForSite($spAdminCtx)
{
    $myOwners = @( "user@domain.onmicrosoft.com" )
    $myGroupParams = New-Object `
		Microsoft.Online.SharePoint.TenantManagement.GroupCreationParams($spAdminCtx)
    $myGroupParams.Owners = $myOwners
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.CreateGroupForSite(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom",
			"GroupForNewSiteCollectionModernPsCsom",
			"GroupForNewSiteCollAlias",
			$true,
			$myGroupParams)

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomSetAdministratorSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.SetSiteAdmin(
        $configFile.AppSettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom",
        "user@domain.onmicrosoft.com",
        $true)

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomRegisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RegisterHubSite(
        $configFile.AppSettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomUnregisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.UnregisterHubSite(
        $onfigFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomGetHubSiteCollectionProperties($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetHubSitePropertiesByUrl(
		$configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    Write-Host($myProps.Title)
}

Function SpPsCsomUpdateHubSiteCollectionProperties($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetHubSitePropertiesByUrl(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    $myProps.Title = $myProps.Title + "_Updated"
    $myProps.Update()

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    Write-Host($myProps.Title)
}

Function SpPsCsomAddSiteToHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.ConnectSiteToHubSite(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteForHub",
		$configFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomremoveSiteFromHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.DisconnectSiteFromHubSite(
        $configFile.appsettings.spBaseUrl + "/sites/NewSiteForHub")

    $spAdminCtx.ExecuteQuery()
}

Function SpPsCsomCreateOneWebInSiteCollection($spCtx)
{
    $mySite = $spCtx.Site

    $myWebCreationInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $myWebCreationInfo.Url = "NewWebSiteModernPsCsom"
    $myWebCreationInfo.Title = "NewWebSiteModernPsCsom"
    $myWebCreationInfo.Description = "NewWebSiteModernPsCsom Description"
    $myWebCreationInfo.UseSamePermissionsAsParentSite = $true
    $myWebCreationInfo.WebTemplate = "STS#3"
    $myWebCreationInfo.Language = 1033

    $myWeb = $mySite.RootWeb.Webs.Add($myWebCreationInfo)
    $spCtx.ExecuteQuery()
}

Function SpPsCsomGetWebsInSiteCollection($spCtx)
{
    $mySite = $spCtx.Site

    $myWebs = $mySite.RootWeb.Webs
    $spCtx.Load($myWebs)
    $spCtx.ExecuteQuery()

    foreach ($oneWeb in $myWebs)
    {
        Write-Host($oneWeb.Title + " - " + $oneWeb.Url + " - " + $oneWeb.Id)
    }
}

Function SpPsCsomGetOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    Write-Host($myWeb.Title + " - " + $myWeb.Url + " - " + $myWeb.Id)
}

Function SpPsCsomUpdateOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.Description = "NewWebSiteModernPsCsom Description Updated"
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}

Function SpPsCsomDeleteOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.DeleteObject()
    $spCtx.ExecuteQuery()
}

Function SpPsCsomBreakSecurityInheritanceWeb()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.BreakRoleInheritance($false, $true)
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}

Function SpPsCsomResetSecurityInheritanceWeb()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.ResetRoleInheritance()
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}

Function SpPsCsomAddUserToSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $roleDefinition = New-Object `
				Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(
							[Microsoft.SharePoint.Client.RoleType]::Reader))
    $myWeb.RoleAssignments.Add($myUser, $roleDefinition)

    $spCtx.ExecuteQuery()
}

Function SpPsCsomUpdateUserSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $roleDefinition = New-Object `
					Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(
						[Microsoft.SharePoint.Client.RoleType]::Administrator))

    $myRoleAssignment = $myWeb.RoleAssignments.GetByPrincipal($myUser)
    $myRoleAssignment.ImportRoleDefinitionBindings($roleDefinition)

    $myRoleAssignment.Update()
    $spCtx.ExecuteQuery()
}

Function SpPsCsomDeleteUserFromSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.spUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.spUserName)
    $myWeb.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsCsom
$spAdminCtx = LoginAdminCsom

#SpPsCsomCreateOneSiteCollection $spAdminCtx
#SpPsCsomFindWebTemplates $spAdminCtx
#SpPsCsomReadAllSiteCollections $spAdminCtx
#SpPsCsomRemoveSiteCollection $spAdminCtx
#SpPsCsomRestoreSiteCollection $spAdminCtx
#SpPsCsomRemoveDeletedSiteCollection $spAdminCtx
#SpPsCsomCreateGroupForSite $spAdminCtx
#SpPsCsomRegisterAsHubSiteCollection $spAdminCtx
#SpPsCsomUnregisterAsHubSiteCollection $spAdminCtx
#SpPsCsomGetHubSiteCollectionProperties $spAdminCtx
#SpPsCsomUpdateHubSiteCollectionProperties $spAdminCtx
#SpPsCsomAddSiteToHubSiteCollection $spAdminCtx
#SpPsCsomremoveSiteFromHubSiteCollection $spAdminCtx

#SpPsCsomCreateOneWebInSiteCollection $spCtx
#SpPsCsomGetWebsInSiteCollection $spCtx
#SpPsCsomGetOneWebInSiteCollection
#SpPsCsomUpdateOneWebInSiteCollection
#SpPsCsomDeleteOneWebInSiteCollection
#SpPsCsomBreakSecurityInheritanceWeb
#SpPsCsomResetSecurityInheritanceWeb
#SpPsCsomAddUserToSecurityRoleInWeb
#SpPsCsomUpdateUserSecurityRoleInWeb
#SpPsCsomDeleteUserFromSecurityRoleInWeb

Write-Host "Done"

