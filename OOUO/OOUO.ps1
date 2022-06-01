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

Function LoginCsom($WebFullUrl)
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext($WebFullUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

Function LoginAdminCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.SiteAdminUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsCsomCreateOneSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myUser = $configFile.appsettings.UserName
    $mySiteCreationProps = New-Object `
				Microsoft.Online.SharePoint.TenantAdministration.SiteCreationProperties
    $mySiteCreationProps.Url = $configFile.appsettings.SiteBaseUrl + 
                                        "/sites/NewSiteCollectionModernPsCsom"
    $mySiteCreationProps.Title = "NewSiteCollectionModernPsCsom"
    $mySiteCreationProps.Owner = $configFile.appsettings.UserName
    $mySiteCreationProps.Template = "STS#3"
    $mySiteCreationProps.StorageMaximumLevel = 100
    $mySiteCreationProps.UserCodeMaximumLevel = 50

    $myOps = $myTenant.CreateSite($mySiteCreationProps)
    $spAdminCtx.Load($myOps)
    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 01

#gavdcodebegin 02
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
#gavdcodeend 02

#gavdcodebegin 03
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
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsCsomRemoveSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsCsomRestoreSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RestoreDeletedSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsCsomRemoveDeletedSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveDeletedSite(
        $onfigFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsCsomCreateGroupForSite($spAdminCtx)
{
    $myOwners = @( "user@domain.onmicrosoft.com" )
    $myGroupParams = New-Object `
		Microsoft.Online.SharePoint.TenantManagement.GroupCreationParams($spAdminCtx)
    $myGroupParams.Owners = $myOwners
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.CreateGroupForSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom",
			"GroupForNewSiteCollectionModernPsCsom",
			"GroupForNewSiteCollAlias",
			$true,
			$myGroupParams)

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsCsomSetAdministratorSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.SetSiteAdmin(
        $configFile.AppSettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom",
        "user@domain.onmicrosoft.com",
        $true)

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsCsomRegisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RegisterHubSite(
        $configFile.AppSettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsCsomUnregisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.UnregisterHubSite(
        $onfigFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsCsomGetHubSiteCollectionProperties($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetHubSitePropertiesByUrl(
		$configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    Write-Host($myProps.Title)
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsCsomUpdateHubSiteCollectionProperties($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetHubSitePropertiesByUrl(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    $myProps.Title = $myProps.Title + "_Updated"
    $myProps.Update()

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    Write-Host($myProps.Title)
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsCsomAddSiteToHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.ConnectSiteToHubSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteForHub",
		$configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsCsomremoveSiteFromHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.DisconnectSiteFromHubSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteForHub")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 14

#gavdcodebegin 15
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
#gavdcodeend 15

#gavdcodebegin 16
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
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsCsomGetOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    Write-Host($myWeb.Title + " - " + $myWeb.Url + " - " + $myWeb.Id)
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsCsomUpdateOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.Description = "NewWebSiteModernPsCsom Description Updated"
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsCsomDeleteOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsCsomBreakSecurityInheritanceWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.BreakRoleInheritance($false, $true)
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsCsomResetSecurityInheritanceWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.ResetRoleInheritance()
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsCsomAddUserToSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
    $roleDefinition = New-Object `
				Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(
							[Microsoft.SharePoint.Client.RoleType]::Reader))
    $myWeb.RoleAssignments.Add($myUser, $roleDefinition)

    $spCtx.ExecuteQuery()
}
#gavdcodeend 22

#gavdcodebegin 23
Function SpPsCsomUpdateUserSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
    $roleDefinition = New-Object `
					Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
    $roleDefinition.Add($myWeb.RoleDefinitions.GetByType(
						[Microsoft.SharePoint.Client.RoleType]::Administrator))

    $myRoleAssignment = $myWeb.RoleAssignments.GetByPrincipal($myUser)
    $myRoleAssignment.ImportRoleDefinitionBindings($roleDefinition)

    $myRoleAssignment.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 23

#gavdcodebegin 24
Function SpPsCsomDeleteUserFromSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
    $myWeb.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 24

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

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
