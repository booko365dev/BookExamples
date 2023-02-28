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

#gavdcodebegin 001
Function SpCsCsom_CreateOneSiteCollection($spAdminCtx)
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
#gavdcodeend 001

#gavdcodebegin 002
Function SpCsCsom_FindWebTemplates($spAdminCtx)
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
#gavdcodeend 002

#gavdcodebegin 003
Function SpCsCsom_ReadAllSiteCollections($spAdminCtx)
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
#gavdcodeend 003

#gavdcodebegin 004
Function SpCsCsom_RemoveSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 004

#gavdcodebegin 005
Function SpCsCsom_RestoreSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RestoreDeletedSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 005

#gavdcodebegin 006
Function SpCsCsom_RemoveDeletedSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveDeletedSite(
        $onfigFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 006

#gavdcodebegin 007
Function SpCsCsom_CreateGroupForSite($spAdminCtx)
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
#gavdcodeend 007

#gavdcodebegin 008
Function SpCsCsom_SetAdministratorSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.SetSiteAdmin(
        $configFile.AppSettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom",
        "user@domain.onmicrosoft.com",
        $true)

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 008

#gavdcodebegin 009
Function SpCsCsom_RegisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RegisterHubSite(
        $configFile.AppSettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 009

#gavdcodebegin 010
Function SpCsCsom_UnregisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.UnregisterHubSite(
        $onfigFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 010

#gavdcodebegin 011
Function SpCsCsom_GetHubSiteCollectionProperties($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myProps = $myTenant.GetHubSitePropertiesByUrl(
		$configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.Load($myProps)
    $spAdminCtx.ExecuteQuery()

    Write-Host($myProps.Title)
}
#gavdcodeend 011

#gavdcodebegin 012
Function SpCsCsom_UpdateHubSiteCollectionProperties($spAdminCtx)
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
#gavdcodeend 012

#gavdcodebegin 013
Function SpCsCsom_AddSiteToHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.ConnectSiteToHubSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteForHub",
		$configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 013

#gavdcodebegin 014
Function SpCsCsom_removeSiteFromHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.DisconnectSiteFromHubSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteForHub")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 014

#gavdcodebegin 015
Function SpCsCsom_CreateOneWebInSiteCollection($spCtx)
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
#gavdcodeend 015

#gavdcodebegin 016
Function SpCsCsom_GetWebsInSiteCollection($spCtx)
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
#gavdcodeend 016

#gavdcodebegin 017
Function SpCsCsom_GetOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    Write-Host($myWeb.Title + " - " + $myWeb.Url + " - " + $myWeb.Id)
}
#gavdcodeend 017

#gavdcodebegin 018
Function SpCsCsom_UpdateOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.Description = "NewWebSiteModernPsCsom Description Updated"
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 018

#gavdcodebegin 019
Function SpCsCsom_DeleteOneWebInSiteCollection()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 019

#gavdcodebegin 020
Function SpCsCsom_BreakSecurityInheritanceWeb()
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
#gavdcodeend 020

#gavdcodebegin 021
Function SpCsCsom_ResetSecurityInheritanceWeb()
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
#gavdcodeend 021

#gavdcodebegin 022
Function SpCsCsom_AddUserToSecurityRoleInWeb()
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
#gavdcodeend 022

#gavdcodebegin 023
Function SpCsCsom_UpdateUserSecurityRoleInWeb()
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
#gavdcodeend 023

#gavdcodebegin 024
Function SpCsCsom_DeleteUserFromSecurityRoleInWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = LoginCsom($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
    $myWeb.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 024

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsCsom
$spAdminCtx = LoginAdminCsom

#SpCsCsom_CreateOneSiteCollection $spAdminCtx
#SpCsCsom_FindWebTemplates $spAdminCtx
#SpCsCsom_ReadAllSiteCollections $spAdminCtx
#SpCsCsom_RemoveSiteCollection $spAdminCtx
#SpCsCsom_RestoreSiteCollection $spAdminCtx
#SpCsCsom_RemoveDeletedSiteCollection $spAdminCtx
#SpCsCsom_CreateGroupForSite $spAdminCtx
#SpCsCsom_RegisterAsHubSiteCollection $spAdminCtx
#SpCsCsom_UnregisterAsHubSiteCollection $spAdminCtx
#SpCsCsom_GetHubSiteCollectionProperties $spAdminCtx
#SpCsCsom_UpdateHubSiteCollectionProperties $spAdminCtx
#SpCsCsom_AddSiteToHubSiteCollection $spAdminCtx
#SpCsCsom_removeSiteFromHubSiteCollection $spAdminCtx

#SpCsCsom_CreateOneWebInSiteCollection $spCtx
#SpCsCsom_GetWebsInSiteCollection $spCtx
#SpCsCsom_GetOneWebInSiteCollection
#SpCsCsom_UpdateOneWebInSiteCollection
#SpCsCsom_DeleteOneWebInSiteCollection
#SpCsCsom_BreakSecurityInheritanceWeb
#SpCsCsom_ResetSecurityInheritanceWeb
#SpCsCsom_AddUserToSecurityRoleInWeb
#SpCsCsom_UpdateUserSecurityRoleInWeb
#SpCsCsom_DeleteUserFromSecurityRoleInWeb

Write-Host "Done"
