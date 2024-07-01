
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


function PsSpCsom_Login
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

function PsSpCsom_LoginWithUrl($WebFullUrl)
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext($WebFullUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

function PsSpCsom_LoginAdmin
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


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpCsom_CreateOneSiteCollection($spAdminCtx)
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
function PsSpCsom_GetWebTemplates($spAdminCtx)
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
function PsSpCsom_ReadAllSiteCollections($spAdminCtx)
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
function PsSpCsom_RemoveSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpCsom_RestoreSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RestoreDeletedSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpCsom_RemoveDeletedSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RemoveDeletedSite(
        $onfigFile.appsettings.spBaseUrl + "/sites/NewSiteCollectionModernPsCsom")
    
	$spAdminCtx.ExecuteQuery()
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpCsom_CreateGroupForSite($spAdminCtx)
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
function PsSpCsom_SetAdministratorSiteCollection($spAdminCtx)
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
function PsSpCsom_RegisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.RegisterHubSite(
        $configFile.AppSettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpCsom_UnregisterAsHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.UnregisterHubSite(
        $onfigFile.appsettings.SiteBaseUrl + "/sites/NewSiteCollectionModernPsCsom")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpCsom_GetHubSiteCollectionProperties($spAdminCtx)
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
function PsSpCsom_UpdateHubSiteCollectionProperties($spAdminCtx)
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
function PsSpCsom_AddSiteToHubSiteCollection($spAdminCtx)
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
function PsSpCsom_removeSiteFromHubSiteCollection($spAdminCtx)
{
	$myTenant = New-Object `
					Microsoft.Online.SharePoint.TenantAdministration.Tenant($spAdminCtx)
    $myTenant.DisconnectSiteFromHubSite(
        $configFile.appsettings.SiteBaseUrl + "/sites/NewSiteForHub")

    $spAdminCtx.ExecuteQuery()
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpCsom_CreateOneWebInSiteCollection($spCtx)
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
function PsSpCsom_GetWebsInSiteCollection($spCtx)
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
function PsSpCsom_GetOneWebInSiteCollection
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    Write-Host($myWeb.Title + " - " + $myWeb.Url + " - " + $myWeb.Id)
}
#gavdcodeend 017

#gavdcodebegin 018
function PsSpCsom_UpdateOneWebInSiteCollection
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.Description = "NewWebSiteModernPsCsom Description Updated"
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 018

#gavdcodebegin 019
function PsSpCsom_DeleteOneWebInSiteCollection
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web
    $myWeb.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 019

#gavdcodebegin 020
function PsSpCsom_BreakSecurityInheritanceWeb()
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.BreakRoleInheritance($false, $true)
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 020

#gavdcodebegin 021
function PsSpCsom_ResetSecurityInheritanceWeb
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web
    $spCtx.Load($myWeb)
    $spCtx.ExecuteQuery()

    $myWeb.ResetRoleInheritance()
    $myWeb.Update()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 021

#gavdcodebegin 022
function PsSpCsom_AddUserToSecurityRoleInWeb
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

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
function PsSpCsom_UpdateUserSecurityRoleInWeb
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

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
function PsSpCsom_DeleteUserFromSecurityRoleInWeb
{
    $myWebFullUrl = $configFile.appsettings.SiteCollUrl + "/NewWebSiteModernPsCsom"
    $spCtx = PsSpCsom_LoginWithUrl($myWebFullUrl)

    $myWeb = $spCtx.Web

    $myUser = $myWeb.EnsureUser($configFile.appsettings.UserName)
    $myWeb.RoleAssignments.GetByPrincipal($myUser).DeleteObject()

    $spCtx.ExecuteQuery()
}
#gavdcodeend 024

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 024 ***

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.Client.Tenant.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = PsSpCsom_Login
$spAdminCtx = PsSpCsom_LoginAdmin

#PsSpCsom_CreateOneSiteCollection $spAdminCtx
#PsSpCsom_GetWebTemplates $spAdminCtx
#PsSpCsom_ReadAllSiteCollections $spAdminCtx
#PsSpCsom_RemoveSiteCollection $spAdminCtx
#PsSpCsom_RestoreSiteCollection $spAdminCtx
#PsSpCsom_RemoveDeletedSiteCollection $spAdminCtx
#PsSpCsom_CreateGroupForSite $spAdminCtx
#PsSpCsom_RegisterAsHubSiteCollection $spAdminCtx
#PsSpCsom_UnregisterAsHubSiteCollection $spAdminCtx
#PsSpCsom_GetHubSiteCollectionProperties $spAdminCtx
#PsSpCsom_UpdateHubSiteCollectionProperties $spAdminCtx
#PsSpCsom_AddSiteToHubSiteCollection $spAdminCtx
#PsSpCsom_removeSiteFromHubSiteCollection $spAdminCtx

#PsSpCsom_CreateOneWebInSiteCollection $spCtx
#PsSpCsom_GetWebsInSiteCollection $spCtx
#PsSpCsom_GetOneWebInSiteCollection
#PsSpCsom_UpdateOneWebInSiteCollection
#PsSpCsom_DeleteOneWebInSiteCollection
#PsSpCsom_BreakSecurityInheritanceWeb
#PsSpCsom_ResetSecurityInheritanceWeb
#PsSpCsom_AddUserToSecurityRoleInWeb
#PsSpCsom_UpdateUserSecurityRoleInWeb
#PsSpCsom_DeleteUserFromSecurityRoleInWeb

Write-Host "Done"
