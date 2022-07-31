
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShellWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnPPowerShellWithAccPw($FullSiteUrl)
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	if($fullSiteUrl -ne $null) {
		[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force

		$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
				-argumentlist $configFile.appsettings.UserName, $securePW
		Connect-PnPOnline -Url $FullSiteUrl -Credentials $myCredentials
	}
}

Function LoginPsPnPPowerShellWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function LoginPsPnPPowerShellWithCertificate
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}

Function LoginPsPnPPowerShellWithCertificateBase64
{
	# Using a Digital Certificate and Azure App Registration (Application)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 01
function SpPsPnpPowerShell_CreateOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	New-PnPList -Title "NewListPsPnp" -Template GenericList
	Disconnect-PnPOnline
}
#gavdcodeend 01

#gavdcodebegin 02
function SpPsPnpPowerShell_ReadAllList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$allLists = Get-PnPList

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title + " - " + $oneList.Id
	}
	Disconnect-PnPOnline
}
#gavdcodeend 02

#gavdcodebegin 03
function SpPsPnpPowerShell_ReadOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$myList = Get-PnPList -Identity "NewListPsPnp"

	Write-Host "List Description -" $myList.Description
	Disconnect-PnPOnline
}
#gavdcodeend 03

#gavdcodebegin 04
function SpPsPnpPowerShell_UpdateOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -Description "New List Description"
	Disconnect-PnPOnline
}
#gavdcodeend 04

#gavdcodebegin 05
function SpPsPnpPowerShell_DeleteOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	 Remove-PnPList -Identity "NewListPsPnp" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 05

#gavdcodebegin 06
function SpPsPnpPowerShell_AddOneFieldToList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$fieldXml = "<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	Add-PnPFieldFromXml -List "NewListPsPnp" -FieldXml $fieldXml
	Disconnect-PnPOnline
}
#gavdcodeend 06

#gavdcodebegin 07
function SpPsPnpPowerShell_ReadAllFieldsFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$allFields = Get-PnPField -List "NewListPsPnp"

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
	Disconnect-PnPOnline
}
#gavdcodeend 07

#gavdcodebegin 08
function SpPsPnpPowerShell_ReadOneFieldFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	$myField = Get-PnPField -List "NewListPsPnp" -Identity "MyMultilineField"

	Write-Host $myField.Id "-" $myField.TypeAsString
	Disconnect-PnPOnline
}
#gavdcodeend 08

#gavdcodebegin 09
function SpPsPnpPowerShell_UpdateOneFieldInList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" `
									-Values @{Description="New Field Description"}
	Disconnect-PnPOnline
}
#gavdcodeend 09

#gavdcodebegin 10
function SpPsPnpPowerShell_DeleteOneFieldFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Remove-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 10

#gavdcodebegin 11
function SpPsPnpPowerShell_BreakInheritanceList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -BreakRoleInheritance
	Disconnect-PnPOnline
}
#gavdcodeend 11

#gavdcodebegin 12
function SpPsPnpPowerShell_RestoreInheritanceList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -ResetRoleInheritance
	Disconnect-PnPOnline
}
#gavdcodeend 12

#gavdcodebegin 13
function SpPsPnpPowerShell_GetPermissionsList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Get-PnPListPermissions -Identity "NewListPsPnp" -PrincipalId 11
	#Get-PnPListPermissions -Identity "NewListPsPnp" `
	#			-PrincipalId (Get-PnPUser | ? Email -eq "user@domain.OnMicrosoft.com").Id
	#Get-PnPListPermissions -Identity "NewListPsPnp" `
	#			-PrincipalId (Get-PnPGroup -Identity "myGroup Members").Id
	Disconnect-PnPOnline
}
#gavdcodeend 13

#gavdcodebegin 14
function SpPsPnpPowerShell_AddPermissionsToList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListPermission -Identity "NewListPsPnp" `
						  -User "user@domain.OnMicrosoft.com" `
						  -AddRole "Full Control"
	Disconnect-PnPOnline
}
#gavdcodeend 14

#gavdcodebegin 15
function SpPsPnpPowerShell_DeletePermissionsFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShellWithAccPwDefault
	Set-PnPListPermission -Identity "NewListPsPnp" `
						  -User "user@domain.OnMicrosoft.com" `
						  -RemoveRole "Full Control"
	Disconnect-PnPOnline
}
#gavdcodeend 15


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#SpPsPnpPowerShell_CreateOneList
#SpPsPnpPowerShell_ReadAllList
#SpPsPnpPowerShell_ReadOneList
#SpPsPnpPowerShell_UpdateOneList
#SpPsPnpPowerShell_DeleteOneList
#SpPsPnpPowerShell_AddOneFieldToList
#SpPsPnpPowerShell_ReadAllFieldsFromList
#SpPsPnpPowerShell_ReadOneFieldFromList
#SpPsPnpPowerShell_UpdateOneFieldInList
#SpPsPnpPowerShell_DeleteOneFieldFromList
#SpPsPnpPowerShell_BreakInheritanceList
#SpPsPnpPowerShell_RestoreInheritanceList
#SpPsPnpPowerShell_GetPermissionsList
#SpPsPnpPowerShell_AddPermissionsToList
#SpPsPnpPowerShell_DeletePermissionsFromList

Write-Host "Done" 
