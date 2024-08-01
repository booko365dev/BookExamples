
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsSpPnpPowerShell_LoginWithAccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function PsSpPnpPowerShell_LoginWithAccPw($FullSiteUrl)
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

Function PsSpPnpPowerShell_LoginWithInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function PsSpPnpPowerShell_LoginWithCertificate
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

Function PsSpPnpPowerShell_LoginWithCertificateBase64
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

#gavdcodebegin 001
function PsSpPnpPowerShell_CreateOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	New-PnPList -Title "NewListPsPnp" -Template GenericList
	Disconnect-PnPOnline
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpPnpPowerShell_ReadAllList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$allLists = Get-PnPList

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title " - " $oneList.Id
	}
	Disconnect-PnPOnline
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpPnpPowerShell_ReadOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$myList = Get-PnPList -Identity "NewListPsPnp"

	Write-Host "List Description -" $myList.Description
	Disconnect-PnPOnline
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpPnpPowerShell_UpdateOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -Description "New List Description"
	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpPnpPowerShell_DeleteOneList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	 Remove-PnPList -Identity "NewListPsPnp" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpPnpPowerShell_AddOneFieldToList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$fieldXml = `
			"<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	Add-PnPFieldFromXml -List "NewListPsPnp" -FieldXml $fieldXml
	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpPnpPowerShell_ReadAllFieldsFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$allFields = Get-PnPField -List "NewListPsPnp"

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpPnpPowerShell_ReadOneFieldFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	$myField = Get-PnPField -List "NewListPsPnp" -Identity "MyMultilineField"

	Write-Host $myField.Id "-" $myField.TypeAsString
	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpPnpPowerShell_UpdateOneFieldInList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSitesWrite.Read
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" `
									-Values @{Description="New Field Description"}
	Disconnect-PnPOnline
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpPnpPowerShell_DeleteOneFieldFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Remove-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" -Force
	Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpPnpPowerShell_BreakInheritanceList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -BreakRoleInheritance
	Disconnect-PnPOnline
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpPnpPowerShell_RestoreInheritanceList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPList -Identity "NewListPsPnp" -ResetRoleInheritance
	Disconnect-PnPOnline
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpPnpPowerShell_GetPermissionsList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Get-PnPListPermissions -Identity "NewListPsPnp" -PrincipalId [IdUser]
	#Get-PnPListPermissions -Identity "NewListPsPnp" `
	#		-PrincipalId (Get-PnPUser | ? Email -eq "user@domain.OnMicrosoft.com").Id
	#Get-PnPListPermissions -Identity "NewListPsPnp" `
	#		-PrincipalId (Get-PnPGroup -Identity "myGroup Members").Id
	Disconnect-PnPOnline
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpPnpPowerShell_AddPermissionsToList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListPermission -Identity "NewListPsPnp" `
						  -User "user@domain.OnMicrosoft.com" `
						  -AddRole "Full Control"
	Disconnect-PnPOnline
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpPnpPowerShell_DeletePermissionsFromList
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = PsSpPnpPowerShell_LoginWithAccPwDefault
	Set-PnPListPermission -Identity "NewListPsPnp" `
						  -User "user@domain.OnMicrosoft.com" `
						  -RemoveRole "Full Control"
	Disconnect-PnPOnline
}
#gavdcodeend 015


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 015

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#PsSpPnpPowerShell_CreateOneList
#PsSpPnpPowerShell_ReadAllList
#PsSpPnpPowerShell_ReadOneList
#PsSpPnpPowerShell_UpdateOneList
#PsSpPnpPowerShell_DeleteOneList
#PsSpPnpPowerShell_AddOneFieldToList
#PsSpPnpPowerShell_ReadAllFieldsFromList
#PsSpPnpPowerShell_ReadOneFieldFromList
#PsSpPnpPowerShell_UpdateOneFieldInList
#PsSpPnpPowerShell_DeleteOneFieldFromList
#PsSpPnpPowerShell_BreakInheritanceList
#PsSpPnpPowerShell_RestoreInheritanceList
#PsSpPnpPowerShell_GetPermissionsList
#PsSpPnpPowerShell_AddPermissionsToList
#PsSpPnpPowerShell_DeletePermissionsFromList

Write-Host "Done" 
