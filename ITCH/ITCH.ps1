##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
Function ConnectPsExo_Interactive
{
    Connect-ExchangeOnline -UserPrincipalName $configFile.appsettings.UserName
}
#gavdcodeend 001

#gavdcodebegin 002
Function ConnectPsExo_Inline
{
    Connect-ExchangeOnline -UserPrincipalName $configFile.appsettings.UserName `
                           -InlineCredential
}
#gavdcodeend 002

#gavdcodebegin 003
Function ConnectPsExo_AccPw
{
    [SecureString]$securePW = ConvertTo-SecureString `
                            -String $configFile.appsettings.UserPw -AsPlainText -Force
    $myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
                            -argumentlist $configFile.appsettings.UserName, $securePW
    Connect-ExchangeOnline -Credential $myCredentials
}
#gavdcodeend 003

#gavdcodebegin 004
Function ConnectPsExo_CertificateThumbprint
{
    Connect-ExchangeOnline `
                -CertificateThumbPrint $configFile.appsettings.CertificateThumbprint `
                -AppID $configFile.appsettings.ClientIdWithCert `
                -Organization $configFile.appsettings.TenantName
}
#gavdcodeend 004

#gavdcodebegin 005
Function ConnectPsExo_CertificateFile
{
    [SecureString]$securePW = ConvertTo-SecureString `
                            -String $configFile.appsettings.UserPw -AsPlainText -Force
    Connect-ExchangeOnline `
                -CertificateFilePath $configFile.appsettings.CertificateFilePath `
                -CertificatePassword $securePW `
                -AppID $configFile.appsettings.ClientIdWithCert `
                -Organization $configFile.appsettings.TenantName
}
#gavdcodeend 005

#gavdcodebegin 006
Function ConnectPsExo_ManagedIdentitySystem
{
    Connect-ExchangeOnline -ManagedIdentity `
                    -Organization $configFile.appsettings.TenantName
}
#gavdcodeend 006

#gavdcodebegin 007
Function ConnectPsExo_ManagedIdentityUser
{
    Connect-ExchangeOnline -ManagedIdentity `
                    -Organization $configFile.appsettings.TenantName
                    -ManagedIdentityAccountId $configFile.appsettings.UserName
}
#gavdcodeend 007

#gavdcodebegin 008
Function ConnectPsExo_ReleaseConnection
{
    Disconnect-ExchangeOnline -Confirm:$false
}
#gavdcodeend 008


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 009
Function ExchangePsExo_GetMailboxesByPropertySet
{
    Get-EXOMailbox -PropertySets Minimum,Policy
}
#gavdcodeend 009

#gavdcodebegin 010
Function ExchangePsExo_GetMailboxesByProperty
{
    Get-EXOMailbox -Properties UserPrincipalName,Alias
}
#gavdcodeend 010

#gavdcodebegin 011
Function ExchangePsExo_GetMailboxesBySetAndProperty()
{
    Get-EXOMailbox -PropertySets Hold,Moderation -Properties UserPrincipalName,Alias
}
#gavdcodeend 011

#gavdcodebegin 012
Function ExchangePsExo_GetClientAccessSettings
{
    Get-EXOCASMailbox -Identity $configFile.appsettings.UserName
}
#gavdcodeend 012

#gavdcodebegin 013
Function ExchangePsExo_GetEmailboxPermissions
{
    Get-EXOMailboxPermission -Identity $configFile.appsettings.UserName
}
#gavdcodeend 013

#gavdcodebegin 014
Function ExchangePsExo_GetEmailboxStatistics
{
    Get-EXOMailboxStatistics -Identity $configFile.appsettings.UserName
}
#gavdcodeend 014

#gavdcodebegin 015
Function ExchangePsExo_GetFolderPermissions
{
    $myIdentifier = $configFile.appsettings.UserName + ":\Inbox"
    Get-EXOMailboxFolderPermission -Identity $myIdentifier
}
#gavdcodeend 015

#gavdcodebegin 016
Function ExchangePsExo_GetFolderStatisticsAll
{
    Get-EXOMailboxFolderStatistics -Identity $configFile.appsettings.UserName
}
#gavdcodeend 016

#gavdcodebegin 017
Function ExchangePsExo_GetFolderStatisticsOne
{
    Get-EXOMailboxFolderStatistics -Identity $configFile.appsettings.UserName `
                                                            -FolderScope Calendar
}
#gavdcodeend 017

#gavdcodebegin 018
Function ExchangePsExo_GetDeviceStatistics
{
    Get-EXOMobileDeviceStatistics -Mailbox $configFile.appsettings.UserName
}
#gavdcodeend 018

#gavdcodebegin 019
Function ExchangePsExo_GetRecipients
{
    Get-EXORecipient -Identity $configFile.appsettings.UserName
}
#gavdcodeend 019

#gavdcodebegin 020
Function ExchangePsExo_GetRecipientPermissions
{
    Get-EXORecipientPermission -ResultSize 5
}
#gavdcodeend 020

#gavdcodebegin 021
Function ExchangePsExo_CreateMailbox
{
	New-Mailbox -Alias somebody -Name Some -FirstName Some -LastName Body `
				-DisplayName "Some Body" `
				-MicrosoftOnlineServicesID somebody@domain.onmicrosoft.com `
				-Password `
                    (ConvertTo-SecureString -String "MyPw1234" -AsPlainText -Force) `
				-ResetPasswordOnNextLogon $true
}
#gavdcodeend 021

#gavdcodebegin 022
Function ExchangePsExo_DeleteMailbox
{
	Remove-Mailbox -Identity "Some Body"
}
#gavdcodeend 022

#gavdcodebegin 023
Function ExchangePsExo_CreateMailcontact
{
	New-MailContact -Name "Some Contact" -ExternalEmailAddress scontact@domain.com
}
#gavdcodeend 023

#gavdcodebegin 024
Function ExchangePsExo_DeleteMailcontact
{
	Remove-MailContact -Identity "Some Contact"
}
#gavdcodeend 024

#gavdcodebegin 025
Function ExchangePsExo_CreateMailuser
{
	New-MailUser -Name "Some Body" -Alias somebody `
				 -ExternalEmailAddress sbody@domain.com `
				 -FirstName Some -LastName Body `
				 -MicrosoftOnlineServicesID sobody@domain.onmicrosoft.com `
				 -Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force)
}
#gavdcodeend 025

#gavdcodebegin 026
Function ExchangePsExo_GetPermissions
{
	$myPerms = Get-ManagementRoleAssignment
	foreach ($onePerm in $myPerms) {
		Write-Host $onePerm.Name - $onePerm.Role - $onePerm.RoleAssigneeType - `
															$onePerm.RoleAssigneeName
	}
}
#gavdcodeend 026

#gavdcodebegin 027
Function ExchangePsExo_EnablePsAccess
{
	Set-User -Identity user@dominio.onmicrosoft.com -RemotePowerShellEnabled $true
}
#gavdcodeend 027

#gavdcodebegin 028
Function ExchangePsExo_GetEnablePsAccess
{
	# For one user
	Get-User -Identity "Name Surname" | Format-List RemotePowerShellEnabled

	# For all users
	Get-User -ResultSize unlimited | Format-Table -Auto `
											Name,DisplayName,RemotePowerShellEnabled
}
#gavdcodeend 028

#gavdcodebegin 029
Function ExchangePsExo_BlockIP
{
	Set-OrganizationConfig -IPListBlocked @{add="111.2222.333.444"}
}
#gavdcodeend 029

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 029 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

#ConnectPsExo_Interactive
#ConnectPsExo_Inline
ConnectPsExo_AccPw
#ConnectPsExo_CertificateThumbprint
#ConnectPsExo_CertificateFile
#ConnectPsExo_ManagedIdentitySystem
#ConnectPsExo_ManagedIdentityUser

#ExchangePsExo_GetMailboxesByPropertySet
#ExchangePsExo_GetMailboxesByProperty
#ExchangePsExo_GetMailboxesBySetAndProperty
#ExchangePsExo_GetClientAccessSettings
#ExchangePsExo_GetEmailboxPermissions
#ExchangePsExo_GetEmailboxStatistics
#ExchangePsExo_GetFolderPermissions
#ExchangePsExo_GetFolderStatisticsAll
#ExchangePsExo_GetFolderStatisticsOne
#ExchangePsExo_GetDeviceStatistics
#ExchangePsExo_GetRecipients
#ExchangePsExo_GetRecipientPermissions
#ExchangePsExo_CreateMailbox
#ExchangePsExo_DeleteMailbox
#ExchangePsExo_CreateMailcontact
#ExchangePsExo_DeleteMailcontact
#ExchangePsExo_CreateMailuser
#ExchangePsExo_BlockIP
#ExchangePsExo_GetPermissions
#ExchangePsExo_EnablePsAccess
#ExchangePsExo_GetEnablePsAccess
#ExchangePsExo_BlockIP

Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Done"  

