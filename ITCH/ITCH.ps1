##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsExo_ConnectInteractive
{
    Connect-ExchangeOnline -UserPrincipalName $cnfUserName
}
#gavdcodeend 001

#gavdcodebegin 002
function PsExo_ConnectInline
{
    Connect-ExchangeOnline -UserPrincipalName $cnfUserName `
                           -InlineCredential
}
#gavdcodeend 002

#gavdcodebegin 003
function PsExo_ConnectAccPw
{
    [SecureString]$securePW = ConvertTo-SecureString `
                            -String $cnfUserPw -AsPlainText -Force
    $myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
                            -argumentlist $cnfUserName, $securePW
    Connect-ExchangeOnline -Credential $myCredentials
}
#gavdcodeend 003

#gavdcodebegin 004
function PsExo_ConnectCertificateThumbprint
{
    Connect-ExchangeOnline `
                -CertificateThumbPrint $cnfCertificateThumbprint `
                -AppID $cnfClientIdWithCert `
                -Organization $cnfTenantName
}
#gavdcodeend 004

#gavdcodebegin 005
function PsExo_ConnectCertificateFile
{
    [SecureString]$securePW = ConvertTo-SecureString `
                            -String $cnfUserPw -AsPlainText -Force
    Connect-ExchangeOnline `
                -CertificateFilePath $cnfCertificateFilePath `
                -CertificatePassword $securePW `
                -AppID $cnfClientIdWithCert `
                -Organization $cnfTenantName
}
#gavdcodeend 005

#gavdcodebegin 006
function PsExo_ConnectManagedIdentitySystem
{
    Connect-ExchangeOnline -ManagedIdentity `
                    -Organization $cnfTenantName
}
#gavdcodeend 006

#gavdcodebegin 007
function PsExo_ConnectManagedIdentityUser
{
    Connect-ExchangeOnline -ManagedIdentity `
                    -Organization $cnfTenantName
                    -ManagedIdentityAccountId $cnfUserName
}
#gavdcodeend 007

#gavdcodebegin 008
function PsExo_ConnectReleaseConnection
{
    Disconnect-ExchangeOnline -Confirm:$false
}
#gavdcodeend 008


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 009
function PsExchangeExo_GetMailboxesByPropertySet
{
    Get-EXOMailbox -PropertySets Minimum,Policy
}
#gavdcodeend 009

#gavdcodebegin 010
function PsExchangeExo_GetMailboxesByProperty
{
    Get-EXOMailbox -Properties UserPrincipalName,Alias
}
#gavdcodeend 010

#gavdcodebegin 011
function PsExchangeExo_GetMailboxesBySetAndProperty()
{
    Get-EXOMailbox -PropertySets Hold,Moderation -Properties UserPrincipalName,Alias
}
#gavdcodeend 011

#gavdcodebegin 012
function PsExchangeExo_GetClientAccessSettings
{
    Get-EXOCASMailbox -Identity $cnfUserName
}
#gavdcodeend 012

#gavdcodebegin 013
function PsExchangeExo_GetEmailboxPermissions
{
    Get-EXOMailboxPermission -Identity $cnfUserName
}
#gavdcodeend 013

#gavdcodebegin 014
function PsExchangeExo_GetEmailboxStatistics
{
    Get-EXOMailboxStatistics -Identity $cnfUserName
}
#gavdcodeend 014

#gavdcodebegin 015
function PsExchangeExo_GetFolderPermissions
{
    $myIdentifier = $cnfUserName + ":\Inbox"
    Get-EXOMailboxFolderPermission -Identity $myIdentifier
}
#gavdcodeend 015

#gavdcodebegin 016
function PsExchangeExo_GetFolderStatisticsAll
{
    Get-EXOMailboxFolderStatistics -Identity $cnfUserName
}
#gavdcodeend 016

#gavdcodebegin 017
function PsExchangeExo_GetFolderStatisticsOne
{
    Get-EXOMailboxFolderStatistics -Identity $cnfUserName `
                                                            -FolderScope Calendar
}
#gavdcodeend 017

#gavdcodebegin 018
function PsExchangeExo_GetDeviceStatistics
{
    Get-EXOMobileDeviceStatistics -Mailbox $cnfUserName
}
#gavdcodeend 018

#gavdcodebegin 019
function PsExchangeExo_GetRecipients
{
    Get-EXORecipient -Identity $cnfUserName
}
#gavdcodeend 019

#gavdcodebegin 020
function PsExchangeExo_GetRecipientPermissions
{
    Get-EXORecipientPermission -ResultSize 5
}
#gavdcodeend 020

#gavdcodebegin 021
function PsExchangeExo_CreateMailbox
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
function PsExchangeExo_DeleteMailbox
{
	Remove-Mailbox -Identity "Some Body"
}
#gavdcodeend 022

#gavdcodebegin 023
function PsExchangeExo_CreateMailcontact
{
	New-MailContact -Name "Some Contact" -ExternalEmailAddress scontact@domain.com
}
#gavdcodeend 023

#gavdcodebegin 024
function PsExchangeExo_DeleteMailcontact
{
	Remove-MailContact -Identity "Some Contact"
}
#gavdcodeend 024

#gavdcodebegin 025
function PsExchangeExo_CreateMailuser
{
	New-MailUser -Name "Some Body" -Alias somebody `
				 -ExternalEmailAddress sbody@domain.com `
				 -FirstName Some -LastName Body `
				 -MicrosoftOnlineServicesID sobody@domain.onmicrosoft.com `
				 -Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force)
}
#gavdcodeend 025

#gavdcodebegin 026
function PsExchangeExo_GetPermissions
{
	$myPerms = Get-ManagementRoleAssignment
	foreach ($onePerm in $myPerms) {
		Write-Host $onePerm.Name - $onePerm.Role - $onePerm.RoleAssigneeType - `
															$onePerm.RoleAssigneeName
	}
}
#gavdcodeend 026

#gavdcodebegin 027
function PsExchangeExo_EnablePsAccess
{
	Set-User -Identity user@dominio.onmicrosoft.com -RemotePowerShellEnabled $true
}
#gavdcodeend 027

#gavdcodebegin 028
function PsExchangeExo_GetEnablePsAccess
{
	# For one user
	Get-User -Identity "Name Surname" | Format-List RemotePowerShellEnabled

	# For all users
	Get-User -ResultSize unlimited | Format-Table -Auto `
											Name,DisplayName,RemotePowerShellEnabled
}
#gavdcodeend 028

#gavdcodebegin 029
function PsExchangeExo_BlockIP
{
	Set-OrganizationConfig -IPListBlocked @{add="111.2222.333.444"}
}
#gavdcodeend 029

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 029 ***

#region ConfigValuesCS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value     # https://domain.onmicrosoft.com
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value   # https://domain.sharepoint.com
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value  # https://domain-admin.sharepoint.com
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value   # https://domain.sharepoint.com/sites/TestSite
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

#PsExo_ConnectInteractive
#PsExo_ConnectInline
PsExo_ConnectAccPw
#PsExo_ConnectCertificateThumbprint
#PsExo_ConnectCertificateFile
#PsExo_ConnectManagedIdentitySystem
#PsExo_ConnectManagedIdentityUser

#PsExchangeExo_GetMailboxesByPropertySet
#PsExchangeExo_GetMailboxesByProperty
#PsExchangeExo_GetMailboxesBySetAndProperty
#PsExchangeExo_GetClientAccessSettings
#PsExchangeExo_GetEmailboxPermissions
#PsExchangeExo_GetEmailboxStatistics
#PsExchangeExo_GetFolderPermissions
#PsExchangeExo_GetFolderStatisticsAll
#PsExchangeExo_GetFolderStatisticsOne
#PsExchangeExo_GetDeviceStatistics
#PsExchangeExo_GetRecipients
#PsExchangeExo_GetRecipientPermissions
#PsExchangeExo_CreateMailbox
#PsExchangeExo_DeleteMailbox
#PsExchangeExo_CreateMailcontact
#PsExchangeExo_DeleteMailcontact
#PsExchangeExo_CreateMailuser
#PsExchangeExo_BlockIP
#PsExchangeExo_GetPermissions
#PsExchangeExo_EnablePsAccess
#PsExchangeExo_GetEnablePsAccess
#PsExchangeExo_BlockIP

Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Done"  

