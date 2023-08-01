
Function ConnectPsOnlBA  #*** LEGACY CODE ***
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList `
				$configFile.appsettings.UserName, $securePW
	$mySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri `
				https://outlook.office365.com/powershell-liveid/ -Authentication Basic `
				-AllowRedirection -Credential $myCredentials
	Import-PSSession $mySession -AllowClobber
}
#-----------------------------------------------------------------------------------------

#gavdcodebegin 001
Function ExchangePsEol_GetPermissions  #*** LEGACY CODE ***
{
	$myPerms = Get-ManagementRoleAssignment
	foreach ($onePerm in $myPerms) {
		Write-Host $onePerm.Name - $onePerm.Role - $onePerm.RoleAssigneeType - `
															$onePerm.RoleAssigneeName
	}
}
#gavdcodeend 001

#gavdcodebegin 002
Function ExchangePsEol_EnablePsAccess  #*** LEGACY CODE ***
{
	Set-User -Identity user@dominio.onmicrosoft.com -RemotePowerShellEnabled $true
}
#gavdcodeend 002

#gavdcodebegin 003
Function ExchangePsEol_EnablePsAccess  #*** LEGACY CODE ***
{
	# For one user
	Get-User -Identity "Name Surname" | Format-List RemotePowerShellEnabled

	# For all users
	Get-User -ResultSize unlimited | Format-Table -Auto `
											Name,DisplayName,RemotePowerShellEnabled
}
#gavdcodeend 003

#gavdcodebegin 004
Function ExchangePsEol_CreateMailbox  #*** LEGACY CODE ***
{
	New-Mailbox -Alias somebody -Name Some -FirstName Some -LastName Body `
				-DisplayName "Some Body" `
				-MicrosoftOnlineServicesID somebody@domain.onmicrosoft.com `
				-Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force) `
				-ResetPasswordOnNextLogon $true}
#gavdcodeend 004

#gavdcodebegin 005
Function ExchangePsEol_DeleteMailbox  #*** LEGACY CODE ***
{
	Remove-MsolUser -UserPrincipalName "Some Body" -RemoveFromRecycleBin true
}
#gavdcodeend 005

#gavdcodebegin 006
Function ExchangePsEol_CreateMailcontact  #*** LEGACY CODE ***
{
	New-MailContact -Name "Some Body" -ExternalEmailAddress sbody@domain.com
}
#gavdcodeend 006

#gavdcodebegin 007
Function ExchangePsEol_CreateMailuser  #*** LEGACY CODE ***
{
	New-MailUser -Name "Some Body" -Alias somebody `
				 -ExternalEmailAddress sbody@domain.com `
				 -FirstName Some -LastName Body `
				 -MicrosoftOnlineServicesID sobody@domain.onmicrosoft.com `
				 -Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force)
}
#gavdcodeend 007

#gavdcodebegin 008
Function ExchangePsEol_BlockIP  #*** LEGACY CODE ***
{
	Set-OrganizationConfig -IPListBlocked @{add="111.2222.333.444"}
}
#gavdcodeend 008


#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

ConnectPsOnlBA

#ExchangePsEol_GetPermissions
#ExchangePsEol_EnablePsAccess
#ExchangePsEol_BlockIP
#ExchangePsEol_CreateMailbox
#ExchangePsEol_DeleteMailbox
#ExchangePsEol_CreateMailcontact
#ExchangePsEol_CreateMailuser

$currentSession = Get-PSSession
Remove-PSSession -Session $currentSession

Write-Host "Done"  
