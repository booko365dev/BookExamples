
Function ConnectPsOnlBA() 
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

#gavdcodebegin 01
Function ExPsEolGetPermissions()
{
	$myPerms = Get-ManagementRoleAssignment
	foreach ($onePerm in $myPerms) {
		Write-Host $onePerm.Name - $onePerm.Role - $onePerm.RoleAssigneeType - `
															$onePerm.RoleAssigneeName
	}
}
#gavdcodeend 01

#gavdcodebegin 02
Function ExPsEolEnablePsAccess()
{
	Set-User -Identity user@dominio.onmicrosoft.com -RemotePowerShellEnabled $true
}
#gavdcodeend 02

#gavdcodebegin 03
Function ExPsEolEnablePsAccess()
{
	# For one user
	Get-User -Identity "Name Surname" | Format-List RemotePowerShellEnabled

	# For all users
	Get-User -ResultSize unlimited | Format-Table -Auto `
											Name,DisplayName,RemotePowerShellEnabled
}
#gavdcodeend 03

#gavdcodebegin 04
Function ExPsEolCreateMailbox()
{
	New-Mailbox -Alias somebody -Name Some -FirstName Some -LastName Body `
				-DisplayName "Some Body" `
				-MicrosoftOnlineServicesID somebody@domain.onmicrosoft.com `
				-Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force) `
				-ResetPasswordOnNextLogon $true}
#gavdcodeend 04

#gavdcodebegin 05
Function ExPsEolDeleteMailbox()
{
	Remove-MsolUser -UserPrincipalName "Some Body" -RemoveFromRecycleBin true
}
#gavdcodeend 05

#gavdcodebegin 06
Function ExPsEolCreateMailcontact()
{
	New-MailContact -Name "Some Body" -ExternalEmailAddress sbody@domain.com
}
#gavdcodeend 06

#gavdcodebegin 07
Function ExPsEolCreateMailuser()
{
	New-MailUser -Name "Some Body" -Alias somebody `
				 -ExternalEmailAddress sbody@domain.com `
				 -FirstName Some -LastName Body `
				 -MicrosoftOnlineServicesID sobody@domain.onmicrosoft.com `
				 -Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force)
}
#gavdcodeend 07

#gavdcodebegin 08
Function ExPsEolBlockIP()
{
	Set-OrganizationConfig -IPListBlocked @{add="111.2222.333.444"}
}
#gavdcodeend 08


#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

ConnectPsOnlBA

#ExPsEolGetPermissions
#ExPsEolEnablePsAccess
#ExPsEolBlockIP
#ExPsEolCreateMailbox
#ExPsEolDeleteMailbox
#ExPsEolCreateMailcontact
#ExPsEolCreateMailuser

$currentSession = Get-PSSession
Remove-PSSession -Session $currentSession

Write-Host "Done"  
