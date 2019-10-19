
Function ConnectPsOnlBA() 
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
				$configFile.appsettings.exUserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential -ArgumentList `
				$configFile.appsettings.exUserName, $securePW
	$mySession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri `
				https://outlook.office365.com/powershell-liveid/ -Authentication Basic `
				-AllowRedirection -Credential $myCredentials
	Import-PSSession $mySession -AllowClobber
}
#-----------------------------------------------------------------------------------------

Function ExPsEolGetPermissions()
{
	$myPerms = Get-ManagementRoleAssignment
	foreach ($onePerm in $myPerms) {
		Write-Host $onePerm.Name - $onePerm.Role - $onePerm.RoleAssigneeType - `
															$onePerm.RoleAssigneeName
	}
}

Function ExPsEolEnablePsAccess()
{
	Set-User -Identity user@dominio.onmicrosoft.com -RemotePowerShellEnabled $true
}

Function ExPsEolEnablePsAccess()
{
	# For one user
	Get-User -Identity "Name Surname" | Format-List RemotePowerShellEnabled

	# For all users
	Get-User -ResultSize unlimited | Format-Table -Auto `
											Name,DisplayName,RemotePowerShellEnabled
}

Function ExPsEolCreateMailbox()
{
	New-Mailbox -Alias somebody -Name Some -FirstName Some -LastName Body `
				-DisplayName "Some Body" `
				-MicrosoftOnlineServicesID somebody@domain.onmicrosoft.com `
				-Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force) `
				-ResetPasswordOnNextLogon $true}

Function ExPsEolDeleteMailbox()
{
	Remove-MsolUser -UserPrincipalName "Some Body" -RemoveFromRecycleBin true
}

Function ExPsEolCreateMailcontact()
{
	New-MailContact -Name "Some Body" -ExternalEmailAddress sbody@domain.com
}

Function ExPsEolCreateMailuser()
{
	New-MailUser -Name "Some Body" -Alias somebody `
				 -ExternalEmailAddress sbody@domain.com `
				 -FirstName Some -LastName Body `
				 -MicrosoftOnlineServicesID sobody@domain.onmicrosoft.com `
				 -Password (ConvertTo-SecureString -String "SecPW56%&" -AsPlainText -Force)
}

Function ExPsEolBlockIP()
{
	Set-OrganizationConfig -IPListBlocked @{add="111.2222.333.444"}
}


#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\exPs.values.config"

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

