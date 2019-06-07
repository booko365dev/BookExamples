
Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

Function SpPsPnpCreateOneList()
{
	New-PnPList -Title "NewListPsPnp" -Template GenericList
}

Function SpPsPnpReadAllList()
{
	$allLists = Get-PnPList

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title + " - " + $oneList.Id
	}
}

Function SpPsPnpReadOneList()
{
	$myList = Get-PnPList -Identity "NewListPsPnp"

	Write-Host "List Description -" $myList.Description
}

Function SpPsPnpUpdateOneList()
{
	Set-PnPList -Identity "NewListPsPnp" -Description "New List Description"
}

Function SpPsPnpDeleteOneList()
{
	 Remove-PnPList -Identity "NewListPsPnp" -Force
}

Function SpPsPnpAddOneFieldToList()
{
	$fieldXml = "<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	Add-PnPFieldFromXml -List "NewListPsPnp" -FieldXml $fieldXml
}

Function SpPsPnpReadAllFieldsFromList()
{
	$allFields = Get-PnPField -List "NewListPsPnp"

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
}

Function SpPsPnpReadOneFieldFromList()
{
	$myField = Get-PnPField -List "NewListPsPnp" -Identity "MyMultilineField"

	Write-Host $myField.Id "-" $myField.TypeAsString
}

Function SpPsPnpUpdateOneFieldInList()
{
	Set-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" `
									-Values @{Description="New Field Description"}
}

Function SpPsPnpDeleteOneFieldFromList()
{
	Remove-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" -Force
}

#----------------------------------------------------------------------------------------

#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsPnP

#SpPsPnpCreateOneList
#SpPsPnpReadAllList
#SpPsPnpReadOneList
#SpPsPnpUpdateOneList
#SpPsPnpDeleteOneList
#SpPsPnpAddOneFieldToList
#SpPsPnpReadAllFieldsFromList
#SpPsPnpReadOneFieldFromList
#SpPsPnpUpdateOneFieldInList
SpPsPnpDeleteOneFieldFromList

Write-Host "Done"
