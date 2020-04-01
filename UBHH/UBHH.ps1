
Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsPnpCreateOneList()
{
	New-PnPList -Title "NewListPsPnp" -Template GenericList
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnpReadAllList()
{
	$allLists = Get-PnPList

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title + " - " + $oneList.Id
	}
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnpReadOneList()
{
	$myList = Get-PnPList -Identity "NewListPsPnp"

	Write-Host "List Description -" $myList.Description
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnpUpdateOneList()
{
	Set-PnPList -Identity "NewListPsPnp" -Description "New List Description"
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnpDeleteOneList()
{
	 Remove-PnPList -Identity "NewListPsPnp" -Force
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnpAddOneFieldToList()
{
	$fieldXml = "<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	Add-PnPFieldFromXml -List "NewListPsPnp" -FieldXml $fieldXml
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsPnpReadAllFieldsFromList()
{
	$allFields = Get-PnPField -List "NewListPsPnp"

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsPnpReadOneFieldFromList()
{
	$myField = Get-PnPField -List "NewListPsPnp" -Identity "MyMultilineField"

	Write-Host $myField.Id "-" $myField.TypeAsString
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnpUpdateOneFieldInList()
{
	Set-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" `
									-Values @{Description="New Field Description"}
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsPnpDeleteOneFieldFromList()
{
	Remove-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" -Force
}
#gavdcodeend 10

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