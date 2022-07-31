
Function LoginPsPnP()  #*** LEGACY CODE *** 
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}
#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsPnp_CreateOneList()  #*** LEGACY CODE *** 
{
	New-PnPList -Title "NewListPsPnp" -Template GenericList
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnp_ReadAllList()  #*** LEGACY CODE *** 
{
	$allLists = Get-PnPList

	foreach ($oneList in $allLists)
	{
		Write-Host $oneList.Title + " - " + $oneList.Id
	}
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnp_ReadOneList()  #*** LEGACY CODE *** 
{
	$myList = Get-PnPList -Identity "NewListPsPnp"

	Write-Host "List Description -" $myList.Description
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnp_UpdateOneList()  #*** LEGACY CODE *** 
{
	Set-PnPList -Identity "NewListPsPnp" -Description "New List Description"
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnp_DeleteOneList()  #*** LEGACY CODE *** 
{
	 Remove-PnPList -Identity "NewListPsPnp" -Force
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnp_AddOneFieldToList()  #*** LEGACY CODE *** 
{
	$fieldXml = "<Field Name='PSCmdletTest' DisplayName='MyMultilineField' Type='Note' />"
	Add-PnPFieldFromXml -List "NewListPsPnp" -FieldXml $fieldXml
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsPnp_ReadAllFieldsFromList()  #*** LEGACY CODE *** 
{
	$allFields = Get-PnPField -List "NewListPsPnp"

	foreach ($oneField in $allFields)
	{
		Write-Host $oneField.Title "-" $oneField.TypeAsString
	}
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsPnp_ReadOneFieldFromList()  #*** LEGACY CODE *** 
{
	$myField = Get-PnPField -List "NewListPsPnp" -Identity "MyMultilineField"

	Write-Host $myField.Id "-" $myField.TypeAsString
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnp_UpdateOneFieldInList()  #*** LEGACY CODE *** 
{
	Set-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" `
									-Values @{Description="New Field Description"}
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsPnp_DeleteOneFieldFromList()  #*** LEGACY CODE *** 
{
	Remove-PnPField -List "NewListPsPnp" -Identity "MyMultilineField" -Force
}
#gavdcodeend 10

#----------------------------------------------------------------------------------------

#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
#Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

$spCtx = LoginPsPnP

#SpPsPnp_CreateOneList
#SpPsPnp_ReadAllList
#SpPsPnp_ReadOneList
#SpPsPnp_UpdateOneList
#SpPsPnp_DeleteOneList
#SpPsPnp_AddOneFieldToList
#SpPsPnp_ReadAllFieldsFromList
#SpPsPnp_ReadOneFieldFromList
#SpPsPnp_UpdateOneFieldInList
#SpPsPnp_DeleteOneFieldFromList

Write-Host "Done"