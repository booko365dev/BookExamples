
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShell_WithAccPw()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnPPowerShell_Certificate()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificatePath "[PathForThePfxCertificateFile]" `
					  -CertificatePassword $securePW
}

Function LoginPsPnPPowerShell_CertificateBase64()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			"myStrongPassword" -AsPlainText -Force

	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -ClientId $configFile.appsettings.ClientIdWithCert `
					  -Tenant "[Domain].onmicrosoft.com" `
					  -CertificateBase64Encoded "[Base64EncodedValue]" `
					  -CertificatePassword $securePW
}

Function LoginPsPnPPowerShell_Interactive()
{
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function LoginPsCLI_WithAccPw()
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
}

Function LoginPsCLI_WithSecret()
{
	m365 login --authType secret `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithSecret `
			   --secret $configFile.appsettings.ClientSecret
}

Function LoginPsCLI_WithCertificate()
{
	m365 login --authType certificate `
			   --tenant $configFile.appsettings.TenantName `
			   --appId $configFile.appsettings.ClientIdWithCert `
			   --certificateFile $configFile.appsettings.CertificateFilePath `
			   --password $configFile.appsettings.CertificateFilePw
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#------- Search --------
#gavdcodebegin 01
Function SpPsRest_ResultsSearchGET
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl +
                "/_api/search/query?querytext='team'"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsRest_ResultsSearchPOST
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
											"/_api/search/query"
	$myPayload = 
		"{
			'request': { 
				'__metadata': { 
					'type': 'Microsoft.Office.Server.Search.REST.SearchRequest' 
				}, 
				'Querytext': 'team', 
				'RowLimit': 20, 
				'ClientType': 'ContentSearchRegular' 
			}
		}"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 02

#------- User Profile --------
#gavdcodebegin 03
Function SpPsRest_GetAllPropertiesUserProfile
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.UserName.Replace("@", "%40");
	$endpointUrl = $configFile.appsettings.SiteBaseUrl +
                "/_api/sp.userprofiles.peoplemanager/" + `
						"getpropertiesfor(@v)?@v='" + $myUser + "'"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsRest_GetAllMyPropertiesUserProfile
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl +
                "/_api/sp.userprofiles.peoplemanager/getmyproperties"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsRest_GetPropertiesUserProfile
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.UserName#.Replace("@", "%40");
	$endpointUrl = $configFile.appsettings.SiteBaseUrl +
					  "/_api/sp.userprofiles.peoplemanager/" + `
					  "getuserprofilepropertyfor" + `
                      "(accountname=@v, propertyname='AboutMe')?@v='" + $myUser + "'"

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Get `
							  -Headers $myHeader `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 05

#------- Site Scripts --------
#gavdcodebegin 06
Function SpPsRest_GenerateListSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
		"GetSiteScriptFromList"

	$myPayload = @{
			listUrl = 'https://[domain].sharepoint.com/sites/Test_Guitaca/Lists/TestList'
			} | ConvertTo-Json

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsRest_GenerateWebSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteScriptFromWeb"

	$myPayload = '{ 
	   "webUrl":"https://m365x25054427.sharepoint.com/sites/Test_Guitaca",
	   "info":{ 
			  "IncludeBranding":true,
			  "IncludedLists":[ 
				 "Shared Documents",
				 "Lists/TestList"
			  ],
			  "IncludeRegionalSettings":true,
			  "IncludeSiteExternalSharingCapability":true,
			  "IncludeTheme":true,
			  "IncludeLinksToExportedItems":true
			}
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsRest_AddSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"CreateSiteScript(Title=@title)?@title='CustomListFromSiteScript'"

	$myPayload = '{
		  "$schema": "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json",
		  "actions": [
			{
			  "verb": "createSPList",
			  "listName": "TestList",
			  "templateType": 100,
			  "color": "1",
			  "icon": "8",
			  "subactions": [
				{
				  "verb": "setDescription",
				  "description": "This is a test list"
				},
				{
				  "verb": "addSPFieldXml",
				  "schemaXml": "<Field ID=\"{fa564e0f-0c70-4ab9-b863-0177e6ddd247}\" 
					Type=\"Text\" Name=\"Title\" DisplayName=\"Title\" Required=\"TRUE\"
					SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" 
					StaticName=\"Title\" FromBaseType=\"TRUE\" MaxLength=\"255\" />"
				},
				{
				  "verb": "addSPFieldXml",
				  "schemaXml": "<Field Description=\"This is a test column\" DisplayName=\"TestColumn\" Format=\"Dropdown\" IsModern=\"TRUE\" 
					MaxLength=\"255\" Name=\"TestColumn\" Required=\"TRUE\" 
					Title=\"TestColumn\" Type=\"Text\" ID=\"{22191eb9-6879-4d8b-8e6b-2c288577ee28}\" 
					StaticName=\"TestColumn\" />"
				},
				{
				  "verb": "addSPFieldXml",
				  "schemaXml": "<Field ID=\"{82642ec8-ef9b-478f-acf9-31f7d45fbc31}\" 
					DisplayName=\"Title\" Description=\"undefined\" Name=\"LinkTitle\" 
					SourceID=\"http://schemas.microsoft.com/sharepoint/v3\" 
					StaticName=\"LinkTitle\" Type=\"Computed\" ReadOnly=\"TRUE\" 
					FromBaseType=\"TRUE\" Width=\"150\" DisplayNameSrcField=\"Title\" 
					Sealed=\"FALSE\"><FieldRefs><FieldRef Name=\"Title\" /><FieldRef 
					Name=\"LinkTitleNoMenu\" /><FieldRef Name=\"_EditMenuTableStart2\" />
					<FieldRef Name=\"_EditMenuTableEnd\" />
					</FieldRefs><DisplayPattern><FieldSwitch><Expr><GetVar 
					Name=\"FreeForm\" /></Expr><Case Value=\"TRUE\"><Field 
					Name=\"LinkTitleNoMenu\" /></Case><Default><HTML>
					<![CDATA[<div class=\"ms-vb itx\" onmouseover=\"OnItem(this)\" 
					CTXName=\"ctx]]></HTML><Field Name=\"_EditMenuTableStart2\" />
					<HTML><![CDATA[\">]]></HTML><Field Name=\"LinkTitleNoMenu\" />
					<HTML><![CDATA[</div>]]></HTML><HTML><![CDATA[<div 
					class=\"s4-ctx\" onmouseover=\"OnChildItem(this.parentNode); 
					return false;\">]]></HTML><HTML><![CDATA[<span>&nbsp;</span>]]>
					</HTML><HTML><![CDATA[<a 
					onfocus=\"OnChildItem(this.parentNode.parentNode); return false;\" 
					onclick=\"PopMenuFromChevron(event); return false;\" 
					href=\"javascript:;\" title=\"Open Menu\"></a>]]></HTML><HTML>
					<![CDATA[<span>&nbsp;</span>]]></HTML><HTML><![CDATA[</div>]]>
					</HTML></Default></FieldSwitch></DisplayPattern></Field>"
				},
				{
				  "verb": "addSPView",
				  "name": "All Items",
				  "viewFields": [
					"LinkTitle",
					"TestColumn"
				  ],
				  "query": "",
				  "rowLimit": 30,
				  "isPaged": true,
				  "makeDefault": true,
				  "formatterJSON": "",
				  "replaceViewFields": true
				}
			  ]
			}
		  ]
		}'	

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsRest_GetAllSiteScripts
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteScripts"

	$myPayload = '{ }'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsRest_GetOneSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteScriptMetadata"

	$myPayload = @{
			id = 'fde0681c-9512-4652-8198-3f9b9934a394'
			} | ConvertTo-Json

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsRest_UpdateSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"UpdateSiteScript"

	$myPayload = '{ 
	   "updateInfo":{ 
			"Id":"fde0681c-9512-4652-8198-3f9b9934a394",  
		    "Title":"CustomListFromSiteScript", 
		    "Description":"Custom List Updated", 
		    "Version": 2
			}
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsRest_DeleteSiteScript
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"DeleteSiteScript"

	$myPayload = '{ 
			"id":"fde0681c-9512-4652-8198-3f9b9934a394"  
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 12

#------- Site Templates --------
#gavdcodebegin 13
Function SpPsRest_AddSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"CreateSiteDesign"

	$myPayload = '{ 
	   "info":{ 
				"Title":"Custom List From Site Design REST",
				"Description":"Creates a Custom List in a site using REST Site Design",
				"SiteScriptIds":["79a5174f-0712-49c7-b6af-5a45918c55ee"],
				"WebTemplate":"64",
				"PreviewImageUrl":"https://[domain].sharepoint.com/SiteAssets/mydesign.png",
				"PreviewImageAltText":"Custom List in a site using REST Site Design"
			}
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsRest_GetAllSiteTemplates
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteDesigns"

	$myPayload = '{ }'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsRest_GetOneSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteDesignMetadata"

	$myPayload = @{
			id = 'c80235ae-b26f-431e-9199-d459be24e89f'
			} | ConvertTo-Json

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsRest_UpdateSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"UpdateSiteDesign"

	$myPayload = '{ 
	   "updateInfo":{ 
			"Id":"c80235ae-b26f-431e-9199-d459be24e89f",  
		    "Title":"Custom List From REST Site Template", 
		    "Description":"Custom List Updated", 
		    "SiteScriptIds":["79a5174f-0712-49c7-b6af-5a45918c55ee"], 
		    "Version": 2
			}
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsRest_DeleteSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"DeleteSiteDesign"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsRest_ApplySiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"ApplySiteDesign"

	$myPayload = '{ 
			"siteDesignId":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"webUrl":"https://[domain].sharepoint.com/sites/Test_Guitaca"  
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsRest_ApplyToSiteSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"AddSiteDesignTaskToCurrentWeb"

	$myPayload = '{ 
			"siteDesignId":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsRest_GetRigthsSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GetSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsRest_GrantRightsSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"GrantSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"principalNames":["[user]@[domain].onmicrosoft.com"],
			"grantedRights":"View"
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsRest_DeleteRightsSiteTemplate
{
	LoginPsPnPPowerShell_WithAccPw
	$myOAuth = Get-PnPAppAuthAccessToken

	$endpointUrl = $configFile.appsettings.SiteBaseUrl + `
		"/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + ` 
		"RevokeSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"principalNames":["[user]@[domain].onmicrosoft.com"]
	}'

	$myHeader = @{ 'Authorization' = "Bearer $($myOAuth)"; `
				   'Accept' = 'application/json;odata=verbose' }
	$data = Invoke-WebRequest -Method Post `
							  -Headers $myHeader `
							  -Body $myPayload `
							  -Uri $endpointUrl `
							  -ContentType "application/json;odata=verbose"

	Write-Host $data
}
#gavdcodeend 22


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#------- Search --------
#SpPsRest_ResultsSearchGET
#SpPsRest_ResultsSearchPOST

#------- User Profile --------
#SpPsRest_GetAllPropertiesUserProfile
#SpPsRest_GetAllMyPropertiesUserProfile
#SpPsRest_GetPropertiesUserProfile

#------- Site Site Scripts --------
#SpPsRest_GenerateListSiteScript
#SpPsRest_GenerateWebSiteScript
#SpPsRest_AddSiteScript
#SpPsRest_GetAllSiteScripts
#SpPsRest_GetOneSiteScript
#SpPsRest_UpdateSiteScript
#SpPsRest_DeleteSiteScript

#------- Site Templates --------
#SpPsRest_AddSiteTemplate
#SpPsRest_GetAllSiteTemplates
#SpPsRest_GetOneSiteTemplate
#SpPsRest_UpdateSiteTemplate
#SpPsRest_DeleteSiteTemplate
#SpPsRest_ApplySiteTemplate
#SpPsRest_ApplyToSiteSiteTemplate
#SpPsRest_GetRigthsSiteTemplate
#SpPsRest_GrantRightsSiteTemplate
#SpPsRest_DeleteRightsSiteTemplate

Write-Host "Done" 
