
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function LoginPsPnPPowerShell_AccPwDefault
{
	# Using the "PnP Management Shell" Azure AD PnP App Registration (Delegated)
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

Function LoginPsPnPPowerShell_AccPw($FullSiteUrl)
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

Function LoginPsPnPPowerShell_UserInteraction
{
	# Using user interaction and the Azure AD PnP App Registration (Delegated)
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl `
					  -Credentials (Get-Credential)
}

Function LoginPsPnPPowerShell_Certificate
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

Function LoginPsPnPPowerShell_CertificateBase64
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

#---- Term Store ----
#gavdcodebegin 001
function SpPsPnpPowerShell_FindTermStore
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTaxSession = Get-PnPTaxonomySession
	Write-Host $myTaxSession.TermStores[0].Name
	Disconnect-PnPOnline
}
#gavdcodeend 001

#gavdcodebegin 002
function SpPsPnpPowerShell_CreateTermGroup
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTermGroup = New-PnPTermGroup -Name "PsPnpTermGroup"
	Write-Host $myTermGroup.Id
	Disconnect-PnPOnline
}
#gavdcodeend 002

#gavdcodebegin 003
function SpPsPnpPowerShell_FindTermGroup
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTermGroups = Get-PnPTermGroup
	foreach ($oneGroup in $myTermGroups) {
		Write-Host $oneGroup.Id
	}
	Disconnect-PnPOnline
}
#gavdcodeend 003

#gavdcodebegin 004
function SpPsPnpPowerShell_CreateTermSet
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTermSet = New-PnPTermSet -Name "PsPnpTermSet" `
								-TermGroup "PsPnpTermGroup"
	Disconnect-PnPOnline
}
#gavdcodeend 004

#gavdcodebegin 005
function SpPsPnpPowerShell_FindTermSet
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTermSets = Get-PnPTermSet -TermGroup "PsPnpTermGroup"
	foreach ($oneSet in $myTermSets) {
		Write-Host $oneSet.Id
	}
	Disconnect-PnPOnline
}
#gavdcodeend 005

#gavdcodebegin 006
function SpPsPnpPowerShell_CreateTerm
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTerm = New-PnPTerm -Name "PsPnpTerm" `
						  -TermGroup "PsPnpTermGroup" `
						  -TermSet "PsPnpTermSet"
	Write-Host $myTerm.Id
	Disconnect-PnPOnline
}
#gavdcodeend 006

#gavdcodebegin 007
function SpPsPnpPowerShell_FindTerm
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTerms = Get-PnPTerm -TermGroup "PsPnpTermGroup" `
						   -TermSet "PsPnpTermSet"
	foreach ($oneTerm in $myTerms) {
		Write-Host $oneTerm.Id
	}
	Disconnect-PnPOnline
}
#gavdcodeend 007

#gavdcodebegin 008
function SpPsPnpPowerShell_DeleteTermGroup
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Remove-PnPTermGroup -GroupName "PsPnpTermGroup"
	Disconnect-PnPOnline
}
#gavdcodeend 008

#gavdcodebegin 009
function SpPsPnpPowerShell_ExportTaxonomy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Export-PnPTaxonomy -Path "C:\Temporary\tax.txt" `
					   -TermSet "7d40eadb-c320-4320-8eb0-da725c8a426f"
	Disconnect-PnPOnline
}
#gavdcodeend 009

#gavdcodebegin 010
function SpPsPnpPowerShell_ImportTaxonomy
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Import-PnPTaxonomy -Path "C:\Temporary\tax.txt"
	Disconnect-PnPOnline
}
#gavdcodeend 010

#gavdcodebegin 011
function SpPsPnpPowerShell_ExportTermGroup
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.ReadWrite.All
	#								Delegated AllSites.ReadWrite
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Export-PnPTermGroupToXml -Out "C:\Temporary\group.xml" -Identity "PsCsomTermGroup"
	Disconnect-PnPOnline
}
#gavdcodeend 011

#gavdcodebegin 012
function SpPsPnpPowerShell_ImportTermGroup
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Import-PnPTermGroupToXml -Path "C:\Temporary\tax.txt"
	Disconnect-PnPOnline
}
#gavdcodeend 012

#---- Search ----
#gavdcodebegin 013
function SpPsPnpPowerShell_Search
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$searchResults = Submit-PnPSearchQuery -Query "teams"
	foreach ($oneResult in $searchResults.ResultRows) {
		Write-Host $oneResult["Title"] " - " $oneResult["Author"] " - " $oneResult["Path"]
	}
	Disconnect-PnPOnline
}
#gavdcodeend 013

#gavdcodebegin 014
function SpPsPnpPowerShell_SearchSiteColls
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteSearchQueryResults
	Disconnect-PnPOnline
}
#gavdcodeend 014

#gavdcodebegin 015
function SpPsPnpPowerShell_SearchCrawl
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSearchCrawlLog
	Disconnect-PnPOnline
}
#gavdcodeend 015

#---- User Profile ----
#gavdcodebegin 016
function SpPsPnpPowerShell_FindUserProfileProperties
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPUserProfileProperty -Account $configFile.appsettings.UserName
	Disconnect-PnPOnline
}
#gavdcodeend 016

#gavdcodebegin 017
function SpPsPnpPowerShell_UpdateUserProfileProperties
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Set-PnPUserProfileProperty -Account $configFile.appsettings.UserName `
							   -Property "AboutMe" `
							   -Value "I am the administrator"
	Disconnect-PnPOnline
}
#gavdcodeend 017

#---- Modern Pages ----
#gavdcodebegin 018
function SpPsPnpPowerShell_CreateModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Add-PnPPage -Name "MyModernPage" `
				-Title "This is my page" `
				-LayoutType Article
	Disconnect-PnPOnline
}
#gavdcodeend 018

#gavdcodebegin 019
function SpPsPnpPowerShell_CreateNewsModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Add-PnPPage -Name "MyNewsModernPage" `
				-Title "This is my News page" `
				-LayoutType Article `
				 -PromoteAs NewsArticle
	Disconnect-PnPOnline
}
#gavdcodeend 019

#gavdcodebegin 020
function SpPsPnpPowerShell_CreateModernPageAsTemplate
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Add-PnPPage -Name "MyTemplateModernPage" `
				-Title "This is my Template page" `
				-LayoutType Article `
				-PromoteAs Template
	Disconnect-PnPOnline
}
#gavdcodeend 020

#gavdcodebegin 021
function SpPsPnpPowerShell_ModernPageToNewsPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Set-PnPPage -Identity "MyModernPage.aspx" `
				-PromoteAs NewsArticle
	Disconnect-PnPOnline
}
#gavdcodeend 021

#gavdcodebegin 022
function SpPsPnpPowerShell_GetModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"

	foreach($oneControl in $myPage.Controls) {
		Write-Host $oneControl.Type
	}
	Disconnect-PnPOnline
}
#gavdcodeend 022

#gavdcodebegin 023
function SpPsPnpPowerShell_AddSectionInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Add-PnPPageSection -Page $myPage `
					   -SectionTemplate "TwoColumnLeft"
	Disconnect-PnPOnline
}
#gavdcodeend 023

#gavdcodebegin 024
function SpPsPnpPowerShell_AddTextWebPartInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Add-PnPPageTextPart -Page $myPage `
						-Text "Some Text" `
						-Section 1 -Column 1 -Order 1
	Disconnect-PnPOnline
}
#gavdcodeend 024

#gavdcodebegin 025
function SpPsPnpPowerShell_AddHeroWebPartInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Add-PnPPageWebPart -Page $myPage `
					   -DefaultWebPartType "Hero" `
					   -Section 1 -Column 1 -Order 1
	Disconnect-PnPOnline
}
#gavdcodeend 025

#gavdcodebegin 026
function SpPsPnpPowerShell_AddNewsWebPartInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Add-PnPPageWebPart -Page $myPage `
					   -DefaultWebPartType "NewsFeed" `
					   -Section 1 -Column 1 -Order 2 `
					   -WebPartProperties @{layoutId="GridNews";title="News"}
	Disconnect-PnPOnline
}
#gavdcodeend 026

#gavdcodebegin 027
function SpPsPnpPowerShell_RemoveOneWebPartFromModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	$myPage.Sections[0].Controls.RemoveAt(1)
	Disconnect-PnPOnline
}
#gavdcodeend 027

#gavdcodebegin 028
function SpPsPnpPowerShell_RemoveOneSectionFromModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	$myPage.Sections.RemoveAt(0)
	Disconnect-PnPOnline
}
#gavdcodeend 028

#gavdcodebegin 029
function SpPsPnpPowerShell_SaveAndPublishModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	$myPage.Save()
	$myPage.Publish()
	Disconnect-PnPOnline
}
#gavdcodeend 029

#gavdcodebegin 030
function SpPsPnpPowerShell_PublishModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Set-PnPPage -Identity $myPage -Publish
	Disconnect-PnPOnline
}
#gavdcodeend 030

#gavdcodebegin 031
function SpPsPnpPowerShell_UpdateModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	Set-PnPPage -Identity $myPage `
				-CommentsEnabled:$false `
				-HeaderType None
	Disconnect-PnPOnline
}
#gavdcodeend 031

#gavdcodebegin 032
function SpPsPnpPowerShell_UpdateTextWebPartInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	$myComponents = Get-PnPPageComponent -Page $myPage

	$typeWP = "OfficeDevPnP.Core.Pages.ClientSideText"
	$contentWP = '<strong>Updated <u>text</u></strong>'
	foreach($oneComponent in $myComponents) {
		if(($oneComponent.Type.ToString() -eq $typeWP) -and `
		   ($oneComponent.Section[0].Order -eq 1) -and `
		   ($oneComponent.Column[0].LayoutIndex -eq 1) -and `
		   ($oneComponent.Order -eq 1)) {
				Set-PnPPageTextPart -Page $myPage `
									-InstanceId $oneComponent.InstanceId `
									-Text $contentWP
		}
	}
	Disconnect-PnPOnline
}
#gavdcodeend 032

#gavdcodebegin 033
function SpPsPnpPowerShell_UpdateWebPartInModernPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPPage -Identity "MyModernPage.aspx"
	$myComponents = Get-PnPPageComponent -Page $myPage

	$typeWP = "OfficeDevPnP.Core.Pages.ClientSideWebPart"
	$contentWP = '{"layoutId":"GridNews","title":"News Updated"}'
	foreach($oneComponent in $myComponents) {
		if(($oneComponent.Type.ToString() -eq $typeWP) -and `
		   ($oneComponent.Section[0].Order -eq 1) -and `
		   ($oneComponent.Column[0].LayoutIndex -eq 1) -and `
		   ($oneComponent.Order -eq 2)) {
				Set-PnPPageWebPart -Page $myPage `
								   -Identity $oneComponent.InstanceId `
								   -PropertiesJson $contentWP
		}
	}
	Disconnect-PnPOnline
}
#gavdcodeend 033

#gavdcodebegin 034
function SpPsPnpPowerShell_PromotePageToNewsPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Set-PnPListItem -List "SitePages" ` 
					-Identity 7 `
					-Values @{`
						"PromotedState" = 2;`
						"FirstPublishedDate" = "2022-10-28T07:00:00Z";`
						"Created" = "2022-10-28T07:00:00Z"}
	Disconnect-PnPOnline
}
#gavdcodeend 034

#gavdcodebegin 035
function SpPsPnpPowerShell_UpdatePage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPListItem -List "SitePages" `
							  -Id 7
	$oldString = $myPage["CanvasContent1"]    
	$newString = $oldString.Replace("<strong>Updated <u>text</u></strong>", "My News")
	Set-PnPListItem -List "SitePages" `
					-Identity 7 `
					-Values @{"CanvasContent1" = $newString}
	Disconnect-PnPOnline
}
#gavdcodeend 035

#gavdcodebegin 036
function SpPsPnpPowerShell_GetAllPropertiesPage
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myPage = Get-PnPListItem -List "SitePages" `
							  -Id 7
	foreach ($oneProp in $myPage.FieldValues) {
		$oneProp
	}
	Disconnect-PnPOnline
}
#gavdcodeend 036

#---- Provisioning ----
#gavdcodebegin 037
function SpPsPnpPowerShell_GenerateSiteTemplateXml
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningSite.xml"
	Disconnect-PnPOnline
}
#gavdcodeend 037

#gavdcodebegin 038
function SpPsPnpPowerShell_GenerateListsTemplate
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningLists.xml" `
						-ListsToExtract "MyCustomList",`
												"7B8f0d6e79-406c-48a9-834e-af0c56489bbe"
	Disconnect-PnPOnline
}
#gavdcodeend 038

#gavdcodebegin 039
function SpPsPnpPowerShell_GenerateTermGroupsTemplate
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningTermGroups.xml" `
						-IncludeAllTermGroups
	Disconnect-PnPOnline
}
#gavdcodeend 039

#gavdcodebegin 040
function SpPsPnpPowerShell_GenerateSiteTemplatePnP
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningSite.pnp"
	Disconnect-PnPOnline
}
#gavdcodeend 040

#gavdcodebegin 041
function SpPsPnpPowerShell_ApplySiteTemplate
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Invoke-PnPSiteTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
	Disconnect-PnPOnline
}
#gavdcodeend 041

#gavdcodebegin 042
function SpPsPnpPowerShell_GenerateTenantTemplateXml
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPTenantTemplate -Out "C:\Temporary\TestProvisioningTenant.xml" `
						  -SiteUrl "https://[domain].sharepoint.com/sites/Test_Guitaca" #`
						  #-Configuration "C:\Temporary\TestConfiguration.xml"
	Disconnect-PnPOnline
}
#gavdcodeend 042

#gavdcodebegin 043
function SpPsPnpPowerShell_ApplyTenantTemplate
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Invoke-PnPTenantTemplate -Path "C:\Temporary\TestProvisioningTenant.xml"
	Disconnect-PnPOnline
}
#gavdcodeend 043

#gavdcodebegin 044
function SpPsPnpPowerShell_GenerateSiteTemplateWithConfig
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Get-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningSiteWithConfig.xml" `
						-Configuration "C:\Temporary\TestConfiguration.xml"
	Disconnect-PnPOnline
}
#gavdcodeend 044

#gavdcodebegin 045
function SpPsPnpPowerShell_GenerateSiteTemplateInMem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = PnPSiteTemplate -OutputInstance
	$myTemplate | ConvertTo-Json
	Disconnect-PnPOnline
}
#gavdcodeend 045

#gavdcodebegin 046
function SpPsPnpPowerShell_GenerateSiteTemplateInMemFromFile
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = Read-PnPSiteTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
	$myTemplate | ConvertTo-Json
	Disconnect-PnPOnline
}
#gavdcodeend 046

#gavdcodebegin 047
function SpPsPnpPowerShell_GenerateSiteTemplateInMemFromScratch
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = New-PnPSiteTemplate
	$myTemplate | ConvertTo-Json
	Disconnect-PnPOnline
}
#gavdcodeend 047

#gavdcodebegin 048
function SpPsPnpPowerShell_SaveSiteTemplateInMemFromScratch
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = New-PnPSiteTemplate
	Save-PnPSiteTemplate -Out "C:\Temporary\TestProvisioningSiteInMem.xml" `
						 -InputInstance $myTemplate
	Disconnect-PnPOnline
}
#gavdcodeend 048

#gavdcodebegin 049
function SpPsPnpPowerShell_ModifySiteTemplateInMem
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = Read-PnPSiteTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
	$myTemplate.DisplayName = "In-memory modified template"
	$myTemplate.Security.AdditionalOwners.Clear()
	$myTemplate | ConvertTo-Json
	Disconnect-PnPOnline
}
#gavdcodeend 049

#gavdcodebegin 050
function SpPsPnpPowerShell_GenerateSiteTemplateInMemFromFilePnP
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	$myTemplate = Read-PnPSiteTemplate -Path "C:\Temporary\TestProvisioningSite.pnp"
	$myTemplate | ConvertTo-Json
	Disconnect-PnPOnline
}
#gavdcodeend 050

#gavdcodebegin 051
function SpPsPnpPowerShell_AddFileSiteTemplateInMemFromFilePnP
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Add-PnPFileToSiteTemplate -Path "C:\Temporary\TestProvisioningSite.pnp" `
							  -Source "C:\Temporary\MyStyles.css" `
							  -Folder "SiteAssets"
	Disconnect-PnPOnline
}
#gavdcodeend 051

#gavdcodebegin 052
function SpPsPnpPowerShell_RemoveFileSiteTemplateInMemFromFilePnP
{
	# App Registration type: Office 365 SharePoint Online 
	# App Registration permissions: Application Sites.Read.All
	#								Delegated AllSites.Read
	
	$spCtx = LoginPsPnPPowerShell_AccPwDefault
	Remove-PnPFileFromSiteTemplate -Path "C:\Temporary\TestProvisioningSite.pnp" `
								   -File "MyStyles.css"
	Disconnect-PnPOnline
}
#gavdcodeend 052


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#---- Term Store ----
#SpPsPnpPowerShell_FindTermStore
#SpPsPnpPowerShell_CreateTermGroup
#SpPsPnpPowerShell_FindTermGroup
#SpPsPnpPowerShell_CreateTermSet
#SpPsPnpPowerShell_FindTermSet
#SpPsPnpPowerShell_CreateTerm
#SpPsPnpPowerShell_FindTerm
#SpPsPnpPowerShell_DeleteTermGroup
#SpPsPnpPowerShell_ExportTaxonomy
#SpPsPnpPowerShell_ImportTaxonomy
#SpPsPnpPowerShell_ExportTermGroup
#SpPsPnpPowerShell_ImportTermGroup

#---- Search ----
#SpPsPnpPowerShell_Search
#SpPsPnpPowerShell_SearchSiteColls
#SpPsPnpPowerShell_SearchCrawl

#---- User Profile ----
#SpPsPnpPowerShell_FindUserProfileProperties
#SpPsPnpPowerShell_UpdateUserProfileProperties

#---- Modern Pages ----
#SpPsPnpPowerShell_CreateModernPage
#SpPsPnpPowerShell_CreateNewsModernPage
#SpPsPnpPowerShell_CreateModernPageAsTemplate
#SpPsPnpPowerShell_ModernPageToNewsPage
#SpPsPnpPowerShell_GetModernPage
#SpPsPnpPowerShell_AddSectionInModernPage
#SpPsPnpPowerShell_AddTextWebPartInModernPage
#SpPsPnpPowerShell_AddHeroWebPartInModernPage
#SpPsPnpPowerShell_AddNewsWebPartInModernPage
#SpPsPnpPowerShell_RemoveOneWebPartFromModernPage
#SpPsPnpPowerShell_RemoveOneSectionFromModernPage
#SpPsPnpPowerShell_SaveAndPublishModernPage
#SpPsPnpPowerShell_PublishModernPage
#SpPsPnpPowerShell_UpdateModernPage
#SpPsPnpPowerShell_UpdateTextWebPartInModernPage
#SpPsPnpPowerShell_UpdateWebPartInModernPage
#SpPsPnpPowerShell_PromotePageToNewsPage
#SpPsPnpPowerShell_UpdatePage
#SpPsPnpPowerShell_GetAllPropertiesPage

#---- Provisioning ----
#SpPsPnpPowerShell_GenerateSiteTemplateXml
#SpPsPnpPowerShell_GenerateListsTemplate
#SpPsPnpPowerShell_GenerateTermGroupsTemplate
#SpPsPnpPowerShell_GenerateSiteTemplatePnP
#SpPsPnpPowerShell_ApplySiteTemplate
#SpPsPnpPowerShell_GenerateTenantTemplateXml
#SpPsPnpPowerShell_ApplyTenantTemplate
#SpPsPnpPowerShell_GenerateSiteTemplateWithConfig
#SpPsPnpPowerShell_GenerateSiteTemplateInMem
#SpPsPnpPowerShell_GenerateSiteTemplateInMemFromFile
#SpPsPnpPowerShell_GenerateSiteTemplateInMemFromScratch
#SpPsPnpPowerShell_SaveSiteTemplateInMemFromScratch
#SpPsPnpPowerShell_ModifySiteTemplateInMem
#SpPsPnpPowerShell_GenerateSiteTemplateInMemFromFilePnP
#SpPsPnpPowerShell_AddFileSiteTemplateInMemFromFilePnP
#SpPsPnpPowerShell_RemoveFileSiteTemplateInMemFromFilePnP

Write-Host "Done" 
