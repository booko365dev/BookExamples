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
Function SpPsPnpCreateModernPage()
{
	Add-PnPClientSidePage -Name "ModernPage" -LayoutType Article
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsPnpCreateNewsModernPage()
{
	Add-PnPClientSidePage -Name "NewsPage" -LayoutType Article -PromoteAs NewsArticle
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsPnpCreateModernPageAsTemplate()
{
	Add-PnPClientSidePage -Name "ModernTemplatePage" -LayoutType Article`
													 -PromoteAs Template
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsPnpModernPageToNewsPage()
{
	Set-PnPClientSidePage -Identity "ModernPage.aspx" -PromoteAs NewsArticle
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsPnpGetModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"

	foreach($oneControl in $myPage.Controls) {
		Write-Host $oneControl.Type
	}
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsPnpAddSectionInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Add-PnPClientSidePageSection -Page $myPage -SectionTemplate "TwoColumnLeft"
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsPnpAddTextWebPartInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Add-PnPClientSideText -Page $myPage -Text "Some Text" -Section 1 -Column 1 -Order 1
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsPnpAddHeroWebPartInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Add-PnPClientSideWebPart -Page $myPage -DefaultWebPartType "Hero" -Section 1 `
																	-Column 2 -Order 1
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsPnpAddNewsWebPartInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Add-PnPClientSideWebPart -Page $myPage -DefaultWebPartType "NewsFeed" -Section 1 `
								-Column 1 -Order 2 `
								-WebPartProperties @{layoutId="GridNews";title="News"}
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsPnpRemoveOneWebPartFromModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	$myPage.Sections[0].Controls.RemoveAt(1)
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsPnpRemoveOneSectionFromModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	$myPage.Sections.RemoveAt(0)
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsPnpSaveAndPublishModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	$myPage.Save()
	$myPage.Publish()
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsPnpPublishModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Set-PnPClientSidePage -Identity $myPage -Publish
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsPnpUpdateModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	Set-PnPClientSidePage -Identity $myPage -CommentsEnabled:$false -HeaderType None
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsPnpUpdateTextWebPartInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	$myComponents = Get-PnPClientSideComponent -Page $myPage

	$typeWP = "OfficeDevPnP.Core.Pages.ClientSideText"
	$contentWP = '<strong>Updated <u>text</u></strong>'
	foreach($oneComponent in $myComponents) {
		if(($oneComponent.Type.ToString() -eq $typeWP) -and `
		   ($oneComponent.Section[0].Order -eq 1) -and `
		   ($oneComponent.Column[0].LayoutIndex -eq 1) -and `
		   ($oneComponent.Order -eq 1)) {
				Set-PnPClientSideText -Page $myPage `
									  -InstanceId $oneComponent.InstanceId `
									  -Text $contentWP
		}
	}
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsPnpUpdateWebPartInModernPage()
{
	$myPage = Get-PnPClientSidePage -Identity "ModernPage.aspx"
	$myComponents = Get-PnPClientSideComponent -Page $myPage

	$typeWP = "OfficeDevPnP.Core.Pages.ClientSideWebPart"
	$contentWP = '{"layoutId":"GridNews","title":"News Updated"}'
	foreach($oneComponent in $myComponents) {
		if(($oneComponent.Type.ToString() -eq $typeWP) -and `
		   ($oneComponent.Section[0].Order -eq 1) -and `
		   ($oneComponent.Column[0].LayoutIndex -eq 1) -and `
		   ($oneComponent.Order -eq 2)) {
				Set-PnPClientSideWebPart -Page $myPage `
									  -Identity $oneComponent.InstanceId `
									  -PropertiesJson $contentWP
		}
	}
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsPnpPromotePageToNewsPage()
{
	Set-PnPListItem -List "SitePages" -Identity 7 -Values @{`
									"PromotedState" = 2;`
									"FirstPublishedDate" = "2019-10-28T07:00:00Z";`
									"Created" = "2019-10-28T07:00:00Z"}
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsPnpUpdatePage()
{
	#Get-PnPListItem -List "SitePages"  # To find the ID of myPage
	$myPage = Get-PnPListItem -List "SitePages" -Id 7
	$oldString = $myPage["CanvasContent1"]    
	$newString = $oldString.Replace("<strong>Updated <u>text</u></strong>", "My News")
	Set-PnPListItem -List "SitePages" -Identity 7 -Values @{"CanvasContent1" = $newString}
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsPnpGetAllPropertiesPage()
{
	$myPage = Get-PnPListItem -List "SitePages" -Id 7
	foreach ($oneProp in $myPage.FieldValues) {
		$oneProp
	}
}
#gavdcodeend 19

#----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

$spCtx = LoginPsPnP

#SpPsPnpCreateModernPage
#SpPsPnpCreateNewsModernPage
#SpPsPnpCreateModernPageAsTemplate
#SpPsPnpModernPageToNewsPage
#SpPsPnpGetModernPage
#SpPsPnpAddSectionInModernPage
#SpPsPnpAddTextWebPartInModernPage
#SpPsPnpAddHeroWebPartInModernPage
#SpPsPnpAddNewsWebPartInModernPage
#SpPsPnpRemoveOneWebPartFromModernPage
#SpPsPnpRemoveOneSectionFromModernPage
#SpPsPnpSaveAndPublishModernPage
#SpPsPnpPublishModernPage
#SpPsPnpUpdateModernPage
#SpPsPnpUpdateTextWebPartInModernPage
#SpPsPnpUpdateWebPartInModernPage
#SpPsPnpPromotePageToNewsPage
#SpPsPnpUpdatePage
SpPsPnpGetAllPropertiesPage

Write-Host "Done"