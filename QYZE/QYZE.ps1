Function LoginPsCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.UserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.SiteCollUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

Function LoginPsSPO()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-SPOService -Url $configFile.appsettings.SiteAdminUrl -Credential $myCredentials
}

#----------------------------------------------------------------------------------------

Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.UserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.UserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.SiteCollUrl -Credentials $myCredentials
}

#----------------------------------------------------------------------------------------

Function Invoke-RestSPO() {
	Param (
		[Parameter(Mandatory=$True)]
		[String]$Url,
 
		[Parameter(Mandatory=$False)]
		[Microsoft.PowerShell.Commands.WebRequestMethod]$Method = `
								[Microsoft.PowerShell.Commands.WebRequestMethod]::Get,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password,
 
		[Parameter(Mandatory=$False)]
		[String]$Metadata,

		[Parameter(Mandatory=$False)]
		[System.Byte[]]$Body,
 
		[Parameter(Mandatory=$False)]
		[String]$RequestDigest,
 
		[Parameter(Mandatory=$False)]
		[String]$ETag,
 
		[Parameter(Mandatory=$False)]
		[String]$XHTTPMethod,

		[Parameter(Mandatory=$False)]
		[System.String]$Accept = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[String]$ContentType = "application/json;odata=verbose",

		[Parameter(Mandatory=$False)]
		[Boolean]$BinaryStringResponseBody = $False
	)
 
	if([string]::IsNullOrEmpty($Password)) {
		$SecurePassword = Read-Host -Prompt "Enter the password" -AsSecureString 
	}
	else {
		$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
	}
 
	$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials(`
															$UserName, $SecurePassword)
	$request = [System.Net.WebRequest]::Create($Url)
	$request.Credentials = $credentials
	$request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
	$request.ContentType = $ContentType
	$request.Accept = $Accept
	$request.Method=$Method
 
	if($RequestDigest) { 
		$request.Headers.Add("X-RequestDigest", $RequestDigest)
	}
	if($ETag) { 
		$request.Headers.Add("If-Match", $ETag)
	}
	if($XHTTPMethod) { 
		$request.Headers.Add("X-HTTP-Method", $XHTTPMethod)
	}
	if($Metadata -or $Body) {
		if($Metadata) {     
			$Body = [byte[]][char[]]$Metadata
		}      
		$request.ContentLength = $Body.Length 
		$stream = $request.GetRequestStream()
		$stream.Write($Body, 0, $Body.Length)
	}
	else {
		$request.ContentLength = 0
	}

	$response = $request.GetResponse()
	try {
		if($BinaryStringResponseBody -eq $False) {    
			$streamReader = New-Object System.IO.StreamReader $response.GetResponseStream()
			try {
				$data=$streamReader.ReadToEnd()
				$results = $data.ToString().Replace("ID", "_ID") | ConvertFrom-Json
				$results.d 
			}
			finally {
				$streamReader.Dispose()
			}
		}
		else {
			$dataStream = New-Object System.IO.MemoryStream
			try {
			Stream-CopyTo -Source $response.GetResponseStream() -Destination $dataStream
			$dataStream.ToArray()
			}
			finally {
				$dataStream.Dispose()
			} 
		}
	}
	finally {
		$response.Dispose()
	}
}
 
Function Get-SPOContextInfo(){
	Param(
		[Parameter(Mandatory=$True)]
		[String]$WebUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$False)]
		[String]$Password
	)
   
	$Url = $WebUrl + "/_api/contextinfo"
	Invoke-RestSPO $Url Post $UserName $Password
}

Function Stream-CopyTo([System.IO.Stream]$Source, [System.IO.Stream]$Destination) {
    $buffer = New-Object Byte[] 8192 
    $bytesRead = 0
    while (($bytesRead = $Source.Read($buffer, 0, $buffer.Length)) -gt 0) {
         $Destination.Write($buffer, 0, $bytesRead)
    }
}

#----------------------------------------------------------------------------------------

#gavdcodebegin 01
Function SpPsCsom_FindTermStore($spCtx)
{
    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $spCtx.Load($myTaxSession.TermStores)
    $spCtx.ExecuteQuery()

    foreach ($oneTermStore in $myTaxSession.TermStores) {
        Write-Host($oneTermStore.Name)
    }
}
#gavdcodeend 01

#gavdcodebegin 02
Function SpPsCsom_CreateTermGroup($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)

	$myNewGuid = New-Guid
    $myTermGroup = $myTermStore.CreateGroup("PsCsomTermGroup", $myNewGuid)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 02

#gavdcodebegin 03
Function SpPsCsom_FindTermGroups($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $spCtx.Load($myTermStore.Groups)
    $spCtx.ExecuteQuery()

    foreach ($oneGroup in $myTermStore.Groups) {
        Write-Host($oneGroup.Name)
    }
}
#gavdcodeend 03

#gavdcodebegin 04
Function SpPsCsom_CreateTermSet($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")

	$myNewGuid = New-Guid
    $myTermSet = $myTermGroup.CreateTermSet("PsCsomTermSet", $myNewGuid, 1033)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 04

#gavdcodebegin 05
Function SpPsCsom_FindTermSets($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")

    $spCtx.Load($myTermGroup.TermSets)
    $spCtx.ExecuteQuery()

    foreach ($oneTermSet in $myTermGroup.TermSets) {
        Write-Host($oneTermSet.Name)
    }
}
#gavdcodeend 05

#gavdcodebegin 06
Function SpPsCsom_CreateTerm($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")
    $myTermSet = $myTermGroup.TermSets.GetByName("PsCsomTermSet")

	$myNewGuid = New-Guid
    $myTerm = $myTermSet.CreateTerm("PsCsomTerm", 1033, $myNewGuid)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 06

#gavdcodebegin 07
Function SpPsCsom_FindTerms($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")
    $myTermSet = $myTermGroup.TermSets.GetByName("PsCsomTermSet")

    $spCtx.Load($myTermSet.Terms)
    $spCtx.ExecuteQuery()

    foreach ($oneTerm in $myTermSet.Terms) {
        Write-Host($oneTerm.Name)
    }
}
#gavdcodeend 07

#gavdcodebegin 08
Function SpPsCsom_FindOneTerm($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")
    $myTermSet = $myTermGroup.TermSets.GetByName("PsCsomTermSet")
    $myTerm = $myTermSet.Terms.GetByName("PsCsomTerm")

    $spCtx.Load($myTerm)
    $spCtx.ExecuteQuery()

    Write-Host($myTerm.Name)
}
#gavdcodeend 08

#gavdcodebegin 09
Function SpPsCsom_UpdateOneTerm($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")
    $myTermSet = $myTermGroup.TermSets.GetByName("PsCsomTermSet")
    $myTerm = $myTermSet.Terms.GetByName("PsCsomTerm")

    $myTerm.Name = "PsCsomTerm_Updated"
    $spCtx.ExecuteQuery()
}
#gavdcodeend 09

#gavdcodebegin 10
Function SpPsCsom_DeleteOneTerm($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermGroup = $myTermStore.Groups.GetByName("PsCsomTermGroup")
    $myTermSet = $myTermGroup.TermSets.GetByName("PsCsomTermSet")
    $myTerm = $myTermSet.Terms.GetByName("PsCsomTerm")

    $myTerm.DeleteObject()
    $spCtx.ExecuteQuery()
}
#gavdcodeend 10

#gavdcodebegin 11
Function SpPsCsom_FindTermSetAndTermById($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)
    $myTermSet = $myTermStore.GetTermSet("529c954a-0235-4202-a739-9b871055427c")
    $myTerm = $myTermStore.GetTerm("23f93354-0659-417c-91b0-b16e9a1666e7")

    $spCtx.Load($myTermSet)
    $spCtx.Load($myTerm)
    $spCtx.ExecuteQuery()

    Write-Host($myTermSet.Name + " - " + $myTerm.Name)
}
#gavdcodeend 11

#gavdcodebegin 12
Function SpPsPnp_FindTermStore()  #*** LEGACY CODE ***
{
	$myTaxSession = Get-PnPTaxonomySession
	Write-Host $myTaxSession.TermStores[0].Name
}
#gavdcodeend 12

#gavdcodebegin 13
Function SpPsPnp_CreateTermGroup()  #*** LEGACY CODE ***
{
	$myTermGroup = New-PnPTermGroup -Name "PsPnpTermGroup"
	Write-Host $myTermGroup.Id
}
#gavdcodeend 13

#gavdcodebegin 14
Function SpPsPnp_FindTermGroup()  #*** LEGACY CODE ***
{
	$myTermGroups = Get-PnPTermGroup
	foreach ($oneGroup in $myTermGroups) {
		Write-Host $oneGroup.Id
	}
}
#gavdcodeend 14

#gavdcodebegin 15
Function SpPsPnp_CreateTermSet()  #*** LEGACY CODE ***
{
	$myTermSet = New-PnPTermSet -Name "PsPnpTermSet" `
								-TermGroup "PsPnpTermGroup"
	Write-Host $myTermSet.Id
}
#gavdcodeend 15

#gavdcodebegin 16
Function SpPsPnp_FindTermSet()  #*** LEGACY CODE ***
{
	$myTermSets = Get-PnPTermSet -TermGroup "PsPnpTermGroup"
	foreach ($oneSet in $myTermSets) {
		Write-Host $oneSet.Id
	}
}
#gavdcodeend 16

#gavdcodebegin 17
Function SpPsPnp_CreateTerm()  #*** LEGACY CODE ***
{
	$myTerm = New-PnPTerm -Name "PsPnpTerm" `
						  -TermGroup "PsPnpTermGroup" `
						  -TermSet "PsPnpTermSet"
	Write-Host $myTerm.Id
}
#gavdcodeend 17

#gavdcodebegin 18
Function SpPsPnp_FindTerm()  #*** LEGACY CODE ***
{
	$myTerms = Get-PnPTerm -TermGroup "PsPnpTermGroup" `
						   -TermSet "PsPnpTermSet"
	foreach ($oneTerm in $myTerms) {
		Write-Host $oneTerm.Id
	}
}
#gavdcodeend 18

#gavdcodebegin 19
Function SpPsPnpDeleteTermGroup()  #*** LEGACY CODE ***
{
	Remove-PnPTermGroup -GroupName "PsPnpTermGroup"
}
#gavdcodeend 19

#gavdcodebegin 20
Function SpPsPnp_ExportTaxonomy()  #*** LEGACY CODE ***
{
	Export-PnPTaxonomy -Path "C:\Temporary\tax.txt" `
					   -TermSet "529c954a-0235-4202-a739-9b871055427c"
}
#gavdcodeend 20

#gavdcodebegin 21
Function SpPsPnp_ImportTaxonomy()  #*** LEGACY CODE ***
{
	Import-PnPTaxonomy -Path "C:\Temporary\tax.txt"
}
#gavdcodeend 21

#gavdcodebegin 22
Function SpPsPnp_ExportTermGroup()  #*** LEGACY CODE ***
{
	Export-PnPTermGroupToXml -Out "C:\Temporary\group.xml" -Identity "PsCsomTermGroup"
}
#gavdcodeend 22

#gavdcodebegin 23
Function SpPsPnpImportTermGroup()  #*** LEGACY CODE ***
{
	Import-PnPTermGroupToXml -Path "C:\Temporary\tax.txt"
}
#gavdcodeend 23

#gavdcodebegin 24
Function SpPsCsom_GetResultsSearch($spCtx)
{
	$keywordQuery = 
				New-Object Microsoft.SharePoint.Client.Search.Query.KeywordQuery($spCtx)
    $keywordQuery.QueryText = "Team"
    $searchExecutor = 
			  New-Object Microsoft.SharePoint.Client.Search.Query.SearchExecutor($spCtx)
    $results = $searchExecutor.ExecuteQuery($keywordQuery)
    $spCtx.ExecuteQuery()

    foreach ($resultRow in $results.Value[0].ResultRows) {
        Write-Host($resultRow["Title"] + " - " + 
                                $resultRow["Path"] + " - " + $resultRow["Write"])
    }
}
#gavdcodeend 24

#gavdcodebegin 25
Function SpPsRest_ResultsSearchGET()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/search/query?querytext='team'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 25

#gavdcodebegin 26
Function SpPsRest_ResultsSearchPOST()    #*** LEGACY CODE ***
{
	$endpointUrl = $webUrl + "/_api/search/query"
	$myPayload = @{
			request = @{
			__metadata = @{ 'type' = 'Microsoft.Office.Server.Search.REST.SearchRequest' }
			Querytext = 'team'
            RowLimit = 20
            ClientType = 'ContentSearchRegular'
			}} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 26

#gavdcodebegin 27
Function SpPsPnp_Search()    #*** LEGACY CODE ***
{
	Submit-PnPSearchQuery -Query "team"
}
#gavdcodeend 27

#gavdcodebegin 28
Function SpPsPnp_SearchSiteColls()    #*** LEGACY CODE ***
{
	Get-PnPSiteSearchQueryResults
}
#gavdcodeend 28

#gavdcodebegin 29
Function SpPsPnpSearchCrawl()    #*** LEGACY CODE ***
{
	Get-PnPSearchCrawlLog
}
#gavdcodeend 29

#gavdcodebegin 30
Function SpPsCsom_GetAllPropertiesUserProfile($spCtx)
{
    $myUser = "i:0#.f|membership|" + $configFile.appsettings.UserName
    $myPeopleManager = New-Object `
						Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spCtx)
    $myUserProperties = $myPeopleManager.GetPropertiesFor($myUser)
    $spCtx.Load($myUserProperties)
    $spCtx.ExecuteQuery()

	$myProfProp = $myUserProperties.UserProfileProperties
    foreach ($oneKey in $myProfProp.Keys) {
        Write-Host($oneKey + " - " + $myProfProp[$oneKey])
    }
}
#gavdcodeend 30

#gavdcodebegin 31
Function SpCsCsom_GetAllMyPropertiesUserProfile($spCtx)
{
    $myPeopleManager = New-Object `
						Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spCtx)
    $myUserProperties = $myPeopleManager.GetMyProperties()
    $spCtx.Load($myUserProperties)
    $spCtx.ExecuteQuery()

	$myProfProp = $myUserProperties.UserProfileProperties
    foreach ($oneKey in $myProfProp.Keys) {
        Write-Host($oneKey + " - " + $myProfProp[$oneKey])
    }
}
#gavdcodeend 31

#gavdcodebegin 32
Function SpPsCsom_GetPropertiesUserProfile($spCtx)
{
    $myUser = "i:0#.f|membership|" + $configFile.appsettings.UserName
    $myPeopleManager = New-Object `
						Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spCtx)
    $myProfPropertyNames = @( "Manager", "Department", "Title" )
    $myProfProperties = New-Object `
			Microsoft.SharePoint.Client.UserProfiles.UserProfilePropertiesForUser(`
												$spCtx, $myUser, $myProfPropertyNames)
    $myProfPropertyValues = `
        $myPeopleManager.GetUserProfilePropertiesFor($myProfProperties)

    $spCtx.Load($myProfProperties)
    $spCtx.ExecuteQuery()

    foreach ($oneValue in $myProfPropertyValues) {
        Write-Host($oneValue)
    }
}
#gavdcodeend 32

#gavdcodebegin 33
Function SpPsCsomUpdateOnePropertyUserProfile($spCtx)
{
    $myPeopleManager = New-Object `
						Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spCtx)
    $myUserProperties = $myPeopleManager.GetMyProperties()
    $spCtx.Load($myUserProperties)
    $spCtx.ExecuteQuery()

    $newValue = "I am also the administrator"
    $myPeopleManager.SetSingleValueProfileProperty(`
            $myUserProperties.AccountName, "AboutMe", $newValue)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 33

#gavdcodebegin 34
Function SpPsCsom_UpdateOneMultPropertyUserProfile($spCtx)
{
    $myPeopleManager = New-Object `
						Microsoft.SharePoint.Client.UserProfiles.PeopleManager($spCtx)
    $myUserProperties = $myPeopleManager.GetMyProperties()
    $spCtx.Load($myUserProperties)
    $spCtx.ExecuteQuery()

    $mySkills = New-Object "System.Collections.Generic.List``1[System.string]"
    $mySkills.Add("OneDrive")
    $mySkills.Add("Teams")
    $myPeopleManager.SetMultiValuedProfileProperty(`
                            $myUserProperties.AccountName, "SPS-Skills", $mySkills)
    $spCtx.ExecuteQuery()
}
#gavdcodeend 34

#gavdcodebegin 35
Function SpPsPnp_FindUserProfileProperties()    #*** LEGACY CODE ***
{
	Get-PnPUserProfileProperty -Account $configFile.appsettings.UserName
}
#gavdcodeend 35

#gavdcodebegin 36
Function SpPsPnp_UpdateUserProfileProperties()    #*** LEGACY CODE ***
{
	Set-PnPUserProfileProperty -Account $configFile.appsettings.UserName `
							   -Property "AboutMe" `
							   -Value "I am not the administrator"
}
#gavdcodeend 36

#gavdcodebegin 37
Function SpPsRest_GetAllPropertiesUserProfile()    #*** LEGACY CODE ***
{
    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.UserName.Replace("@", "%40");
    $endpointUrl = $webUrl + "/_api/sp.userprofiles.peoplemanager/" + `
						"getpropertiesfor(@v)?@v='" + $myUser + "'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 37

#gavdcodebegin 38
Function SpPsRest_GetAllMyPropertiesUserProfile()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/sp.userprofiles.peoplemanager/getmyproperties"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 38

#gavdcodebegin 39
Function SpPsRest_GetPropertiesUserProfile()    #*** LEGACY CODE ***
{
    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.UserName#.Replace("@", "%40");
    $endpointUrl = $webUrl + "/_api/sp.userprofiles.peoplemanager/" + `
					  "getuserprofilepropertyfor" + `
                      "(accountame=@v, propertyname='AboutMe')?@v='" + $myUser + "'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}
#gavdcodeend 39

#gavdcodebegin 40
Function SpPsSpo_GenerateListSiteScript
{
	$mySourceListUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca/Lists/TestList"
	Get-SPOSiteScriptFromList -ListUrl $mySourceListUrl
}
#gavdcodeend 40

#gavdcodebegin 41
Function SpPsSpo_GenerateWebSiteScript
{
	$mySourceWebUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"
	Get-SPOSiteScriptFromWeb -WebUrl $mySourceWebUrl `
						     -IncludeBranding `
							 -IncludeTheme `
							 -IncludeRegionalSettings `
							 -IncludeSiteExternalSharingCapability `
							 -IncludeLinksToExportedItems `
							 -IncludedLists ("Shared Documents", "Lists/TestList")
}
#gavdcodeend 41

#gavdcodebegin 42
Function SpPsSpo_AddSiteScript
{
	$myScript = Get-Content "C:\Temporary\TestListSiteScript.json" -Raw
	Add-SPOSiteScript -Title "CustomListFromSiteScript" `
					  -Content $myScript `
					  -Description "Creates a Custom List using SPO"
}
#gavdcodeend 42

#gavdcodebegin 43
Function SpPsSpo_GetAllSiteScripts
{
	Get-SPOSiteScript
}
#gavdcodeend 43

#gavdcodebegin 44
Function SpPsSpo_GetOneSiteScript
{
	$myScriptId = "83a75409-c005-4125-b7b1-f8b288bb3374"
	Get-SPOSiteScript -Identity $myScriptId
}
#gavdcodeend 44

#gavdcodebegin 45
Function SpPsSpo_UpdateSiteScript
{
	$myScriptId = "83a75409-c005-4125-b7b1-f8b288bb3374"
	$myScript = Get-Content "C:\Temporary\TestListSiteScript.json" -Raw
	Set-SPOSiteScript -Identity $myScriptId `
					  -Title "CustomerListFromSiteScript" `
					  -Content $myScript `
					  -Description "Creates a Custom List updated"
}
#gavdcodeend 45

#gavdcodebegin 46
Function SpPsSpo_DeleteOneSiteScript
{
	$myScriptId = "83a75409-c005-4125-b7b1-f8b288bb3374"
	Remove-SPOSiteScript -Identity $myScriptId
}
#gavdcodeend 46

#gavdcodebegin 47
Function SpPsRest_GenerateListSiteScript()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteScriptFromList"
	$myPayload = @{
			listUrl = 'https://[domain].sharepoint.com/sites/Test_Guitaca/Lists/TestList'
			} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 47

#gavdcodebegin 48
Function SpPsRest_GenerateWebSiteScript()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteScriptFromWeb"

	$myPayload = '{ 
	   "webUrl":"https://[domain].sharepoint.com/sites/Test_Guitaca",
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

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 48

#gavdcodebegin 49
Function SpPsRest_AddSiteScript()    #*** LEGACY CODE ***
{
	$myPayload = '
		{
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

	$endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"CreateSiteScript(Title=@title)?@title='CustomListFromSiteScript'"

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 49

#gavdcodebegin 50
Function SpPsRest_GetAllSiteScripts()    #*** LEGACY CODE ***
{
	$endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteScripts"

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 50

#gavdcodebegin 51
Function SpPsRest_GetOneSiteScript()    #*** LEGACY CODE ***
{
	$endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteScriptMetadata"

	$myPayload = @{
			id = 'fde0681c-9512-4652-8198-3f9b9934a394'
			} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 51

#gavdcodebegin 52
Function SpPsRest_UpdateSiteScript()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"UpdateSiteScript"

	$myPayload = '{ 
	   "updateInfo":{ 
			"Id":"fde0681c-9512-4652-8198-3f9b9934a394",  
		    "Title":"CustomListFromSiteScript", 
		    "Description":"Custom List Updated", 
		    "Version": 2
			}
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 52

#gavdcodebegin 53
Function SpPsRest_DeleteSiteScript()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"DeleteSiteScript"

	$myPayload = '{ 
			"id":"fde0681c-9512-4652-8198-3f9b9934a394"  
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 53

#gavdcodebegin 54
Function SpPsSpo_AddSiteTemplate
{
	Add-SPOSiteDesign -Title "Custom List From Site Design SPO" `
					  -WebTemplate "64" `
					  -SiteScripts "79a5174f-0712-49c7-b6af-5a45918c55ee" `
					  -Description "Creates a Custom List in a site using SPO Site Design"
}
#gavdcodeend 54

#gavdcodebegin 55
Function SpPsSpo_GetAllSiteTemplates
{
	Get-SPOSiteDesign
}
#gavdcodeend 55

#gavdcodebegin 56
Function SpPsSpo_GetOneSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"

	Get-SPOSiteDesign -Identity $myDesignId
}
#gavdcodeend 56

#gavdcodebegin 57
Function SpPsSpo_GetRunsSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	$mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"

	Get-SPOSiteDesignRun -SiteDesignId $myDesignId -WebUrl $mySiteUrl
}
#gavdcodeend 57

#gavdcodebegin 58
Function SpPsSpo_GetRunStatusSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	$mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"
	
	$myRuns = Get-SPOSiteDesignRun -SiteDesignId $myDesignId -WebUrl $mySiteUrl
	Get-SPOSiteDesignRunStatus -Run $myRuns
}
#gavdcodeend 58

#gavdcodebegin 59
Function SpPsSpo_DeleteSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	
	Remove-SPOSiteDesign -Identity $myDesignId
}
#gavdcodeend 59

#gavdcodebegin 60
Function SpPsSpo_InvokeSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	$mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"

	Invoke-SPOSiteDesign -Identity $myDesignId -WebUrl $mySiteUrl
}
#gavdcodeend 60

#gavdcodebegin 61
Function SpPsSpo_AddTaskSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	$mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"

	Add-SPOSiteDesignTask -SiteDesignId $myDesignId -WebUrl $mySiteUrl
}
#gavdcodeend 61

#gavdcodebegin 62
Function SpPsSpo_GetTaskSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"
	$mySiteUrl = "https://[domain].sharepoint.com/sites/Test_Guitaca"

	Get-SPOSiteDesignTask -Identity $myDesignId -WebUrl $mySiteUrl
}
#gavdcodeend 62

#gavdcodebegin 63
Function SpPsSpoDeleteTaskSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"

	Remove-SPOSiteDesignTask -Identity $myDesignId
}
#gavdcodeend 63

#gavdcodebegin 64
Function SpPsSpo_GrantRightsSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"

	Grant-SPOSiteDesignRights -Identity $myDesignId `
							  -Principals "[user]@[domain].onmicrosoft.com" `
							  -Rights View
}
#gavdcodeend 64

#gavdcodebegin 65
Function SpPsSpo_GetRightsSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"

	Get-SPOSiteDesignRights -Identity $myDesignId
}
#gavdcodeend 65

#gavdcodebegin 66
Function SpPsSpo_DeleteRightsSiteTemplate
{
	$myDesignId = "f155ed5e-d8f9-4ba6-9385-b5f702502540"

	Revoke-SPOSiteDesignRights -Identity $myDesignId `
							   -Principals "[user]@[domain].onmicrosoft.com" `
}
#gavdcodeend 66

#gavdcodebegin 67
Function SpPsRest_AddSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
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

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 67

#gavdcodebegin 68
Function SpPsRest_GetAllSiteTemplates()    #*** LEGACY CODE ***
{
	$endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteDesigns"

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 68

#gavdcodebegin 69
Function SpPsRest_GetOneSiteTemplate()    #*** LEGACY CODE ***
{
	$endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteDesignMetadata"

	$myPayload = @{
			id = 'c80235ae-b26f-431e-9199-d459be24e89f'
			} | ConvertTo-Json
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 69

#gavdcodebegin 70
Function SpPsRest_UpdateSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"UpdateSiteDesign"

	$myPayload = '{ 
	   "updateInfo":{ 
			"Id":"c80235ae-b26f-431e-9199-d459be24e89f",  
		    "Title":"Custom List From REST Site Design", 
		    "Description":"Custom List Updated", 
		    "SiteScriptIds":["79a5174f-0712-49c7-b6af-5a45918c55ee"], 
		    "Version": 2
			}
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 70

#gavdcodebegin 71
Function SpPsRest_DeleteSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"DeleteSiteDesign"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 71

#gavdcodebegin 72
Function SpPsRest_ApplySiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"ApplySiteDesign"

	$myPayload = '{ 
			"siteDesignId":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"webUrl":"https://[domain].sharepoint.com/sites/Test_Guitaca"  
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 72

#gavdcodebegin 73
Function SpPsRest_ApplyToSiteSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"AddSiteDesignTaskToCurrentWeb"

	$myPayload = '{ 
			"siteDesignId":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 73

#gavdcodebegin 74
Function SpPsRest_GetRigthsSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GetSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f"  
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 74

#gavdcodebegin 75
Function SpPsRest_GrantRightsSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"GrantSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"principalNames":["[user]@[domain].onmicrosoft.com"],
			"grantedRights":"View"
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 75

#gavdcodebegin 76
Function SpPsRest_DeleteRightsSiteTemplate()    #*** LEGACY CODE ***
{
    $endpointUrl = $webUrl + "/_api/" + 
			"Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility." + 
			"RevokeSiteDesignRights"

	$myPayload = '{ 
			"id":"c80235ae-b26f-431e-9199-d459be24e89f",  
			"principalNames":["[user]@[domain].onmicrosoft.com"]
	}'

	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
	$data = Invoke-RestSPO -Url $endpointUrl -Method POST -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
	$data | ConvertTo-Json
}
#gavdcodeend 76

#gavdcodebegin 77
Function SpPsPnp_GenerateSiteTemplateXml    #*** LEGACY CODE ***
{
	Get-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningSite.xml"
}
#gavdcodeend 77

#gavdcodebegin 78
Function SpPsPnp_GenerateListsTemplate    #*** LEGACY CODE ***
{
	Get-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningLists.xml" `
								-ListsToExtract "MyCustomList",`
												"7B8f0d6e79-406c-48a9-834e-af0c56489bbe"
}
#gavdcodeend 78

#gavdcodebegin 79
Function SpPsPnp_GenerateTemplateTermGroups    #*** LEGACY CODE ***
{
	Get-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningTermGroups.xml" `
								-IncludeAllTermGroups
}
#gavdcodeend 79

#gavdcodebegin 80
Function SpPsPnp_GenerateSiteTemplatePnP    #*** LEGACY CODE ***
{
	Get-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningSite.pnp"
}
#gavdcodeend 80

#gavdcodebegin 81
Function SpPsPnpApplySiteTemplate    #*** LEGACY CODE ***
{
	Apply-PnPProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
}
#gavdcodeend 81

#gavdcodebegin 82
Function SpPsPnpTenantTemplateConnect    #*** LEGACY CODE ***
{
	Connect-PnPOnline -Graph
}
#gavdcodeend 82

#gavdcodebegin 83
Function SpPsPnp_GenerateTenantTemplateXml    #*** LEGACY CODE ***
{
	Get-PnPTenantTemplate -Out "C:\Temporary\TestProvisioningTenant.xml" `
						  -SiteUrl "https://[domain].sharepoint.com/sites/Test_Guitaca" `
						  -Configuration "C:\Temporary\TestConfiguration.xml"
}
#gavdcodeend 83

#gavdcodebegin 84
Function SpPsPnp_ApplyTenantTemplate    #*** LEGACY CODE ***
{
	Apply-PnPTenantTemplate -Path "C:\Temporary\TestProvisioningTenant.xml"
}
#gavdcodeend 84

#gavdcodebegin 85
Function SpPsPnp_GenerateSiteTemplateWithConfig    #*** LEGACY CODE ***
{
	Get-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningSiteWithConfig.xml" `
						  -Configuration "C:\Temporary\TestConfiguration.xml"
}
#gavdcodeend 85

#gavdcodebegin 86
Function SpPsPnp_GenerateSiteTemplateInMem    #*** LEGACY CODE ***
{
	$myTemplate = PnPProvisioningTemplate -OutputInstance
	$myTemplate | ConvertTo-Json
}
#gavdcodeend 86

#gavdcodebegin 87
Function SpPsPnp_GenerateSiteTemplateInMemFromFile    #*** LEGACY CODE ***
{
	$myTemplate = Read-PnPProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
	$myTemplate | ConvertTo-Json
}
#gavdcodeend 87

#gavdcodebegin 88
Function SpPsPnp_GenerateSiteTemplateInMemFromScratch    #*** LEGACY CODE ***
{
	$myTemplate = New-PnPProvisioningTemplate
	$myTemplate | ConvertTo-Json
}
#gavdcodeend 88

#gavdcodebegin 89
Function SpPsPnp_SaveSiteTemplateInMemFromScratch    #*** LEGACY CODE ***
{
	$myTemplate = New-PnPProvisioningTemplate
	Save-PnPProvisioningTemplate -Out "C:\Temporary\TestProvisioningSiteInMem.xml" `
								 -InputInstance $myTemplate
}
#gavdcodeend 89

#gavdcodebegin 90
Function SpPsPnp_ModifySiteTemplateInMem    #*** LEGACY CODE ***
{
	$myTemplate = Read-PnPProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.xml"
	$myTemplate.DisplayName = "In-memory modified template"
	$myTemplate.Security.AdditionalOwners.Clear()
	$myTemplate | ConvertTo-Json
}
#gavdcodeend 90

#gavdcodebegin 91
Function SpPsPnp_GenerateSiteTemplateInMemFromFilePnP    #*** LEGACY CODE ***
{
	$myTemplate = Read-PnPProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.pnp"
	$myTemplate | ConvertTo-Json
}
#gavdcodeend 91

#gavdcodebegin 92
Function SpPsPnp_AddFileSiteTemplateInMemFromFilePnP    #*** LEGACY CODE ***
{
	Add-PnPFileToProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.pnp" `
									  -Source "C:\Temporary\MyStyles.css" `
								      -Folder "SiteAssets"
}
#gavdcodeend 92

#gavdcodebegin 93
Function SpPsPnp_RemoveFileSiteTemplateInMemFromFilePnP    #*** LEGACY CODE ***
{
	Remove-PnPFileFromProvisioningTemplate -Path "C:\Temporary\TestProvisioningSite.pnp" `
										   -File "MyStyles.css"
}
#gavdcodeend 93

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Search.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"

[xml]$configFile = get-content "C:\Projects\ConfigValuesPS.config"

# CSOM PowerShell Term Store
#$spCtx = LoginPsCsom
#SpPsCsom_FindTermStore $spCtx
#SpPsCsom_CreateTermGroup $spCtx
#SpPsCsom_FindTermGroups $spCtx
#SpPsCsom_CreateTermSet $spCtx
#SpPsCsom_FindTermSets $spCtx
#SpPsCsom_CreateTerm $spCtx
#SpPsCsom_FindTerms $spCtx
#SpPsCsom_FindOneTerm $spCtx
#SpPsCsom_UpdateOneTerm $spCtx
#SpPsCsom_DeleteOneTerm $spCtx
#SpPsCsom_FindTermSetAndTermById $spCtx

# PnP PowerShell Term Store
#$spCtx = LoginPsPnP
#SpPsPnp_FindTermStore
#SpPsPnp_CreateTermGroup
#SpPsPnp_FindTermGroup
#SpPsPnp_CreateTermSet
#SpPsPnp_FindTermSet
#SpPsPnp_CreateTerm
#SpPsPnp_FindTerm
#SpPsPnp_DeleteTermGroup
#Export-PnPTaxonomy
#SpPsPnp_ImportTaxonomy
#SpPsPnp_ExportTermGroup
#SpPsPnp_ImportTermGroup

# CSOM PowerShell Search
#$spCtx = LoginPsCsom
#SpPsCsom_GetResultsSearch $spCtx

# REST PowerShell Search
#$webUrl = $configFile.appsettings.SiteCollUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#SpPsRest_ResultsSearchGET
#SpPsRest_ResultsSearchPOST

# PnP PowerShell Search
#$spCtx = LoginPsPnP
#SpPsPnp_Search
#SpPsPnp_SearchSiteColls
#SpPsPnp_SearchCrawl

# CSOM PowerShell User Profile
#$spCtx = LoginPsCsom
#SpPsCsom_GetAllPropertiesUserProfile $spCtx
#SpCsCsom_GetAllMyPropertiesUserProfile $spCtx
#SpPsCsom_GetPropertiesUserProfile $spCtx
#SpPsCsom_UpdateOnePropertyUserProfile $spCtx
#SpPsCsom_UpdateOneMultPropertyUserProfile $spCtx

# PnP PowerShell User Profile
#$spCtx = LoginPsPnP
#SpPsPnp_FindUserProfileProperties
#SpPsPnp_UpdateUserProfileProperties

# REST PowerShell User Profile
#$webUrl = $configFile.appsettings.SiteCollUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#SpPsRest_GetAllPropertiesUserProfile
#SpPsRest_GetAllMyPropertiesUserProfile
#SpPsRest_GetPropertiesUserProfile

# SPO PowerShell Site Scripts
#LoginPsSPO
#SpPsSpo_GenerateListSiteScript
#SpPsSpo_GenerateWebSiteScript
#SpPsSpo_AddSiteScript
#SpPsSpo_GetAllSiteScripts
#SpPsSpo_GetOneSiteScript
#SpPsSpo_UpdateSiteScript
#SpPsSpo_DeleteOneSiteScript
#Disconnect-SPOService

# REST PowerShell Site Scripts
#$webUrl = $configFile.appsettings.SiteCollUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
#SpPsRest_GenerateListSiteScript
#SpPsRest_GenerateWebSiteScript
#SpPsRest_AddSiteScript
#SpPsRest_GetAllSiteScripts
#SpPsRest_GetOneSiteScript
#SpPsRest_UpdateSiteScript
#SpPsRest_DeleteSiteScript

# SPO PowerShell Site Templates
#LoginPsSPO
#SpPsSpo_AddSiteTemplate
#SpPsSpo_GetAllSiteTemplates
#SpPsSpo_GetOneSiteTemplate
#SpPsSpo_GetRunsSiteTemplate
#SpPsSpo_GetRunStatusSiteTemplate
#SpPsSpo_DeleteSiteTemplate
#SpPsSpo_InvokeSiteTemplate
#SpPsSpo_AddTaskSiteTemplate
#SpPsSpo_GetTaskSiteTemplate
#SpPsSpo_DeleteTaskSiteTemplate
#SpPsSpo_GrantRightsSiteTemplate
#SpPsSpo_GetRightsSiteTemplate
#SpPsSpo_DeleteRightsSiteTemplate
#Disconnect-SPOService

# REST PowerShell Site Designs
#$webUrl = $configFile.appsettings.SiteCollUrl
#$userName = $configFile.appsettings.UserName
#$password = $configFile.appsettings.UserPw
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

# PnP PowerShell Provisioning
#$spCtx = LoginPsPnP
#SpPsPnp_GenerateSiteTemplateXml
#SpPsPnp_GenerateListsTemplate
#SpPsPnp_GenerateTemplateTermGroups
#SpPsPnp_GenerateSiteTemplatePnP
#SpPsPnp_TenantTemplateConnect
#SpPsPnp_GenerateTenantTemplateXml
#SpPsPnp_ApplySiteTemplate
#SpPsPnp_ApplyTenantTemplate
#SpPsPnp_GenerateSiteTemplateWithConfig
#SpPsPnp_GenerateSiteTemplateInMem
#SpPsPnp_GenerateSiteTemplateInMemFromFile
#SpPsPnp_GenerateSiteTemplateInMemFromScratch
#SpPsPnp_SaveSiteTemplateInMemFromScratch
#SpPsPnp_ModifySiteTemplateInMem
#SpPsPnp_GenerateSiteTemplateInMemFromFilePnP
#SpPsPnp_AddFileSiteTemplateInMemFromFilePnP
#SpPsPnp_RemoveFileSiteTemplateInMemFromFilePnP

Write-Host "Done"
