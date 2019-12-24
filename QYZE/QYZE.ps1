Function LoginPsCsom()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials `
			($configFile.appsettings.spUserName, $securePW)
	$rtnContext = New-Object Microsoft.SharePoint.Client.ClientContext `
			($configFile.appsettings.spUrl) 
	$rtnContext.Credentials = $myCredentials

	return $rtnContext
}

#----------------------------------------------------------------------------------------

Function LoginPsPnP()
{
	[SecureString]$securePW = ConvertTo-SecureString -String `
			$configFile.appsettings.spUserPw -AsPlainText -Force

	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
			-argumentlist $configFile.appsettings.spUserName, $securePW
	Connect-PnPOnline -Url $configFile.appsettings.spUrl -Credentials $myCredentials
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

Function SpPsCsomFindTermStore($spCtx)
{
    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $spCtx.Load($myTaxSession.TermStores)
    $spCtx.ExecuteQuery()

    foreach ($oneTermStore in $myTaxSession.TermStores) {
        Write-Host($oneTermStore.Name)
    }
}

Function SpPsCsomCreateTermGroup($spCtx)
{
    $termStoreName = "Taxonomy_hVIOdhme2obc+5zqZXqqUQ=="

    $myTaxSession = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::`
															GetTaxonomySession($spCtx)
    $myTermStore = $myTaxSession.TermStores.GetByName($termStoreName)

	$myNewGuid = New-Guid
    $myTermGroup = $myTermStore.CreateGroup("PsCsomTermGroup", $myNewGuid)
    $spCtx.ExecuteQuery()
}

Function SpPsCsomFindTermGroups($spCtx)
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

Function SpPsCsomCreateTermSet($spCtx)
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

Function SpPsCsomFindTermSets($spCtx)
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

Function SpPsCsomCreateTerm($spCtx)
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

Function SpPsCsomFindTerms($spCtx)
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

Function SpPsCsomFindOneTerm($spCtx)
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

Function SpPsCsomUpdateOneTerm($spCtx)
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

Function SpPsCsomDeleteOneTerm($spCtx)
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

Function SpPsCsomFindTermSetAndTermById($spCtx)
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

Function SpPsPnpFindTermStore()
{
	$myTaxSession = Get-PnPTaxonomySession
	Write-Host $myTaxSession.TermStores[0].Name
}

Function SpPsPnpCreateTermGroup()
{
	$myTermGroup = New-PnPTermGroup -Name "PsPnpTermGroup"
	Write-Host $myTermGroup.Id
}

Function SpPsPnpFindTermGroup()
{
	$myTermGroups = Get-PnPTermGroup
	foreach ($oneGroup in $myTermGroups) {
		Write-Host $oneGroup.Id
	}
}

Function SpPsPnpCreateTermSet()
{
	$myTermSet = New-PnPTermSet -Name "PsPnpTermSet" `
								-TermGroup "PsPnpTermGroup"
	Write-Host $myTermSet.Id
}

Function SpPsPnpFindTermSet()
{
	$myTermSets = Get-PnPTermSet -TermGroup "PsPnpTermGroup"
	foreach ($oneSet in $myTermSets) {
		Write-Host $oneSet.Id
	}
}

Function SpPsPnpCreateTerm()
{
	$myTerm = New-PnPTerm -Name "PsPnpTerm" `
						  -TermGroup "PsPnpTermGroup" `
						  -TermSet "PsPnpTermSet"
	Write-Host $myTerm.Id
}

Function SpPsPnpFindTerm()
{
	$myTerms = Get-PnPTerm -TermGroup "PsPnpTermGroup" `
						   -TermSet "PsPnpTermSet"
	foreach ($oneTerm in $myTerms) {
		Write-Host $oneTerm.Id
	}
}

Function SpPsPnpDeleteTermGroup()
{
	Remove-PnPTermGroup -GroupName "PsPnpTermGroup"
}

Function SpPsPnpExportTaxonomy()
{
	Export-PnPTaxonomy -Path "C:\Temporary\tax.txt" `
					   -TermSet "529c954a-0235-4202-a739-9b871055427c"
}

Function SpPsPnpImportTaxonomy()
{
	Import-PnPTaxonomy -Path "C:\Temporary\tax.txt"
}

Function SpPsPnpExportTermGroup()
{
	Export-PnPTermGroupToXml -Out "C:\Temporary\group.xml" -Identity "PsCsomTermGroup"
}

Function SpPsPnpImportTermGroup()
{
	Import-PnPTermGroupToXml -Path "C:\Temporary\tax.txt"
}

Function SpPsCsomGetResultsSearch($spCtx)
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

Function SpPsRestResultsSearchGET()
{
    $endpointUrl = $webUrl + "/_api/search/query?querytext='team'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}

Function SpPsRestResultsSearchPOST()
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

Function SpPsPnpSearch()
{
	Submit-PnPSearchQuery -Query "team"
}

Function SpPsPnpSearchSiteColls()
{
	Get-PnPSiteSearchQueryResults
}

Function SpPsPnpSearchCrawl()
{
	Get-PnPSearchCrawlLog
}

Function SpPsCsomGetAllPropertiesUserProfile ($spCtx)
{
    $myUser = "i:0#.f|membership|" + $configFile.appsettings.spUserName
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

Function SpCsCsomGetAllMyPropertiesUserProfile($spCtx)
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

Function SpPsCsomGetPropertiesUserProfile($spCtx)
{
    $myUser = "i:0#.f|membership|" + $configFile.appsettings.spUserName
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

Function SpPsCsomUpdateOneMultPropertyUserProfile($spCtx)
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

Function SpPsPnpFindUserProfileProperties()
{
	Get-PnPUserProfileProperty -Account $configFile.appsettings.spUserName
}

Function SpPsPnpUpdateUserProfileProperties()
{
	Set-PnPUserProfileProperty -Account $configFile.appsettings.spUserName `
							   -Property "AboutMe" `
							   -Value "I am not the administrator"
}

Function SpPsRestGetAllPropertiesUserProfile()
{
    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.spUserName.Replace("@", "%40");
    $endpointUrl = $webUrl + "/_api/sp.userprofiles.peoplemanager/" + `
						"getpropertiesfor(@v)?@v='" + $myUser + "'"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}

Function SpPsRestGetAllMyPropertiesUserProfile()
{
    $endpointUrl = $webUrl + "/_api/sp.userprofiles.peoplemanager/getmyproperties"
	$contextInfo = Get-SPOContextInfo -WebUrl $webUrl -UserName $userName `
																-Password $password
    $data = Invoke-RestSPO -Url $endpointUrl -Method GET -UserName $userName -Password `
						$password -Metadata $myPayload -RequestDigest `
						$contextInfo.GetContextWebInformation.FormDigestValue 
    $data | ConvertTo-Json
}

Function SpPsRestGetPropertiesUserProfile()
{
    $myUser = "i%3A0%23.f%7Cmembership%7C" + `
                     $configFile.appsettings.spUserName#.Replace("@", "%40");
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

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Search.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"

[xml]$configFile = get-content "C:\Projects\spPs.values.config"

# CSOM PowerShell Term Store
#$spCtx = LoginPsCsom
#SpPsCsomFindTermStore $spCtx
#SpPsCsomCreateTermGroup $spCtx
#SpPsCsomFindTermGroups $spCtx
#SpPsCsomCreateTermSet $spCtx
#SpPsCsomFindTermSets $spCtx
#SpPsCsomCreateTerm $spCtx
#SpPsCsomFindTerms $spCtx
#SpPsCsomFindOneTerm $spCtx
#SpPsCsomUpdateOneTerm $spCtx
#SpPsCsomDeleteOneTerm $spCtx
#SpPsCsomFindTermSetAndTermById $spCtx

# PnP PowerShell Term Store
#$spCtx = LoginPsPnP
#SpPsPnpFindTermStore
#SpPsPnpCreateTermGroup
#SpPsPnpFindTermGroup
#SpPsPnpCreateTermSet
#SpPsPnpFindTermSet
#SpPsPnpCreateTerm
#SpPsPnpFindTerm
#SpPsPnpDeleteTermGroup
#Export-PnPTaxonomy
#SpPsPnpImportTaxonomy
#SpPsPnpExportTermGroup
#SpPsPnpImportTermGroup

# CSOM PowerShell Search
#$spCtx = LoginPsCsom
#SpPsCsomGetResultsSearch $spCtx

# REST PowerShell Search
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRestResultsSearchGET
#SpPsRestResultsSearchPOST

# PnP PowerShell Search
#$spCtx = LoginPsPnP
#SpPsPnpSearch
#SpPsPnpSearchSiteColls
#SpPsPnpSearchCrawl

# CSOM PowerShell User Profile
#$spCtx = LoginPsCsom
#SpPsCsomGetAllPropertiesUserProfile $spCtx
#SpCsCsomGetAllMyPropertiesUserProfile $spCtx
#SpPsCsomGetPropertiesUserProfile $spCtx
#SpPsCsomUpdateOnePropertyUserProfile $spCtx
#SpPsCsomUpdateOneMultPropertyUserProfile $spCtx

# PnP PowerShell User Profile
#$spCtx = LoginPsPnP
#SpPsPnpFindUserProfileProperties
#SpPsPnpUpdateUserProfileProperties

# REST PowerShell User Profile
#$webUrl = $configFile.appsettings.spUrl
#$userName = $configFile.appsettings.spUserName
#$password = $configFile.appsettings.spUserPw
#SpPsRestGetAllPropertiesUserProfile
#SpPsRestGetAllMyPropertiesUserProfile
#SpPsRestGetPropertiesUserProfile

Write-Host "Done"

