##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function PsGraphSDK_LoginWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $UserName, $securePW

	$myToken = Get-MsalToken -TenantId $TenantName `
							 -ClientId $ClientId `
							 -UserCredential $myCredentials 

	[SecureString]$secureToken = ConvertTo-SecureString -String `
											$myToken.AccessToken -AsPlainText -Force

	Connect-Graph -AccessToken $secureToken
}

Function PsCliM365_LoginWithAccPw
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientIdWithAccPw
	)

	m365 login --authType password `
			   --appId $ClientIdWithAccPw `
			   --userName $UserName `
			   --password $UserPw
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

##----> Adaptive Cards

function JsonAdaptiveCard_01 
{ 

@"
#gavdcodebegin 001
{
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.0",
    "body": [
        {
            "type": "TextBlock",
            "text": "Visit the Guitaca site",
            "size": "Large"
        },
        {
            "type": "TextBlock",
            "text": "The online place for this book"
        },
        {
            "type": "Input.Text",
            "id": "feedbackText",
            "placeholder": "Let us know your thoughts about it..."
        }
    ],
    "actions": [
        {
            "type": "Action.OpenUrl",
            "title": "Guitaca Site",
            "url": "https://guitaca.com"
        }
    ],
    "padding": "None"
}
#gavdcodeend 001
"@

}

#gavdcodebegin 002
function PsAdaptiveCards_SendCardToWebhook
{
    $myWebHookUrl = "https://domain.webhook.office.com/webhookb2/dd12..."
    $myWebHookText = ' 
        { 
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json", 
            "type": "AdaptiveCard", 
            "version": "1.0", 
            "summary": "This is a test",
            "body": [ 
                { 
                    "type": "TextBlock", 
                    "text": "Hello!", 
                    "size": "large" 
                }, 
                { 
                    "type": "TextBlock", 
                    "text": "This is an Adaptive Card example for Teams.", 
                    "wrap": true 
                } 
            ] 
        }'

    Invoke-RestMethod -Method POST `
                      -ContentType "Application/Json" `
                      -Body $myWebHookText `
                      -Uri $myWebHookUrl
}
#gavdcodeend 002

#gavdcodebegin 003
function PsCliAdaptiveCards_SendCardToWebhook
{
    m365 adaptivecard send `
            --url "https://domain.webhook.office.com/webhookb2/dd12..." `
            --title "Adaptive Card using the CLI for M365" `
            --description "This is an example of Adaptive Cards" `
            --imageUrl "https://guitaca.com/Pics/GuitacaLogoWebTop.png" `
            --actionUrl "https://guitaca.com"
}
#gavdcodeend 003

#gavdcodebegin 004
function PsCliAdaptiveCards_SendCardToWebhookWithJson
{
    $myJsonCard = '{"type":"AdaptiveCard","body":[{"type":"TextBlock",...}'
    
    m365 adaptivecard send `
            --url "https://domain.webhook.office.com/webhookb2/dd12..." `
            --card $myJsonCard

}
#gavdcodeend 004

##-----------------------------------------------------------

##----> Licensing

#gavdcodebegin 005
Function PsLicensingGraphSdk_UserGetLicenseList
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgUser -UserId "user@domain.onmicrosoft.com" `
               -Property Id, displayName, assignedLicenses | Select `
               -ExpandProperty AssignedLicenses

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsLicensingGraphSdk_UserGetLicenseListDetailSku
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgUserLicenseDetail -UserId "user@tenant.onmicrosoft.com" | fl

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsLicensingGraphSdk_UserGetLicenseListDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    Get-MgUserLicenseDetail -UserId "user@domain.onmicrosoft.com" | ? `
        {$_.SkuId -eq "c42b9cae-xxxx-4ab7-9717-81576235ccac"} | `
        Select-Object -ExpandProperty ServicePlans

	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsLicensingGraphSdk_GetSku
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    $oneSku = Get-MgSubscribedSku -All | Where "SkuPartNumber" -eq "DEVELOPERPACK_E5"

    Write-Host "SkuId-" $oneSku.SkuId

	Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsLicensingGraphSdk_GetUsersLicenseList
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    Get-MgUser -Filter `
            "assignedLicenses/any(x:x/skuId eq c42b9cae-xxxx-4ab7-9717-81576235ccac)" `
            -All

	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsLicensingGraphSdk_UserAddLicense
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    Set-MgUserLicense -UserId "user@domain.onmicrosoft.com" `
                      -Addlicenses @{SkuId = "c42b9cae-xxxx-4ab7-9717-81576235ccac"} `
                      -RemoveLicenses @()

	Disconnect-MgGraph
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsLicensingGraphSdk_UserDeleteLicense
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    Set-MgUserLicense -UserId "user@domain.onmicrosoft.com" `
                      -Addlicenses @{} `
                      -RemoveLicenses @("c42b9cae-xxxx-4ab7-9717-81576235ccac")

	Disconnect-MgGraph
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsLicensingGraphSdk_UserDisableLicenseAndPlans
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

    $myLicenseOptions = @{SkuId = "c42b9cae-xxxx-4ab7-9717-81576235ccac"; `
                          DisabledPlans = @("a1ace008-xxxx-4ea0-8dac-33b3a23a2472", `
                                            "199a5c09-xxxx-4e37-8f7c-b05d533e1ea2")}

    Set-MgUserLicense -UserId "user@domain.onmicrosoft.com" `
                      -Addlicenses @($myLicenseOptions) `
                      -RemoveLicenses @()


	Disconnect-MgGraph
}
#gavdcodeend 012

#gavdcodebegin 013
function PsLicensingCli365_LicenseGetList
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 entra license list

	m365 logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsLicensingCli365_UserGetLicenseListByName
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 entra user license list --userName "user@domain.onmicrosoft.com"

	m365 logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsLicensingCli365_UserGetLicenseListById
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 entra user license list --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0"

	m365 logout
}
#gavdcodeend 015

#gavdcodebegin 016
function PsLicensingCli365_UserAddLicenseListById
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 entra user license add --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0" `
                              --ids "c42b9cae-xxxx-4ab7-9717-81576235ccac"

	m365 logout
}
#gavdcodeend 016

#gavdcodebegin 017
function PsLicensingCli365_UserDeleteLicenseListById
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 entra user license remove --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0" `
                                 --ids "c42b9cae-xxxx-4ab7-9717-81576235ccac"

	m365 logout
}
#gavdcodeend 017

##-----------------------------------------------------------

##----> Reporting

#gavdcodebegin 018
Function PsReportingGraphSdk_EmailActivityCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportEmailActivityCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivity.csv"

	Disconnect-MgGraph
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsReportingGraphSdk_EmailActivityUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportEmailActivityUserCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivityUser.csv"

	Disconnect-MgGraph
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsReportingGraphSdk_EmailActivityUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportEmailActivityUserDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsReportingGraphSdk_MailboxUsageDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportMailboxUsageDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxUsageDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsReportingGraphSdk_MailboxUsageCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportMailboxUsageMailboxCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxUsageCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsReportingGraphSdk_MailboxQuota
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportMailboxUsageQuotaStatusMailboxCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxQuota.csv"

	Disconnect-MgGraph
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsReportingGraphSdk_MailboxStorage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportMailboxUsageStorage `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 024

#gavdcodebegin 025
function PsReportingCli365_EmailActivityCountToScreen
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailactivitycounts --period "D90" `
                    --output "csv"

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function PsReportingCli365_EmailActivityCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailactivitycounts --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivity.csv"

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function PsReportingCli365_EmailActivityUserCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailactivityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivityUser.csv"

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function PsReportingCli365_EmailActivityUserDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailactivityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivityUserDetail.csv"

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
function PsReportingCli365_MailboxUsageDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailboxusagedetail --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxUsageDetail.csv"

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
function PsReportingCli365_MailboxUsageCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailboxusagemailboxcount --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxUsageCount.csv"

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
function PsReportingCli365_MailboxQuotaToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailboxusagequotastatusmailboxcounts --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxQuota.csv"

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
function PsReportingCli365_MailboxStorageToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 outlook report mailboxusagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxStorage.csv"

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
Function PsReportingGraphSdk_M365ActivationCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365ActivationCount `
                               -OutFile "C:\Temporary\M365ActivationCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsReportingGraphSdk_M365ActivationUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365ActivationUserCount `
                               -OutFile "C:\Temporary\M365ActivationUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsReportingGraphSdk_M365ActivationUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365ActivationUserDetail `
                               -OutFile "C:\Temporary\M365ActivationUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsReportingGraphSdk_M365ActiveUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365ActiveUserCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\M365ActivUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsReportingGraphSdk_M365ActiveUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365ActiveUserDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\M365ActivUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 037

#gavdcodebegin 038
function PsReportingCli365_M365ActivationCountToScreen
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report office365activationcounts

	m365 logout
}
#gavdcodeend 038

#gavdcodebegin 039
function PsReportingCli365_M365ActivationCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report office365activationcounts `
                    --output "csv" > "C:\Temporary\M365ActivationCount.csv"

	m365 logout
}
#gavdcodeend 039

#gavdcodebegin 040
function PsReportingCli365_M365ActivationUserDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report office365activationsuserdetail `
                    --output "csv" > "C:\Temporary\M365ActivationUserDetail.csv"

	m365 logout
}
#gavdcodeend 040

#gavdcodebegin 041
function PsReportingCli365_M365ActivationUserCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report office365activationsusercounts `
                    --output "csv" > "C:\Temporary\M365ActivationUserCount.csv"

	m365 logout
}
#gavdcodeend 041

#gavdcodebegin 042
function PsReportingCli365_M365ActiveUserCountToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report activeusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\M365ActiveUserCount.csv"

	m365 logout
}
#gavdcodeend 042

#gavdcodebegin 043
function PsReportingCli365_M365ActiveUserDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 tenant report activeuserdetail  --period "D90" `
                    --output "csv" > "C:\Temporary\M365ActiveUserDetail.csv"

	m365 logout
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsReportingGraphSdk_GroupActivityCount #==> Error in cmd: Gets only reports about Yammer
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365GroupActivityCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 044

#gavdcodebegin 045
Function PsReportingGraphSdk_GroupActivityDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365GroupActivityDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsReportingGraphSdk_GroupActivityFileCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365GroupActivityFileCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 046

#gavdcodebegin 047
Function PsReportingGraphSdk_GroupActivityGroupCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365GroupActivityGroupCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityGroupCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 047

#gavdcodebegin 048
Function PsReportingGraphSdk_GroupActivityStorage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOffice365GroupActivityStorage `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 048

#gavdcodebegin 049
Function PsReportingGraphSdk_OneDriveActivityFileCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveActivityFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 049

#gavdcodebegin 050
Function PsReportingGraphSdk_OneDriveActivityUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 050

#gavdcodebegin 051
Function PsReportingGraphSdk_OneDriveActivityUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 051

#gavdcodebegin 052
Function PsReportingGraphSdk_OneDriveUsageAccountCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveUsageAccountCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageAccountCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 052

#gavdcodebegin 053
Function PsReportingGraphSdk_OneDriveUsageAccountDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveUsageAccountDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageAccountDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 053

#gavdcodebegin 054
Function PsReportingGraphSdk_OneDriveUsageFileCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveUsageFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 054

#gavdcodebegin 055
Function PsReportingGraphSdk_OneDriveUsageStorage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportOneDriveUsageStorage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 055

#gavdcodebegin 056
function PsReportingCli365_OneDriveActivityFileCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report activityfilecounts --period "D90"

	m365 logout
}
#gavdcodeend 056

#gavdcodebegin 057
function PsReportingCli365_OneDriveActivityFileCountsToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report activityfilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityFileCounts.csv"

	m365 logout
}
#gavdcodeend 057

#gavdcodebegin 058
function PsReportingCli365_OneDriveActivityUserCountsToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report activityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserCounts.csv"

	m365 logout
}
#gavdcodeend 058

#gavdcodebegin 059
function PsReportingCli365_OneDriveActivityUserDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report activityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserDetail.csv"

	m365 logout
}
#gavdcodeend 059

#gavdcodebegin 060
function PsReportingCli365_OneDriveActivityUserDetailByDateToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report activityuserdetail --date "2023-12-16" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserDetailByDate.csv"

	m365 logout
}
#gavdcodeend 060

#gavdcodebegin 061
function PsReportingCli365_OneDriveUsageAccountCountsToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report usageaccountcounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageAccountCounts.csv"

	m365 logout
}
#gavdcodeend 061

#gavdcodebegin 062
function PsReportingCli365_OneDriveUsageAccountDetailToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report usageaccountdetail --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageAccountDetail.csv"

	m365 logout
}
#gavdcodeend 062

#gavdcodebegin 063
function PsReportingCli365_OneDriveUsageFileCountsToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report usagefilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageFileCounts.csv"

	m365 logout
}
#gavdcodeend 063

#gavdcodebegin 064
function PsReportingCli365_OneDriveUsageStorageToFile
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 onedrive report usagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageStorage.csv"

	m365 logout
}
#gavdcodeend 064

#gavdcodebegin 065
Function PsReportingGraphSdk_SharePointActivityFileCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointActivityFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 065

#gavdcodebegin 066
Function PsReportingGraphSdk_SharePointActivityPage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointActivityPage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityPage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 066

#gavdcodebegin 067
Function PsReportingGraphSdk_SharePointActivityUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 067

#gavdcodebegin 068
Function PsReportingGraphSdk_SharePointActivityUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 068

#gavdcodebegin 069
Function PsReportingGraphSdk_SharePointSiteUsageDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointSiteUsageDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 069

#gavdcodebegin 070
Function PsReportingGraphSdk_SharePointSiteUsageFileCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointSiteUsageFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 070

#gavdcodebegin 071
Function PsReportingGraphSdk_SharePointSiteUsagePage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointSiteUsagePage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsagePage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 071

#gavdcodebegin 072
Function PsReportingGraphSdk_SharePointSiteUsageSiteCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointSiteUsageSiteCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageSiteCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 072

#gavdcodebegin 073
Function PsReportingGraphSdk_SharePointSiteUsageStorage
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportSharePointSiteUsageStorage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 073

#gavdcodebegin 074
function PsReportingCli365_SharePointActivityFileCount
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report activityfilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityFileCount.csv"

	m365 logout
}
#gavdcodeend 074

#gavdcodebegin 075
function PsReportingCli365_SharePointActivityPages
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report activitypages --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityPages.csv"

	m365 logout
}
#gavdcodeend 075

#gavdcodebegin 076
function PsReportingCli365_SharePointActivityUserCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report activityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityUserCounts.csv"

	m365 logout
}
#gavdcodeend 076

#gavdcodebegin 077
function PsReportingCli365_SharePointActivityUserDetail
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report activityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityUserDetail.csv"

	m365 logout
}
#gavdcodeend 077

#gavdcodebegin 078
function PsReportingCli365_SharePointSiteUsageDetail
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report siteusagedetail --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteusageDetail.csv"

	m365 logout
}
#gavdcodeend 078

#gavdcodebegin 079
function PsReportingCli365_SharePointSiteUsageFileCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report siteusagefilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageFileCounts.csv"

	m365 logout
}
#gavdcodeend 079

#gavdcodebegin 080
function PsReportingCli365_SharePointSiteSiteUsagePages
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report siteusagepages --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsagePages.csv"

	m365 logout
}
#gavdcodeend 080

#gavdcodebegin 081
function PsReportingCli365_SharePointSiteUsageSiteCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report siteusagesitecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageSiteCounts.csv"

	m365 logout
}
#gavdcodeend 081

#gavdcodebegin 082
function PsReportingCli365_SharePointSiteUsageStorage
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 spo report siteusagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageStorage.csv"

	m365 logout
}
#gavdcodeend 082

#gavdcodebegin 083
Function PsReportingGraphSdk_TeamActivityCount   #==> Does not exist
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamActivityCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 083

#gavdcodebegin 084
Function PsReportingGraphSdk_TeamActivityDetail   #==> Does not exist
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamActivityDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 084

#gavdcodebegin 085
Function PsReportingGraphSdk_TeamActivityDistributionCount   #==> Does not exist
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamActivityDistributionCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityDistributionCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 085

#gavdcodebegin 086
Function PsReportingGraphSdk_TeamCount   #==> Does not exist
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 086

#gavdcodebegin 087
Function PsReportingGraphSdk_TeamUserActivityCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamUserActivityCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 087

#gavdcodebegin 088
Function PsReportingGraphSdk_TeamUserActivityUserCount
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamUserActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 088

#gavdcodebegin 089
Function PsReportingGraphSdk_TeamUserActivityUserDetail
{
	PsGraphSDK_LoginWithAccPw -TenantName $cnfTenantName  `
						      -ClientID $cnfClientIdWithAccPw `
						      -UserName $cnfUserName `
						      -UserPw $cnfUserPw

	Get-MgReportTeamUserActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 089

#gavdcodebegin 090
function PsReportingCli365_TeamDirectRoutingCalls #==> Needs CallRecords.Read.All permission
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 teams report directroutingcalls --debug `
					--fromDateTime 2023-12-01 --toDateTime 2023-12-17 `
                    --output "csv" > "C:\Temporary\DirectRoutingCalls.csv"

	m365 logout
}
#gavdcodeend 090

#gavdcodebegin 091
function PsReportingCli365_TeamPstnCalls #==> Needs CallRecords.Read.All permission
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 teams report pstncalls `
					--fromDateTime 2023-12-01 --toDateTime 2023-12-17 `
                    --output "csv" > "C:\Temporary\PstnCalls.csv"

	m365 logout
}
#gavdcodeend 091

#gavdcodebegin 092
function PsReportingCli365_TeamUserActivityCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 teams report useractivitycounts --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityCounts.csv"

	m365 logout
}
#gavdcodeend 092

#gavdcodebegin 093
function PsReportingCli365_TeamUserActivityUserCounts
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 teams report useractivityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityUserCounts.csv"

	m365 logout
}
#gavdcodeend 093

#gavdcodebegin 094
function PsReportingCli365_TeamUserActivityUserDetail
{
	PsCliM365_LoginWithAccPw $cnfUserName $cnfUserPw $cnfClientIdWithAccPw
	
    m365 teams report useractivityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityUserDetail.csv"

	m365 logout
}
#gavdcodeend 094

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 094 ***

#region ConfigValuesCS.config
[xml]$config = Get-Content -Path "C:\Projects\ConfigValuesCS.config"
$cnfUserName               = $config.SelectSingleNode("//add[@key='UserName']").value
$cnfUserPw                 = $config.SelectSingleNode("//add[@key='UserPw']").value
$cnfTenantUrl              = $config.SelectSingleNode("//add[@key='TenantUrl']").value     # https://domain.onmicrosoft.com
$cnfSiteBaseUrl            = $config.SelectSingleNode("//add[@key='SiteBaseUrl']").value   # https://domain.sharepoint.com
$cnfSiteAdminUrl           = $config.SelectSingleNode("//add[@key='SiteAdminUrl']").value  # https://domain-admin.sharepoint.com
$cnfSiteCollUrl            = $config.SelectSingleNode("//add[@key='SiteCollUrl']").value   # https://domain.sharepoint.com/sites/TestSite
$cnfTenantName             = $config.SelectSingleNode("//add[@key='TenantName']").value
$cnfClientIdWithAccPw      = $config.SelectSingleNode("//add[@key='ClientIdWithAccPw']").value
$cnfClientIdWithSecret     = $config.SelectSingleNode("//add[@key='ClientIdWithSecret']").value
$cnfClientSecret           = $config.SelectSingleNode("//add[@key='ClientSecret']").value
$cnfClientIdWithCert       = $config.SelectSingleNode("//add[@key='ClientIdWithCert']").value
$cnfCertificateThumbprint  = $config.SelectSingleNode("//add[@key='CertificateThumbprint']").value
$cnfCertificateFilePath    = $config.SelectSingleNode("//add[@key='CertificateFilePath']").value
$cnfCertificateFilePw      = $config.SelectSingleNode("//add[@key='CertificateFilePw']").value
#endregion ConfigValuesCS.config

#Adaptive Cards
#JsonAdaptiveCard_01
#PsAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhookWithJson

#Licensing
#PsLicensingGraphSdk_UserGetLicenseList
#PsLicensingGraphSdk_UserGetLicenseListDetailSku
#PsLicensingGraphSdk_UserGetLicenseListDetail
#PsLicensingGraphSdk_GetSku
#PsLicensingGraphSdk_GetUsersLicenseList
#PsLicensingGraphSdk_UserAddLicense
#PsLicensingGraphSdk_UserDeleteLicense
#PsLicensingGraphSdk_UserDisableLicenseAndPlans
#PsLicensingCli365_LicenseGetList
#PsLicensingCli365_UserGetLicenseListByName
#PsLicensingCli365_UserGetLicenseListById
#PsLicensingCli365_UserAddLicenseListById
#PsLicensingCli365_UserDeleteLicenseListById

#Reporting
#PsReportingGraphSdk_EmailActivityCount
#PsReportingGraphSdk_EmailActivityUserCount
#PsReportingGraphSdk_EmailActivityUserDetail
#PsReportingGraphSdk_MailboxUsageDetail
#PsReportingGraphSdk_MailboxUsageCount
#PsReportingGraphSdk_MailboxQuota
#PsReportingGraphSdk_MailboxStorage
#PsReportingCli365_EmailActivityCountToScreen
#PsReportingCli365_EmailActivityCountToFile
#PsReportingCli365_EmailActivityUserCountToFile
#PsReportingCli365_EmailActivityUserDetailToFile
#PsReportingCli365_MailboxUsageDetailToFile
#PsReportingCli365_MailboxUsageCountToFile
#PsReportingCli365_MailboxQuotaToFile
#PsReportingCli365_MailboxStorageToFile
#PsReportingGraphSdk_M365ActivationCount
#PsReportingGraphSdk_M365ActivationUserCount
#PsReportingGraphSdk_M365ActivationUserDetail
#PsReportingGraphSdk_M365ActiveUserCount
#PsReportingGraphSdk_M365ActiveUserDetail
#PsReportingCli365_M365ActivationCountToScreen
#PsReportingCli365_M365ActivationCountToFile
#PsReportingCli365_M365ActivationUserDetailToFile
#PsReportingCli365_M365ActivationUserCountToFile
#PsReportingCli365_M365ActiveUserCountToFile
#PsReportingCli365_M365ActiveUserDetailToFile
#PsReportingGraphSdk_GroupActivityCount  #==> Error in cmd: Gets only reports about Yammer
#PsReportingGraphSdk_GroupActivityDetail
#PsReportingGraphSdk_GroupActivityFileCount
#PsReportingGraphSdk_GroupActivityGroupCount
#PsReportingGraphSdk_GroupActivityStorage
#PsReportingGraphSdk_OneDriveActivityFileCount
#PsReportingGraphSdk_OneDriveActivityUserCount
#PsReportingGraphSdk_OneDriveActivityUserDetail
#PsReportingGraphSdk_OneDriveUsageAccountCount
#PsReportingGraphSdk_OneDriveUsageAccountDetail
#PsReportingGraphSdk_OneDriveUsageFileCount
#PsReportingGraphSdk_OneDriveUsageStorage
#PsReportingCli365_OneDriveActivityFileCounts
#PsReportingCli365_OneDriveActivityFileCountsToFile
#PsReportingCli365_OneDriveActivityUserCountsToFile
#PsReportingCli365_OneDriveActivityUserDetailToFile
#PsReportingCli365_OneDriveActivityUserDetailByDateToFile
#PsReportingCli365_OneDriveUsageAccountCountsToFile
#PsReportingCli365_OneDriveUsageAccountDetailToFile
#PsReportingCli365_OneDriveUsageFileCountsToFile
#PsReportingCli365_OneDriveUsageStorageToFile
#PsReportingGraphSdk_SharePointActivityFileCount
#PsReportingGraphSdk_SharePointActivityPage
#PsReportingGraphSdk_SharePointActivityUserCount
#PsReportingGraphSdk_SharePointActivityUserDetail
#PsReportingGraphSdk_SharePointSiteUsageDetail
#PsReportingGraphSdk_SharePointSiteUsageFileCount
#PsReportingGraphSdk_SharePointSiteUsagePage
#PsReportingGraphSdk_SharePointSiteUsageSiteCount
#PsReportingGraphSdk_SharePointSiteUsageStorage
#PsReportingCli365_SharePointActivityFileCount
#PsReportingCli365_SharePointActivityPages
#PsReportingCli365_SharePointActivityUserCounts
#PsReportingCli365_SharePointActivityUserDetail
#PsReportingCli365_SharePointSiteUsageDetail
#PsReportingCli365_SharePointSiteUsageFileCounts
#PsReportingCli365_SharePointSiteSiteUsagePages
#PsReportingCli365_SharePointSiteUsageSiteCounts
#PsReportingCli365_SharePointSiteUsageStorage
#PsReportingGraphSdk_TeamActivityCount
#PsReportingGraphSdk_TeamActivityDetail
#PsReportingGraphSdk_TeamActivityDistributionCount
#PsReportingGraphSdk_TeamCount
#PsReportingGraphSdk_TeamUserActivityCount
#PsReportingGraphSdk_TeamUserActivityUserCount
#PsReportingGraphSdk_TeamUserActivityUserDetail
#PsReportingCli365_TeamDirectRoutingCalls
#PsReportingCli365_TeamPstnCalls
#PsReportingCli365_TeamUserActivityCounts
#PsReportingCli365_TeamUserActivityUserCounts
#PsReportingCli365_TeamUserActivityUserDetail

Write-Host "Done" 
