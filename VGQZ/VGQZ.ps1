##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

Function GrPsLoginGraphSDKWithAccPw
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

	Connect-Graph -AccessToken $myToken.AccessToken
}

Function LoginPsCLI()
{
	m365 login --authType password `
			   --userName $configFile.appsettings.UserName `
			   --password $configFile.appsettings.UserPw
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
Function PsGraphSdkLicensing_UserGetLicenseList
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgUser -UserId "user@domain.onmicrosoft.com" `
               -Property Id, displayName, assignedLicenses | Select `
               -ExpandProperty AssignedLicenses

	Disconnect-MgGraph
}
#gavdcodeend 005

#gavdcodebegin 006
Function PsGraphSdkLicensing_UserGetLicenseListDetailSku
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgUserLicenseDetail -UserId "user@domain.onmicrosoft.com" | fl

	Disconnect-MgGraph
}
#gavdcodeend 006

#gavdcodebegin 007
Function PsGraphSdkLicensing_UserGetLicenseListDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

    Get-MgUserLicenseDetail -UserId "user@domain.onmicrosoft.com" | ? `
        {$_.SkuId -eq "c42b9cae-xxxx-4ab7-9717-81576235ccac"} | `
        Select-Object -ExpandProperty ServicePlans

	Disconnect-MgGraph
}
#gavdcodeend 007

#gavdcodebegin 008
Function PsGraphSdkLicensing_GetSku
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

    $oneSku = Get-MgSubscribedSku -All | Where "SkuPartNumber" -eq "DEVELOPERPACK_E5"

    Write-Host "SkuId-" $oneSku.SkuId

	Disconnect-MgGraph
}
#gavdcodeend 008

#gavdcodebegin 009
Function PsGraphSdkLicensing_GetUsersLicenseList
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

    Get-MgUser -Filter `
            "assignedLicenses/any(x:x/skuId eq c42b9cae-xxxx-4ab7-9717-81576235ccac)" `
            -All

	Disconnect-MgGraph
}
#gavdcodeend 009

#gavdcodebegin 010
Function PsGraphSdkLicensing_UserAddLicense
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

    Set-MgUserLicense -UserId "user@domain.onmicrosoft.com" `
                      -Addlicenses @{SkuId = "c42b9cae-xxxx-4ab7-9717-81576235ccac"} `
                      -RemoveLicenses @()

	Disconnect-MgGraph
}
#gavdcodeend 010

#gavdcodebegin 011
Function PsGraphSdkLicensing_UserDeleteLicense
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

    Set-MgUserLicense -UserId "user@domain.onmicrosoft.com" `
                      -Addlicenses @{} `
                      -RemoveLicenses @("c42b9cae-xxxx-4ab7-9717-81576235ccac")

	Disconnect-MgGraph
}
#gavdcodeend 011

#gavdcodebegin 012
Function PsGraphSdkLicensing_UserDisableLicenseAndPlans
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

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
function PsCliLicensing_LicenseGetList
{
    m365 aad license list
}
#gavdcodeend 013

#gavdcodebegin 014
function PsCliLicensing_UserGetLicenseListByName
{
    m365 aad user license list --userName "user@domain.onmicrosoft.com"
}
#gavdcodeend 014

#gavdcodebegin 015
function PsCliLicensing_UserGetLicenseListById
{
    m365 aad user license list --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0"
}
#gavdcodeend 015

#gavdcodebegin 016
function PsCliLicensing_UserAddLicenseListById
{
    m365 aad user license add --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0" `
                              --ids "c42b9cae-xxxx-4ab7-9717-81576235ccac"
}
#gavdcodeend 016

#gavdcodebegin 017
function PsCliLicensing_UserDeleteLicenseListById
{
    m365 aad user license remove --userId "e4ab0702-xxxx-4a17-974f-2d9d0449a7c0" `
                                 --ids "c42b9cae-xxxx-4ab7-9717-81576235ccac"
}
#gavdcodeend 017

##-----------------------------------------------------------

##----> Reporting

#gavdcodebegin 018
Function PsGraphSdkReporting_EmailActivityCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportEmailActivityCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivity.csv"

	Disconnect-MgGraph
}
#gavdcodeend 018

#gavdcodebegin 019
Function PsGraphSdkReporting_EmailActivityUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportEmailActivityUserCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivityUser.csv"

	Disconnect-MgGraph
}
#gavdcodeend 019

#gavdcodebegin 020
Function PsGraphSdkReporting_EmailActivityUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportEmailActivityUserDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\EmailActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 020

#gavdcodebegin 021
Function PsGraphSdkReporting_MailboxUsageDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportMailboxUsageDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxUsageDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 021

#gavdcodebegin 022
Function PsGraphSdkReporting_MailboxUsageCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportMailboxUsageMailboxCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxUsageCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 022

#gavdcodebegin 023
Function PsGraphSdkReporting_MailboxQuota
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportMailboxUsageQuotaStatusMailboxCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxQuota.csv"

	Disconnect-MgGraph
}
#gavdcodeend 023

#gavdcodebegin 024
Function PsGraphSdkReporting_MailboxStorage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportMailboxUsageStorage `
                               -Period "D90" `
                               -OutFile "C:\Temporary\MailboxStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 024

#gavdcodebegin 025
function PsCliReporting_EmailActivityCountToScreen
{
    m365 outlook report mailactivitycounts --period "D90" `
                    --output "csv"
}
#gavdcodeend 025

#gavdcodebegin 026
function PsCliReporting_EmailActivityCountToFile
{
    m365 outlook report mailactivitycounts --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivity.csv"
}
#gavdcodeend 026

#gavdcodebegin 027
function PsCliReporting_EmailActivityUserCountToFile
{
    m365 outlook report mailactivityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivityUser.csv"
}
#gavdcodeend 027

#gavdcodebegin 028
function PsCliReporting_EmailActivityUserDetailToFile
{
    m365 outlook report mailactivityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\EmailActivityUserDetail.csv"
}
#gavdcodeend 028

#gavdcodebegin 029
function PsCliReporting_MailboxUsageDetailToFile
{
    m365 outlook report mailboxusagedetail --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxUsageDetail.csv"
}
#gavdcodeend 029

#gavdcodebegin 030
function PsCliReporting_MailboxUsageCountToFile
{
    m365 outlook report mailboxusagemailboxcount --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxUsageCount.csv"
}
#gavdcodeend 030

#gavdcodebegin 031
function PsCliReporting_MailboxQuotaToFile
{
    m365 outlook report mailboxusagequotastatusmailboxcounts --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxQuota.csv"
}
#gavdcodeend 031

#gavdcodebegin 032
function PsCliReporting_MailboxStorageToFile
{
    m365 outlook report mailboxusagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\MailboxStorage.csv"
}
#gavdcodeend 032

#gavdcodebegin 033
Function PsGraphSdkReporting_M365ActivationCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365ActivationCount `
                               -OutFile "C:\Temporary\M365ActivationCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 033

#gavdcodebegin 034
Function PsGraphSdkReporting_M365ActivationUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365ActivationUserCount `
                               -OutFile "C:\Temporary\M365ActivationUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 034

#gavdcodebegin 035
Function PsGraphSdkReporting_M365ActivationUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365ActivationUserDetail `
                               -OutFile "C:\Temporary\M365ActivationUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 035

#gavdcodebegin 036
Function PsGraphSdkReporting_M365ActiveUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365ActiveUserCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\M365ActivUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 036

#gavdcodebegin 037
Function PsGraphSdkReporting_M365ActiveUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365ActiveUserDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\M365ActivUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 037

#gavdcodebegin 038
function PsCliReporting_M365ActivationCountToScreen
{
    m365 tenant report office365activationcounts
}
#gavdcodeend 038

#gavdcodebegin 039
function PsCliReporting_M365ActivationCountToFile
{
    m365 tenant report office365activationcounts `
                    --output "csv" > "C:\Temporary\M365ActivationCount.csv"
}
#gavdcodeend 039

#gavdcodebegin 040
function PsCliReporting_M365ActivationUserDetailToFile
{
    m365 tenant report office365activationsuserdetail `
                    --output "csv" > "C:\Temporary\M365ActivationUserDetail.csv"
}
#gavdcodeend 040

#gavdcodebegin 041
function PsCliReporting_M365ActivationUserCountToFile
{
    m365 tenant report office365activationsusercounts `
                    --output "csv" > "C:\Temporary\M365ActivationUserCount.csv"
}
#gavdcodeend 041

#gavdcodebegin 042
function PsCliReporting_M365ActiveUserCountToFile
{
    m365 tenant report activeusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\M365ActiveUserCount.csv"
}
#gavdcodeend 042

#gavdcodebegin 043
function PsCliReporting_M365ActiveUserDetailToFile
{
    m365 tenant report activeuserdetail  --period "D90" `
                    --output "csv" > "C:\Temporary\M365ActiveUserDetail.csv"
}
#gavdcodeend 043

#gavdcodebegin 044
Function PsGraphSdkReporting_GroupActivityCount #==> Error in cmd: Gets only reports about Yammer
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365GroupActivityCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 044

#gavdcodebegin 045
Function PsGraphSdkReporting_GroupActivityDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365GroupActivityDetail `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 045

#gavdcodebegin 046
Function PsGraphSdkReporting_GroupActivityFileCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365GroupActivityFileCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 046

#gavdcodebegin 047
Function PsGraphSdkReporting_GroupActivityGroupCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365GroupActivityGroupCount `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityGroupCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 047

#gavdcodebegin 048
Function PsGraphSdkReporting_GroupActivityStorage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOffice365GroupActivityStorage `
                               -Period "D90" `
                               -OutFile "C:\Temporary\GroupActivityStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 048

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 048 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

# Connect to M365 using the CLI
$spCtx = LoginPsCLI

#Adaptive Cards
#JsonAdaptiveCard_01
#PsAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhookWithJson

#Licensing
#PsGraphSdkLicensing_UserGetLicenseList
#PsGraphSdkLicensing_UserGetLicenseListDetailSku
#PsGraphSdkLicensing_UserGetLicenseListDetail
#PsGraphSdkLicensing_GetSku
#PsGraphSdkLicensing_GetUsersLicenseList
#PsGraphSdkLicensing_UserAddLicense
#PsGraphSdkLicensing_UserDeleteLicense
#PsGraphSdkLicensing_UserDisableLicenseAndPlans
#PsCliLicensing_LicenseGetList
#PsCliLicensing_UserGetLicenseListByName
#PsCliLicensing_UserGetLicenseListById
#PsCliLicensing_UserAddLicenseListById
#PsCliLicensing_UserDeleteLicenseListById

#Reporting
#PsGraphSdkReporting_EmailActivityCount
#PsGraphSdkReporting_EmailActivityUserCount
#PsGraphSdkReporting_EmailActivityUserDetail
#PsGraphSdkReporting_MailboxUsageDetail
#PsGraphSdkReporting_MailboxUsageCount
#PsGraphSdkReporting_MailboxQuota
#PsGraphSdkReporting_MailboxStorage
#PsCliReporting_EmailActivityCountToScreen
#PsCliReporting_EmailActivityCountToFile
#PsCliReporting_EmailActivityUserCountToFile
#PsCliReporting_EmailActivityUserDetailToFile
#PsCliReporting_MailboxUsageDetailToFile
#PsCliReporting_MailboxUsageCountToFile
#PsCliReporting_MailboxQuotaToFile
#PsCliReporting_MailboxStorageToFile
#PsGraphSdkReporting_M365ActivationCount
#PsGraphSdkReporting_M365ActivationUserCount
#PsGraphSdkReporting_M365ActivationUserDetail
#PsGraphSdkReporting_M365ActiveUserCount
#PsGraphSdkReporting_M365ActiveUserDetail
#PsCliReporting_M365ActivationCountToScreen
#PsCliReporting_M365ActivationCountToFile
#PsCliReporting_M365ActivationUserDetailToFile
#PsCliReporting_M365ActivationUserCountToFile
#PsCliReporting_M365ActiveUserCountToFile
#PsCliReporting_M365ActiveUserDetailToFile
#PsGraphSdkReporting_GroupActivityCount  #==> Error in cmd: Gets only reports about Yammer
#PsGraphSdkReporting_GroupActivityDetail
#PsGraphSdkReporting_GroupActivityFileCount
#PsGraphSdkReporting_GroupActivityGroupCount
#PsGraphSdkReporting_GroupActivityStorage

m365 logout
Write-Host "Done" 



