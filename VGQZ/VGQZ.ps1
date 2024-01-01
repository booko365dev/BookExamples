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

#gavdcodebegin 049
Function PsGraphSdkReporting_OneDriveActivityFileCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveActivityFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 049

#gavdcodebegin 050
Function PsGraphSdkReporting_OneDriveActivityUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 050

#gavdcodebegin 051
Function PsGraphSdkReporting_OneDriveActivityUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 051

#gavdcodebegin 052
Function PsGraphSdkReporting_OneDriveUsageAccountCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveUsageAccountCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageAccountCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 052

#gavdcodebegin 053
Function PsGraphSdkReporting_OneDriveUsageAccountDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveUsageAccountDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageAccountDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 053

#gavdcodebegin 054
Function PsGraphSdkReporting_OneDriveUsageFileCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveUsageFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 054

#gavdcodebegin 055
Function PsGraphSdkReporting_OneDriveUsageStorage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportOneDriveUsageStorage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\OneDriveUsageStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 055

#gavdcodebegin 056
function PsCliReporting_OneDriveActivityFileCounts
{
    m365 onedrive report activityfilecounts --period "D90"
}
#gavdcodeend 056

#gavdcodebegin 057
function PsCliReporting_OneDriveActivityFileCountsToFile
{
    m365 onedrive report activityfilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityFileCounts.csv"
}
#gavdcodeend 057

#gavdcodebegin 058
function PsCliReporting_OneDriveActivityUserCountsToFile
{
    m365 onedrive report activityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserCounts.csv"
}
#gavdcodeend 058

#gavdcodebegin 059
function PsCliReporting_OneDriveActivityUserDetailToFile
{
    m365 onedrive report activityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserDetail.csv"
}
#gavdcodeend 059

#gavdcodebegin 060
function PsCliReporting_OneDriveActivityUserDetailByDateToFile
{
    m365 onedrive report activityuserdetail --date "2023-12-16" `
                    --output "csv" > "C:\Temporary\OneDriveActivityUserDetailByDate.csv"
}
#gavdcodeend 060

#gavdcodebegin 061
function PsCliReporting_OneDriveUsageAccountCountsToFile
{
    m365 onedrive report usageaccountcounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageAccountCounts.csv"
}
#gavdcodeend 061

#gavdcodebegin 062
function PsCliReporting_OneDriveUsageAccountDetailToFile
{
    m365 onedrive report usageaccountdetail --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageAccountDetail.csv"
}
#gavdcodeend 062

#gavdcodebegin 063
function PsCliReporting_OneDriveUsageFileCountsToFile
{
    m365 onedrive report usagefilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageFileCounts.csv"
}
#gavdcodeend 063

#gavdcodebegin 064
function PsCliReporting_OneDriveUsageStorageToFile
{
    m365 onedrive report usagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\OneDriveUsageStorage.csv"
}
#gavdcodeend 064

#gavdcodebegin 065
Function PsGraphSdkReporting_SharePointActivityFileCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointActivityFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 065

#gavdcodebegin 066
Function PsGraphSdkReporting_SharePointActivityPage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointActivityPage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityPage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 066

#gavdcodebegin 067
Function PsGraphSdkReporting_SharePointActivityUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 067

#gavdcodebegin 068
Function PsGraphSdkReporting_SharePointActivityUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 068

#gavdcodebegin 069
Function PsGraphSdkReporting_SharePointSiteUsageDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointSiteUsageDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 069

#gavdcodebegin 070
Function PsGraphSdkReporting_SharePointSiteUsageFileCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointSiteUsageFileCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageFileCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 070

#gavdcodebegin 071
Function PsGraphSdkReporting_SharePointSiteUsagePage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointSiteUsagePage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsagePage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 071

#gavdcodebegin 072
Function PsGraphSdkReporting_SharePointSiteUsageSiteCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointSiteUsageSiteCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageSiteCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 072

#gavdcodebegin 073
Function PsGraphSdkReporting_SharePointSiteUsageStorage
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportSharePointSiteUsageStorage `
							   -Period "D90" `
							   -OutFile "C:\Temporary\SharePointSiteUsageStorage.csv"

	Disconnect-MgGraph
}
#gavdcodeend 073

#gavdcodebegin 074
function PsCliReporting_SharePointActivityFileCount
{
    m365 spo report activityfilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityFileCount.csv"
}
#gavdcodeend 074

#gavdcodebegin 075
function PsCliReporting_SharePointActivityPages
{
    m365 spo report activitypages --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityPages.csv"
}
#gavdcodeend 075

#gavdcodebegin 076
function PsCliReporting_SharePointActivityUserCounts
{
    m365 spo report activityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityUserCounts.csv"
}
#gavdcodeend 076

#gavdcodebegin 077
function PsCliReporting_SharePointActivityUserDetail
{
    m365 spo report activityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointActivityUserDetail.csv"
}
#gavdcodeend 077

#gavdcodebegin 078
function PsCliReporting_SharePointSiteUsageDetail
{
    m365 spo report siteusagedetail --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteusageDetail.csv"
}
#gavdcodeend 078

#gavdcodebegin 079
function PsCliReporting_SharePointSiteUsageFileCounts
{
    m365 spo report siteusagefilecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageFileCounts.csv"
}
#gavdcodeend 079

#gavdcodebegin 080
function PsCliReporting_SharePointSiteSiteUsagePages
{
    m365 spo report siteusagepages --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsagePages.csv"
}
#gavdcodeend 080

#gavdcodebegin 081
function PsCliReporting_SharePointSiteUsageSiteCounts
{
    m365 spo report siteusagesitecounts --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageSiteCounts.csv"
}
#gavdcodeend 081

#gavdcodebegin 082
function PsCliReporting_SharePointSiteUsageStorage
{
    m365 spo report siteusagestorage --period "D90" `
                    --output "csv" > "C:\Temporary\SharePointSiteUsageStorage.csv"
}
#gavdcodeend 082

#gavdcodebegin 083
Function PsGraphSdkReporting_TeamActivityCount   #==> Does not exist
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamActivityCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 083

#gavdcodebegin 084
Function PsGraphSdkReporting_TeamActivityDetail   #==> Does not exist
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamActivityDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 084

#gavdcodebegin 085
Function PsGraphSdkReporting_TeamActivityDistributionCount   #==> Does not exist
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamActivityDistributionCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamActivityDistributionCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 085

#gavdcodebegin 086
Function PsGraphSdkReporting_TeamCount   #==> Does not exist
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 086

#gavdcodebegin 087
Function PsGraphSdkReporting_TeamUserActivityCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamUserActivityCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 087

#gavdcodebegin 088
Function PsGraphSdkReporting_TeamUserActivityUserCount
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamUserActivityUserCount `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityUserCount.csv"

	Disconnect-MgGraph
}
#gavdcodeend 088

#gavdcodebegin 089
Function PsGraphSdkReporting_TeamUserActivityUserDetail
{
	GrPsLoginGraphSDKWithAccPw -TenantName $configFile.appsettings.TenantName `
							   -ClientID $configFile.appsettings.ClientIdWithAccPw `
							   -UserName $configFile.appsettings.UserName `
							   -UserPw $configFile.appsettings.UserPw

	Get-MgReportTeamUserActivityUserDetail `
							   -Period "D90" `
							   -OutFile "C:\Temporary\TeamUserActivityUserDetail.csv"

	Disconnect-MgGraph
}
#gavdcodeend 089

#gavdcodebegin 090
function PsCliReporting_TeamDirectRoutingCalls #==> Needs CallRecords.Read.All permission
{
    m365 teams report directroutingcalls --debug `
					--fromDateTime 2023-12-01 --toDateTime 2023-12-17 `
                    --output "csv" > "C:\Temporary\DirectRoutingCalls.csv"
}
#gavdcodeend 090

#gavdcodebegin 091
function PsCliReporting_TeamPstnCalls #==> Needs CallRecords.Read.All permission
{
    m365 teams report pstncalls `
					--fromDateTime 2023-12-01 --toDateTime 2023-12-17 `
                    --output "csv" > "C:\Temporary\PstnCalls.csv"
}
#gavdcodeend 091

#gavdcodebegin 092
function PsCliReporting_TeamUserActivityCounts
{
    m365 teams report useractivitycounts --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityCounts.csv"
}
#gavdcodeend 092

#gavdcodebegin 093
function PsCliReporting_TeamUserActivityUserCounts
{
    m365 teams report useractivityusercounts --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityUserCounts.csv"
}
#gavdcodeend 093

#gavdcodebegin 094
function PsCliReporting_TeamUserActivityUserDetail
{
    m365 teams report useractivityuserdetail --period "D90" `
                    --output "csv" > "C:\Temporary\UserActivityUserDetail.csv"
}
#gavdcodeend 094

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 094 ***

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
#PsGraphSdkReporting_OneDriveActivityFileCount
#PsGraphSdkReporting_OneDriveActivityUserCount
#PsGraphSdkReporting_OneDriveActivityUserDetail
#PsGraphSdkReporting_OneDriveUsageAccountCount
#PsGraphSdkReporting_OneDriveUsageAccountDetail
#PsGraphSdkReporting_OneDriveUsageFileCount
#PsGraphSdkReporting_OneDriveUsageStorage
#PsCliReporting_OneDriveActivityFileCounts
#PsCliReporting_OneDriveActivityFileCountsToFile
#PsCliReporting_OneDriveActivityUserCountsToFile
#PsCliReporting_OneDriveActivityUserDetailToFile
#PsCliReporting_OneDriveActivityUserDetailByDateToFile
#PsCliReporting_OneDriveUsageAccountCountsToFile
#PsCliReporting_OneDriveUsageAccountDetailToFile
#PsCliReporting_OneDriveUsageFileCountsToFile
#PsCliReporting_OneDriveUsageStorageToFile
#PsGraphSdkReporting_SharePointActivityFileCount
#PsGraphSdkReporting_SharePointActivityPage
#PsGraphSdkReporting_SharePointActivityUserCount
#PsGraphSdkReporting_SharePointActivityUserDetail
#PsGraphSdkReporting_SharePointSiteUsageDetail
#PsGraphSdkReporting_SharePointSiteUsageFileCount
#PsGraphSdkReporting_SharePointSiteUsagePage
#PsGraphSdkReporting_SharePointSiteUsageSiteCount
#PsGraphSdkReporting_SharePointSiteUsageStorage
#PsCliReporting_SharePointActivityFileCount
#PsCliReporting_SharePointActivityPages
#PsCliReporting_SharePointActivityUserCounts
#PsCliReporting_SharePointActivityUserDetail
#PsCliReporting_SharePointSiteUsageDetail
#PsCliReporting_SharePointSiteUsageFileCounts
#PsCliReporting_SharePointSiteSiteUsagePages
#PsCliReporting_SharePointSiteUsageSiteCounts
#PsCliReporting_SharePointSiteUsageStorage
#PsGraphSdkReporting_TeamActivityCount
#PsGraphSdkReporting_TeamActivityDetail
#PsGraphSdkReporting_TeamActivityDistributionCount
#PsGraphSdkReporting_TeamCount
#PsGraphSdkReporting_TeamUserActivityCount
#PsGraphSdkReporting_TeamUserActivityUserCount
#PsGraphSdkReporting_TeamUserActivityUserDetail
#PsCliReporting_TeamDirectRoutingCalls
#PsCliReporting_TeamPstnCalls
#PsCliReporting_TeamUserActivityCounts
#PsCliReporting_TeamUserActivityUserCounts
#PsCliReporting_TeamUserActivityUserDetail

m365 logout
Write-Host "Done" 



