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


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 017 ***

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

m365 logout
Write-Host "Done" 



