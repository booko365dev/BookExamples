##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

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



##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 003 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

# Connect to M365 using the CLI
#$spCtx = LoginPsCLI

#JsonAdaptiveCard_01
#PsAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhook
#PsCliAdaptiveCards_SendCardToWebhookWithJson

#m365 logout
Write-Host "Done" 



