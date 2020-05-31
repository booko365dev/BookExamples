
#gavdcodebegin 01
Function ConnectPsExoV2()
{
    Connect-ExchangeOnline -UserPrincipalName $configFile.appsettings.exUserName
}
#gavdcodeend 01

#-----------------------------------------------------------------------------------------

#gavdcodebegin 02
Function ExPsExoV2GetMailboxesByPropertySet()
{
    Get-EXOMailbox -PropertySets Minimum,Policy
}
#gavdcodeend 02

#gavdcodebegin 03
Function ExPsExoV2GetMailboxesByProperty()
{
    Get-EXOMailbox -Properties UserPrincipalName,Alias
}
#gavdcodeend 03

#gavdcodebegin 04
Function ExPsExoV2GetMailboxesBySetAndProperty()
{
    Get-EXOMailbox -PropertySets Hold,Moderation -Properties UserPrincipalName,Alias
}
#gavdcodeend 04

#gavdcodebegin 05
Function ExPsExoV2GetClientAccessSettings()
{
    Get-EXOCASMailbox -Identity $configFile.appsettings.exUserName
}
#gavdcodeend 05

#gavdcodebegin 06
Function ExPsExoV2GetEmailboxPermissions()
{
    Get-EXOMailboxPermission -Identity $configFile.appsettings.exUserName
}
#gavdcodeend 06

#gavdcodebegin 07
Function ExPsExoV2GetEmailboxStatistics()
{
    Get-EXOMailboxStatistics -Identity $configFile.appsettings.exUserName
}
#gavdcodeend 07

#gavdcodebegin 08
Function ExPsExoV2GetFolderPermissions()
{
    $myIdentifier = $configFile.appsettings.exUserName + ":\Inbox"
    Get-EXOMailboxFolderPermission -Identity $myIdentifier
}
#gavdcodeend 08

#gavdcodebegin 09
Function ExPsExoV2GetFolderStatisticsAll()
{
    Get-EXOMailboxFolderStatistics -Identity $configFile.appsettings.exUserName
}
#gavdcodeend 09

#gavdcodebegin 10
Function ExPsExoV2GetFolderStatisticsOne()
{
    Get-EXOMailboxFolderStatistics -Identity $configFile.appsettings.exUserName `
                                                            -FolderScope Calendar
}
#gavdcodeend 10

#gavdcodebegin 11
Function ExPsExoV2GetDeviceStatistics()
{
    Get-EXOMobileDeviceStatistics -Mailbox $configFile.appsettings.exUserName
}
#gavdcodeend 11

#gavdcodebegin 12
Function ExPsExoV2GetRecipients()
{
    Get-EXORecipient -Identity $configFile.appsettings.exUserName
}
#gavdcodeend 12

#gavdcodebegin 13
Function ExPsExoV2GetRecipientPermissions()
{
    Get-EXORecipientPermission -ResultSize 5
}
#gavdcodeend 13

#-----------------------------------------------------------------------------------------

[xml]$configFile = get-content "C:\Projects\exPs.values.config"

ConnectPsExoV2

#ExPsExoV2GetMailboxesByPropertySet
#ExPsExoV2GetMailboxesByProperty
#ExPsExoV2GetMailboxesBySetAndProperty
#ExPsExoV2GetClientAccessSettings
#ExPsExoV2GetEmailboxPermissions
#ExPsExoV2GetEmailboxStatistics
#ExPsExoV2GetFolderPermissions
#ExPsExoV2GetFolderStatisticsAll
#ExPsExoV2GetFolderStatisticsOne
#ExPsExoV2GetDeviceStatistics
#ExPsExoV2GetRecipients
#ExPsExoV2GetRecipientPermissions

Write-Host "Done"  

