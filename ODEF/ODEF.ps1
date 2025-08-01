##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------


function PsEws_Connect  #*** LEGACY CODE ***
{
	$ExService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
	$ExService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(`
		$cnfUserName, $cnfUserPw)
	$ExService.Url = new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx");
	#$ExService.TraceEnabled = $true
	#$ExService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
	$ExService.AutodiscoverUrl($cnfUserName, {$true})

	return $ExService
}

function PsCliM365_LoginWithAccPw
{
	m365 login --authType password `
			   --appId $cnfClientIdWithAccPw `
			   --userName $cnfUserName `
			   --password $cnfUserPw
}

function PsGraphPnP_LoginWithAccPwAndClientId
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantUrl,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientId,
 
		[Parameter(Mandatory=$True)]
		[String]$UserName,
 
		[Parameter(Mandatory=$True)]
		[String]$UserPw
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$UserPw -AsPlainText -Force
	$myCredentials = New-Object System.Management.Automation.PSCredential `
								-argumentlist $UserName, $securePW

	Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Credentials $myCredentials
}

function PsGraphPowerShellSdk_LoginWithAccPw
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

function PsGraphPowerShellSdk_LoginWithSecret
{
	Param(
		[Parameter(Mandatory=$True)]
		[String]$TenantName,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientID,
 
		[Parameter(Mandatory=$True)]
		[String]$ClientSecret
	)

	[SecureString]$securePW = ConvertTo-SecureString -String `
									$ClientSecret -AsPlainText -Force
	$myCredentials = New-Object -TypeName System.Management.Automation.PSCredential `
							-argumentlist $ClientID, $securePW

	Connect-MgGraph -TenantId $TenantName -ClientSecretCredential $myCredentials
}

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

##==> Routines for EWS

#gavdcodebegin 001
function PsExchangeEws_CreateOneFolder($ExService)  #*** LEGACY CODE ***
{
    $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($ExService)
    $newFolder.DisplayName = "My Custom Folder PowerShell"

    $newFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}
#gavdcodeend 001

#gavdcodebegin 002
function PsExchangeEws_GetAllFolders($ExService)  #*** LEGACY CODE ***
{
    $myView = [Microsoft.Exchange.WebServices.Data.FolderView]100
    $isHidden = `
			New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(`
				[System.Int32]0x10f4, `
				[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
    $myView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
				[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, `
				[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, `
				$isHidden)
    $myView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $allFolders = $ExService.FindFolders(`
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, `
				$myView)

    foreach ($oneFolder in $allFolders) {
        $strHidden = $oneFolder.ExtendedProperties[0].Value
        Write-Host $oneFolder.DisplayName " - Hidden: " $strHidden
    }
}
#gavdcodeend 002

#gavdcodebegin 003
function PsExchangeEws_FindOneFolder($ExService)  #*** LEGACY CODE ***
{
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
			$ExService, `
			[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $rootFolder.Load()
    $subjectFilter = `
			New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(`
					[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, `
					"my custom folder powershell", `
					[Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, `
					[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)

    $myFolderId = $null
    $myView = [Microsoft.Exchange.WebServices.Data.FolderView]1
    foreach ($oneFolder in $rootFolder.FindFolders($subjectFilter, $myView))
    {
        $myFolderId = $oneFolder.Id
		Write-Host $myFolderId
    }
}
#gavdcodeend 003

#gavdcodebegin 004
function PsExchangeEws_UpdateOneFolder($ExService)  #*** LEGACY CODE ***
{
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
			$ExService, `
			[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $rootFolder.Load()
    $subjectFilter = `
			New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(`
					[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, `
					"my custom folder powershell", `
					[Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, `
					[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)

    $myFolderId = $null
    $myView = [Microsoft.Exchange.WebServices.Data.FolderView]1
    foreach ($oneFolder in $rootFolder.FindFolders($subjectFilter, $myView))
    {
        $myFolderId = $oneFolder.Id
    }

    $folderToUpdate = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
			$ExService, $myFolderId)
    $folderToUpdate.DisplayName = "New Folder Name"
    $folderToUpdate.Update()
}
#gavdcodeend 004

#gavdcodebegin 005
function PsExchangeEws_DeleteOneFolder($ExService)  #*** LEGACY CODE ***
{
    $rootFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
			$ExService, `
			[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
    $rootFolder.Load()
    $subjectFilter = `
			New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(`
					[Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, `
					"New Folder Name", `
					[Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring, `
					[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase)

    $myFolderId = $null
    $myView = [Microsoft.Exchange.WebServices.Data.FolderView]1
    foreach ($oneFolder in $rootFolder.FindFolders($subjectFilter, $myView))
    {
        $myFolderId = $oneFolder.Id
    }

    $folderToDelete = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(`
			$ExService, $myFolderId);
    $folderToDelete.Delete(`
			[Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete);
}
#gavdcodeend 005

#gavdcodebegin 006
function PsExchangeEws_CreateAndSendEmail($ExService)  #*** LEGACY CODE ***
{
    $newEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($ExService)
    $newEmail.Subject = "Email send by EWS PowerShell"
    $newEmail.Body = "To Whom It May Concern"
    $newEmail.ToRecipients.Add("user01@domain.com")
    $newEmail.BccRecipients.Add("user02@domain.com")
    $newEmail.CcRecipients.Add("user03@domain.com")
    #$newEmail.From = "user04@domain.com"
    $newEmail.Importance = [Microsoft.Exchange.WebServices.Data.Importance]::Normal
            
    $newEmail.SendAndSaveCopy();
    #$newEmail.Send();
}
#gavdcodeend 006

#gavdcodebegin 007
function PsExchangeEws_GetUnreadEmails($ExService)  #*** LEGACY CODE ***
{
	$mySearchFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
					[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead, `
					$false)
    $myFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(`
					[Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, `
					$mySearchFilter)
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
    $findResults = $ExService.FindItems(
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, `
				$myFilter, $myView)

    Write-Host $findResults.TotalCount
}
#gavdcodeend 007

#gavdcodebegin 008
function PsExchangeEws_ReplyToEmail($ExService)  #*** LEGACY CODE ***
{
	$mySearchFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
					[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, `
					"Email asking for replay")
    $myFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(`
					[Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, `
					$mySearchFilter)
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
    $findResults = $ExService.FindItems(
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, `
				$myFilter, $myView)

    $myEmailId = $null
    foreach ($oneItem in $findResults) {
        $myEmailId = $oneItem.Id
    }

    $emailToReply = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind(`
				$ExService, $myEmailId)

    $myReply = "Reply body"
    $emailToReply.Reply($myReply, $true);
}
#gavdcodeend 008

#gavdcodebegin 009
function PsExchangeEws_DeleteOneEmail($ExService)  #*** LEGACY CODE ***
{
	$mySearchFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
					[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, `
					"Email asking for replay")
    $myFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(`
					[Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, `
					$mySearchFilter)
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
    $findResults = $ExService.FindItems(
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, `
				$myFilter, $myView)

    $myEmailId = $null
    foreach ($oneItem in $findResults) {
        $myEmailId = $oneItem.Id
    }

    $myPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
				[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, `
		        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, `
				[Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ParentFolderId)
    $emailToDelete = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind(`
				$ExService, $myEmailId, $myPropSet)

    $emailToDelete.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete)
}
#gavdcodeend 009

#gavdcodebegin 010
function PsExchangeEws_CreateOneContact($ExService)  #*** LEGACY CODE ***
{
    $newContact = New-Object Microsoft.Exchange.WebServices.Data.Contact($ExService)
    $newContact.GivenName = "Somename"
    $newContact.MiddleName = "Mymiddle"
    $newContact.Surname = "Hersurname"
    $newContact.FileAsMapping = `
		[Microsoft.Exchange.WebServices.Data.FileAsMapping]::SurnameCommaGivenName
    $newContact.CompanyName = "Mycompany"
    $newContact.PhoneNumbers[`
		[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = "1234567890";
    $newContact.PhoneNumbers[`
		[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = "0987654321";
    $newContact.PhoneNumbers[`
		[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::CarPhone] = "1029384756";
    $newContact.EmailAddresses[`
		[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = `
		New-Object Microsoft.Exchange.WebServices.Data.EmailAddress("somename@domain.com");
    $newContact.EmailAddresses[`
		[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = `
		New-Object Microsoft.Exchange.WebServices.Data.EmailAddress(`
			"somename.mymiddle@domain.com");
    $newContact.ImAddresses[`
		[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = `
	"somenameIM1@domain.com";
    $newContact.ImAddresses[`
		[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress2] = `
			"somenameIM2@domain.com";

    $paHome = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
    $paHome.Street = "123 Somewhere Street"
    $paHome.City = "Here"
    $paHome.State = "AZ"
    $paHome.PostalCode = "92835"
    $paHome.CountryOrRegion = "Europe"

    $newContact.PhysicalAddresses[`
		[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Home] = $paHome

    $paBusiness = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
    $paBusiness.Street = "456 Somewhere Avenue"
    $paBusiness.City = "There"
    $paBusiness.State = "ZA"
    $paBusiness.PostalCode = "53829"
    $paBusiness.CountryOrRegion = "Europe"

    $newContact.PhysicalAddresses[`
		[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = $paBusiness

    $newContact.Save()
}
#gavdcodeend 010

#gavdcodebegin 011
function PsExchangeEws_FindAllContacts($ExService)  #*** LEGACY CODE ***
{
    $myContactsfolder = [Microsoft.Exchange.WebServices.Data.ContactsFolder]::Bind(`
		$ExService, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts)

    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]50
    $myView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
		[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, `
		[Microsoft.Exchange.WebServices.Data.ContactSchema]::DisplayName)
    $allContacts = $ExService.FindItems(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, $myView)

    foreach ($oneContact in $allContacts) {
        Write-Host $oneContact.DisplayName
    }
}
#gavdcodeend 011

#gavdcodebegin 012
function PsExchangeEws_FindOneContactByName($ExService)  #*** LEGACY CODE ***
{
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
	$myFilter = `
		New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
				[Microsoft.Exchange.WebServices.Data.ContactSchema]::GivenName, `
				"Somename")
    $allFound = $ExService.FindItems(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, `
		$myFilter, $myView)

    foreach ($oneFound in $allFound) {
        Write-Host $oneFound.CompleteName.FullName + " - " + $onefound.CompanyName
    }
}
#gavdcodeend 012

#gavdcodebegin 013
function PsExchangeEws_UpdateOneContact($ExService)  #*** LEGACY CODE ***
{
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
	$myFilter = `
		New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
				[Microsoft.Exchange.WebServices.Data.ContactSchema]::GivenName, `
				"Somename")
    $allFound = $ExService.FindItems(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, `
		$myFilter, $myView)

    $myContactId = $null
    foreach ($oneFound in $allFound) {
        $myContactId = $oneFound.Id
    }

    $myContact = `
		[Microsoft.Exchange.WebServices.Data.Contact]::Bind($ExService, $myContactId)

    $myContact.Surname = "Hissurname"
    $myContact.CompanyName = "Hiscompany"
    $myContact.PhoneNumbers[`
		[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = `
			"32143254321"
    $myContact.EmailAddresses[`
		[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress2] = `
		New-Object Microsoft.Exchange.WebServices.Data.EmailAddress("somebody@domain.com")
    $myContact.ImAddresses[`
		[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = `
			"otherM1@domain.com"

    $paBusiness = New-Object Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
    $paBusiness.Street = "987 Somewhere Way"
    $paBusiness.City = "Noidea"
    $paBusiness.State = "ZZ"
    $paBusiness.PostalCode = "66666"
    $paBusiness.CountryOrRegion = "Europe"

    $myContact.PhysicalAddresses[`
		[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = $paBusiness

    $myContact.Update(`
		[Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
}
#gavdcodeend 013

#gavdcodebegin 014
function PsExchangeEws_DeleteOneContact($ExService)  #*** LEGACY CODE ***
{
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
	$myFilter = `
		New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
				[Microsoft.Exchange.WebServices.Data.ContactSchema]::GivenName, `
				"Somename")
    $allFound = $ExService.FindItems(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts, `
		$myFilter, $myView)

    $myContactId = $null
    foreach ($oneFound in $allFound) {
        $myContactId = $oneFound.Id
    }

    $myContact = `
		[Microsoft.Exchange.WebServices.Data.Contact]::Bind($ExService, $myContactId)

    $myContact.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems);
}
#gavdcodeend 014

#gavdcodebegin 015
function PsExchangeEws_CreateAppointment($ExService)  #*** LEGACY CODE ***
{
    $myDt = (Get-Date).AddDays(+1)
    $newAppointment = New-Object `
			Microsoft.Exchange.WebServices.Data.Appointment($ExService)
    $newAppointment.Subject = "This is a new meeting from PowerShell"
    $newAppointment.Body = "To Whom It May Concern"
    $newAppointment.Start = New-Object System.DateTime($myDt.Year, $myDt.Month, `
									$myDt.Day, $myDt.Hour, $myDt.Minute, $myDt.Second)
    $newAppointment.Location = "Somewhere"
    $newAppointment.End = $newAppointment.Start.AddHours(1)
    $newAppointment.RequiredAttendees.Add("user1@domain.com")
    $newAppointment.OptionalAttendees.Add("user2@domain.com")

    $newAppointment.Save(`
		[Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToNone)
}
#gavdcodeend 015

#gavdcodebegin 016
function PsExchangeEws_FindAppointmentsByDate($ExService)  #*** LEGACY CODE ***
{
	$myView = New-Object `
		Microsoft.Exchange.WebServices.Data.CalendarView((Get-Date), `
			(Get-Date).AddDays(7))
    $allAppointments = $ExService.FindAppointments(`
		[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, $myView)

    foreach ($oneAppointment in $allAppointments) {
        Write-Host "Subject: " $oneAppointment.Subject
        Write-Host "Start: " $oneAppointment.Start
        Write-Host "Duration: " $oneAppointment.Duration
    }
}
#gavdcodeend 016

#gavdcodebegin 017
function PsExchangeEws_UpdateOneAppointment($ExService)  #*** LEGACY CODE ***
{
	$mySearchFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
					[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, `
					"This is a new meeting from PowerShell")
    $myFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(`
					[Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, `
					$mySearchFilter)
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
    $findResults = $ExService.FindItems(`
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, `
				$myFilter, $myView)

    $myAppointmentId = $null
    foreach ($oneItem in $findResults) {
        $myAppointmentId = $oneItem.Id
    }

    $myPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
				[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, `
				[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, `
				[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location)
    $myAppointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind(`
				$ExService, $myAppointmentId, $myPropSet)

    $myAppointment.Location = "Other place"
    $myAppointment.RequiredAttendees.Add("user2@domain.com")

    $myAppointment.Update(`
		[Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, `
		[Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::`
				SendToAllAndSaveCopy)
}
#gavdcodeend 017

#gavdcodebegin 018
function PsExchangeEws_DeleteOneAppointment($ExService)  #*** LEGACY CODE ***
{
	$mySearchFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(`
					[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, `
					"This is a new meeting from PowerShell")
    $myFilter = New-Object `
				Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(`
					[Microsoft.Exchange.WebServices.Data.LogicalOperator]::And, `
					$mySearchFilter)
    $myView = [Microsoft.Exchange.WebServices.Data.ItemView]1
    $findResults = $ExService.FindItems(`
				[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar, `
				$myFilter, $myView)

    $myAppointmentId = $null
    foreach ($oneItem in $findResults) {
        $myAppointmentId = $oneItem.Id
    }

    $myPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(`
				[Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, `
				[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Subject, `
				[Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Location)
    $myAppointment = [Microsoft.Exchange.WebServices.Data.Appointment]::Bind(`
				$ExService, $myAppointmentId, $myPropSet)

    # Using the Delete method (use only one of the code lines)
    $myAppointment.Delete(`
		[Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
    $myAppointment.Delete(`
		[Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems, `
		[Microsoft.Exchange.WebServices.Data.SendCancellationsMode]::SendOnlyToAll)

    # Using the Cancel method (use only one of the code lines)
    $myAppointment.CancelMeeting()
    $myAppointment.CancelMeeting("Meeting canceled")

    $cancelMessage = $myAppointment.CreateCancelMeetingMessage()
    $cancelMessage.Body = New-Object `
		Microsoft.Exchange.WebServices.Data.MessageBody("Meeting canceled")
    $cancelMessage.IsReadReceiptRequested = $true
    $cancelMessage.SendAndSaveCopy()
}
#gavdcodeend 018

#-----------------------------------------------------------------------------------------

##==> Routines for CLI

#gavdcodebegin 019
function PsExchangeCli_GetAllMessages
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook message list --folderName "Inbox"

	m365 logout
}
#gavdcodeend 019

#gavdcodebegin 020
function PsExchangeCli_GetMessageById
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook message list `
            --folderName "Inbox" `
            --query "[?id == 'AAMkAGE0ODQ3NTc1L...MAvtBLB-F9SJ2ZDb7Xo-OrAAByF4KBAAA=']"

	m365 logout
}
#gavdcodeend 020

#gavdcodebegin 034
function PsExchangeCli_GetOneMessageById
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook message get `
            --id "AAMkAGE0ODQ3NTc1L...MAvtBLB-F9SJ2ZDb7Xo-OrAAByF4KBAAA="

	m365 logout
}
#gavdcodeend 034

#gavdcodebegin 021
function PsExchangeCli_MoveMessage
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook message move `
            --sourceFolderName "Inbox" `
            --targetFolderName "MyFolder" `
            --messageId "AAMkAGE0ODQ3NTc1L...MAvtBLB-F9SJ2ZDb7Xo-OrAAByF4KBAAA="

	m365 logout
}
#gavdcodeend 021

#gavdcodebegin 022
function PsExchangeCli_SendMessage
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook mail send `
            --to "user@domain.OnMicrosoft.com" `
            --subject "Test Email from CLI" `
            --bodyContents "This is a <b>test</b> message" `
            --bodyContentType HTML

	m365 logout
}
#gavdcodeend 022

#gavdcodebegin 023
function PsExchangeCli_ReportActivity
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailactivitycounts --period D7 --output json

	m365 logout
}
#gavdcodeend 023

#gavdcodebegin 024
function PsExchangeCli_ReportActivityUser
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailactivityusercounts --period D7 --output json

	m365 logout
}
#gavdcodeend 024

#gavdcodebegin 025
function PsExchangeCli_ReportActivityUserDetails
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailactivityuserdetail --period D7 --output json

	m365 logout
}
#gavdcodeend 025

#gavdcodebegin 026
function PsExchangeCli_ReportActivityByAppTotals
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailappusageappsusercounts --period D7 --output json

	m365 logout
}
#gavdcodeend 026

#gavdcodebegin 027
function PsExchangeCli_ReportActivityByApp
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailappusageusercounts --period D7 --output json

	m365 logout
}
#gavdcodeend 027

#gavdcodebegin 028
function PsExchangeCli_ReportActivityByAppAndUserDetails
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailappusageuserdetail --period D7 --output json

	m365 logout
}
#gavdcodeend 028

#gavdcodebegin 029
function PsExchangeCli_ReportActivityByAppVersions
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailappusageversionsusercounts --period D7 --output json

	m365 logout
}
#gavdcodeend 029

#gavdcodebegin 030
function PsExchangeCli_ReportUsageDetail
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailboxusagedetail --period D7 --output json

	m365 logout
}
#gavdcodeend 030

#gavdcodebegin 031
function PsExchangeCli_ReportUsageMailboxes
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailboxusagemailboxcount --period D7 --output json

	m365 logout
}
#gavdcodeend 031

#gavdcodebegin 032
function PsExchangeCli_ReportQuotaMailboxes
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailboxusagequotastatusmailboxcounts --period D7 --output json

	m365 logout
}
#gavdcodeend 032

#gavdcodebegin 033
function PsExchangeCli_ReportStorageMailboxes
{
	PsCliM365_LoginWithAccPw
	
	m365 outlook report mailboxusagestorage --period D7 --output json

	m365 logout
}
#gavdcodeend 033

#-----------------------------------------------------------------------------------------

##==> Routines for PnP PowerShell

#gavdcodebegin 034
function PsExchangePnP_SendEmailWithFrom
{
	PsGraphPnP_LoginWithAccPwAndClientId `
					-TenantUrl $cnfSiteBaseUrl `
					-ClientId $cnfClientIdWithAccPw `
					-UserName $cnfUserName `
					-UserPw $cnfUserPw
	
	Send-PnPMail -From "user@domain.onmicrosoft.com" `
				 -To "user@domain.com" `
				 -Subject "Test message from me" `
				 -Body "This is a test message using PnP from me"

	Disconnect-PnPOnline
}
#gavdcodeend 034

#gavdcodebegin 035
function PsExchangePnP_SendEmailWithoutFrom
{
	PsGraphPnP_LoginWithAccPwAndClientId `
					-TenantUrl $cnfSiteBaseUrl `
					-ClientId $cnfClientIdWithAccPw `
					-UserName $cnfUserName `
					-UserPw $cnfUserPw
	
	Send-PnPMail -To "user@domain.com" `
				 -Subject "Test message from MS" `
				 -Body "This is a test message using PnP from Microsoft"

	Disconnect-PnPOnline
}
#gavdcodeend 035

#gavdcodebegin 036
function PsExchangePnP_SendEmailFromSmtpMs
{
	PsGraphPnP_LoginWithAccPwAndClientId `
					-TenantUrl $cnfSiteBaseUrl `
					-ClientId $cnfClientIdWithAccPw `
					-UserName $cnfUserName `
					-UserPw $cnfUserPw
	
	Send-PnPMail -From "user@domain.onmicrosoft.com" `
				 -To "user@domain.onmicrosoft.com" `
				 -Subject "Test message from SMTP Microsoft" `
				 -Body "This is a test message using PnP and SMTP Microsoft" `
				 -Server "domain.mail.protection.outlook.com"

	Disconnect-PnPOnline
}
#gavdcodeend 036

#gavdcodebegin 037
function PsExchangePnP_SendEmailFromSmtpAll
{
	PsGraphPnP_LoginWithAccPwAndClientId `
					-TenantUrl $cnfSiteBaseUrl `
					-ClientId $cnfClientIdWithAccPw `
					-UserName $cnfUserName `
					-UserPw $cnfUserPw
	
	Send-PnPMail -From "user@domain.onmicrosoft.com" `
				 -To "user@domain.onmicrosoft.com" `
				 -Subject "Test message from SMTP" `
				 -Body "This is a test message using PnP and SMTP" `
				 -Server "any.server.isp.com" `
				 -Port 587 `
				 -EnableSsl:$true `
				 -Username "user-smtp" `
				 -Password "password-smtp"

	Disconnect-PnPOnline
}
#gavdcodeend 037

#-----------------------------------------------------------------------------------------

##==> Routines for Graph SDK PowerShell

#gavdcodebegin 038
function PsExchangeGraphSdk_GetMessages
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `
	
	Get-MgUserMessage -UserId $cnfUserName -All

	Disconnect-MgGraph
}
#gavdcodeend 038

#gavdcodebegin 039
function PsExchangeGraphSdk_GetOneMessageById
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Get-MgUserMessage -UserId $cnfUserName `
					  -MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA="

	Disconnect-MgGraph
}
#gavdcodeend 039

#gavdcodebegin 040
function PsExchangeGraphSdk_GetMessageContent
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Get-MgUserMessageContent -UserId $cnfUserName `
							 -MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA=" `
							 -OutFile "C:\Temporary\myEmail.txt"

	Disconnect-MgGraph
}
#gavdcodeend 040

#gavdcodebegin 041
function PsExchangeGraphSdk_GetOneMessageAttachments
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Get-MgUserMessageAttachment -UserId $cnfUserName `
								-MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA="

	Disconnect-MgGraph
}
#gavdcodeend 041

#gavdcodebegin 042
function PsExchangeGraphSdk_GetOneMessageOneAttachment
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Get-MgUserMessageAttachment -UserId $cnfUserName `
								-MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA=" `
								-AttachmentId "slkjfs...lkjfs"

	Disconnect-MgGraph
}
#gavdcodeend 042

#gavdcodebegin 043
function PsExchangeGraphSdk_CreateMessageDraft
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myParams = @{
		Subject = "Test Email Graph SDK"
		Importance = "Low"
		Body = @{
			ContentType = "HTML"
			Content = "Test Email sent by <b>Graph SDK</b>"
		}
		ToRecipients = @(
			@{
				EmailAddress = @{
					Address = "user@domain.onmicrosoft.com"
				}
			}
		)
	}

	New-MgUserMessage -UserId $cnfUserName @myParams

	Disconnect-MgGraph
}
#gavdcodeend 043

#gavdcodebegin 044
function PsExchangeGraphSdk_UpdateMessageContent
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Set-MgUserMessageContent -UserId $cnfUserName `
							 -MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA=" `
							 -InFile "C:\Temporary\myEmail.txt"

	Disconnect-MgGraph
}
#gavdcodeend 044

#gavdcodebegin 045
function PsExchangeGraphSdk_UpdateMessage
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myParams = @{
		Subject = "Test Email Graph SDK updated"
		Importance = "Low"
		Body = @{
			ContentType = "HTML"
			Content = "Test Email sent by <b>Graph SDK</b> updated"
		}
		ToRecipients = @(
			@{
				EmailAddress = @{
					Address = "user@domain.onmicrosoft.com"
				}
			}
		)
	}

	Update-MgUserMessage -UserId $cnfUserName `
						 -MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA=" `
						 @myParams

	Disconnect-MgGraph
}
#gavdcodeend 045

#gavdcodebegin 046
function PsExchangeGraphSdk_DeleteMessage
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Remove-MgUserMessage -UserId $cnfUserName `
						 -MessageId "AAMkAGE0OD...7Xo-OrAABoXpDqAAA="

	Disconnect-MgGraph
}
#gavdcodeend 046

#gavdcodebegin 047
function PsExchangeGraphSdk_SendMessage
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myParams = @{
		Message = @{
			Subject = "This is a test email"
			Body = @{
				ContentType = "Text"
				Content = "This is the content of the email"
			}
			ToRecipients = @(
				@{
					EmailAddress = @{
						Address = "usera@domain.onmicrosoft.com"
					}
				}
			)
			CcRecipients = @(
				@{
					EmailAddress = @{
						Address = "userb@domain.onmicrosoft.com"
					}
				}
			)
		}
		SaveToSentItems = "false"
	}

	Send-MgUserMail -UserId $cnfUserName `
					-BodyParameter $myParams

	Disconnect-MgGraph
}
#gavdcodeend 047

#gavdcodebegin 048
function PsExchangeGraphSdk_GetMailFolder
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Get-MgUserMailFolder -UserId $cnfUserName

	Disconnect-MgGraph
}
#gavdcodeend 048

#gavdcodebegin 049
function PsExchangeGraphSdk_GetMailFolderProperties
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myFolders = Get-MgUserMailFolder -UserId $cnfUserName
	foreach($oneFolder in $myFolders)
	{
		Write-Host "--Name - " $oneFolder.DisplayName "--Id - " $oneFolder.Id
	}

	Disconnect-MgGraph
}
#gavdcodeend 049

#gavdcodebegin 050
function PsExchangeGraphSdk_GetOneMailFolder
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `  
										 -ClientID $cnfClientIdWithSecret `  
										 -ClientSecret $cnfClientSecret `  

	Get-MgUserMailFolderMessage -UserId $cnfUserName `
								-MailFolderId "AAMkAGE0ODQ3NTc...7Xo-OrAAAAAAEMAAA="

	Disconnect-MgGraph
}
#gavdcodeend 050

#gavdcodebegin 051
function PsExchangeGraphSdk_CreateOneMailFolder
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myParams = @{
		displayName = "MyNewFolder"
		isHidden = $false
	}

	New-MgUserMailFolder -UserId $cnfUserName `
						 -BodyParameter $myParams

	Disconnect-MgGraph
}
#gavdcodeend 051

#gavdcodebegin 052
function PsExchangeGraphSdk_UpdateFolder
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	$myParams = @{
		DisplayName = "MyNewFolderUpdated"
	}

	Update-MgUserMailFolder -UserId $cnfUserName `
							-MailFolderId "AAMkAGE0ODQ3NTc1LTZkM2...OrAADqDEOrAAA=" `
							@myParams

	Disconnect-MgGraph
}
#gavdcodeend 052

#gavdcodebegin 053
function PsExchangeGraphSdk_DeleteFolder
{
	PsGraphPowerShellSdk_LoginWithSecret -TenantName $cnfTenantName `
										 -ClientID $cnfClientIdWithSecret `
										 -ClientSecret $cnfClientSecret `

	Remove-MgUserMailFolder -UserId $cnfUserName `
							-MailFolderId "AAMkAGE0ODQ3NTc1LTZkM2I...rAADqDEOrAAA="

	Disconnect-MgGraph
}
#gavdcodeend 053

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 053 ***

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

##==> EWS
#Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
#$ExService = PsEws_Connect  #*** LEGACY CODE ***

#PsExchangeEws_CreateOneFolder $ExService  #*** LEGACY CODE ***
#PsExchangeEws_GetAllFolders $ExService  #*** LEGACY CODE ***
#PsExchangeEws_FindOneFolder $ExService  #*** LEGACY CODE ***
#PsExchangeEws_UpdateOneFolder $ExService  #*** LEGACY CODE ***
#PsExchangeEws_DeleteOneFolder $ExService  #*** LEGACY CODE ***
#PsExchangeEws_CreateAndSendEmail $ExService  #*** LEGACY CODE ***
#PsExchangeEws_GetUnreadEmails $ExService  #*** LEGACY CODE ***
#PsExchangeEws_ReplyToEmail $ExService  #*** LEGACY CODE ***
#PsExchangeEws_DeleteOneEmail $ExService  #*** LEGACY CODE ***
#PsExchangeEws_CreateOneContact $ExService  #*** LEGACY CODE ***
#PsExchangeEws_FindAllContacts $ExService  #*** LEGACY CODE ***
#PsExchangeEws_FindOneContactByName $ExService  #*** LEGACY CODE ***
#PsExchangeEws_UpdateOneContact $ExService  #*** LEGACY CODE ***
#PsExchangeEws_DeleteOneContact $ExService  #*** LEGACY CODE ***
#PsExchangeEws_CreateAppointment $ExService  #*** LEGACY CODE ***
#PsExchangeEws_FindAppointmentsByDate $ExService  #*** LEGACY CODE ***
#PsExchangeEws_UpdateOneAppointment $ExService  #*** LEGACY CODE ***
#PsExchangeEws_DeleteOneAppointment $ExService  #*** LEGACY CODE ***

##==> CLI
#PsExchangeCli_GetAllMessages
#PsExchangeCli_GetMessageById
#PsExchangeCli_GetOneMessageById
#PsExchangeCli_MoveMessage
#PsExchangeCli_SendMessage
#PsExchangeCli_ReportActivity
#PsExchangeCli_ReportActivityUser
#PsExchangeCli_ReportActivityUserDetails
#PsExchangeCli_ReportActivityByAppTotals
#PsExchangeCli_ReportActivityByApp
#PsExchangeCli_ReportActivityByAppAndUserDetails
#PsExchangeCli_ReportActivityByAppVersions
#PsExchangeCli_ReportUsageDetail
#PsExchangeCli_ReportUsageMailboxes
#PsExchangeCli_ReportQuotaMailboxes
#PsExchangeCli_ReportStorageMailboxes

##==> PnP PowerShell
#PsExchangePnP_SendEmailWithFrom
#PsExchangePnP_SendEmailWithoutFrom
#PsExchangePnP_SendEmailFromSmtpMs
#PsExchangePnP_SendEmailFromSmtpAll

##==> Graph SDK PowerShell
#PsExchangeGraphSdk_GetMessages
#PsExchangeGraphSdk_GetOneMessageById
#PsExchangeGraphSdk_GetMessageContent
#PsExchangeGraphSdk_GetOneMessageAttachments
#PsExchangeGraphSdk_GetOneMessageOneAttachment
#PsExchangeGraphSdk_CreateMessageDraft
#PsExchangeGraphSdk_UpdateMessageContent
#PsExchangeGraphSdk_UpdateMessage
#PsExchangeGraphSdk_DeleteMessage
#PsExchangeGraphSdk_SendMessage
#PsExchangeGraphSdk_GetMailFolder
#PsExchangeGraphSdk_GetMailFolderProperties
#PsExchangeGraphSdk_GetOneMailFolder
#PsExchangeGraphSdk_CreateOneMailFolder
#PsExchangeGraphSdk_UpdateFolder
#PsExchangeGraphSdk_DeleteFolder

Write-Host "Done"  

