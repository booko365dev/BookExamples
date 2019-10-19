
Function ConnectPsEwsBA()
{
	$ExService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
	$ExService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(`
		$configFile.appsettings.exUserName, $configFile.appsettings.exUserPw)
	$ExService.Url = new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx");
	#$ExService.TraceEnabled = $true
	#$ExService.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All
	$ExService.AutodiscoverUrl($configFile.appsettings.exUserName, {$true})

	return $ExService
}
#-----------------------------------------------------------------------------------------

Function ExPsEwsCreateOneFolder($ExService)
{
    $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($ExService)
    $newFolder.DisplayName = "My Custom Folder PowerShell"

    $newFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
}

Function ExPsEwsGetAllFolders($ExService) {
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

Function ExPsEwsFindOneFolder($ExService)
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

Function ExPsEwsUpdateOneFolder($ExService)
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

Function ExPsEwsDeleteOneFolder($ExService)
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

Function ExPsEwsCreateAndSendEmail($ExService)
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

Function ExPsEwsGetUnreadEmails($ExService)
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

Function ExPsEwsReplyToEmail($ExService)
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

Function ExPsEwsDeleteOneEmail($ExService)
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

Function ExPsEwsCreateOneContact($ExService)
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

Function ExPsEwsFindAllContacts($ExService)
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

Function ExPsEwsFindOneContactByName($ExService)
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

Function ExPsEwsUpdateOneContact($ExService)
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

Function ExPsEwsDeleteOneContact($ExService)
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

Function ExPsEwsCreateAppointment($ExService)
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

Function ExPsEwsFindAppointmentsByDate($ExService)
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

Function ExPsEwsUpdateOneAppointment($ExService)
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

Function ExPsEwsDeleteOneAppointment($ExService)
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

#-----------------------------------------------------------------------------------------

Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

[xml]$configFile = get-content "C:\Projects\exPs.values.config"

$ExService = ConnectPsEwsBA

#ExPsEwsCreateOneFolder $ExService
#ExPsEwsGetAllFolders $ExService
#ExPsEwsFindOneFolder $ExService
#ExPsEwsUpdateOneFolder $ExService
#ExPsEwsDeleteOneFolder $ExService
#ExPsEwsCreateAndSendEmail $ExService
#ExPsEwsGetUnreadEmails $ExService
#ExPsEwsReplyToEmail $ExService
#ExPsEwsDeleteOneEmail $ExService
#ExPsEwsCreateOneContact $ExService
#ExPsEwsFindAllContacts $ExService
#ExPsEwsFindOneContactByName $ExService
#ExPsEwsUpdateOneContact $ExService
#ExPsEwsDeleteOneContact $ExService
#ExPsEwsCreateAppointment $ExService
#ExPsEwsFindAppointmentsByDate $ExService
#ExPsEwsUpdateOneAppointment $ExService
#ExPsEwsDeleteOneAppointment $ExService

Write-Host "Done"  


