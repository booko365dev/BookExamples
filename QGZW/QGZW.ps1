
##---------------------------------------------------------------------------------------
## ------**** ATTENTION **** This is a PowerShell solution ****--------------------------
##---------------------------------------------------------------------------------------

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Login routines ***---------------------------
##---------------------------------------------------------------------------------------

function PsSpGraphCli_LoginWithCertificate
{
	mgc login --tenant-id $configFile.appsettings.TenantName `
			  --client-id $configFile.appsettings.ClientIdWithCert `
			  --certificate-thumb-print $configFile.appsettings.CertificateThumbprint `
			  --strategy ClientCertificate
}


##---------------------------------------------------------------------------------------
##***-----------------------------------*** Example routines ***-------------------------
##---------------------------------------------------------------------------------------

#gavdcodebegin 001
function PsSpGraphCli_GetAllListsInSite
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"

	mgc sites lists list --site-id $mySiteId

	mgc logout
}
#gavdcodeend 001

#gavdcodebegin 002
function PsSpGraphCli_GetOneListInSite
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"

	mgc sites lists get --site-id $mySiteId `
						--list-id $myListId

	mgc logout
}
#gavdcodeend 002

#gavdcodebegin 003
function PsSpGraphCli_CreateOneListInSite
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myBody = @{	displayName = "List Created with GraphCLI"
					columns = @(
						@{ name = "MyTextField"; text = @{} }
						@{ name = "MyNumberField"; number = @{} }
					)
					list = @{ template = "genericList" }
				} | ConvertTo-Json -Depth 5

	mgc sites lists create --site-id $mySiteId `
						   --body $myBody

	mgc logout
}
#gavdcodeend 003

#gavdcodebegin 004
function PsSpGraphCli_UpdateOneListInSite
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myBody = @{	description = "List description updated"
				} | ConvertTo-Json -Depth 5

	mgc sites lists patch --site-id $mySiteId `
						  --list-id $myListId `
						  --body $myBody

	mgc logout
}
#gavdcodeend 004

#gavdcodebegin 005
function PsSpGraphCli_DeleteOneListFromSite
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"

	mgc sites lists delete --site-id $mySiteId `
						   --list-id $myListId

	mgc logout
}
#gavdcodeend 005

#gavdcodebegin 006
function PsSpGraphCli_GetAllColumnsInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"

	mgc sites lists columns list --site-id $mySiteId `
								 --list-id $myListId

	mgc logout
}
#gavdcodeend 006

#gavdcodebegin 007
function PsSpGraphCli_GetOneColumnInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myColumnId = "de8ac4e8-2145-4bdd-ba9b-3450baecffdc"

	mgc sites lists columns get --site-id $mySiteId `
								--list-id $myListId `
								--column-definition-id $myColumnId

	mgc logout
}
#gavdcodeend 007

#gavdcodebegin 008
function PsSpGraphCli_CreateOneColumnInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myBody = @{	name = "Column Created with GraphCLI"
					enforceUniqueValues = $false
					hidden = $false
					indexed = $false
					description = "Column Created with GraphCLI"
					text = @{   allowMultipleLines = $false
								appendChangesToExistingText = $false
								linesForEditing = 0
								maxLength = 255
							}
					} | ConvertTo-Json -Depth 5

	mgc sites lists columns create --site-id $mySiteId `
								   --list-id $myListId `
								   --body $myBody

	mgc logout
}
#gavdcodeend 008

#gavdcodebegin 009
function PsSpGraphCli_UpdateOneColumnInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myColumnId = "07129402-8171-4151-b331-c382a9b42f6a"
	$myBody = @{	description = "Column description updated"
				} | ConvertTo-Json -Depth 5

	mgc sites lists columns patch --site-id $mySiteId `
								  --list-id $myListId `
								  --column-definition-id $myColumnId `
								  --body $myBody

	mgc logout
}
#gavdcodeend 009

#gavdcodebegin 010
function PsSpGraphCli_DeleteOneColumnInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myColumnId = "07129402-8171-4151-b331-c382a9b42f6a"

	mgc sites lists columns delete --site-id $mySiteId `
								   --list-id $myListId `
								   --column-definition-id $myColumnId

	mgc logout
}
#gavdcodeend 010

#gavdcodebegin 011
function PsSpGraphCli_GetAllContentTypesInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"

	mgc sites lists content-types list --site-id $mySiteId `
									   --list-id $myListId

	mgc logout
}
#gavdcodeend 011

#gavdcodebegin 012
function PsSpGraphCli_GetOneContentTypeInList
{
	# Requires Delegated rights: Sites.Read.All, Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myContentTypeId = "0x0100C48E769F09462748923DDAAD58A2A72C"

	mgc sites lists content-types get --site-id $mySiteId `
									  --list-id $myListId `
									  --content-type-id $myContentTypeId

	mgc logout
}
#gavdcodeend 012

#gavdcodebegin 013
function PsSpGraphCli_CreateOneContentTypeInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myBody = @{	name = "Content Type from GraphCLI"
					description = "My description"
					base = @{
						name = "Document"
						id = "0x010101"
					}
					group = "Document Content Types"
				} | ConvertTo-Json -Depth 5

	mgc sites lists content-types create --site-id $mySiteId `
										 --list-id $myListId `
										 --body $myBody

	mgc logout
}
#gavdcodeend 013

#gavdcodebegin 014
function PsSpGraphCli_UpdateOneContentTypeInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myContentTypeId = "0x0100C48E769F09462748923DDAAD58A2A72C"
	$myBody = @{	description = "ContentType description updated"
				} | ConvertTo-Json -Depth 5

	mgc sites lists content-types patch --site-id $mySiteId `
										--list-id $myListId `
										--content-type-id $myContentTypeId `
										--body $myBody

	mgc logout
}
#gavdcodeend 014

#gavdcodebegin 015
function PsSpGraphCli_DeleteOneContentTypeInList
{
	# Requires Delegated rights: Sites.ReadWrite.All

	PsSpGraphCli_LoginWithCertificate

	$mySiteId = "91ee115a-8a5b-49ad-9627-99dae04394ab"
	$myListId = "1becd33f-b69d-4e07-99b4-fdcb7ecf2429"
	$myContentTypeId = "0x0100C48E769F09462748923DDAAD58A2A72C"

	mgc sites lists content-types delete --site-id $mySiteId `
										 --list-id $myListId `
										 --content-type-id $myContentTypeId

	mgc logout
}
#gavdcodeend 015

##---------------------------------------------------------------------------------------
##***-----------------------------------*** Running the routines ***---------------------
##---------------------------------------------------------------------------------------

# *** Latest Source Code Index: 015 ***

[xml]$configFile = get-content "C:\Projects\ConfigValuesPs.config"

#*** Using the MS Graph CLI
#		ATTENTION: There is a Windows Environment Variable already configured in the computer
#					to redirect the commands to the mgc.exe directory (see instructions in the book)
#PsSpGraphCli_GetAllListsInSite
#PsSpGraphCli_GetOneListInSite
#PsSpGraphCli_CreateOneListInSite
#PsSpGraphCli_UpdateOneListInSite
#PsSpGraphCli_DeleteOneListFromSite
#PsSpGraphCli_GetAllColumnsInList
#PsSpGraphCli_GetOneColumnInList
#PsSpGraphCli_CreateOneColumnInList
#PsSpGraphCli_UpdateOneColumnInList
#PsSpGraphCli_DeleteOneColumnInList
#PsSpGraphCli_GetAllContentTypesInList
#PsSpGraphCli_GetOneContentTypeInList
#PsSpGraphCli_CreateOneContentTypeInList
#PsSpGraphCli_UpdateOneContentTypeInList
#PsSpGraphCli_DeleteOneContentTypeInList

Write-Host "Done" 

