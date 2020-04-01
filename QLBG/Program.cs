using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QLBG
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["exUserName"],
                                ConfigurationManager.AppSettings["exUserPw"]);

            //CreateOneContact(myExService);
            //CreateOneContactWithPhoto(myExService);
            //FindAllContacts(myExService);
            //FindOneContactByName(myExService);
            //FindContactsByPartialName(myExService);
            //FindContactsByPartialNameFiltered(myExService);
            //GetPhotoOneContact(myExService);
            //UpdateOneContact(myExService);
            //DeleteOneContact(myExService);
            //ExportContacts(myExService);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void CreateOneContact(ExchangeService ExService)
        {
            Contact newContact = new Contact(ExService)
            {
                GivenName = "Somename",
                MiddleName = "Mymiddle",
                Surname = "Hersurname",
                FileAsMapping = FileAsMapping.SurnameCommaGivenName,
                CompanyName = "Mycompany"
            };

            newContact.PhoneNumbers[PhoneNumberKey.BusinessPhone] = "1234567890";
            newContact.PhoneNumbers[PhoneNumberKey.HomePhone] = "0987654321";
            newContact.PhoneNumbers[PhoneNumberKey.CarPhone] = "1029384756";

            newContact.EmailAddresses[EmailAddressKey.EmailAddress1] = new 
                                                    EmailAddress("somename@domain.com");
            newContact.EmailAddresses[EmailAddressKey.EmailAddress2] = new 
                                            EmailAddress("somename.mymiddle@domain.com");

            newContact.ImAddresses[ImAddressKey.ImAddress1] = "somenameIM1@domain.com";
            newContact.ImAddresses[ImAddressKey.ImAddress2] = "somenameIM2@domain.com";

            PhysicalAddressEntry paHome = new PhysicalAddressEntry
            {
                Street = "123 Somewhere Street",
                City = "Here",
                State = "AZ",
                PostalCode = "92835",
                CountryOrRegion = "Europe"
            };
            newContact.PhysicalAddresses[PhysicalAddressKey.Home] = paHome;

            PhysicalAddressEntry paBusiness = new PhysicalAddressEntry
            {
                Street = "456 Somewhere Avenue",
                City = "There",
                State = "ZA",
                PostalCode = "53829",
                CountryOrRegion = "Europe"
            };
            newContact.PhysicalAddresses[PhysicalAddressKey.Business] = paBusiness;

            newContact.Save();
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void FindAllContacts(ExchangeService ExService)
        {
            ContactsFolder myContactsfolder = ContactsFolder.Bind(
                                                ExService, WellKnownFolderName.Contacts);

            int ammountContacts = myContactsfolder.TotalCount < 50 ? 
                                                        myContactsfolder.TotalCount : 50;

            ItemView myView = new ItemView(ammountContacts);
            myView.PropertySet = new PropertySet(BasePropertySet.IdOnly, 
                                                            ContactSchema.DisplayName);
            FindItemsResults<Item> allContacts = ExService.FindItems(
                                                    WellKnownFolderName.Contacts, myView);

            foreach (Item oneContact in allContacts)
            {
                if (oneContact is Contact myContact)
                {
                    Console.WriteLine(myContact.DisplayName);
                }
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void FindOneContactByName(ExchangeService ExService)
        {
            FindItemsResults<Item> allFound =
                        ExService.FindItems(WellKnownFolderName.Contacts,
                            new SearchFilter.IsEqualTo(ContactSchema.GivenName, "Somename"),
                            new ItemView(5));

            foreach (Item oneFound in allFound)
            {
                if (oneFound is Contact foundContact)
                {
                    Console.WriteLine(foundContact.CompleteName.FullName + " - " +
                                      foundContact.CompanyName);
                }
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void UpdateOneContact(ExchangeService ExService)
        {
            FindItemsResults<Item> allFound =
                        ExService.FindItems(WellKnownFolderName.Contacts,
                            new SearchFilter.IsEqualTo(ContactSchema.GivenName, "Somename"),
                            new ItemView(5));

            ItemId myContactId = null;
            foreach (Item oneFound in allFound)
            {
                if (oneFound is Contact foundContact)
                {
                    myContactId = foundContact.Id;
                }
            }

            Contact myContact = Contact.Bind(ExService, myContactId);

            myContact.Surname = "Hissurname";
            myContact.CompanyName = "Hiscompany";
            myContact.PhoneNumbers[PhoneNumberKey.BusinessPhone] = "32143254321";
            myContact.EmailAddresses[EmailAddressKey.EmailAddress2] = new 
                                                    EmailAddress("somebody@domain.com");
            myContact.ImAddresses[ImAddressKey.ImAddress1] = "otherM1@domain.com";

            PhysicalAddressEntry paBusiness = new PhysicalAddressEntry
            {
                Street = "987 Somewhere Way",
                City = "Noidea",
                State = "ZZ",
                PostalCode = "66666",
                CountryOrRegion = "Europe"
            };
            myContact.PhysicalAddresses[PhysicalAddressKey.Business] = paBusiness;

            myContact.Update(ConflictResolutionMode.AlwaysOverwrite);
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void DeleteOneContact(ExchangeService ExService)
        {
            FindItemsResults<Item> allFound =
                        ExService.FindItems(WellKnownFolderName.Contacts,
                            new SearchFilter.IsEqualTo(ContactSchema.GivenName, "Somename"),
                            new ItemView(5));

            ItemId myContactId = null;
            foreach (Item oneFound in allFound)
            {
                if (oneFound is Contact foundContact)
                {
                    myContactId = foundContact.Id;
                }
            }

            Contact myContact = Contact.Bind(ExService, myContactId);

            myContact.Delete(DeleteMode.MoveToDeletedItems);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void FindContactsByPartialName(ExchangeService ExService)
        {
            NameResolutionCollection resolvedNames = ExService.ResolveName("Mymiddle");

            foreach (NameResolution oneName in resolvedNames)
            {
                Console.WriteLine("Name: " + oneName.Mailbox.Name);
                Console.WriteLine("Email: " + oneName.Mailbox.Address);
                Console.WriteLine("Id: " + oneName.Mailbox.Id);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void FindContactsByPartialNameFiltered(ExchangeService ExService)
        {
            NameResolutionCollection resolvedNames = ExService.ResolveName(
                            "Mymiddle", ResolveNameSearchLocation.ContactsOnly, false);

            foreach (NameResolution oneName in resolvedNames)
            {
                Console.WriteLine("Name: " + oneName.Mailbox.Name);
                Console.WriteLine("Email: " + oneName.Mailbox.Address);
                Console.WriteLine("Id: " + oneName.Mailbox.Id);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void CreateOneContactWithPhoto(ExchangeService ExService)
        {
            Contact newContact = new Contact(ExService)
            {
                GivenName = "Somename",
                Surname = "Withphoto",
                FileAsMapping = FileAsMapping.SurnameCommaGivenName,
            };

            FileAttachment contactPhoto =
                    newContact.Attachments.AddFileAttachment(@"C:\Temporary\MyPhoto.jpg");
            contactPhoto.IsContactPhoto = true;

            newContact.Save();
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void GetPhotoOneContact(ExchangeService ExService)
        {
            FindItemsResults<Item> allFound =
                        ExService.FindItems(WellKnownFolderName.Contacts,
                            new SearchFilter.IsEqualTo(ContactSchema.Surname, "Withphoto"),
                            new ItemView(5));

            ItemId myContactId = null;
            foreach (Item oneFound in allFound)
            {
                if (oneFound is Contact foundContact)
                {
                    myContactId = foundContact.Id;
                }
            }

            Contact myContact = Contact.Bind(ExService, myContactId);

            myContact.Load(new PropertySet(ContactSchema.Attachments));
            foreach (Attachment oneAttachment in myContact.Attachments)
            {
                if ((oneAttachment as FileAttachment).IsContactPhoto)
                {
                    oneAttachment.Load();
                }
            }

            FileAttachment contactPhoto = myContact.GetContactPictureAttachment();
            using (FileStream myPhotoStream = new FileStream(
                                @"C:\Temporary\Photo_" + contactPhoto.Name, 
                                FileMode.Create, System.IO.FileAccess.Write))
            {
                contactPhoto.Load(myPhotoStream);
            }
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        private static void ExportContacts(ExchangeService ExService)
        {
            FindItemsResults<Item> findResults =
                        ExService.FindItems(WellKnownFolderName.Contacts,
                            new SearchFilter.IsEqualTo(ContactSchema.Surname, "Withphoto"),
                            new ItemView(1));

            ItemId myContactId = null;
            foreach (Item oneItem in findResults)
            {
                if (oneItem is Contact foundContact)
                {
                    myContactId = foundContact.Id;
                }
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                                                  ContactSchema.MimeContent);
            Contact contactToExport = Contact.Bind(ExService, myContactId, myPropSet);

            string vcfFileName = @"C:\Temporary\myContact.vcf";
            using (FileStream myContactStream = new FileStream(
                            vcfFileName,
                            FileMode.Create, FileAccess.Write))
            {
                myContactStream.Write(contactToExport.MimeContent.Content, 0,
                                             contactToExport.MimeContent.Content.Length);
            }
        }
        //gavdcodeend 10

        //-------------------------------------------------------------------------------
        static ExchangeService ConnectBA(string userEmail, string userPW)
        {
            ExchangeService exService = new ExchangeService
            {
                Credentials = new WebCredentials(userEmail, userPW)
            };

            //exService.TraceEnabled = true;
            //exService.TraceFlags = TraceFlags.All;

            exService.AutodiscoverUrl(userEmail, RedirectionUrlValidationCallback);

            return exService;
        }

        static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool validationResult = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                validationResult = true;
            }

            return validationResult;
        }
    }
}
