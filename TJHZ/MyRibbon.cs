using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace TJHZ
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //gavdcodebegin 001
        Outlook.Application myApplication;
        Outlook.Inspector myInspector;

        private void btnGetSelectedText_Click(object sender, RibbonControlEventArgs e)
        {
            myApplication = Globals.ThisAddIn.Application;
            myInspector = myApplication.ActiveInspector();

            Outlook.MailItem myMailItem = myInspector.CurrentItem as Outlook.MailItem;
            if (myMailItem != null)
            {
                Word.Document myWordDocument = (Word.Document)myInspector.WordEditor;
                string selectedText = myWordDocument.Application.Selection.Text;

                MessageBox.Show(selectedText);
            }
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        private void btnCreateEmail_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.MailItem myMailItem = (Outlook.MailItem)
                            myApplication.CreateItem(Outlook.OlItemType.olMailItem);
            myMailItem.Subject = "Subject of the new Email";
            myMailItem.To = "emailto@somewhere.com";
            myMailItem.Body = "Body of the Email";
            myMailItem.Importance = Outlook.OlImportance.olImportanceLow;
            myMailItem.Display(false);
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        private void btnAddAttachment_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.MailItem myMailItem = (Outlook.MailItem)
                            myApplication.CreateItem(Outlook.OlItemType.olMailItem);
            myMailItem.Subject = "Email with attachment";
            myMailItem.To = "emailto@somewhere.com";
            myMailItem.Body = "Body of the Email";
            myMailItem.Importance = Outlook.OlImportance.olImportanceLow;

            OpenFileDialog myAttachment = new OpenFileDialog();
            myAttachment.Title = "Select a file for the attachment";
            myAttachment.ShowDialog();

            if (myAttachment.FileName.Length > 0)
            {
                myMailItem.Attachments.Add(myAttachment.FileName,
                                           Outlook.OlAttachmentType.olByValue,
                                           1,
                                           myAttachment.FileName);
            }

            ((Outlook._MailItem)myMailItem).Send();
        }
        //gavdcodeend 003

        //gavdcodebegin 004
        private void btnGetEmails_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.MAPIFolder myInbox = myApplication.ActiveExplorer().Session.
                                GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Items myUnreadItems = myInbox.Items.Restrict("[Unread]=true");

            MessageBox.Show("Unread items = " + myUnreadItems.Count.ToString());
        }
        //gavdcodeend 004

        //gavdcodebegin 005
        private void btnSaveAttachments_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;
            Outlook.Inspector myInspector = myApplication.ActiveInspector();

            Outlook.MailItem myMailItem = myInspector.CurrentItem as Outlook.MailItem;
            if (myMailItem != null)
            {
                Outlook.Attachments allItemAttachments = myMailItem.Attachments;
                foreach (Outlook.Attachment oneItemAttachment in allItemAttachments)
                {
                    oneItemAttachment.SaveAsFile(@"C:\Temporary\" +
                            oneItemAttachment.FileName);
                }
            }
        }
        //gavdcodeend 005

        //gavdcodebegin 006
        private void btnFindContact_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.NameSpace outlookNameSpace = myApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder myContactsFolder = outlookNameSpace.GetDefaultFolder(
                                Outlook.OlDefaultFolders.olFolderContacts);

            Outlook.Items myContactItems = myContactsFolder.Items;
            Outlook.ContactItem myContact = (Outlook.ContactItem)myContactItems.
                Find("[FirstName]='a' and [LastName]='b'");
            if (myContact != null)
            {
                myContact.Display(true);
            }
            else
            {
                MessageBox.Show("Contact not found");
            }
        }
        //gavdcodeend 006

        //gavdcodebegin 007
        private void btnAddContact_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.ContactItem newContact = (Outlook.ContactItem)
                            myApplication.CreateItem(Outlook.OlItemType.olContactItem);

            newContact.FirstName = "He Is";
            newContact.LastName = "Somebody";
            newContact.Email1Address = "he.is.somebody@somewhere.com";
            newContact.PrimaryTelephoneNumber = "(123)4567890";
            newContact.MailingAddressStreet = "Here 123";
            newContact.MailingAddressCity = "There";
            newContact.Save();
            newContact.Display(true);
        }
        //gavdcodeend 007

        //gavdcodebegin 008
        private void btnDeleteContact_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.NameSpace outlookNameSpace = myApplication.GetNamespace("MAPI");
            Outlook.MAPIFolder myContactsFolder = outlookNameSpace.GetDefaultFolder(
                                Outlook.OlDefaultFolders.olFolderContacts);

            Outlook.Items myContactItems = myContactsFolder.Items;
            Outlook.ContactItem myContact = (Outlook.ContactItem)myContactItems.
                Find("[FirstName]='a' and [LastName]='b'");
            if (myContact != null)
            {
                myContact.Delete();
            }
        }
        //gavdcodeend 008

        //gavdcodebegin 009
        private void btnCreateAppointment_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.AppointmentItem newAppointment = (Outlook.AppointmentItem)
                        myApplication.CreateItem(Outlook.OlItemType.olAppointmentItem);
            
            newAppointment.Start = DateTime.Now.AddHours(1);
            newAppointment.End = DateTime.Now.AddHours(2);
            newAppointment.Location = "An Empty Room";
            newAppointment.Body = "This is a test appointment";
            newAppointment.AllDayEvent = false;
            newAppointment.Subject = "My test";
            newAppointment.Recipients.Add("Somebody, He Is");

            Outlook.Recipients sentAppointmentTo = newAppointment.Recipients;
            Outlook.Recipient sentInvite = null;
            sentInvite = sentAppointmentTo.Add("b, bbb b");
            sentInvite.Type = (int)Outlook.OlMeetingRecipientType.olRequired;
            sentInvite = sentAppointmentTo.Add("c, ccc c");
            sentInvite.Type = (int)Outlook.OlMeetingRecipientType.olOptional;
            sentAppointmentTo.ResolveAll();
            newAppointment.Save();
            newAppointment.Display(true);
        }
        //gavdcodeend 009

        //gavdcodebegin 010
        private void btnDeleteAppointment_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.MAPIFolder myCalendar = myApplication.Session.GetDefaultFolder(
                                            Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items calendarItems = myCalendar.Items;
            Outlook.AppointmentItem oneCalendarItem =
                    (Outlook.AppointmentItem)calendarItems["Test Appointment to delete"];

            if (oneCalendarItem != null)
            {
                oneCalendarItem.Delete();
            }
        }
        //gavdcodeend 010

        //gavdcodebegin 011
        private void btnCreateCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            const string newCalendarName = "MyCalendar";
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.MAPIFolder myCalendar = (Outlook.MAPIFolder)myApplication.
                                            ActiveExplorer().Session.GetDefaultFolder
                                            (Outlook.OlDefaultFolders.olFolderCalendar);
            bool thereIsNoFolder = true;
            foreach (Outlook.MAPIFolder oneCalendarFolder in myCalendar.Folders)
            {
                if (oneCalendarFolder.Name == newCalendarName)
                {
                    thereIsNoFolder = false;
                    break;
                }
            }

            if (thereIsNoFolder)
            {
                Outlook.MAPIFolder newCalendar = myCalendar.Folders.Add(
                            newCalendarName, Outlook.OlDefaultFolders.olFolderCalendar);
                Outlook.AppointmentItem newEvent = newCalendar.Items.Add(
                                                Outlook.OlItemType.olAppointmentItem)
                                                as Outlook.AppointmentItem;
                newEvent.Start = DateTime.Now.AddHours(1);
                newEvent.End = DateTime.Now.AddHours(1.25);
                newEvent.Subject = "Test in new calendar";
                newEvent.Body = "This is a new meeting in a new calendar";
                newEvent.Save();
            }
            myApplication.ActiveExplorer().SelectFolder(myCalendar.
                                                            Folders[newCalendarName]);
            myApplication.ActiveExplorer().CurrentFolder.Display();
        }
        //gavdcodeend 011

        //gavdcodebegin 012
        private void btnCreateFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            Outlook.Folder myInBox = (Outlook.Folder)myApplication.ActiveExplorer().
                        Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            //string userName = (string)myApplication.ActiveExplorer()
            //    .Session.CurrentUser.Name;
            string userName = "My New Folder";
            Outlook.Folder myNewFolder = null;
            myNewFolder = (Outlook.Folder)myInBox.Folders.Add(userName,
                                            Outlook.OlDefaultFolders.olFolderInbox);
            myInBox.Folders[userName].Display();
        }
        //gavdcodeend 012

        //gavdcodebegin 013
        private void btnSelectFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            string folderToFind = "My New Folder";
            Outlook.Folder myInBox = (Outlook.Folder)myApplication.ActiveExplorer().
                        Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            try
            {
                myApplication.ActiveExplorer().CurrentFolder = 
                                                            myInBox.Folders[folderToFind];
                myApplication.ActiveExplorer().CurrentFolder.Display();
            }
            catch
            {
                MessageBox.Show("There is no folder " + folderToFind);
            }
        }
        //gavdcodeend 013

        //gavdcodebegin 014
        private void btnDeleteFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application myApplication = Globals.ThisAddIn.Application;

            string folderToFind = "My New Folder";
            Outlook.Folder myInBox = (Outlook.Folder)myApplication.ActiveExplorer().
                        Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            Outlook.Folder myFolder = (Outlook.Folder)myInBox.Folders[folderToFind];

            if (myFolder != null)
            {
                myFolder.Delete();
            }
        }
        //gavdcodeend 014
    }
}
