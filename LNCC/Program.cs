﻿using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace LNCC
{
    class Program
    {
        static void Main(string[] args)  //*** LEGACY CODE ***
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

            //CreateAndSendEmail(myExService);
            //CreateDraftEmail(myExService);
            //SendDraftEmail(myExService);
            //SendDelayedEmail(myExService);
            //ReplyToEmail(myExService);
            //ForwardEmail(myExService);
            //GetUnreadEmails(myExService);
            //MoveOneEmail(myExService);
            //CopyOneEmail(myExService);
            //DeleteOneEmail(myExService);
            //EntityExtractionFromEmail(myExService);
            //ExportEmail(myExService);
            //ImportEmail(myExService);
            //GetOutOfOfficeConfig(myExService);
            //SetOutOfOfficeConfig(myExService);
            //CreateAndSendEmailWithAttachment(myExService);
            //GetAttachments(myExService);
            //RemoveAttachmentsFromEmail(myExService);
            //SetEmailAsJunk(myExService);
            //CreateInboxRule(myExService);
            //GetInboxRules(myExService);
            //UpdateInboxRule(myExService);
            //DeleteInboxRule(myExService);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void CreateAndSendEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            EmailMessage newEmail = new EmailMessage(ExService)
            {
                Subject = "Email send by EWS",
                Body = "To Whom It May Concern"
            };

            newEmail.ToRecipients.Add("user01@domain.com");
            newEmail.BccRecipients.Add("user02@domain.com");
            newEmail.CcRecipients.Add("user03@domain.com");
            //newEmail.From = "user04@domain.com";
            newEmail.Importance = Importance.Normal;
            
            newEmail.SendAndSaveCopy();
            //newEmail.Send();
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void CreateDraftEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            EmailMessage newEmail = new EmailMessage(ExService)
            {
                Subject = "Draft Email created by EWS",
                Body = "To Whom It May Concern"
            };
            newEmail.ToRecipients.Add("user@domain.com");

            newEmail.Save(WellKnownFolderName.Drafts);
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SendDraftEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And, 
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject, 
                                        "Draft Email created by EWS"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Drafts, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        EmailMessageSchema.Subject, EmailMessageSchema.ToRecipients);
            EmailMessage newEmail = EmailMessage.Bind(ExService, myEmailId, myPropSet);

            newEmail.SendAndSaveCopy();
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SendDelayedEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            EmailMessage newEmail = new EmailMessage(ExService);

            ExtendedPropertyDefinition PR_DEFERRED_SEND_TIME = new 
                                                ExtendedPropertyDefinition(16367,
                                                MapiPropertyType.SystemTime);
            string sendTime = DateTime.Now.AddMinutes(2).ToUniversalTime().ToString();

            newEmail.SetExtendedProperty(PR_DEFERRED_SEND_TIME, sendTime);

            newEmail.ToRecipients.Add("user@domain.net");
            newEmail.Subject = "Delayed Emails sent by EWS";

            newEmail.Body = "Submitted at " + DateTime.Now.ToString() +
                           " - Sent at " + sendTime;
            newEmail.SendAndSaveCopy();
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void ReplyToEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Email asking for replay"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            EmailMessage emailToReply = EmailMessage.Bind(ExService, myEmailId);

            string myReply = "Reply body";
            emailToReply.Reply(myReply, true);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void ForwardEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Email to forward"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            EmailMessage emailToReply = EmailMessage.Bind(ExService, myEmailId);

            EmailAddress[] forwardAddresses = new EmailAddress[1];
            forwardAddresses[0] = new EmailAddress("user@domain.com");
            string myForward = "Forward body";
            emailToReply.Forward(myForward, forwardAddresses);
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void GetUnreadEmails(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And, new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.IsRead, false));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            Console.WriteLine(findResults.TotalCount.ToString());
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void MoveOneEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Email created by EWS"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);
            EmailMessage emailToMove = EmailMessage.Bind(ExService, myEmailId, myPropSet);

            emailToMove.Move(WellKnownFolderName.JunkEmail);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void CopyOneEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Email created by EWS"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);
            EmailMessage emailToCopy = EmailMessage.Bind(ExService, myEmailId, myPropSet);

            emailToCopy.Copy(WellKnownFolderName.Drafts);
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void DeleteOneEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Email created by EWS"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        EmailMessageSchema.Subject, EmailMessageSchema.ParentFolderId);
            EmailMessage emailToDelete = EmailMessage.Bind(ExService, myEmailId, myPropSet);

            emailToDelete.Delete(DeleteMode.SoftDelete);
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void EntityExtractionFromEmail(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                LogicalOperator.And,
                                new SearchFilter.IsEqualTo(
                                        EmailMessageSchema.Subject,
                                        "Entity extraction email"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                                                    ItemSchema.EntityExtractionResult);
            Item emailWithEntities = 
                                    EmailMessage.Bind(ExService, myEmailId, myPropSet);

            if (emailWithEntities.EntityExtractionResult != null)
            {
                if (emailWithEntities.EntityExtractionResult.MeetingSuggestions != null)
                {
                    Console.WriteLine("Entity Meetings");
                    foreach (MeetingSuggestion oneMeeting in 
                             emailWithEntities.EntityExtractionResult.MeetingSuggestions)
                    {
                        Console.WriteLine("Meeting: " + oneMeeting.MeetingString);
                    }
                    Console.WriteLine(Environment.NewLine);
                }

                if (emailWithEntities.EntityExtractionResult.EmailAddresses != null)
                {
                    Console.WriteLine("Entity Emails");
                    foreach (EmailAddressEntity oneEmail in
                                 emailWithEntities.EntityExtractionResult.EmailAddresses)
                    {
                        Console.WriteLine("Email address: " + oneEmail.EmailAddress);
                    }
                    Console.WriteLine(Environment.NewLine);
                }

                if (emailWithEntities.EntityExtractionResult.PhoneNumbers != null)
                {
                    Console.WriteLine("Entity Phones");
                    foreach (PhoneEntity onePhone in 
                                   emailWithEntities.EntityExtractionResult.PhoneNumbers)
                    {
                        Console.WriteLine("Phone: " + onePhone.PhoneString);
                    }
                    Console.WriteLine(Environment.NewLine);
                }
            }
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        private static void ExportEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    new SearchFilter.IsEqualTo(
                            EmailMessageSchema.Subject,
                            "Email created by EWS"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                                                    EmailMessageSchema.MimeContent);
            EmailMessage emailToExport =
                                    EmailMessage.Bind(ExService, myEmailId, myPropSet);

            string emlFileName = @"C:\Temporary\myEmail.eml";
            using (FileStream myStream = new FileStream(
                                        emlFileName, FileMode.Create, FileAccess.Write))
            {
                myStream.Write(emailToExport.MimeContent.Content, 0, 
                                        emailToExport.MimeContent.Content.Length);
            }

            string mhtFileName = @"C:\Temporary\myEmail.mht";
            using (FileStream myStream = new FileStream(
                                        mhtFileName, FileMode.Create, FileAccess.Write))
            {
                myStream.Write(emailToExport.MimeContent.Content, 0, 
                                        emailToExport.MimeContent.Content.Length);
            }
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        private static void ImportEmail(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            EmailMessage emailToImport = new EmailMessage(ExService);

            string emlFileName = @"C:\Temporary\myEmail.eml";
            using (FileStream myStream = new FileStream(
                                        emlFileName, FileMode.Open, FileAccess.Read))
            {
                byte[] mailBytes = new byte[myStream.Length];
                int bytesToRead = (int)myStream.Length;
                int bytesRead = 0;
                while (bytesToRead > 0)
                {
                    int myBlock = myStream.Read(mailBytes, bytesRead, bytesToRead);
                    if (myBlock == 0)
                        break;
                    bytesRead += myBlock;
                    bytesToRead -= myBlock;
                }
                emailToImport.MimeContent = new MimeContent("UTF-8", mailBytes);
            }

            ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new 
                            ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
            emailToImport.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);

            emailToImport.Save(WellKnownFolderName.Inbox);
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void GetOutOfOfficeConfig(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            OofSettings myOOFConfig = ExService.GetUserOofSettings("user@domain.com");

            OofExternalAudience myAllowedExternalAudience = myOOFConfig.AllowExternalOof;
            Console.WriteLine(myAllowedExternalAudience.ToString());

            TimeWindow myOOFDuration = myOOFConfig.Duration;
            if(myOOFDuration != null)
                Console.WriteLine(myOOFDuration.StartTime.ToLocalTime() + " - " + 
                                  myOOFDuration.EndTime.ToLocalTime());

            OofExternalAudience myExternalAudience = myOOFConfig.ExternalAudience;
            Console.WriteLine(myExternalAudience.ToString());

            OofReply myExternalReply = myOOFConfig.ExternalReply;
            if(myExternalReply != null)
                Console.WriteLine(myExternalReply.ToString());

            OofReply myInternalReply = myOOFConfig.InternalReply;
            if(myInternalReply != null)
                Console.WriteLine(myInternalReply.ToString());

            OofState myOofState = myOOFConfig.State;
            Console.WriteLine(myOofState.ToString());
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        static void SetOutOfOfficeConfig(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            OofSettings myOOFConfig = new OofSettings
            {
                State = OofState.Enabled,
                Duration = new TimeWindow(
                                    DateTime.Now.AddDays(4), DateTime.Now.AddDays(2)),
                ExternalAudience = OofExternalAudience.All,
                InternalReply = new OofReply("Out of office internal reply"),
                ExternalReply = new OofReply("Out of the office external reply")
            };
            myOOFConfig.ExternalAudience = OofExternalAudience.Known;

            ExService.SetUserOofSettings("user@domain.com", myOOFConfig);
        }
        //gavdcodeend 15

        //gavdcodebegin 16
        static void CreateAndSendEmailWithAttachment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            EmailMessage newEmail = new EmailMessage(ExService)
            {
                Subject = "Email with Attachments",
                Body = "This is an email with attachments"
            };
            newEmail.ToRecipients.Add("user@domain.com");

            newEmail.Attachments.AddFileAttachment(@"C:\Temporary\file_01.jpg");
            newEmail.Attachments.AddFileAttachment("SecondAttachment.jpg", 
                                                    @"C:\Temporary\file_01.jpg");

            byte[] attBytes = File.ReadAllBytes(@"C:\Temporary\file_01.jpg");
            newEmail.Attachments.AddFileAttachment("ThirdAttachment.jpg", attBytes);

            ItemAttachment<EmailMessage> emailAttachment = 
                                newEmail.Attachments.AddItemAttachment<EmailMessage>();
            emailAttachment.Name = "Attached email";
            emailAttachment.Item.Subject = "Attached email Subject";
            emailAttachment.Item.Body = "Attached email Body";
            emailAttachment.Item.ToRecipients.Add("user@domain.com");

            newEmail.SendAndSaveCopy();
        }
        //gavdcodeend 16

        //gavdcodebegin 17
        static void GetAttachments(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    new SearchFilter.IsEqualTo(
                            EmailMessageSchema.Subject,
                            "Email with Attachments"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                                                    ItemSchema.Attachments);
            EmailMessage emailWithAttachments =
                                    EmailMessage.Bind(ExService, myEmailId, myPropSet);

            foreach (Attachment oneAttachment in emailWithAttachments.Attachments)
            {
                if (oneAttachment is FileAttachment) // Attachment is a File
                {
                    FileAttachment myAttachment = oneAttachment as FileAttachment;

                    FileStream attStream = new FileStream(@"C:\Temporary\Attch_" + 
                        myAttachment.Name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
                    myAttachment.Load(attStream);
                    attStream.Close();
                    attStream.Dispose();
                }
                else // Attachment is an Item
                {
                    ItemAttachment itemAttachment = oneAttachment as ItemAttachment;
                    itemAttachment.Load();
                    Console.WriteLine("Subject: " + itemAttachment.Item.Subject);
                }
            }
        }
        //gavdcodeend 17

        //gavdcodebegin 18
        static void RemoveAttachmentsFromEmail(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    new SearchFilter.IsEqualTo(
                            EmailMessageSchema.Subject,
                            "Email with Attachments"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.FirstClassProperties,
                                                    ItemSchema.Attachments);
            EmailMessage emailWithAttachments =
                                    EmailMessage.Bind(ExService, myEmailId, myPropSet);

            // Delete the second attachment
            if (emailWithAttachments.Attachments.Count > 1)
            {
                emailWithAttachments.Attachments.RemoveAt(1);
            }

            // Delete the attachment named "ThirdAttachment.jpg"
            foreach (Attachment oneAttachment in emailWithAttachments.Attachments)
            {
                if (oneAttachment.Name.ToLower() == "thirdattachment.jpg")
                {
                    emailWithAttachments.Attachments.Remove(oneAttachment);
                    break;
                }
            }

            // Delete all attachments
            emailWithAttachments.Attachments.Clear();  

            emailWithAttachments.Update(ConflictResolutionMode.AlwaysOverwrite);
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        private static void SetEmailAsJunk(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                    LogicalOperator.And,
                    new SearchFilter.IsEqualTo(
                            EmailMessageSchema.Subject,
                            "This is junk email"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Inbox, myFilter, myView);

            ItemId myEmailId = null;
            foreach (Item oneItem in findResults)
            {
                myEmailId = oneItem.Id;
            }

            List<ItemId> junkItemIds = new List<ItemId>();
            junkItemIds.Add(myEmailId);
            ExService.MarkAsJunk(junkItemIds, true, true);
        }
        //gavdcodeend 19

        //gavdcodebegin 20
        static void CreateInboxRule(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            Rule newRule = new Rule
            {
                DisplayName = "MoveEmailToJunk",
                Priority = 1,
                IsEnabled = true
            };
            newRule.Conditions.ContainsSubjectStrings.Add("ItIsJunk");
            newRule.Actions.MoveToFolder = WellKnownFolderName.JunkEmail;

            CreateRuleOperation myOperation = new CreateRuleOperation(newRule);
            ExService.UpdateInboxRules(new RuleOperation[] { myOperation }, true);
        }
        //gavdcodeend 20

        //gavdcodebegin 21
        static void GetInboxRules(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            RuleCollection allRules = ExService.GetInboxRules("user@domain.com");

            foreach (Rule oneRule in allRules)
            {
                Console.WriteLine(oneRule.DisplayName + " - " + oneRule.Id);
            }
        }
        //gavdcodeend 21

        //gavdcodebegin 22
        static void UpdateInboxRule(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            RuleCollection allRules = ExService.GetInboxRules("user@domain.com");

            foreach (Rule oneRule in allRules)
            {
                if (oneRule.DisplayName == "MoveEmailToJunk")
                {
                    oneRule.Conditions.ContainsSubjectStrings.Clear();
                    oneRule.Conditions.ContainsSubjectStrings.Add("It Is Junk");

                    SetRuleOperation myOperation = new SetRuleOperation(oneRule);
                    ExService.UpdateInboxRules(new RuleOperation[] { myOperation }, true);
                }
            }
        }
        //gavdcodeend 22

        //gavdcodebegin 23
        static void DeleteInboxRule(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            RuleCollection allRules = ExService.GetInboxRules("user@domain.com");

            foreach (Rule oneRule in allRules)
            {
                if (oneRule.DisplayName == "MoveEmailToJunk")
                {
                    oneRule.Conditions.ContainsSubjectStrings.Clear();
                    oneRule.Conditions.ContainsSubjectStrings.Add("It Is Junk");

                    DeleteRuleOperation myOperation = new DeleteRuleOperation(oneRule.Id);
                    ExService.UpdateInboxRules(new RuleOperation[] { myOperation }, true);
                }
            }
        }
        //gavdcodeend 23

        //-------------------------------------------------------------------------------
        static ExchangeService ConnectBA(string userEmail, string userPW)  //*** LEGACY CODE ***
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

        static bool RedirectionUrlValidationCallback(string redirectionUrl)  //*** LEGACY CODE ***
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
