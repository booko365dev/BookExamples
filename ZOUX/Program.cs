using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace ZOUX
{
    class Program
    {
        static void Main(string[] args)  //*** LEGACY CODE ***
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["UserName"],
                                ConfigurationManager.AppSettings["UserPw"]);

            //CreateAppointment(myExService);
            //CreateRecurrentAppointment(myExService);
            //FindAppointmentsByDate(myExService);
            //FindAppointmentsByUser(myExService);
            //FindAppointmentsBySubject(myExService);
            //FindRecurrentAppointmentsByDate(myExService);
            //UpdateOneAppointment(myExService);
            //UpdateOneRecurrentAppointment(myExService);
            //AcceptDeclineOneAppointment(myExService);
            //ForwardOneAppointment(myExService);
            //TrackResponsesOneAppointment(myExService);
            //DeleteOneAppointment(myExService);
            //DeleteOneRecurrentAppointment(myExService);
            //ExportOneAppointment(myExService);
            //ImportOneAppointment(myExService);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void CreateAppointment(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            DateTime myDt = DateTime.Now.AddDays(1);
            Appointment newAppointment = new Appointment(ExService)
            {
                Subject = "This is a new meeting",
                Body = "To Whom It May Concern",
                Start = new DateTime(myDt.Year, myDt.Month, myDt.Day,
                                     myDt.Hour, myDt.Minute, myDt.Second),
                Location = "Somewhere"
            };
            newAppointment.End = newAppointment.Start.AddHours(1);
            newAppointment.RequiredAttendees.Add("user1@domain.com");
            newAppointment.OptionalAttendees.Add("user2@domain.com");

            newAppointment.Save(SendInvitationsMode.SendToNone);
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void CreateRecurrentAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            // Fixed date recurrent appointment
            DateTime myDt = DateTime.Now.AddDays(2);
            Appointment newAppointment = new Appointment(ExService)
            {
                Subject = "Monthly Meeting (second Thursday)",
                Body = "To Whom It May Concern monthly",
                Start = new DateTime(myDt.Year, myDt.Month, myDt.Day,
                                     myDt.Hour, myDt.Minute, myDt.Second),
                Location = "Somewhere"
            };
            newAppointment.End = newAppointment.Start.AddHours(1);
            newAppointment.RequiredAttendees.Add("user@domain.com");

            newAppointment.Recurrence = new Recurrence.MonthlyPattern(
                                    newAppointment.Start.Date, 1, 20)
            {
                StartDate = newAppointment.Start.Date,
                EndDate = new DateTime(2019, 12, 31)
            };

            newAppointment.Save(SendInvitationsMode.SendToNone);

            // Relative date recurrent appointment. 
            DateTime myRDt = DateTime.Now.AddDays(3);
            Appointment newRelAppointment = new Appointment(ExService)
            {
                Subject = "Bi-Monthly Appointment (last Friday)",
                Body = "To Whom It May Concern bi-monthly",
                Start = new DateTime(myRDt.Year, myRDt.Month, myRDt.Day,
                                     myRDt.Hour, myRDt.Minute, myRDt.Second),
                Location = "Somewhere"
            };
            newRelAppointment.End = newRelAppointment.Start.AddHours(1);

            newRelAppointment.Recurrence = new Recurrence.RelativeMonthlyPattern(
                                    newRelAppointment.Start.Date, 2,
                                    DayOfTheWeek.Friday, DayOfTheWeekIndex.First)
            {
                StartDate = newRelAppointment.Start.Date,
                EndDate = new DateTime(2019, 12, 31)
            };

            newRelAppointment.Save(SendInvitationsMode.SendToNone);
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void FindAppointmentsByDate(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            FindItemsResults<Appointment> allAppointments =
                ExService.FindAppointments(WellKnownFolderName.Calendar,
                    new CalendarView(DateTime.Now, DateTime.Now.AddDays(7)));

            foreach (Appointment oneAppointment in allAppointments)
            {
                Console.WriteLine("Subject: " + oneAppointment.Subject);
                Console.WriteLine("Start: " + oneAppointment.Start);
                Console.WriteLine("Duration: " + oneAppointment.Duration);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        private static void FindAppointmentsByUser(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            List<AttendeeInfo> accountsToScan = new List<AttendeeInfo>
            {
                new AttendeeInfo()
                {
                    SmtpAddress = "user@domain.com",
                    AttendeeType = MeetingAttendeeType.Organizer
                }
            };

            AvailabilityOptions myOptions = new AvailabilityOptions
            {
                MeetingDuration = 30,
                RequestedFreeBusyView = FreeBusyViewType.FreeBusy
            };

            GetUserAvailabilityResults scanResults = ExService.GetUserAvailability(
                                accountsToScan,
                                new TimeWindow(DateTime.Now, DateTime.Now.AddDays(1)),
                                    AvailabilityData.FreeBusy,
                                    myOptions);

            foreach (AttendeeAvailability accountAvailability in
                                                        scanResults.AttendeesAvailability)
            {
                foreach (CalendarEvent oneCalendarEvent in
                                                       accountAvailability.CalendarEvents)
                {
                    Console.WriteLine("Status: " + oneCalendarEvent.FreeBusyStatus);
                    Console.WriteLine("Start time: " + oneCalendarEvent.StartTime);
                    Console.WriteLine("End time: " + oneCalendarEvent.EndTime);
                }
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void FindRecurrentAppointmentsByDate(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter.SearchFilterCollection myFilter = 
                                                new SearchFilter.SearchFilterCollection
            {
                new SearchFilter.IsGreaterThanOrEqualTo(AppointmentSchema.Start, 
                                                        DateTime.Now)
            };

            ItemView myView = new ItemView(100)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, 
                                              AppointmentSchema.Subject, 
                                              AppointmentSchema.Start, 
                                              AppointmentSchema.Duration, 
                                              AppointmentSchema.AppointmentType)
            };

            FindItemsResults<Item> foundResults = ExService.FindItems(
                                    WellKnownFolderName.Calendar, myFilter, myView);

            foreach (Item oneItem in foundResults.Items)
            {
                Appointment oneAppointment = oneItem as Appointment;
                if (oneAppointment.AppointmentType == AppointmentType.RecurringMaster)
                {
                    Console.WriteLine("Subject: " + oneAppointment.Subject);
                    Console.WriteLine("Start: " + oneAppointment.Start);
                    Console.WriteLine("Duration: " + oneAppointment.Duration);
                }
            }
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void FindAppointmentsBySubject(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService, 
                                                        myAppointmentId, myPropSet);

            Console.WriteLine(myAppointment.Subject + " - " + myAppointment.Location);
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void UpdateOneAppointment(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService, 
                                                            myAppointmentId, myPropSet);

            myAppointment.Location = "Other place";
            myAppointment.RequiredAttendees.Add("user2@domain.com");

            myAppointment.Update(ConflictResolutionMode.AlwaysOverwrite, 
                                SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy);
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void UpdateOneRecurrentAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                   "Monthly Meeting (second Thursday)"));

            ItemView myView = new ItemView(1)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly,
                                              AppointmentSchema.AppointmentType)
            };

            FindItemsResults<Item> foundResults = ExService.FindItems(
                                    WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in foundResults.Items)
            {
                Appointment oneAppointment = oneItem as Appointment;
                if (oneAppointment.AppointmentType == AppointmentType.RecurringMaster)
                {
                    myAppointmentId = oneAppointment.Id;
                }
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            myAppointment.Location = "Other place";
            myAppointment.RequiredAttendees.Add("user2@domain.com");

            myAppointment.Update(ConflictResolutionMode.AlwaysOverwrite,
                                SendInvitationsOrCancellationsMode.SendToNone);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void AcceptDeclineOneAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            myAppointment.Accept(true);
            //myAppointment.AcceptTentatively(true);
            //myAppointment.Decline(true);

            AcceptMeetingInvitationMessage responseMessage = 
                                                myAppointment.CreateAcceptMessage(true);
            responseMessage.Body = new MessageBody("Meeting acknowledged");
            responseMessage.Sensitivity = Sensitivity.Private;
            responseMessage.Send();
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void ForwardOneAppointment(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            // Using the Forward method
            EmailAddress[] fwdAccounts = new EmailAddress[1];
            fwdAccounts[0] = new EmailAddress("user1@domain.com");
            myAppointment.Forward("Forwarding meeting invitation", fwdAccounts);

            // Using the CreateForward method
            ResponseMessage fwdMessage = myAppointment.CreateForward();
            fwdMessage.ToRecipients.Add("user1@domain.com");
            fwdMessage.BodyPrefix = "Forwarding meeting invitation";
            fwdMessage.IsDeliveryReceiptRequested = true;
            fwdMessage.SendAndSaveCopy();
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void TrackResponsesOneAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            for (int i = 0; i < myAppointment.RequiredAttendees.Count; i++)
            {
                Console.WriteLine(myAppointment.RequiredAttendees[i].Address + " - " + 
                    myAppointment.RequiredAttendees[i].ResponseType.Value.ToString());
            }

            for (int i = 0; i < myAppointment.OptionalAttendees.Count; i++)
            {
                Console.WriteLine(myAppointment.OptionalAttendees[i].Address + " - " + 
                    myAppointment.OptionalAttendees[i].ResponseType.Value.ToString());
            }

            for (int i = 0; i < myAppointment.Resources.Count; i++)
            {
                Console.WriteLine(myAppointment.Resources[i].Address + " - " + 
                    myAppointment.Resources[i].ResponseType.Value.ToString());
            }
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void DeleteOneAppointment(ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                            "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            // Using the Delete method (use only one of the code lines)
            myAppointment.Delete(DeleteMode.MoveToDeletedItems);
            myAppointment.Delete(DeleteMode.MoveToDeletedItems,
                                                SendCancellationsMode.SendOnlyToAll);

            // Using the Cancel method
            myAppointment.CancelMeeting();
            myAppointment.CancelMeeting("Meeting canceled");

            CancelMeetingMessage cancelMessage = myAppointment.CreateCancelMeetingMessage();
            cancelMessage.Body = new MessageBody("Meeting canceled");
            cancelMessage.IsReadReceiptRequested = true;
            cancelMessage.SendAndSaveCopy();
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void DeleteOneRecurrentAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                                    LogicalOperator.And,
                                                    new SearchFilter.IsEqualTo(
                                                            AppointmentSchema.Subject,
                                                   "Monthly Meeting (second Thursday)"));

            ItemView myView = new ItemView(1)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly,
                                              AppointmentSchema.AppointmentType)
            };

            FindItemsResults<Item> foundResults = ExService.FindItems(
                                    WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in foundResults.Items)
            {
                Appointment oneAppointment = oneItem as Appointment;
                if (oneAppointment.AppointmentType == AppointmentType.RecurringMaster)
                {
                    myAppointmentId = oneAppointment.Id;
                }
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                        AppointmentSchema.Subject, AppointmentSchema.Location);
            Appointment myAppointment = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            // Using the Delete method (use only one of the code lines)
            myAppointment.Delete(DeleteMode.MoveToDeletedItems);
            myAppointment.Delete(DeleteMode.MoveToDeletedItems,
                                                SendCancellationsMode.SendOnlyToAll);

            // Using the Cancel method
            myAppointment.CancelMeeting();
            myAppointment.CancelMeeting("Meeting canceled");

            CancelMeetingMessage cancelMessage = myAppointment.CreateCancelMeetingMessage();
            cancelMessage.Body = new MessageBody("Meeting canceled");
            cancelMessage.IsReadReceiptRequested = true;
            cancelMessage.SendAndSaveCopy();
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        private static void ExportOneAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            SearchFilter myFilter = new SearchFilter.SearchFilterCollection(
                                        LogicalOperator.And,
                                        new SearchFilter.IsEqualTo(
                                                AppointmentSchema.Subject,
                                                "This is a new meeting"));
            ItemView myView = new ItemView(1);
            FindItemsResults<Item> findResults = ExService.FindItems(
                                        WellKnownFolderName.Calendar, myFilter, myView);

            ItemId myAppointmentId = null;
            foreach (Item oneItem in findResults)
            {
                myAppointmentId = oneItem.Id;
            }

            PropertySet myPropSet = new PropertySet(BasePropertySet.IdOnly,
                                                    AppointmentSchema.MimeContent);
            Appointment appointmentToExport = Appointment.Bind(ExService,
                                                            myAppointmentId, myPropSet);

            string apmFileName = @"C:\Temporary\myAppointment.ics";
            using (FileStream myStream = new FileStream(
                                        apmFileName, FileMode.Create, FileAccess.Write))
            {
                myStream.Write(appointmentToExport.MimeContent.Content, 0,
                                        appointmentToExport.MimeContent.Content.Length);
            }
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        private static void ImportOneAppointment(
                                        ExchangeService ExService)  //*** LEGACY CODE ***
        {
            Appointment appointmentToImport = new Appointment(ExService);

            string iCalFileName = @"C:\Temporary\myAppointment.ics";
            using (FileStream myStream = new FileStream(
                                        iCalFileName, FileMode.Open, FileAccess.Read))
            {
                byte[] appBytes = new byte[myStream.Length];
                int bytesToRead = (int)myStream.Length;
                int bytesRead = 0;
                while (bytesToRead > 0)
                {
                    int myBlock = myStream.Read(appBytes, bytesRead, bytesToRead);
                    if (myBlock == 0)
                        break;
                    bytesRead += myBlock;
                    bytesToRead -= myBlock;
                }
                appointmentToImport.MimeContent = new MimeContent("UTF-8", appBytes);
            }

            appointmentToImport.Save(WellKnownFolderName.Calendar);
        }
        //gavdcodeend 15

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
