using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;

namespace ZOUX
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["exUserName"],
                                ConfigurationManager.AppSettings["exUserPw"]);

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

        static void CreateAppointment(ExchangeService ExService)
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

        static void CreateRecurrentAppointment(ExchangeService ExService)
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

        static void FindAppointmentsByDate(ExchangeService ExService)
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

        private static void FindAppointmentsByUser(ExchangeService ExService)
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

        static void FindRecurrentAppointmentsByDate(ExchangeService ExService)
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

        static void FindAppointmentsBySubject(ExchangeService ExService)
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

        static void UpdateOneAppointment(ExchangeService ExService)
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

        static void UpdateOneRecurrentAppointment(ExchangeService ExService)
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

        static void AcceptDeclineOneAppointment(ExchangeService ExService)
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

        static void ForwardOneAppointment(ExchangeService ExService)
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

        static void TrackResponsesOneAppointment(ExchangeService ExService)
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

        static void DeleteOneAppointment(ExchangeService ExService)
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

        static void DeleteOneRecurrentAppointment(ExchangeService ExService)
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

        private static void ExportOneAppointment(ExchangeService ExService)
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

        private static void ImportOneAppointment(ExchangeService ExService)
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

