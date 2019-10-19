using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;

namespace RCAQ
{
    class Program
    {
        static void Main(string[] args)
        {
            ExchangeService myExService = ConnectBA(
                                ConfigurationManager.AppSettings["exUserName"],
                                ConfigurationManager.AppSettings["exUserPw"]);

            //GetFolders(myExService);
            //GetOneRootFolder(myExService);
            //FindOneFolder(myExService);
            //CreateOneFolder(myExService);
            //CopyOneFolder(myExService);
            //MoveOneFolder(myExService);
            //UpdateOneFolder(myExService);
            //EmptyOneFolder(myExService);
            //HideOneFolder(myExService);
            //DeleteOneFolder(myExService);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void GetFolders(ExchangeService ExService)
        {
            FolderView myView = new FolderView(100);
            ExtendedPropertyDefinition isHidden = new 
                            ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
            myView.PropertySet = new PropertySet(BasePropertySet.IdOnly, 
                                                    FolderSchema.DisplayName, isHidden);
            myView.Traversal = FolderTraversal.Deep;
            FindFoldersResults allFolders = ExService.FindFolders(
                                            WellKnownFolderName.MsgFolderRoot, myView);

            foreach (Folder oneFolder in allFolders)
            {
                string strHidden = oneFolder.ExtendedProperties[0].Value.ToString();
                Console.WriteLine(oneFolder.DisplayName + " - Hidden: " + strHidden);
            }
        }

        static void GetOneRootFolder(ExchangeService ExService)
        {
            Folder myInboxFolder = Folder.Bind(ExService, WellKnownFolderName.Inbox);

            Console.WriteLine(myInboxFolder.DisplayName + " - " + 
                                    myInboxFolder.ChildFolderCount + " child folder");

            myInboxFolder.Load();
            foreach (Folder oneFolder in myInboxFolder.FindFolders(new FolderView(100)))
            {
                Console.WriteLine("-- " + oneFolder.DisplayName + " - Id: " + oneFolder.Id);
            }
        }

        static void FindOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Inbox);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "my custom folder",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }
        }

        static void CreateOneFolder(ExchangeService ExService)
        {
            Folder newFolder = new Folder(ExService)
            {
                DisplayName = "My Custom Folder"
                //newFolder.FolderClass = "IPF.MyCustomFolderClass";
            };

            newFolder.Save(WellKnownFolderName.Inbox);
        }

        static void CopyOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Inbox);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter = 
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "my custom folder", 
                        ContainmentMode.Substring, 
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            Folder folderToCopy = Folder.Bind(ExService, myFolderId);
            folderToCopy.Copy(WellKnownFolderName.JunkEmail);
        }

        static void MoveOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Inbox);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "my custom folder",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            Folder folderToMove = Folder.Bind(ExService, myFolderId);
            folderToMove.Move(WellKnownFolderName.Drafts);
        }

        static void UpdateOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Drafts);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "my custom folder",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            Folder folderToUpdate = Folder.Bind(ExService, myFolderId);
            folderToUpdate.DisplayName = "New Folder Name";
            folderToUpdate.Update();
        }

        static void EmptyOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Drafts);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "new folder name",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            Folder folderToEmpty = Folder.Bind(ExService, myFolderId);
            folderToEmpty.Empty(DeleteMode.HardDelete, true);
        }

        static void HideOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.JunkEmail);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "my custom folder",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            ExtendedPropertyDefinition isHiddenProp = new 
                        ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
            PropertySet propSet = new PropertySet(isHiddenProp);

            Folder folderToHide = Folder.Bind(ExService, myFolderId, propSet);
            folderToHide.SetExtendedProperty(isHiddenProp, true);
            folderToHide.Update();
        }

        static void DeleteOneFolder(ExchangeService ExService)
        {
            Folder rootFolder = Folder.Bind(ExService, WellKnownFolderName.Drafts);
            rootFolder.Load();
            SearchFilter.ContainsSubstring subjectFilter =
                    new SearchFilter.ContainsSubstring(
                        FolderSchema.DisplayName,
                        "new folder name",
                        ContainmentMode.Substring,
                        ComparisonMode.IgnoreCase);

            FolderId myFolderId = null;
            foreach (Folder oneFolder in rootFolder.FindFolders(
                                                    subjectFilter, new FolderView(1)))
            {
                myFolderId = oneFolder.Id;
            }

            Folder folderToDelete = Folder.Bind(ExService, myFolderId);
            folderToDelete.Delete(DeleteMode.HardDelete);
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

