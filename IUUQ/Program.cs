using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.IO;
using System.Security;

namespace IUUQ
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginCsom();

            //SpCsCsomCreateOneItem(spCtx);
            //SpCsCsomUploadOneDoc(spCtx);
            //SpCsCsomUploadOneDocumentFileCrInfo(spCtx);
            //SpCsCsomDownloadOneDoc(spCtx);
            //SpCsCsomReadAllListItems(spCtx);
            //SpCsCsomReadOneListItem(spCtx);
            //SpCsCsomUpdateOneListItem(spCtx);
            //SpCsCsomDeleteOneListItem(spCtx);
            //SpCsCsomReadAllLibraryDocs(spCtx);
            //SpCsCsomReadOneLibraryDoc(spCtx);
            //SpCsCsomUpdateOneLibraryDoc(spCtx);
            //SpCsCsomDeleteOneLibraryDoc(spCtx);
            //SpCsCsomCreateMultipleItem(spCtx);
            //SpCsCsomUploadMultipleDocs(spCtx);
            //SpCsCsomDownloadMultipleDocs(spCtx);
            //SpCsCsomDeleteAllListItems(spCtx);
            //SpCsCsomDeleteAllLibraryDocs(spCtx);
            //SpCsCsomBreakSecurityInheritanceListItem(spCtx);
            //SpCsCsomResetSecurityInheritanceListItem(spCtx);
            //SpCsCsomAddUserToSecurityRoleInListItem(spCtx);
            //SpCsCsomUpdateUserSecurityRoleInListItem(spCtx);
            //SpCsCsomDeleteUserFromSecurityRoleInListItem(spCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void SpCsCsomCreateOneItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            ListItemCreationInformation myListItemCreationInfo =
                                            new ListItemCreationInformation();
            ListItem newListItem = myList.AddItem(myListItemCreationInfo);
            newListItem["Title"] = "NewListItemCsCsom";

            newListItem.Update();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomCreateMultipleItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            for (int intCounter = 0; intCounter < 4; intCounter++)
            {
                ListItemCreationInformation myListItemCreationInfo =
                                                new ListItemCreationInformation();
                ListItem newListItem = myList.AddItem(myListItemCreationInfo);
                newListItem["Title"] = intCounter.ToString() + "-NewListItemCsCsom";
                newListItem.Update();
            }

            spCtx.ExecuteQuery();
        }

        static void SpCsCsomUploadOneDocument(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

            string filePath = @"C:\Temporary\";
            string fileName = @"TestDocument.docx";

            using (FileStream myFileStream = new
                                    FileStream(filePath + fileName, FileMode.Open))
            {
                FileInfo myFileInfo = new FileInfo(fileName);
                spCtx.Load(myList.RootFolder);
                spCtx.ExecuteQuery();

                string fileUrl = String.Format("{0}/{1}",
                                myList.RootFolder.ServerRelativeUrl, myFileInfo.Name);
                Microsoft.SharePoint.Client.File.
                                SaveBinaryDirect(spCtx, fileUrl, myFileStream, true);
            }
        }

        static void SpCsCsomUploadOneDocumentFileCrInfo(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

            string filePath = @"C:\Temporary\";
            string fileName = @"TestDocument01.docx";

            using (FileStream myFileStream = new
                                    FileStream(filePath + fileName, FileMode.Open))
            {
                spCtx.Load(myList.RootFolder);
                spCtx.ExecuteQuery();

                FileCreationInformation myFileCreationInfo = new FileCreationInformation
                {
                    Overwrite = true,
                    ContentStream = myFileStream,
                    Url = fileName
                };

                Microsoft.SharePoint.Client.File newFile = 
                                        myList.RootFolder.Files.Add(myFileCreationInfo);
	            spCtx.Load(newFile);
                spCtx.ExecuteQuery();
            }
        }

        static void SpCsCsomUploadMultipleDocs(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

            string filesPath = @"C:\Temporary\";

            string[] myFiles = Directory.GetFiles(filesPath);

            foreach (string oneFile in myFiles)
            {
                using (FileStream myFileStream = new
                                        FileStream(oneFile, FileMode.Open))
                {
                    FileInfo myFileInfo = new FileInfo(oneFile.Replace(filesPath, ""));
                    spCtx.Load(myList.RootFolder);
                    spCtx.ExecuteQuery();

                    string fileUrl = String.Format("{0}/{1}",
                                    myList.RootFolder.ServerRelativeUrl, myFileInfo.Name);
                    Microsoft.SharePoint.Client.File.
                                    SaveBinaryDirect(spCtx, fileUrl, myFileStream, true);
                }
            }
        }

        static void SpCsCsomDownloadOneDoc(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

            string filePath = @"C:\Temporary\";

            int listItemId = 1;
            ListItem myListItem = myList.GetItemById(listItemId);
            spCtx.Load(myListItem);
            spCtx.Load(myListItem, itm => itm.File);
            spCtx.ExecuteQuery();

            string fileRef = myListItem.File.ServerRelativeUrl;
            FileInformation myFileInfo = Microsoft.SharePoint.Client.File.
                                                    OpenBinaryDirect(spCtx, fileRef);
            string fileName = Path.Combine(filePath, (string)myListItem.File.Name);
            using (FileStream myFileStream = System.IO.File.Create(fileName))
            {
                myFileInfo.Stream.CopyTo(myFileStream);
            }
        }

        static void SpCsCsomDownloadMultipleDocs(ClientContext spCtx)
        {
            string filePath = @"C:\Temporary\";

            FileCollection myFiles = spCtx.Web.GetFolderByServerRelativeUrl(
                                                                 "TestLibrary").Files;
            spCtx.Load(myFiles);
            spCtx.ExecuteQuery();

            foreach (Microsoft.SharePoint.Client.File oneFile in myFiles)
            {
                string fileRef = oneFile.ServerRelativeUrl;
                FileInformation myFileInfo = Microsoft.SharePoint.Client.File.
                                                        OpenBinaryDirect(spCtx, fileRef);
                string fileName = Path.Combine(filePath, (string)oneFile.Name);
                using (FileStream myFileStream = System.IO.File.Create(fileName))
                {
                    myFileInfo.Stream.CopyTo(myFileStream);
                }
            }
        }

        static void SpCsCsomReadAllListItems(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItemCollection allItems = myList.GetItems(CamlQuery.CreateAllItemsQuery());
            spCtx.Load(allItems, itms => itms.Include(itm => itm["Title"],
                                                     itm => itm.Id));
            spCtx.ExecuteQuery();

            foreach (ListItem oneItem in allItems)
            {
                Console.WriteLine(oneItem["Title"] + " - " + oneItem.Id);
            }
        }

        static void SpCsCsomReadAllLibraryDocs(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
            ListItemCollection allItems = myList.GetItems(CamlQuery.CreateAllItemsQuery());
            spCtx.Load(allItems, itms => itms.Include(itm => itm["FileLeafRef"],
                                                     itm => itm.Id));
            spCtx.ExecuteQuery();

            foreach (ListItem oneItem in allItems)
            {
                Console.WriteLine(oneItem["FileLeafRef"] + " - " + oneItem.Id);
            }
        }

        static void SpCsCsomReadOneListItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            int filterField = 1;
            int rowLimit = 10;
            string myViewXml = string.Format(@"
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='ID' />
                                <Value Type='Number'>{0}</Value>
                            </Eq>
                        </Where>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='Title' />
                    </ViewFields>
                    <RowLimit>{1}</RowLimit>
                </View>", filterField, rowLimit);

            CamlQuery myCamlQuery = new CamlQuery();
            myCamlQuery.ViewXml = myViewXml;
            ListItemCollection allItems = myList.GetItems(myCamlQuery);
            spCtx.Load(allItems, itms => itms.Include(itm => itm["Title"],
                                                     itm => itm.Id));
            spCtx.ExecuteQuery();

            Console.WriteLine("Item Title - " + allItems[0]["Title"]);
        }

        static void SpCsCsomReadOneLibraryDoc(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");

            int filterField = 1;
            int rowLimit = 10;
            string myViewXml = string.Format(@"
                <View>
                    <Query>
                        <Where>
                            <Eq>
                                <FieldRef Name='ID' />
                                <Value Type='Number'>{0}</Value>
                            </Eq>
                        </Where>
                    </Query>
                    <ViewFields>
                        <FieldRef Name='FileLeafRef' />
                    </ViewFields>
                    <RowLimit>{1}</RowLimit>
                </View>", filterField, rowLimit);

            CamlQuery myCamlQuery = new CamlQuery();
            myCamlQuery.ViewXml = myViewXml;
            ListItemCollection allItems = myList.GetItems(myCamlQuery);
            spCtx.Load(allItems, itms => itms.Include(itm => itm["FileLeafRef"],
                                                     itm => itm.Id));
            spCtx.ExecuteQuery();

            Console.WriteLine("Item Title - " + allItems[0]["FileLeafRef"]);
        }

        static void SpCsCsomUpdateOneListItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(43);
            myListItem["Title"] = "NewListItemCsCsomUpdated";

            myListItem.Update();
            spCtx.Load(myListItem);
            spCtx.ExecuteQuery();

            Console.WriteLine("Item Title - " + myListItem["Title"]);
        }

        static void SpCsCsomUpdateOneLibraryDoc(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
            ListItem myListItem = myList.GetItemById(1);
            myListItem["FileLeafRef"] = "LibraryDocCsCsomUpdated.docx";

            myListItem.Update();
            spCtx.Load(myListItem);
            spCtx.ExecuteQuery();

            Console.WriteLine("Item Title - " + myListItem["FileLeafRef"]);
        }

        static void SpCsCsomDeleteOneListItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(1);
            myListItem.DeleteObject();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomDeleteAllListItems(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItemCollection myListItems = myList.GetItems(
                                                    CamlQuery.CreateAllItemsQuery());
            spCtx.Load(myListItems);
            spCtx.ExecuteQuery();

            foreach (ListItem oneItem in myListItems)
            {
                ListItem oneItemToDelete = myList.GetItemById(oneItem.Id);
                oneItemToDelete.DeleteObject();
            }

            spCtx.ExecuteQuery();
        }

        static void SpCsCsomDeleteOneLibraryDoc(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
            ListItem myListItem = myList.GetItemById(1);
            myListItem.DeleteObject();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomDeleteAllLibraryDocs(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestLibrary");
            ListItemCollection myListItems = myList.GetItems(
                                                    CamlQuery.CreateAllItemsQuery());
            spCtx.Load(myListItems);
            spCtx.ExecuteQuery();

            foreach (ListItem oneItem in myListItems)
            {
                ListItem oneItemToDelete = myList.GetItemById(oneItem.Id);
                oneItemToDelete.DeleteObject();
            }

            spCtx.ExecuteQuery();
        }

        static void SpCsCsomBreakSecurityInheritanceListItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(1);
            spCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myListItem.HasUniqueRoleAssignments == false)
            {
                myListItem.BreakRoleInheritance(false, true);
            }
            myListItem.Update();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomResetSecurityInheritanceListItem(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(1);
            spCtx.Load(myListItem, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myListItem.HasUniqueRoleAssignments == true)
            {
                myListItem.ResetRoleInheritance();
            }
            myListItem.Update();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomAddUserToSecurityRoleInListItem(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(43);

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
            myListItem.RoleAssignments.Add(myUser, roleDefinition);

            spCtx.ExecuteQuery();
        }

        static void SpCsCsomUpdateUserSecurityRoleInListItem(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(1);

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Administrator));

            RoleAssignment myRoleAssignment = myListItem.RoleAssignments.GetByPrincipal(
                                                                                myUser);
            myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

            myRoleAssignment.Update();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomDeleteUserFromSecurityRoleInListItem(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("TestList");
            ListItem myListItem = myList.GetItemById(1);

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            myListItem.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

            spCtx.ExecuteQuery();
            spCtx.Dispose();
        }

        //-------------------------------------------------------------------------------
        static ClientContext LoginCsom()
        {
            ClientContext rtnContext = new ClientContext(
                ConfigurationManager.AppSettings["spUrl"]);

            SecureString securePw = new SecureString();
            foreach (
                char oneChar in ConfigurationManager.AppSettings["spUserPw"].ToCharArray())
            {
                securePw.AppendChar(oneChar);
            }
            rtnContext.Credentials = new SharePointOnlineCredentials(
                ConfigurationManager.AppSettings["spUserName"], securePw);

            return rtnContext;
        }
    }
}

