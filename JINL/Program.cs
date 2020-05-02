using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace JINL
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext psCtx = LoginPnPCore();

            //SpCsPnpcoreCreatePropertyBag(psCtx);
            //SpCsPnpcoreReadPropertyBag(psCtx);
            //SpCsPnpcorePropertyBagExists(psCtx);
            //SpCsPnpcorePropertyBagIndex(psCtx);
            //SpCsPnpcoreDeletePropertyBag(psCtx);
            //SpCsPnpcoreDownloadFile(psCtx);
            //SpCsPnpcoreDownloadFileAsString(psCtx);
            //SpCsPnpcoreFindFiles(psCtx);
            //SpCsPnpcoreRequireUpload(psCtx);
            //SpCsPnpcoreResetVersion(psCtx);
            //SpCsPnpcoreCreateFolder(psCtx);
            //SpCsPnpcoreEnsureFolder(psCtx);
            //SpCsPnpcoreCreateSubFolder(psCtx);
            //SpCsPnpcoreFolderExistsBool(psCtx);
            //SpCsPnpcoreFolderExistsFolder(psCtx);
            //SpCsPnpcoreFolderExistsWeb(psCtx);
            //SpCsPnpcoreCreateSubFolderFromFolder(psCtx);
            //SpCsPnpcoreUploadFileToFolder(psCtx);
            //SpCsPnpcoreDownloadFileToFolder(psCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsPnpcoreCreatePropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.SetPropertyBagValue("myKey", "myValueString");
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsPnpcoreReadPropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            string myKeyValue = myList.GetPropertyBagValueString("myKey", "");
            Console.WriteLine(myKeyValue);
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsPnpcorePropertyBagExists(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            bool myKeyExists = myList.PropertyBagContainsKey("myKey");
            Console.WriteLine(myKeyExists.ToString());
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsPnpcorePropertyBagIndex(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.AddIndexedPropertyBagKey("myKey");
            IEnumerable<string> myIndexedPropertyBagKeys =
                                                myList.GetIndexedPropertyBagKeys();

            foreach (string oneKey in myIndexedPropertyBagKeys)
            {
                Console.WriteLine(oneKey);
            }
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsPnpcoreDeletePropertyBag(ClientContext spCtx)
        {
            List myList = spCtx.Web.Lists.GetByTitle("TestList");

            myList.RemovePropertyBagValue("myKey");
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsPnpcoreDownloadFile(ClientContext spCtx)
        {
            string pathRelative =
                    "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";
            string pathLocal = @"C:\Temporary";
            string fileName = "TestDocument01.docx";

            spCtx.Web.SaveFileToLocal(pathRelative, pathLocal, fileName);
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsPnpcoreDownloadFileAsString(ClientContext spCtx)
        {
            string pathRelative =
                    "/sites/Test_Guitaca/TestDocuments/FirstLevelFolderPS/TestDocument01.docx";

            string myFileAsText = spCtx.Web.GetFileAsString(pathRelative);

            Console.WriteLine(myFileAsText);
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsPnpcoreFindFiles(ClientContext spCtx)
        {
            List<File> allFiles = spCtx.Web.FindFiles("*.docx");

            foreach (File oneFile in allFiles)
            {
                Console.WriteLine(oneFile.Name + " - " + oneFile.ServerRelativeUrl);
            }
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsPnpcoreRequireUpload(ClientContext spCtx)
        {
            List<File> myFiles = spCtx.Web.FindFiles("TestDocument01.docx");
            File myFile = myFiles[0];

            bool requireUpload = myFile.VerifyIfUploadRequired(@"C:\Temporary\TestDocument01.docx");

            Console.WriteLine("Require Upload " + requireUpload.ToString());
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsPnpcoreResetVersion(ClientContext spCtx)
        {
            string pathRelative =
                    "/sites/[SiteName]/[LibraryName]/[FolderName]/[DocumentName].docx";

            spCtx.Web.ResetFileToPreviousVersion(pathRelative,
                                        CheckinType.MinorCheckIn, "Done by PnPCore");
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsPnpcoreCreateFolder(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.CreateFolder("PnPCoreFolder");
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsPnpcoreEnsureFolder(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.EnsureFolder("PnPCoreEnsureFolder");
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsPnpcoreCreateSubFolder(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.EnsureFolder("PnPCoreEnsureFolder");
            Folder mySubFolder = myFolder.EnsureFolder("PnPCoreSubFolder");
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void SpCsPnpcoreFolderExistsBool(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            bool fldExists = myList.RootFolder.FolderExists("PnPCoreEnsureFolder");
            Console.WriteLine("Folder exists - " + fldExists.ToString());
        }
        //gavdcodeend 14

        //gavdcodebegin 15
        static void SpCsPnpcoreFolderExistsFolder(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder fldExists = myList.RootFolder.ResolveSubFolder("PnPCoreEnsureFolder");
            Console.WriteLine("Folder exists - " + fldExists.ServerRelativeUrl);
        }
        //gavdcodeend 15

        //gavdcodebegin 16
        static void SpCsPnpcoreFolderExistsWeb(ClientContext spCtx)
        {
            string pathRelative =
                    "/sites/[SiteName]/[LibraryName]/[FolderName]";

            bool fldExists = spCtx.Site.RootWeb.DoesFolderExists(pathRelative);

            Console.WriteLine("Folder exists - " + fldExists.ToString());
        }
        //gavdcodeend 16

        //gavdcodebegin 17
        static void SpCsPnpcoreCreateSubFolderFromFolder(ClientContext spCtx)
        {
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.ResolveSubFolder("PnPCoreEnsureFolder");
            Folder mySubFolder = myFolder.CreateFolder("PnPCoreSubFolder02");
        }
        //gavdcodeend 17

        //gavdcodebegin 18
        static void SpCsPnpcoreUploadFileToFolder(ClientContext spCtx)
        {
            string pathLocal = @"C:\Temporary\TestDocument01.docx";
            string fileName = "TestDocument01.docx";
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.EnsureFolder("PnPCoreEnsureFolder");
            File myFile = myFolder.UploadFile(fileName, pathLocal, true);
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        static void SpCsPnpcoreDownloadFileToFolder(ClientContext spCtx)
        {
            string pathLocal = @"C:\Temporary\TestDocument01.docx";
            string spFileName = "TestDocument01.docx";
            List myList = spCtx.Site.RootWeb.GetListByTitle("TestDocuments");

            Folder myFolder = myList.RootFolder.EnsureFolder("PnPCoreEnsureFolder");
            File myFile = myFolder.GetFile(spFileName);

            ClientResult<System.IO.Stream> myStream = myFile.OpenBinaryStream();
            spCtx.ExecuteQueryRetry();
            using (System.IO.FileStream fileStream = new System.IO.FileStream(
                        pathLocal, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                myStream.Value.CopyTo(fileStream);
            }
        }
        //gavdcodeend 19

        //-------------------------------------------------------------------------------
        static ClientContext LoginPnPCore()
        {
            OfficeDevPnP.Core.AuthenticationManager pnpAuthMang =
                new OfficeDevPnP.Core.AuthenticationManager();
            ClientContext rtnContext =
                        pnpAuthMang.GetSharePointOnlineAuthenticatedContextTenant
                            (ConfigurationManager.AppSettings["spUrl"],
                             ConfigurationManager.AppSettings["spUserName"],
                             ConfigurationManager.AppSettings["spUserPw"]);

            return rtnContext;
        }
    }
}
