using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Enums;

namespace ZJLE
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext psCtx = LoginPnPCore();

            //SpCsPnpcore_CreateOneList(psCtx);
            //SpCsPnpcore_ReadOneList(psCtx);
            //SpCsPnpcore_ListExists(psCtx);
            //SpCsPnpcore_AddUserToSecurityRoleInList(psCtx);
            //SpCsPnpcore_AddOneFieldToList(psCtx);
            //SpCsPnpcore_ReadFilteredFieldsFromList(psCtx);
            //SpCsPnpcore_ReadOneFieldFromList(psCtx);
            //SpCsPnpcore_GetContentTypeList(psCtx);
            //SpCsPnpcore_AddContentTypeToList(psCtx);
            //SpCsPnpcore_RemoveContentTypeFromList(psCtx);
            //SpCsPnpcore_GetViewList(psCtx);
            //SpCsPnpcore_AddViewToList(psCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsPnpcore_CreateOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            ListTemplateType myTemplate = ListTemplateType.GenericList;
            string listName = "NewListPnPCore";
            bool enableVersioning = false;
            List newList = spCtx.Web.CreateList(myTemplate, listName, enableVersioning);
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsPnpcore_ReadOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            Console.WriteLine("List title - " + myList.Title);
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsPnpcore_ListExists(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            bool blnListExists = myWeb.ListExists("NewListPnPCore");

            Console.WriteLine("List exists - " + blnListExists);
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsPnpcore_AddUserToSecurityRoleInList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            myList.SetListPermission(BuiltInIdentity.Everyone, RoleType.Editor);
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsPnpcore_AddOneFieldToList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");

            FieldType fieldType = FieldType.Text;
            OfficeDevPnP.Core.Entities.FieldCreationInformation newFieldInfo = 
                new OfficeDevPnP.Core.Entities.FieldCreationInformation(fieldType)
            {
                DisplayName = "NewFieldPnPCoreUsingInfo",
                InternalName = "NewFieldPnPCoreInfo",
                Id = new Guid()
            };
            myList.CreateField(newFieldInfo);

            string fieldXml = "<Field DisplayName='NewFieldPnPCoreUsingXml' " +
                "Type='Note' Required='FALSE' Name='NewFieldPnPCoreXml' />";
            myList.CreateField(fieldXml);
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsPnpcore_ReadFilteredFieldsFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");

            string[] fieldsToFind = new string[] 
                    { "NewFieldPnPCoreXml", "NewFieldPnPCoreInfo" };
            IEnumerable<Field> allFields = myList.GetFields(fieldsToFind);

            foreach (Field oneField in allFields)
            {
                Console.WriteLine(oneField.Title + " - " + oneField.TypeAsString);
            }
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsPnpcore_ReadOneFieldFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");

            Field myField = myList.GetFieldById
                    (new Guid("b0b75b9d-b358-49e6-b7fe-b2e35295f4bc"));

            Console.WriteLine(myField.InternalName + " - " + myField.TypeAsString);
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsPnpcore_GetContentTypeList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            ContentType myContentType = myList.GetContentTypeByName("Item");

            Console.WriteLine(myContentType.Description);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsPnpcore_AddContentTypeToList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.AddContentTypeToListByName("Comment");
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsPnpcore_RemoveContentTypeFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.RemoveContentTypeByName("Comment");
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsPnpcore_GetViewList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            View myView = myList.GetViewByName("All Items");

            Console.WriteLine(myView.ListViewXml);
        }
        //gavdcodeend 11

        //gavdcodebegin 12
        static void SpCsPnpcore_AddViewToList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.CreateView("NewView", ViewType.Html, null, 30, false);
        }
        //gavdcodeend 12

        //----------------------------------------------------------------------------------------
        static ClientContext LoginPnPCore()  //*** LEGACY CODE ***
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
