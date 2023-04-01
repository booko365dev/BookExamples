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

        //gavdcodebegin 001
        //*** LEGACY CODE ***
        static void SpCsPnpcore_CreateOneList(ClientContext spCtx)
        {
            ListTemplateType myTemplate = ListTemplateType.GenericList;
            string listName = "NewListPnPCore";
            bool enableVersioning = false;
            List newList = spCtx.Web.CreateList(myTemplate, listName, enableVersioning);
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        //*** LEGACY CODE ***
        static void SpCsPnpcore_ReadOneList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            Console.WriteLine("List title - " + myList.Title);
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        //*** LEGACY CODE ***
        static void SpCsPnpcore_ListExists(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            bool blnListExists = myWeb.ListExists("NewListPnPCore");

            Console.WriteLine("List exists - " + blnListExists);
        }
        //gavdcodeend 003

        //gavdcodebegin 004
        //*** LEGACY CODE ***
        static void SpCsPnpcore_AddUserToSecurityRoleInList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            myList.SetListPermission(BuiltInIdentity.Everyone, RoleType.Editor);
        }
        //gavdcodeend 004

        //gavdcodebegin 005
        //*** LEGACY CODE ***
        static void SpCsPnpcore_AddOneFieldToList(ClientContext spCtx)
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
        //gavdcodeend 005

        //gavdcodebegin 006
        //*** LEGACY CODE ***
        static void SpCsPnpcore_ReadFilteredFieldsFromList(ClientContext spCtx)
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
        //gavdcodeend 006

        //gavdcodebegin 007
        //*** LEGACY CODE ***
        static void SpCsPnpcore_ReadOneFieldFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");

            Field myField = myList.GetFieldById
                    (new Guid("b0b75b9d-b358-49e6-b7fe-b2e35295f4bc"));

            Console.WriteLine(myField.InternalName + " - " + myField.TypeAsString);
        }
        //gavdcodeend 007

        //gavdcodebegin 008
        //*** LEGACY CODE ***
        static void SpCsPnpcore_GetContentTypeList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            ContentType myContentType = myList.GetContentTypeByName("Item");

            Console.WriteLine(myContentType.Description);
        }
        //gavdcodeend 008

        //gavdcodebegin 009
        //*** LEGACY CODE ***
        static void SpCsPnpcore_AddContentTypeToList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.AddContentTypeToListByName("Comment");
        }
        //gavdcodeend 009

        //gavdcodebegin 010
        //*** LEGACY CODE ***
        static void SpCsPnpcore_RemoveContentTypeFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.RemoveContentTypeByName("Comment");
        }
        //gavdcodeend 010

        //gavdcodebegin 011
        //*** LEGACY CODE ***
        static void SpCsPnpcore_GetViewList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            View myView = myList.GetViewByName("All Items");

            Console.WriteLine(myView.ListViewXml);
        }
        //gavdcodeend 011

        //gavdcodebegin 012
        //*** LEGACY CODE ***
        static void SpCsPnpcore_AddViewToList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.CreateView("NewView", ViewType.Html, null, 30, false);
        }
        //gavdcodeend 012

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
