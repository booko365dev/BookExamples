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

            //SpCsPnpcoreCreateOneList(psCtx);
            //SpCsPnpcoreReadOneList(psCtx);
            //SpCsPnpcoreListExists(psCtx);
            //SpCsPnpcoreAddUserToSecurityRoleInList(psCtx);
            //SpCsPnpcoreAddOneFieldToList(psCtx);
            //SpCsPnpcoreReadFilteredFieldsFromList(psCtx);
            //SpCsPnpcoreReadOneFieldFromList(psCtx);
            //SpCsPnpcoreGetContentTypeList(psCtx);
            //SpCsPnpcoreAddContentTypeToList(psCtx);
            //SpCsPnpcoreRemoveContentTypeFromList(psCtx);
            //SpCsPnpcoreGetViewList(psCtx);
            //SpCsPnpcoreAddViewToList(psCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void SpCsPnpcoreCreateOneList(ClientContext spCtx)
        {
            ListTemplateType myTemplate = ListTemplateType.GenericList;
            string listName = "NewListPnPCore";
            bool enableVersioning = false;
            List newList = spCtx.Web.CreateList(myTemplate, listName, enableVersioning);
        }

        static void SpCsPnpcoreReadOneList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            Console.WriteLine("List title - " + myList.Title);
        }

        static void SpCsPnpcoreListExists(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            bool blnListExists = myWeb.ListExists("NewListPnPCore");

            Console.WriteLine("List exists - " + blnListExists);
        }

        static void SpCsPnpcoreAddUserToSecurityRoleInList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.GetListByTitle("NewListPnPCore");

            myList.SetListPermission(BuiltInIdentity.Everyone, RoleType.Editor);
        }

        static void SpCsPnpcoreAddOneFieldToList(ClientContext spCtx)
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

        static void SpCsPnpcoreReadFilteredFieldsFromList(ClientContext spCtx)
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

        static void SpCsPnpcoreReadOneFieldFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");

            Field myField = myList.GetFieldById
                    (new Guid("b0b75b9d-b358-49e6-b7fe-b2e35295f4bc"));

            Console.WriteLine(myField.InternalName + " - " + myField.TypeAsString);
        }

        static void SpCsPnpcoreGetContentTypeList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            ContentType myContentType = myList.GetContentTypeByName("Item");

            Console.WriteLine(myContentType.Description);
        }

        static void SpCsPnpcoreAddContentTypeToList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.AddContentTypeToListByName("Comment");
        }

        static void SpCsPnpcoreRemoveContentTypeFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.RemoveContentTypeByName("Comment");
        }

        static void SpCsPnpcoreGetViewList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            View myView = myList.GetViewByName("All Items");

            Console.WriteLine(myView.ListViewXml);
        }

        static void SpCsPnpcoreAddViewToList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListPnPCore");
            myList.CreateView("NewView", ViewType.Html, null, 30, false);
        }

        //----------------------------------------------------------------------------------------
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

