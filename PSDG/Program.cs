using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Enums;
using System.Configuration;
using System.Security;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 6.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext LoginPnPFramework_WithAccPw()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithCertificate()
{
    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_PnPManagementShell()
{
    SecureString mySecurePw = new SecureString();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new
        AuthenticationManager(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext LoginPnPFramework_WithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query
    ClientContext rtnContext = new
        AuthenticationManager().GetACSAppOnlyContext(
                        ConfigurationManager.AppSettings["SiteCollUrl"],
                        ConfigurationManager.AppSettings["ClientIdWithSecret"],
                        ConfigurationManager.AppSettings["ClientSecret"]);

    return rtnContext;
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void SpCsPnpFramework_CreateOneList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        ListTemplateType myTemplate = ListTemplateType.GenericList;
        string listName = "NewListPnPFramework";
        bool enableVersioning = false;
        List newList = spPnpCtx.Web.CreateList(myTemplate, listName, enableVersioning);
    }
}
//gavdcodeend 001

//gavdcodebegin 002
static void SpCsPnpFramework_ReadOneList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        Console.WriteLine("List title - " + myList.Title);
    }
}
//gavdcodeend 002

//gavdcodebegin 003
static void SpCsPnpFramework_ListExists()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        bool blnListExists = myWeb.ListExists("NewListPnPFramework");

        Console.WriteLine("List exists - " + blnListExists);
    }
}
//gavdcodeend 003

//gavdcodebegin 004
static void SpCsPnpFramework_AddUserToSecurityRoleInList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        myList.SetListPermission(BuiltInIdentity.Everyone, RoleType.Editor);
    }
}
//gavdcodeend 004

//gavdcodebegin 005
static void SpCsPnpFramework_AddOneFieldToList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

        FieldType fieldType = FieldType.Text;
        PnP.Framework.Entities.FieldCreationInformation newFieldInfo =
                new PnP.Framework.Entities.FieldCreationInformation(fieldType)
                {
                    DisplayName = "NewFieldPnPFrameworkUsingInfo",
                    InternalName = "NewFieldPnPFrameworkInfo",
                    Id = new Guid()
                };
        myList.CreateField(newFieldInfo);

        string fieldXml = "<Field DisplayName='NewFieldPnPFrameworkUsingXml' " +
            "Type='Note' Required='FALSE' Name='NewFieldPnPFrameworkXml' />";
        myList.CreateField(fieldXml);
    }
}
//gavdcodeend 005

//gavdcodebegin 006
static void SpCsPnpFramework_ReadFilteredFieldsFromList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

        string[] fieldsToFind = new string[]
                { "NewFieldPnPFrameworkXml", "NewFieldPnPFrameworkInfo" };
        IEnumerable<Field> allFields = myList.GetFields(fieldsToFind);

        foreach (Field oneField in allFields)
        {
            Console.WriteLine(oneField.Title + " - " + oneField.TypeAsString);
        }
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void SpCsPnpFramework_ReadOneFieldFromList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

        Field myField = myList.GetFieldById
                    (new Guid("b0b75b9d-b358-49e6-b7fe-b2e35295f4bc"));

        Console.WriteLine(myField.InternalName + " - " + myField.TypeAsString);
    }
}
//gavdcodeend 007

//gavdcodebegin 008
static void SpCsPnpFramework_GetContentTypeList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
        ContentType myContentType = myList.GetContentTypeByName("Item");

        Console.WriteLine(myContentType.Description);
    }
}
//gavdcodeend 008

//gavdcodebegin 009
static void SpCsPnpFramework_AddContentTypeToList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
        myList.AddContentTypeToListByName("Comment");
    }
}
//gavdcodeend 009

//gavdcodebegin 010
static void SpCsPnpFramework_RemoveContentTypeFromList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
        myList.RemoveContentTypeByName("Comment");
    }
}
//gavdcodeend 010

//gavdcodebegin 011
static void SpCsPnpFramework_GetViewList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
        View myView = myList.GetViewByName("All Items");

        Console.WriteLine(myView.ListViewXml);
    }
}
//gavdcodeend 011

//gavdcodebegin 012
static void SpCsPnpFramework_AddViewToList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
        myList.CreateView("NewView", ViewType.Html, null, 30, false);
    }
}
//gavdcodeend 012

//gavdcodebegin 013
static void SpCsPnpFramework_BreakRoleInheritanceList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        myList.BreakRoleInheritance(true, false);
    }
}
//gavdcodeend 013

//gavdcodebegin 014
static void SpCsPnpFramework_RestoreRoleInheritanceList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        myList.ResetRoleInheritance();
    }
}
//gavdcodeend 014

//gavdcodebegin 015
static void SpCsPnpFramework_UpdateList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        myList.Description = "New List created with PnP Framework";
        myList.Update();
    }
}
//gavdcodeend 015

//gavdcodebegin 016
static void SpCsPnpFramework_DeleteList()
{
    using (ClientContext spPnpCtx = LoginPnPFramework_WithAccPw())
    {
        Web myWeb = spPnpCtx.Web;
        List myList = myWeb.GetListByTitle("NewListPnPFramework");

        myList.DeleteObject();
        myWeb.Update();
    }
}
//gavdcodeend 016


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//SpCsPnpFramework_CreateOneList();
//SpCsPnpFramework_ReadOneList();
//SpCsPnpFramework_ListExists();
//SpCsPnpFramework_UpdateList();
//SpCsPnpFramework_DeleteList();
//SpCsPnpFramework_AddUserToSecurityRoleInList();
//SpCsPnpFramework_BreakRoleInheritanceList();
//SpCsPnpFramework_RestoreRoleInheritanceList();
//SpCsPnpFramework_AddOneFieldToList();
//SpCsPnpFramework_ReadFilteredFieldsFromList();
//SpCsPnpFramework_ReadOneFieldFromList();
//SpCsPnpFramework_GetContentTypeList();
//SpCsPnpFramework_AddContentTypeToList();
//SpCsPnpFramework_RemoveContentTypeFromList();
//SpCsPnpFramework_GetViewList();
//SpCsPnpFramework_AddViewToList();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable

