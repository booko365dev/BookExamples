using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Enums;
using System.Configuration;
using System.Security;

//---------------------------------------------------------------------------------------
// ------**** ATTENTION **** This is a DotNet Core 8.0 Console Application ****----------
//---------------------------------------------------------------------------------------
#nullable disable
#pragma warning disable CS8321 // Local function is declared but never used

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Login routines ***---------------------------
//---------------------------------------------------------------------------------------

static ClientContext CsPnPFramework_LoginWithAccPw()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithAccPw"],
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsPnPFramework_LoginWithCertificate()
{
    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["ClientIdWithCert"],
                            @"[PathForThePfxCertificateFile]",
                            "[PasswordForTheCertificate]",
                            "[Domain].onmicrosoft.com");

    ClientContext rtnContext = myAuthManager.GetContext(
                                     ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsPnPFramework_LoginPnPManagementShell()
{
    SecureString mySecurePw = new();
    foreach (char oneChr in ConfigurationManager.AppSettings["UserPw"])
    { mySecurePw.AppendChar(oneChr); }

    AuthenticationManager myAuthManager = new(
                            ConfigurationManager.AppSettings["UserName"],
                            mySecurePw);

    ClientContext rtnContext = myAuthManager.GetContext(
                            ConfigurationManager.AppSettings["SiteCollUrl"]);

    return rtnContext;
}

static ClientContext CsPnPFramework_LoginWithSecret()  //*** LEGACY CODE ***
{
    // NOTE: Microsoft stopped AzureAD App access for authentication of SharePoint
    //  using secrets. This method does not work anymore for any SharePoint query
    ClientContext rtnContext = new AuthenticationManager().GetACSAppOnlyContext(
                        ConfigurationManager.AppSettings["SiteCollUrl"],
                        ConfigurationManager.AppSettings["ClientIdWithSecret"],
                        ConfigurationManager.AppSettings["ClientSecret"]);

    return rtnContext;
}


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Example routines ***-------------------------
//---------------------------------------------------------------------------------------

//gavdcodebegin 001
static void CsSpPnpFramework_CreateOneList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    ListTemplateType myTemplate = ListTemplateType.GenericList;
    string listName = "NewListPnPFramework";
    bool enableVersioning = false;

    List newList = spPnpCtx.Web.CreateList(myTemplate, listName, enableVersioning);
}
//gavdcodeend 001

//gavdcodebegin 002
static void CsSpPnpFramework_ReadOneList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    Console.WriteLine("List title - " + myList.Title);
}
//gavdcodeend 002

//gavdcodebegin 003
static void CsSpPnpFramework_ListExists()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    bool blnListExists = myWeb.ListExists("NewListPnPFramework");

    Console.WriteLine("List exists - " + blnListExists);
}
//gavdcodeend 003

//gavdcodebegin 004
static void CsSpPnpFramework_AddUserToSecurityRoleInList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    myList.SetListPermission(BuiltInIdentity.Everyone, RoleType.Editor);
}
//gavdcodeend 004

//gavdcodebegin 005
static void CsSpPnpFramework_AddOneFieldToList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    FieldType fieldType = FieldType.Text;
    PnP.Framework.Entities.FieldCreationInformation newFieldInfo = new(fieldType)
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
//gavdcodeend 005

//gavdcodebegin 006
static void CsSpPnpFramework_ReadFilteredFieldsFromList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    string[] fieldsToFind = ["NewFieldPnPFrameworkXml", "NewFieldPnPFrameworkInfo"];
    IEnumerable<Field> allFields = myList.GetFields(fieldsToFind);

    foreach (Field oneField in allFields)
    {
        Console.WriteLine(oneField.Title + " - " + oneField.TypeAsString);
    }
}
//gavdcodeend 006

//gavdcodebegin 007
static void CsSpPnpFramework_ReadOneFieldFromList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    Field myField = myList.GetFieldById
                (new Guid("b0b75b9d-b358-49e6-b7fe-b2e35295f4bc"));

    Console.WriteLine(myField.InternalName + " - " + myField.TypeAsString);
}
//gavdcodeend 007

//gavdcodebegin 008
static void CsSpPnpFramework_GetContentTypeList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
    ContentType myContentType = myList.GetContentTypeByName("Item");

    Console.WriteLine(myContentType.Description);
}
//gavdcodeend 008

//gavdcodebegin 009
static void CsSpPnpFramework_AddContentTypeToList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    myList.AddContentTypeToListByName("Comment");
}
//gavdcodeend 009

//gavdcodebegin 010
static void CsSpPnpFramework_RemoveContentTypeFromList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    myList.RemoveContentTypeByName("Comment");
}
//gavdcodeend 010

//gavdcodebegin 011
static void CsSpPnpFramework_GetViewList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");
    View myView = myList.GetViewByName("All Items");

    Console.WriteLine(myView.ListViewXml);
}
//gavdcodeend 011

//gavdcodebegin 012
static void CsSpPnpFramework_AddViewToList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.Lists.GetByTitle("NewListPnPFramework");

    myList.CreateView("NewView", ViewType.Html, null, 30, false);
}
//gavdcodeend 012

//gavdcodebegin 013
static void CsSpPnpFramework_BreakRoleInheritanceList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    myList.BreakRoleInheritance(true, false);
}
//gavdcodeend 013

//gavdcodebegin 014
static void CsSpPnpFramework_RestoreRoleInheritanceList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    myList.ResetRoleInheritance();
}
//gavdcodeend 014

//gavdcodebegin 015
static void CsSpPnpFramework_UpdateList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();

    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    myList.Description = "New List created with PnP Framework";
    myList.Update();
}
//gavdcodeend 015

//gavdcodebegin 016
static void CsSpPnpFramework_DeleteList()
{
    using ClientContext spPnpCtx = CsPnPFramework_LoginWithAccPw();
    
    Web myWeb = spPnpCtx.Web;
    List myList = myWeb.GetListByTitle("NewListPnPFramework");

    myList.DeleteObject();
    myWeb.Update();
}
//gavdcodeend 016


//---------------------------------------------------------------------------------------
//***-----------------------------------*** Running the routines ***---------------------
//---------------------------------------------------------------------------------------

//# *** Latest Source Code Index: 016 ***

//CsSpPnpFramework_CreateOneList();
//CsSpPnpFramework_ReadOneList();
//CsSpPnpFramework_ListExists();
//CsSpPnpFramework_UpdateList();
//CsSpPnpFramework_DeleteList();
//CsSpPnpFramework_BreakRoleInheritanceList();
//CsSpPnpFramework_AddUserToSecurityRoleInList();
//CsSpPnpFramework_RestoreRoleInheritanceList();
//CsSpPnpFramework_AddOneFieldToList();
//CsSpPnpFramework_ReadFilteredFieldsFromList();
//CsSpPnpFramework_ReadOneFieldFromList();
//CsSpPnpFramework_GetContentTypeList();
//CsSpPnpFramework_AddContentTypeToList();
//CsSpPnpFramework_RemoveContentTypeFromList();
//CsSpPnpFramework_GetViewList();
//CsSpPnpFramework_AddViewToList();

Console.WriteLine("Done");

//---------------------------------------------------------------------------------------
//***-----------------------------------*** Class routines ***---------------------------
//---------------------------------------------------------------------------------------


#nullable enable
#pragma warning restore CS8321 // Local function is declared but never used

