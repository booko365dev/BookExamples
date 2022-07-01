using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Security;

namespace BURF
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext spCtx = LoginCsom();

            //SpCsCsom_CreateOneList(spCtx);
            //SpCsCsom_ReadAllList(spCtx);
            //SpCsCsom_ReadOneList(spCtx);
            //SpCsCsom_UpdateOneList(spCtx);
            //SpCsCsom_DeleteOneList(spCtx);
            //SpCsCsom_AddOneFieldToList(spCtx);
            //SpCsCsom_ReadAllFieldsFromList(spCtx);
            //SpCsCsom_ReadOneFieldFromList(spCtx);
            //SpCsCsom_UpdateOneFieldInList(spCtx);
            //SpCsCsom_DeleteOneFieldFromList(spCtx);
            //SpCsCsom_BreakSecurityInheritanceList(spCtx);
            //SpCsCsom_AddUserToSecurityRoleInList(spCtx);
            //SpCsCsom_UpdateUserSecurityRoleInList(spCtx);
            //SpCsCsom_DeleteUserFromSecurityRoleInList(spCtx);
            //SpCsCsom_ResetSecurityInheritanceList(spCtx);
            //SpCsCsom_FieldCreateText(spCtx);
            //SpCsCsom_ReadAllSiteColumns(spCtx);
            //SpCsCsom_AddOneSiteColumn(spCtx);
            //SpCsCsom_ColumnIndex(spCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        //gavdcodebegin 01
        static void SpCsCsom_CreateOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            ListCreationInformation myListCreationInfo = new ListCreationInformation();
            myListCreationInfo.Title = "NewListCsCsom";
            myListCreationInfo.Description = "New List created using CSharp CSOM";
            myListCreationInfo.TemplateType = (int)ListTemplateType.GenericList;

            List newList = spCtx.Web.Lists.Add(myListCreationInfo);
            newList.OnQuickLaunch = true;
            newList.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 01

        //gavdcodebegin 02
        static void SpCsCsom_ReadAllList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            ListCollection allLists = myWeb.Lists;
            spCtx.Load(allLists, lsts => lsts.Include(lst => lst.Title,
                                                     lst => lst.Id));
            spCtx.ExecuteQuery();

            foreach (List oneList in allLists)
            {
                Console.WriteLine(oneList.Title + " - " + oneList.Id);
            }
        }
        //gavdcodeend 02

        //gavdcodebegin 03
        static void SpCsCsom_ReadOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            spCtx.Load(myList);
            spCtx.ExecuteQuery();

            Console.WriteLine("List description - " + myList.Description);
        }
        //gavdcodeend 03

        //gavdcodebegin 04
        static void SpCsCsom_UpdateOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            myList.Description = "New List Description";
            myList.Update();
            spCtx.Load(myList);
            spCtx.ExecuteQuery();

            Console.WriteLine("List description - " + myList.Description);
        }
        //gavdcodeend 04

        //gavdcodebegin 05
        static void SpCsCsom_DeleteOneList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            myList.DeleteObject();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 05

        //gavdcodebegin 06
        static void SpCsCsom_AddOneFieldToList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            string fieldXml = "<Field DisplayName='MyMultilineField' Type='Note' />";
            Field myField = myList.Fields.AddFieldAsXml(fieldXml,
                                                       true,
                                                       AddFieldOptions.DefaultValue);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 06

        //gavdcodebegin 07
        static void SpCsCsom_ReadAllFieldsFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            FieldCollection allFields = myList.Fields;
            spCtx.Load(allFields, flds => flds.Include(fld => fld.Title,
                                                       fld => fld.TypeAsString));
            spCtx.ExecuteQuery();

            foreach (Field oneField in allFields)
            {
                Console.WriteLine(oneField.Title + " - " + oneField.TypeAsString);
            }
        }
        //gavdcodeend 07

        //gavdcodebegin 08
        static void SpCsCsom_ReadOneFieldFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            Field myField = myList.Fields.GetByTitle("MyMultilineField");
            spCtx.Load(myField);
            spCtx.ExecuteQuery();

            Console.WriteLine(myField.Id + " - " + myField.TypeAsString);
        }
        //gavdcodeend 08

        //gavdcodebegin 09
        static void SpCsCsom_UpdateOneFieldInList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            Field myField = myList.Fields.GetByTitle("MyMultilineField");

            FieldMultiLineText myFieldNote = spCtx.CastTo<FieldMultiLineText>(myField);
            myFieldNote.Description = "New Field Description";
            myFieldNote.Hidden = false;
            myFieldNote.NumberOfLines = 3;

            myField.Update();
            spCtx.Load(myField);
            spCtx.ExecuteQuery();

            Console.WriteLine(myField.Description);
        }
        //gavdcodeend 09

        //gavdcodebegin 10
        static void SpCsCsom_DeleteOneFieldFromList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            Field myField = myList.Fields.GetByTitle("MyMultilineField");
            myField.DeleteObject();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 10

        //gavdcodebegin 11
        static void SpCsCsom_BreakSecurityInheritanceList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            spCtx.Load(myList, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myList.HasUniqueRoleAssignments == false)
            {
                myList.BreakRoleInheritance(false, true);
            }
            myList.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 11

        //gavdcodebegin 15
        static void SpCsCsom_ResetSecurityInheritanceList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            spCtx.Load(myList, hura => hura.HasUniqueRoleAssignments);
            spCtx.ExecuteQuery();

            if (myList.HasUniqueRoleAssignments == true)
            {
                myList.ResetRoleInheritance();
            }
            myList.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 15

        //gavdcodebegin 12
        static void SpCsCsom_AddUserToSecurityRoleInList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Reader));
            myList.RoleAssignments.Add(myUser, roleDefinition);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 12

        //gavdcodebegin 13
        static void SpCsCsom_UpdateUserSecurityRoleInList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            RoleDefinitionBindingCollection roleDefinition =
                    new RoleDefinitionBindingCollection(spCtx);
            roleDefinition.Add(myWeb.RoleDefinitions.GetByType(RoleType.Administrator));

            RoleAssignment myRoleAssignment = myList.RoleAssignments.GetByPrincipal(myUser);
            myRoleAssignment.ImportRoleDefinitionBindings(roleDefinition);

            myRoleAssignment.Update();
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 13

        //gavdcodebegin 14
        static void SpCsCsom_DeleteUserFromSecurityRoleInList(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            myList.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 14

        //gavdcodebegin 16
        static void SpCsCsom_FieldCreateText(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            Guid myGuid = Guid.NewGuid();
            string schemaField = "<Field ID='" + myGuid + "' Type='Text' " +
                "Name='myTextCol' StaticName='myTextCol' DisplayName='My Text Col' />";
            Field myField = myList.Fields.AddFieldAsXml(schemaField, true,
                        AddFieldOptions.AddFieldInternalNameHint |
                        AddFieldOptions.AddToDefaultContentType);

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 16

        //gavdcodebegin 17
        static void SpCsCsom_ReadAllSiteColumns(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            FieldCollection allSiteColls = myWeb.Fields;

            spCtx.Load(allSiteColls, flds => flds.Include(fld => fld.Title,
                                                       fld => fld.Group));
            spCtx.ExecuteQuery();

            foreach (Field oneColl in allSiteColls)
            {
                Console.WriteLine(oneColl.Title + " - " + oneColl.Group);
            }
        }
        //gavdcodeend 17

        //gavdcodebegin 18
        static void SpCsCsom_AddOneSiteColumn(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;

            string fieldXml = "<Field DisplayName='MySiteColMultilineField' " +
                                                    "Type='Note' Group='MyGroup' />";
            Field myField = myWeb.Fields.AddFieldAsXml(fieldXml,
                                                       true,
                                                       AddFieldOptions.DefaultValue);
            spCtx.ExecuteQuery();
        }
        //gavdcodeend 18

        //gavdcodebegin 19
        static void SpCsCsom_ColumnIndex(ClientContext spCtx)  //*** LEGACY CODE ***
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            string myColumn = "My Text Col";
            Field myField = myList.Fields.GetByTitle(myColumn);
            myField.Indexed = true;
            myField.Update();

            spCtx.ExecuteQuery();
        }
        //gavdcodeend 19

        //-------------------------------------------------------------------------------
        static ClientContext LoginCsom()  //*** LEGACY CODE ***
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
