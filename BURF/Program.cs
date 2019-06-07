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

            //SpCsCsomCreateOneList(spCtx);
            //SpCsCsomReadAllList(spCtx);
            //SpCsCsomReadOneList(spCtx);
            //SpCsCsomUpdateOneList(spCtx);
            //SpCsCsomDeleteOneList(spCtx);
            //SpCsCsomAddOneFieldToList(spCtx);
            //SpCsCsomReadAllFieldsFromList(spCtx);
            //SpCsCsomReadOneFieldFromList(spCtx);
            //SpCsCsomUpdateOneFieldInList(spCtx);
            //SpCsCsomDeleteOneFieldFromList(spCtx);
            //SpCsCsomBreakSecurityInheritanceList(spCtx);
            //SpCsCsomAddUserToSecurityRoleInList(spCtx);
            //SpCsCsomUpdateUserSecurityRoleInList(spCtx);
            //SpCsCsomDeleteUserFromSecurityRoleInList(spCtx);
            //SpCsCsomResetSecurityInheritanceList(spCtx);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        static void SpCsCsomCreateOneList(ClientContext spCtx)
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

        static void SpCsCsomReadAllList(ClientContext spCtx)
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

        static void SpCsCsomReadOneList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            spCtx.Load(myList);
            spCtx.ExecuteQuery();

            Console.WriteLine("List description - " + myList.Description);
        }

        static void SpCsCsomUpdateOneList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            myList.Description = "New List Description";
            myList.Update();
            spCtx.Load(myList);
            spCtx.ExecuteQuery();

            Console.WriteLine("List description - " + myList.Description);
        }

        static void SpCsCsomDeleteOneList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            myList.DeleteObject();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomAddOneFieldToList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            string fieldXml = "<Field DisplayName='MyMultilineField' Type='Note' />";
            Field myField = myList.Fields.AddFieldAsXml(fieldXml,
                                                       true,
                                                       AddFieldOptions.DefaultValue);
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomReadAllFieldsFromList(ClientContext spCtx)
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

        static void SpCsCsomReadOneFieldFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            Field myField = myList.Fields.GetByTitle("MyMultilineField");
            spCtx.Load(myField);
            spCtx.ExecuteQuery();

            Console.WriteLine(myField.Id + " - " + myField.TypeAsString);
        }

        static void SpCsCsomUpdateOneFieldInList(ClientContext spCtx)
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

        static void SpCsCsomDeleteOneFieldFromList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");
            Field myField = myList.Fields.GetByTitle("MyMultilineField");
            myField.DeleteObject();
            spCtx.ExecuteQuery();
        }

        static void SpCsCsomBreakSecurityInheritanceList(ClientContext spCtx)
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

        static void SpCsCsomResetSecurityInheritanceList(ClientContext spCtx)
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

        static void SpCsCsomAddUserToSecurityRoleInList(ClientContext spCtx)
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

        static void SpCsCsomUpdateUserSecurityRoleInList(ClientContext spCtx)
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

        static void SpCsCsomDeleteUserFromSecurityRoleInList(ClientContext spCtx)
        {
            Web myWeb = spCtx.Web;
            List myList = myWeb.Lists.GetByTitle("NewListCsCsom");

            User myUser = myWeb.EnsureUser(ConfigurationManager.AppSettings["spUserName"]);
            myList.RoleAssignments.GetByPrincipal(myUser).DeleteObject();

            spCtx.ExecuteQuery();
            spCtx.Dispose();
        }

        //----------------------------------------------------------------------------------------
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
