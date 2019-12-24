using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace WQEM
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Outlook.Application myApplication = this.Application;
            Outlook.Inspectors myInspectors = myApplication.Inspectors;

            myInspectors.NewInspector +=
                                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                                                            Inspectors_AddTextToNewMail);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, 
            //    see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        void Inspectors_AddTextToNewMail(Outlook.Inspector inspector)
        {
            string userName = (string)this.Application.ActiveExplorer().Session.
                                                                        CurrentUser.Name;

            Outlook.MailItem myMailItem = inspector.CurrentItem as Outlook.MailItem;
            if (myMailItem != null)
            {
                if (myMailItem.EntryID == null)
                {
                    myMailItem.Subject = "Email created by " + userName;
                    myMailItem.Body = DateTime.Now + Environment.NewLine +
                                        "To Whom It May Concern,";
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

