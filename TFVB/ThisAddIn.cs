namespace TFVB
{
    public partial class ThisAddIn
    {
        //gavdcodebegin 02
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ShowPanel();
        }

        private Microsoft.Office.Tools.CustomTaskPane customPanel;
        public void ShowPanel()
        {
            UserControlExcelAddIn panelObject = new UserControlExcelAddIn();
            customPanel = this.CustomTaskPanes.Add(panelObject, "My Panel");
            customPanel.Width = panelObject.Width;
            customPanel.Visible = true;
        }
        //gavdcodeend 02

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
