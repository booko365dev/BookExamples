using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace RLRC
{
    public partial class ThisAddIn
    {
        //gavdcodebegin 001
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationSave +=
                    new PowerPoint.EApplication_PresentationSaveEventHandler(
                                    Application_PresentationSave);
        }

        void Application_PresentationSave(PowerPoint.Presentation Prs)
        {
            Prs.ApplyTheme(
                @"C:\Program Files\Microsoft Office\root\Document Themes 16\Wisp.thmx");
            //or, for Office 32 bits
            //C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Wisp.tmx

            PowerPoint.CustomLayout pptLayout = Prs.Slides[1].CustomLayout;
            Prs.Slides.AddSlide(1, pptLayout);

            Prs.RemovePersonalInformation = Office.MsoTriState.msoTrue;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        //gavdcodeend 001

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
