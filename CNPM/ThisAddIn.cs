using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace CNPM
{
    public partial class ThisAddIn
    {
        //gavdcodebegin 01
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave +=
                                new Excel.AppEvents_WorkbookBeforeSaveEventHandler(
                                                Application_WorkbookBeforeSave);
        }

        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI,
                                                                    ref bool Cancel)
        {
            Excel.Worksheet activeWorksheet = (Excel.Worksheet)Application.ActiveSheet;

            Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            activeCell.Interior.Color = Color.Aqua;
            activeCell.Borders.Color = Color.Red;
            activeCell.Borders.LineStyle = Excel.XlLineStyle.xlDouble;
            activeCell.Columns.AutoFit();

            Excel.Range firstRow = activeWorksheet.get_Range("A1");
            firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
            newFirstRow.Value2 = DateTime.Now.ToShortDateString();

            Excel.Range secondRow = activeWorksheet.get_Range("A2");
            secondRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            Excel.Range newSecondRow = activeWorksheet.get_Range("A2");
            newSecondRow.Value2 = "To Whom It May Concern";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        //gavdcodeend 01

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
