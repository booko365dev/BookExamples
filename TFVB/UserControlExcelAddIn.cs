using Microsoft.Office.Interop.Excel;
using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TFVB
{
    public partial class UserControlExcelAddIn : UserControl
    {
        public UserControlExcelAddIn()
        {
            InitializeComponent();
        }

        //gavdcodebegin 01
        private void BtnAddPicture_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.ValidateNames = true;

            fileDialog.Filter = "All|*.*|Bitmap|*.bmp|Gif|*.gif|JPEG|*.jpeg|Png|*.png";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                Bitmap bmPicture = new Bitmap(fileDialog.FileName);

                Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddPicture(
                        fileDialog.FileName,
                        Microsoft.Office.Core.MsoTriState.msoFalse,
                        Microsoft.Office.Core.MsoTriState.msoCTrue,
                        10, 10, bmPicture.Width, bmPicture.Height);
            }
        }
        //gavdcodeend 01

        //gavdcodebegin 03
        private void BtnSaveAsCsv_Click(object sender, EventArgs e)
        {
            SaveFileDialog exportDialog = new SaveFileDialog();
            exportDialog.ValidateNames = true;

            if (exportDialog.ShowDialog() == DialogResult.OK)
            {
                Worksheet mySheet = (Excel.Worksheet)Globals.ThisAddIn.Application.
                                                            ActiveWorkbook.ActiveSheet;

                mySheet.SaveAs(
                                exportDialog.FileName,
                                Excel.XlFileFormat.xlCSV,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);
            }
        }
        //gavdcodeend 03

        //gavdcodebegin 05
        private void BtnGetTime_Click(object sender, EventArgs e)
        {
            FormTime newForm = new FormTime();
            newForm.Show();
        }
        //gavdcodeend 05
    }
}
