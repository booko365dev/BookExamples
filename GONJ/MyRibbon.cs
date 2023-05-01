using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace GONJ
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        //gavdcodebegin 001
        private void BtnAddPicture_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.ValidateNames = true;

            fileDialog.Filter = "All|*.*|Bitmap|*.bmp|Gif|*.gif|JPEG|*.jpeg|Png|*.png";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.Shapes.AddPicture(
                                                            fileDialog.FileName);
            }
        }
        //gavdcodeend 001

        //gavdcodebegin 002
        private void BtnAddTable_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(
                Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0), 3, 4);
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.
                Shading.BackgroundPatternColor = Word.WdColor.wdColorAqua;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Font.Size = 12;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Rows.Borders.Enable = 1;
        }
        //gavdcodeend 002

        //gavdcodebegin 003
        private void BtnSaveAsPdf_Click(object sender, RibbonControlEventArgs e)
        {
            SaveFileDialog exportDialog = new SaveFileDialog();
            exportDialog.ValidateNames = true;

            if (exportDialog.ShowDialog() == DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(
                            exportDialog.FileName,
                            Word.WdExportFormat.wdExportFormatPDF,
                            OpenAfterExport: true);
            }
        }
        //gavdcodeend 003

        //gavdcodebegin 004
        private void BtnGetTime_Click(object sender, RibbonControlEventArgs e)
        {
            FormTime newForm = new FormTime();
            newForm.Show();
        }
        //gavdcodeend 004

    }
}
