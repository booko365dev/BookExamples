using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace RMXX
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelEpplusCreateDocument();
            //ExcelEpplusReadSpreedshet();
            //ExcelEpplusUpdateCellValue();
            //ExcelEpplusInsertLineChart();
            //ExcelEpplusInsertPieChart();
            //ExcelEpplusStyleSheet();
            //ExcelEpplusAddFormules();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        public static void ExcelEpplusCreateDocument()
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "Guitaca";
                excelPackage.Workbook.Properties.Title = "Test EPPlus Excel Document";
                excelPackage.Workbook.Properties.Subject = "EPPlus demo";
                excelPackage.Workbook.Properties.Created = DateTime.Now;

                // Create the WorkSheet
                ExcelWorksheet myWorksheet =
                                        excelPackage.Workbook.Worksheets.Add("NewSheet");

                // Add some text to cell A1
                myWorksheet.Cells["A1"].Value = "First spreadsheet";
                // You could also use [line, column] notation:
                myWorksheet.Cells[1, 2].Value = "This is cell B1";

                // Save the file
                FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
                excelPackage.SaveAs(myFileInfo);
            }
        }

        public static void ExcelEpplusReadSpreedshet()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                // Get a WorkSheet by index. The EPPlus indexes are base 1, not base 0
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[1];

                // Get a WorkSheet by name
                ExcelWorksheet namedWorksheet =
                                        excelPackage.Workbook.Worksheets["NewSheet"];

                // Get a WorkSheet by name using LINQ
                ExcelWorksheet anotherWorksheet =
                    excelPackage.Workbook.Worksheets.FirstOrDefault(ws => ws.Name ==
                                                                            "NewSheet");

                // The content from cells A1 and B1 as string (two different notations)
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();

                Console.WriteLine(valA1 + " - " + valB1);
            }
        }

        public static void ExcelEpplusUpdateCellValue()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                ExcelWorksheet myWorksheet = excelPackage.Workbook.Worksheets["NewSheet"];
                myWorksheet.Cells[1, 2].Value = 10;

                excelPackage.SaveAs(myFileInfo);
            }
        }

        public static void ExcelEpplusInsertLineChart()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                // Create a WorkSheet
                ExcelWorksheet myWorksheet =
                                excelPackage.Workbook.Worksheets.Add("LineChartSheet");

                // Fill cell data with a loop. The row and column indexes start at 1
                Random rnd = new Random();
                for (int counter = 2; counter <= 11; counter++)
                {
                    myWorksheet.Cells[1, counter].Value = "Value " + (counter - 1);
                    myWorksheet.Cells[2, counter].Value = rnd.Next(5, 25);
                    myWorksheet.Cells[3, counter].Value = rnd.Next(5, 25);
                }
                myWorksheet.Cells[2, 1].Value = "Age 1";
                myWorksheet.Cells[3, 1].Value = "Age 2";

                // Create a new chart of type Line
                ExcelLineChart myLineChart = myWorksheet.Drawings.AddChart(
                                    "lineChart", eChartType.Line) as ExcelLineChart;
                myLineChart.Title.Text = "LineChart Example";

                // Create and add the ranges for the chart
                ExcelRange rangeLabel = myWorksheet.Cells["B1:K1"];
                ExcelRange range1 = myWorksheet.Cells["B2:K2"];
                ExcelRange range2 = myWorksheet.Cells["B3:K3"];

                myLineChart.Series.Add(range1, rangeLabel);
                myLineChart.Series.Add(range2, rangeLabel);

                // Set the properties of the chart
                myLineChart.Series[0].Header = myWorksheet.Cells["A2"].Value.ToString();
                myLineChart.Series[1].Header = myWorksheet.Cells["A3"].Value.ToString();
                myLineChart.Legend.Position = eLegendPosition.Right;
                myLineChart.SetSize(600, 300);
                myLineChart.SetPosition(5, 0, 1, 0);

                excelPackage.SaveAs(myFileInfo);
            }
        }

        public static void ExcelEpplusInsertPieChart()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                // Create a WorkSheet
                ExcelWorksheet myWorksheet =
                                excelPackage.Workbook.Worksheets.Add("PieChartSheet");

                // Fill cell data with a loop. The row and column indexes start at 1
                Random rnd = new Random();
                for (int counter = 1; counter <= 10; counter++)
                {
                    myWorksheet.Cells[1, counter].Value = "Value " + counter;
                    myWorksheet.Cells[2, counter].Value = rnd.Next(5, 15);
                }

                // Create a new pie chart of type Pie3D
                ExcelPieChart myPieChart = myWorksheet.Drawings.AddChart("pieChart",
                                                    eChartType.Pie3D) as ExcelPieChart;

                // Set the properties of the chart
                myPieChart.Title.Text = "PieChart Example";
                myPieChart.Series.Add(ExcelRange.GetAddress(2, 1, 2, 10),
                                                ExcelRange.GetAddress(1, 1, 1, 10));
                myPieChart.Legend.Position = eLegendPosition.Bottom;
                myPieChart.DataLabel.ShowPercent = true;
                myPieChart.SetSize(500, 400);
                myPieChart.SetPosition(4, 0, 2, 0);

                excelPackage.SaveAs(myFileInfo);
            }
        }

        public static void ExcelEpplusStyleSheet()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                // Create the WorkSheet
                ExcelWorksheet myWorksheet =
                                    excelPackage.Workbook.Worksheets.Add("StyleSheet");

                // Add some dummy data. The row and column indexes start at 1
                for (int counter1 = 1; counter1 <= 30; counter1++)
                    for (int counter2 = 1; counter2 <= 15; counter2++)
                        myWorksheet.Cells[counter1, counter2].Value =
                                            "Row " + counter1 + ", Column " + counter2;

                // Fill column A with solid red color
                myWorksheet.Column(1).Style.Fill.PatternType = ExcelFillStyle.Solid;
                myWorksheet.Column(1).Style.Fill.BackgroundColor.SetColor(
                                                ColorTranslator.FromHtml("#FF0000"));

                // Set the font type for cells C1 - C30
                myWorksheet.Cells["C1:C30"].Style.Font.Size = 13;
                myWorksheet.Cells["C1:C30"].Style.Font.Name = "Calibri";
                myWorksheet.Cells["C1:C30"].Style.Font.Bold = true;
                myWorksheet.Cells["C1:C30"].Style.Font.Color.SetColor(Color.Blue);

                // Fill row 4 with striped orange background
                myWorksheet.Row(4).Style.Fill.PatternType = ExcelFillStyle.DarkHorizontal;
                myWorksheet.Row(4).Style.Fill.BackgroundColor.SetColor(Color.Orange);

                // Make the borders of cell F6 thick
                myWorksheet.Cells[6, 6].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                myWorksheet.Cells[6, 6].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                myWorksheet.Cells[6, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                myWorksheet.Cells[6, 6].Style.Border.Left.Style = ExcelBorderStyle.Thick;

                // Make the borders of cells A18 - J18 double and with a purple color
                myWorksheet.Cells["A18:J18"].Style.Border.Top.Style =
                                                                ExcelBorderStyle.Double;
                myWorksheet.Cells["A18:J18"].Style.Border.Bottom.Style =
                                                                ExcelBorderStyle.Double;
                myWorksheet.Cells["A18:J18"].Style.Border.Top.Color.SetColor(Color.Purple);
                myWorksheet.Cells["A18:J18"].Style.Border.Bottom.Color.
                                                                SetColor(Color.Purple);

                // Make all text fit the cells
                myWorksheet.Cells[myWorksheet.Dimension.Address].AutoFitColumns();

                // Make all columns just a bit wider
                for (int col = 1; col <= myWorksheet.Dimension.End.Column; col++)
                    myWorksheet.Column(col).Width = myWorksheet.Column(col).Width + 1;

                // Make column H wider and set the text align to the top and right
                myWorksheet.Column(8).Width = 25;
                myWorksheet.Column(8).Style.HorizontalAlignment =
                                                        ExcelHorizontalAlignment.Right;
                myWorksheet.Column(8).Style.VerticalAlignment =
                                                        ExcelVerticalAlignment.Top;

                // Get the image from disk
                using (System.Drawing.Image myImage = System.Drawing.Image.
                                            FromFile(@"C:\Temporary\MyPicture.jpg"))
                {
                    ExcelPicture myExcelImage = myWorksheet.Drawings.AddPicture(
                                                                    "My Logo", myImage);

                    // Add the image to row 20, column E
                    myExcelImage.SetPosition(20, 0, 5, 0);
                }

                excelPackage.SaveAs(myFileInfo);
            }
        }

        public static void ExcelEpplusAddFormules()
        {
            FileInfo myFileInfo = new FileInfo(@"C:\Temporary\ExcelEPPlus01.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(myFileInfo))
            {
                // Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.
                                                                Add("FormuleSheet");

                // Set the calculation mode to automatic
                excelPackage.Workbook.CalcMode = ExcelCalcMode.Automatic;

                // Fill cell data with a loop. The row and column indexes start at 1
                for (int counter1 = 1; counter1 <= 25; counter1++)
                    for (int counter2 = 1; counter2 <= 10; counter2++)
                        worksheet.Cells[counter1, counter2].Value =
                                                            (counter1 + counter2) - 1;

                // Set the total value of cells in range A1 - A25 into A27
                worksheet.Cells["A27"].Formula = "=SUM(A1:A25)";

                // Set the number of cells with content in range C1 - C25 into C27
                worksheet.Cells["C27"].Formula = "=COUNT(C1:C25)";

                // Fill column K with the sum of each row, range A - J
                for (int counter3 = 1; counter3 <= 25; counter3++)
                {
                    ExcelRange myCell = worksheet.Cells[counter3, 12];
                    myCell.Formula = "=SUM(" + worksheet.Cells[counter3, 1].Address +
                                    ":" + worksheet.Cells[counter3, 10].Address + ")";
                }

                // Calculate the quartile of range E1 - E25 into E27
                worksheet.Cells[27, 5].Formula = "=QUARTILE(E1:E25,1)";

                // Calculate all the values of the formulas in the Excel file
                excelPackage.Workbook.Calculate();

                excelPackage.SaveAs(myFileInfo);
            }
        }
    }
}

