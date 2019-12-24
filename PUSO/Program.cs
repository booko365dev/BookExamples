using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DrCh = DocumentFormat.OpenXml.Drawing.Charts;
using DrSp = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace PUSO
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelOpenXmlCreateDocument();
            //ExcelOpenXmlInsertTextInCell();
            //ExcelOpenXmlInsertFormulaInCell();
            //ExcelOpenXmlFindValueInOneCell();
            //ExcelOpenXmlFindAllValuesInCells();
            //ExcelOpenXmlFindAllValuesInCellsDataTable();
            //ExcelOpenXmlUpdateCellValue();
            //ExcelOpenXmlFindAllWorksheets();
            //ExcelOpenXmlFindAllHiddenWorksheets();
            //ExcelOpenXmlFindHiddenRowsAndCols();
            //ExcelOpenXmlInsertChart();

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        public static void ExcelOpenXmlCreateDocument()
        {
            using (SpreadsheetDocument myExcelDoc =
                SpreadsheetDocument.Create(@"C:\Temporary\ExcelDoc01.xlsx",
                                                    SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.AddWorkbookPart();
                myWorkbookPart.Workbook = new Workbook();

                WorksheetPart myWorksheetPart = myWorkbookPart.AddNewPart<WorksheetPart>();
                myWorksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets mySheets = myExcelDoc.WorkbookPart.Workbook.AppendChild<Sheets>(
                                                                new Sheets());
                Sheet oneSheet = new Sheet()
                {
                    Id = myExcelDoc.WorkbookPart.GetIdOfPart(myWorksheetPart),
                    SheetId = 1,
                    Name = "NewSheet"
                };
                mySheets.Append(oneSheet);

                myWorkbookPart.Workbook.Save();
            }
        }

        public static void ExcelOpenXmlInsertTextInCell()
        {
            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", true))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                WorksheetPart myWorksheetPart = myWorkbookPart.WorksheetParts.First();
                SheetData mySheetData = myWorksheetPart.Worksheet.Elements<SheetData>().
                                                                            First();

                Row newRow = mySheetData.Elements<Row>().FirstOrDefault();

                Cell newCell = new Cell()
                {
                    DataType = CellValues.InlineString,
                    CellReference = "C3"
                };
                InlineString myInlineString = new InlineString();
                Text myText = new Text()
                {
                    Text = "Text in Cell"
                };
                myInlineString.AppendChild(myText);
                newCell.AppendChild(myInlineString);

                newRow.Append(newCell);
            }
        }

        public static void ExcelOpenXmlInsertFormulaInCell()
        {
            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", true))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                WorksheetPart myWorksheetPart = myWorkbookPart.WorksheetParts.First();
                SheetData mySheetData = myWorksheetPart.Worksheet.Elements<SheetData>().
                                                                            First();

                Row newRow = mySheetData.Elements<Row>().FirstOrDefault();

                Cell newCell = new Cell()
                {
                    CellReference = "B4"
                };
                CellFormula myCellformula = new CellFormula();
                myCellformula.Text = "=RAND()";
                newCell.Append(myCellformula);

                newRow.Append(newCell);
            }
        }

        public static void ExcelOpenXmlFindValueInOneCell()
        {
            string cellValue = null;

            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                // Use the sheet name
                Sheet mySheet = myWorkbookPart.Workbook.Descendants<Sheet>().
                                    Where(sh => sh.Name == "NewSheet").FirstOrDefault();
                if (mySheet == null)
                {
                    throw new ArgumentException("Sheet not found");
                }

                WorksheetPart myWorksheetPart =
                                (WorksheetPart)(myWorkbookPart.GetPartById(mySheet.Id));

                // Use the cell address
                Cell myCell = myWorksheetPart.Worksheet.Descendants<Cell>().
                                Where(cl => cl.CellReference == "B2").FirstOrDefault();
                if (myCell != null)
                {
                    cellValue = myCell.InnerText;
                    if (myCell.DataType != null)
                    {
                        switch (myCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var stringTable =
                                    myWorkbookPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();
                                if (stringTable != null)
                                {
                                    cellValue =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(cellValue)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (cellValue)
                                {
                                    case "0":
                                        cellValue = "FALSE";
                                        break;
                                    default:
                                        cellValue = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }

            Console.WriteLine("Value found - " + cellValue);
        }

        public static void ExcelOpenXmlFindAllValuesInCellsDataTable()
        {
            DataTable myDataTable = new DataTable();

            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                IEnumerable<Sheet> mySheets = myWorkbookPart.Workbook.
                                            GetFirstChild<Sheets>().Elements<Sheet>();
                string firstSheet = mySheets.First().Id.Value;
                WorksheetPart myWorksheetPart = (WorksheetPart)myWorkbookPart.
                                                            GetPartById(firstSheet);
                Worksheet myWorkSheet = myWorksheetPart.Worksheet;
                SheetData mySheetData = myWorkSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> myRows = mySheetData.Descendants<Row>();

                foreach (Cell oneCell in myRows.ElementAt(0))
                {
                    myDataTable.Columns.Add(GetCellValue(myExcelDoc, oneCell));
                }

                foreach (Row oneRow in myRows)
                {
                    DataRow myDataRow = myDataTable.NewRow();
                    for (int rowIndex = 0; rowIndex <
                                        oneRow.Descendants<Cell>().Count(); rowIndex++)
                    {
                        Cell myCell = oneRow.Descendants<Cell>().ElementAt(rowIndex);
                        int myCellIndex = GetCellIndex(myCell);
                        myDataRow[myCellIndex] = GetCellValue(myExcelDoc, myCell);
                    }

                    myDataTable.Rows.Add(myDataRow);
                }
            }

            foreach (DataRow oneRow in myDataTable.Rows)
            {
                foreach (var oneColumn in oneRow.ItemArray)
                {
                    Console.WriteLine(oneColumn);
                }
            }
        }

        private static int GetCellIndex(Cell CellToFind)
        {
            int cellIndex = 0;

            string myCellRef = CellToFind.CellReference.ToString().ToUpper();
            foreach (char oneChar in myCellRef)
            {
                if (Char.IsLetter(oneChar))
                {
                    int charValue = (int)oneChar - (int)'A';
                    cellIndex = (cellIndex == 0) ? charValue : ((cellIndex + 1) * 26) +
                                                                            charValue;
                }
                else
                {
                    return cellIndex;
                }
            }

            return cellIndex;
        }

        private static string GetCellValue(SpreadsheetDocument ExcelDoc, Cell CellToFind)
        {
            SharedStringTablePart mySSTablePart =
                                            ExcelDoc.WorkbookPart.SharedStringTablePart;
            string cellValue = CellToFind.CellValue.InnerXml;

            if (CellToFind.DataType != null &&
                CellToFind.DataType.Value == CellValues.SharedString)
            {
                return mySSTablePart.SharedStringTable.ChildElements[
                                                    Int32.Parse(cellValue)].InnerText;
            }
            else
            {
                return cellValue;
            }
        }

        public static void ExcelOpenXmlFindAllValuesInCells()
        {
            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                Workbook myWorkbook = myWorkbookPart.Workbook;

                IEnumerable<Sheet> mySheets = myWorkbook.Descendants<Sheet>();
                foreach (Sheet oneSheet in mySheets)
                {
                    WorksheetPart myWorksheetPart = (WorksheetPart)myWorkbookPart.
                                                                GetPartById(oneSheet.Id);
                    SharedStringTablePart mySSTablePart = myWorkbookPart.
                                                                SharedStringTablePart;
                    SharedStringItem[] ssValues = mySSTablePart.SharedStringTable.
                                                Elements<SharedStringItem>().ToArray();

                    IEnumerable<Cell> myCells = myWorksheetPart.
                                                        Worksheet.Descendants<Cell>();
                    foreach (Cell oneCell in myCells)
                    {
                        Console.Write(oneCell.CellReference);

                        if (oneCell.DataType != null &&
                            oneCell.DataType.Value == CellValues.SharedString)
                        {
                            var cellIndex = int.Parse(oneCell.CellValue.Text);
                            var cellValue = ssValues[cellIndex].InnerText;
                            Console.Write(" - " + cellValue);
                        }
                        else
                        {
                            Console.WriteLine(" - " + oneCell.CellValue.Text);
                        }

                        if (oneCell.CellFormula != null)
                        {
                            Console.WriteLine(" - " + oneCell.CellFormula.Text);
                        }

                        Console.WriteLine("");
                    }
                }
            }
        }

        public static void ExcelOpenXmlUpdateCellValue()
        {
            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", true))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                // Use the sheet name
                Sheet sheet = myWorkbookPart.Workbook.Descendants<Sheet>().Where(
                    sh => sh.Name == "NewSheet").FirstOrDefault();
                WorksheetPart myWorksheetPart = (WorksheetPart)myWorkbookPart.
                                                            GetPartById(sheet.Id.Value);

                // Use the address of the cell
                Cell myCell = GetCell(myWorksheetPart.Worksheet, "B", 2);

                // Use the new value for the cell
                myCell.CellValue = new CellValue("9876");
                myCell.DataType = new EnumValue<CellValues>(CellValues.Number);

                myWorksheetPart.Worksheet.Save();
            }
        }

        private static Cell GetCell(Worksheet ExcelWorkSheet,
                                                    string ColumnIndex, uint RowIndex)
        {
            Row myRow = ExcelWorkSheet.GetFirstChild<SheetData>().
                           Elements<Row>().FirstOrDefault(rw => rw.RowIndex == RowIndex);
            if (myRow == null) return null;

            var myFirstRow = myRow.Elements<Cell>().Where(cl => string.Compare
                                                (cl.CellReference.Value, ColumnIndex +
                                                RowIndex, true) == 0).FirstOrDefault();

            if (myFirstRow == null) return null;

            return myFirstRow;
        }

        public static void ExcelOpenXmlFindAllWorksheets()
        {
            Sheets mySheets = null;

            using (SpreadsheetDocument myExcelDoc =
                SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                mySheets = myWorkbookPart.Workbook.Sheets;
            }

            foreach (Sheet oneSheet in mySheets)
            {
                Console.WriteLine(oneSheet.Name);
            }
        }

        public static void ExcelOpenXmlFindAllHiddenWorksheets()
        {
            IEnumerable<Sheet> myHiddenSheets = null;

            using (SpreadsheetDocument myExcelDoc =
                SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;
                IEnumerable<Sheet> mySheets = myWorkbookPart.Workbook.Descendants<Sheet>();

                myHiddenSheets = mySheets.Where((sh) => sh.State != null &&
                                         sh.State.HasValue &&
                                        (sh.State.Value == SheetStateValues.Hidden ||
                                         sh.State.Value == SheetStateValues.VeryHidden));
            }

            foreach (Sheet oneSheet in myHiddenSheets)
            {
                Console.WriteLine(oneSheet.Name);
            }
        }

        public static void ExcelOpenXmlFindHiddenRowsAndCols()
        {
            List<uint> hiddenRows = new List<uint>();
            List<uint> myHiddenCols = new List<uint>();

            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", false))
            {
                WorkbookPart myWorkbookPart = myExcelDoc.WorkbookPart;

                Sheet mySheet = myWorkbookPart.Workbook.Descendants<Sheet>().
                                    Where((sh) => sh.Name == "NewSheet").FirstOrDefault();
                if (mySheet == null)
                {
                    throw new Exception();
                }
                else
                {
                    WorksheetPart myWorksheetPart =
                                (WorksheetPart)(myWorkbookPart.GetPartById(mySheet.Id));
                    Worksheet myWorksheet = myWorksheetPart.Worksheet;

                    Console.WriteLine("Hidden Rows");
                    hiddenRows = myWorksheet.Descendants<Row>().
                                Where((rw) => rw.Hidden != null && rw.Hidden.Value).
                                Select(rw => rw.RowIndex.Value).ToList<uint>();
                    foreach (uint oneRow in hiddenRows)
                    {
                        Console.WriteLine(oneRow.ToString());
                    }

                    Console.WriteLine("Hidden Cols");
                    var hiddenCols = myWorksheet.Descendants<Column>().
                                Where((cl) => cl.Hidden != null && cl.Hidden.Value);
                    foreach (Column oneCol in hiddenCols)
                    {
                        for (uint clIndex = oneCol.Min.Value;
                                                clIndex <= oneCol.Max.Value; clIndex++)
                        {
                            myHiddenCols.Add(clIndex);
                        }
                    }
                    foreach (uint oneCol in myHiddenCols)
                    {
                        Console.WriteLine(oneCol.ToString());
                    }
                }
            }
        }

        private static void ExcelOpenXmlInsertChart()
        {
            Dictionary<string, int> chartData = new Dictionary<string, int>();
            chartData.Add("abc", 1);
            chartData.Add("def", 2);
            chartData.Add("ghi", 1);

            using (SpreadsheetDocument myExcelDoc =
                        SpreadsheetDocument.Open(@"C:\Temporary\ExcelDoc01.xlsx", true))
            {
                // Use the name of the sheet
                IEnumerable<Sheet> mySheets = myExcelDoc.WorkbookPart.Workbook.
                                  Descendants<Sheet>().Where(s => s.Name == "NewSheet");
                if (mySheets.Count() == 0)
                {
                    return;
                }
                WorksheetPart myWorksheetPart =
                                    (WorksheetPart)myExcelDoc.WorkbookPart.
                                    GetPartById(mySheets.First().Id);

                // Add a new drawing to the worksheet
                DrawingsPart myDrawingsPart = myWorksheetPart.AddNewPart<DrawingsPart>();
                myWorksheetPart.Worksheet.Append(new
                                            DocumentFormat.OpenXml.Spreadsheet.Drawing()
                {
                    Id = myWorksheetPart.GetIdOfPart(myDrawingsPart)
                });
                myWorksheetPart.Worksheet.Save();

                // Add a new chart and set the chart language to English-US
                ChartPart myChartPart = myDrawingsPart.AddNewPart<ChartPart>();
                myChartPart.ChartSpace = new DrCh.ChartSpace();
                myChartPart.ChartSpace.Append(new DrCh.EditingLanguage()
                {
                    Val = new StringValue("en-US")
                });
                DrCh.Chart myChart =
                       myChartPart.ChartSpace.AppendChild<DrCh.Chart>(new DrCh.Chart());

                // Create a new clustered column chart
                DrCh.PlotArea plotArea = myChart.AppendChild<DrCh.PlotArea>(
                                                                new DrCh.PlotArea());
                DrCh.Layout layout = plotArea.AppendChild<DrCh.Layout>(new DrCh.Layout());
                DrCh.BarChart barChart = plotArea.AppendChild<DrCh.BarChart>(
                                            new DrCh.BarChart(new DrCh.BarDirection()
                                            {
                                              Val = new EnumValue<DrCh.BarDirectionValues>
                                            (DrCh.BarDirectionValues.Column)
                                            },
                                            new DrCh.BarGrouping()
                                            {
                                               Val = new EnumValue<DrCh.BarGroupingValues>
                                                (DrCh.BarGroupingValues.Clustered)
                                            }));

                uint myIndex = 0;

                foreach (string oneKey in chartData.Keys)
                {
                    DrCh.BarChartSeries barChartSeries = barChart.AppendChild
                          <DrCh.BarChartSeries>(new DrCh.BarChartSeries(new DrCh.Index()
                          {
                              Val = new UInt32Value(myIndex)
                          },
                        new DrCh.Order() { Val = new UInt32Value(myIndex) },
                        new DrCh.SeriesText(new DrCh.NumericValue() { Text = oneKey })));

                    DrCh.StringLiteral strLit = barChartSeries.AppendChild<DrCh.
                            CategoryAxisData>(new DrCh.CategoryAxisData()).
                            AppendChild<DrCh.StringLiteral>(new DrCh.StringLiteral());
                    strLit.Append(new DrCh.PointCount() { Val = new UInt32Value(1U) });
                    // Use the title for the graphic
                    strLit.AppendChild<DrCh.StringPoint>(new DrCh.StringPoint()
                    {
                        Index = new UInt32Value(0U)
                    }).
                        Append(new DrCh.NumericValue("My New Graphic"));

                    DrCh.NumberLiteral numLit = barChartSeries.AppendChild
                            <DocumentFormat.OpenXml.Drawing.Charts.Values>(
                            new DocumentFormat.OpenXml.Drawing.Charts.Values()).
                            AppendChild<DrCh.NumberLiteral>(new DrCh.NumberLiteral());
                    numLit.Append(new DrCh.FormatCode("General"));
                    numLit.Append(new DrCh.PointCount() { Val = new UInt32Value(1U) });
                    numLit.AppendChild<DrCh.NumericPoint>(new DrCh.NumericPoint()
                    {
                        Index = new UInt32Value(0u)
                    }).Append
                            (new DrCh.NumericValue(chartData[oneKey].ToString()));

                    myIndex++;
                }

                barChart.Append(new DrCh.AxisId() { Val = new UInt32Value(48650112u) });
                barChart.Append(new DrCh.AxisId() { Val = new UInt32Value(48672768u) });

                // Add the Category Axis.
                DrCh.CategoryAxis catAx = plotArea.AppendChild<DrCh.CategoryAxis>
                        (new DrCh.CategoryAxis(new DrCh.AxisId()
                        {
                            Val = new UInt32Value(48650112u)
                        },
                                  new DrCh.Scaling(new DrCh.Orientation()
                                  {
                                      Val = new EnumValue<DocumentFormat.
                                  OpenXml.Drawing.Charts.OrientationValues>(
                                  DrCh.OrientationValues.MinMax)
                                  }),
                    new DrCh.AxisPosition()
                    {
                        Val = new EnumValue<DrCh.AxisPositionValues>
                                    (DrCh.AxisPositionValues.Bottom)
                    },
                                  new DrCh.TickLabelPosition()
                                  {
                                      Val = new EnumValue<DrCh.TickLabelPositionValues>
                                      (DrCh.TickLabelPositionValues.NextTo)
                                  },
                                  new DrCh.CrossingAxis()
                                  {
                                      Val = new UInt32Value(48672768U)
                                  },
                                  new DrCh.Crosses()
                                  {
                                      Val = new EnumValue<DrCh.CrossesValues>(
                                          DrCh.CrossesValues.AutoZero)
                                  },
                                  new DrCh.AutoLabeled()
                                  {
                                      Val = new BooleanValue(true)
                                  },
                                  new DrCh.LabelAlignment()
                                  {
                                      Val = new EnumValue<DrCh.LabelAlignmentValues>(
                                          DrCh.LabelAlignmentValues.Center)
                                  },
                                  new DrCh.LabelOffset()
                                  {
                                      Val = new UInt16Value((ushort)100)
                                  }));

                // Add the Value Axis.
                DrCh.ValueAxis valAx = plotArea.AppendChild<DrCh.ValueAxis>(
                    new DrCh.ValueAxis(new DrCh.AxisId()
                    {
                        Val = new UInt32Value(48672768u)
                    },
                        new DrCh.Scaling(new DrCh.Orientation()
                        {
                            Val = new EnumValue<DrCh.OrientationValues>(
                                DrCh.OrientationValues.MinMax)
                        }),
                        new DrCh.AxisPosition()
                        {
                            Val = new EnumValue<DrCh.AxisPositionValues>(
                                DrCh.AxisPositionValues.Left)
                        },
                        new DrCh.MajorGridlines(),
                        new DrCh.NumberingFormat()
                        {
                            FormatCode = new StringValue("General"),
                            SourceLinked = new BooleanValue(true)
                        },
                        new DrCh.TickLabelPosition()
                        {
                            Val = new EnumValue<DrCh.TickLabelPositionValues>
                                (DrCh.TickLabelPositionValues.NextTo)
                        },
                        new DrCh.CrossingAxis() { Val = new UInt32Value(48650112U) },
                        new DrCh.Crosses()
                        {
                            Val = new EnumValue<DrCh.CrossesValues>
                                (DrCh.CrossesValues.AutoZero)
                        },
                        new DrCh.CrossBetween()
                        {
                            Val = new EnumValue<DrCh.CrossBetweenValues>
                                (DrCh.CrossBetweenValues.Between)
                        }));

                // Add the chart Legend.
                DrCh.Legend myLegend = myChart.AppendChild<DrCh.Legend>(
                    new DrCh.Legend(new DrCh.LegendPosition()
                    {
                        Val = new EnumValue<DrCh.LegendPositionValues>
                        (DrCh.LegendPositionValues.Right)
                    },
                    new DrCh.Layout()));

                myChart.Append(new DrCh.PlotVisibleOnly()
                {
                    Val = new BooleanValue(true)
                });

                myChartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object.
                myDrawingsPart.WorksheetDrawing = new DrSp.WorksheetDrawing();
                DrSp.TwoCellAnchor twoCellAnchor = myDrawingsPart.WorksheetDrawing.
                                        AppendChild<DrSp.TwoCellAnchor>(
                    new DrSp.TwoCellAnchor());
                twoCellAnchor.Append(new DrSp.FromMarker(new DrSp.ColumnId("9"),
                    new DrSp.ColumnOffset("581025"),
                    new DrSp.RowId("17"),
                    new DrSp.RowOffset("114300")));
                twoCellAnchor.Append(new DrSp.ToMarker(new DrSp.ColumnId("17"),
                    new DrSp.ColumnOffset("276225"),
                    new DrSp.RowId("32"),
                    new DrSp.RowOffset("0")));

                // Append a GraphicFrame to the TwoCellAnchor object.
                DrSp.GraphicFrame myGraphicFrame =
                    twoCellAnchor.AppendChild<DrSp.GraphicFrame>(new DrSp.GraphicFrame());
                myGraphicFrame.Macro = "";

                myGraphicFrame.Append(new DrSp.NonVisualGraphicFrameProperties(
                    new DrSp.NonVisualDrawingProperties()
                    {
                        Id = new UInt32Value(2u),
                        Name = "Chart 1"
                    },
                    new DrSp.NonVisualGraphicFrameDrawingProperties()));

                myGraphicFrame.Append(new DrSp.Transform(
                    new DocumentFormat.OpenXml.Drawing.Offset()
                    {
                        X = 0L,
                        Y = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Extents()
                    {
                        Cx = 0L,
                        Cy = 0L
                    }));

                myGraphicFrame.Append(
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DrCh.ChartReference()
                            {
                                Id = myDrawingsPart.GetIdOfPart(myChartPart)
                            })
                        {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                        }));

                twoCellAnchor.Append(new DrSp.ClientData());

                myDrawingsPart.WorksheetDrawing.Save();
            }
        }
    }
}

