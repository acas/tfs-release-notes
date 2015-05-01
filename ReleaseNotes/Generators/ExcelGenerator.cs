﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Data;
using System.IO;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using ReleaseNotes.Utility;
using System.Drawing;

namespace ReleaseNotes
{
    class ExcelGenerator : ReleaseNotesGenerator
    {
        // excel persistent objects
        private Excel.Application app;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;

        // excel positioning vars
        private int starterRow = 1;
        private int currentRow = 1;
        private int currentColumnCount = 4;
        private int currentColumnOffset = 2;
        private const int totalAllowedColumns = 24;

        // alphabet (for columns)
        private char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        /// <summary>
        /// Generates an excel instance
        /// </summary>
        /// <param name="settings"></param>
        private ExcelGenerator(NamedLookup settings, bool silent) : base(settings, silent)
        {
            app = new Excel.Application();
            app.Visible = !this.silent;
            app.UserControl = !this.silent;
            workbook = (Excel.Workbook)app.Workbooks.Add();
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "Release Notes";
            app.ActiveWindow.DisplayGridlines = false;
        }

        /// <summary>
        /// Generates an Excel instance
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static ExcelGenerator ExcelGeneratorFactory(NamedLookup settings, bool silent)
        {
            try
            {
                return new ExcelGenerator(settings, silent);
            }
            catch (COMException e)
            {
                (new Logger())
                    .SetLoggingType(Logger.Type.Error)
                    .SetMessage(e.Message + "Excel not initialized. \n Are you trying to run this server-side?...")
                    .Display();
                throw;
            }
        }

        /// <summary>
        /// Creates a vertical style table
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="headerText"></param>
        /// <param name="header"></param>
        public override void CreateVerticalTable(DataTable dataTable, string headerText, bool header)
        {
            // set the column count
            this.currentColumnCount = dataTable.Columns.Count;
            
            // create header
            if (header)
                CreateHeader(headerText);

            // set the current column count
            this.starterRow = currentRow;

            // add header row
            AddVerticalTableRow(Utilities.TableColumnsToStringArray(dataTable), false);

            // add table information
            foreach (DataRow row in dataTable.Rows)
                AddVerticalTableRow(Utilities.TableRowToStringArray(row), false);

            // set sizing and theming
            SetDefaultTheme(header);
            SetBasicTheme(false);
            AdvanceRow();
        }

        /// <summary>
        /// Creates a horizontal data table in Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="splits"></param>
        /// <param name="header"></param>
        public override void CreateHorizontalTable(NamedLookup data, int splits, bool header)
        {
            // 2 splits = 4 columns
            this.currentColumnCount = 2 * splits;
            this.starterRow = this.currentRow + 1;

            // if header needed
            if (header)
                CreateHeader(data.GetName());

            // get a list of the keys
            List<string> tableKeys = data.GetLookup().Keys.ToList();

            // determine the optimal number of rows for the table
            int optimalRowNumber = (tableKeys.Count() / splits) + (tableKeys.Count() % splits);

            // counter variable
            int counter = 0;

            for (int i = 1; i <= optimalRowNumber; i++)
            {
                for (int j = 1; j <= this.currentColumnCount; j++)
                {
                    Excel.Range cellRange = GetSingleCellRange(worksheet, j + currentColumnOffset - 1, currentRow);

                    string currentKey = "";
                    if (counter != tableKeys.Count())
                        currentKey = tableKeys.ElementAt(counter);

                    cellRange.RowHeight = 18;
                    cellRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    cellRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    cellRange.Font.Bold = 1;
                    cellRange.Font.Size = 10;
                    cellRange.Font.Name = "Arial";

                    if (j % 2 != 0)
                    {
                        cellRange.Interior.Color = Excel.XlRgbColor.rgbLightGrey;
                        cellRange.Value = currentKey;
                    }
                    else
                    {
                        if (counter != tableKeys.Count())
                        {
                            cellRange.Font.Color = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                            if (currentKey.Equals("Source"))
                            {
                                // hyperlink
                                cellRange.Font.Name = "Arial";
                                cellRange.Font.Size = 10;
                                cellRange.Font.Bold = 1;
                                cellRange.Hyperlinks.Add(cellRange, settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems", Type.Missing, "Work Items",
                                    settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems" + Environment.NewLine + data[currentKey]);
                            }
                            else
                            {
                                cellRange.Value = data[currentKey];
                            }
                            counter++;
                        }
                        else
                        {
                            cellRange.Value = "";
                        }
                    }
                }

                if (i == optimalRowNumber)
                {
                    // style with basic theme
                    SetBasicTheme(true);
                }
                AdvanceRow(0);
            }

            // insert final table split
            AdvanceRow();
        }

        /// <summary>
        /// Creates a title for this Excel document
        /// </summary>
        /// <param name="titleText"></param>
        public override void CreateTitle(string titleText)
        {
            // don't start at the top of the table
            AdvanceRow(0);

            // get the range of the title
            Excel.Range titleRowRange = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);

            // merge
            titleRowRange.Merge();
            titleRowRange.RowHeight = 30;

            // set the title
            titleRowRange.Cells.Font.Name = "Times New Roman";
            titleRowRange.Cells.Font.Size = 14;
            titleRowRange.Cells.Font.Bold = 1;
            titleRowRange.Cells.Font.Color = Excel.XlRgbColor.rgbBlack;
            // titleRowRange.Cells.Interior.Color = Excel.XlRgbColor.rgbLightGray;
            // titleRowRange.Cells.Borders.Color = Excel.XlRgbColor.rgbBlack;
            // titleRowRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            titleRowRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRowRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
            titleRowRange.Cells.Value = titleText;

            // increment the current row
            AdvanceRow();
            AutoSizeWorksheet();
        }

        /// <summary>
        /// Creates an Excel table heading
        /// </summary>
        /// <param name="headingText"></param>
        public override void CreateHeader(string headingText)
        {
            // get the range of the title
            Excel.Range titleRowRange = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);

            // merge
            titleRowRange.Merge();
            titleRowRange.RowHeight = 15;

            // set the title
            titleRowRange.Cells.Font.Name = "Calibri";
            titleRowRange.Cells.Font.Size = 14;
            titleRowRange.Cells.Font.Bold = 1;
            titleRowRange.Cells.Font.Color = Excel.XlRgbColor.rgbBlack;
            titleRowRange.Cells.Interior.Color = Excel.XlRgbColor.rgbWhite;
            // titleRowRange.Cells.Borders.Color = Excel.XlRgbColor.rgbBlack;
            // titleRowRange.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            titleRowRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            titleRowRange.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            titleRowRange.Cells.Value = headingText;

            AdvanceRow(0);
            AutoSizeWorksheet();
        }

        /// <summary>
        /// Creates the header graphic for this table
        /// </summary>
        /// <param name="path"></param>
        public override void CreateHeaderGraphic(string path)
        {
            // add header graphics
            // save the graphic before its path can be referenced
            if (!File.Exists(Utilities.GetExecutingPath() + "ACAS.jpg"))
            {
                Resources.Resources.ACAS.Save(@"ACAS.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }

            // height
            int height = 70;
            int width = 125;

            // add a picture to the worksheet
            worksheet.Shapes.AddPicture(Utilities.GetExecutingPath() + "ACAS.jpg", 
                Microsoft.Office.Core.MsoTriState.msoFalse, 
                Microsoft.Office.Core.MsoTriState.msoCTrue, 5, 5, width, height);

            // resize the first row to avoid a border issue
            Excel.Range range = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);
            range.RowHeight = height + 1;

            // resize the first columnm
            range = GetSingleCellRange(worksheet, currentColumnOffset, 1);
            range.ColumnWidth = width + 1;

            // go forward 
            AdvanceRow(0);
        }

        /// <summary>
        /// Creates an error message in the table
        /// </summary>
        /// <param name="message"></param>
        public override void CreateErrorMessage(string message)
        {
            // autosize and theme, with error message
            AddVerticalTableRow(new string[] { message }, true);

            // set sizing and theming
            AutoSizeWorksheet();
            SetDefaultTheme(false);
        }

        /// <summary>
        /// Add a row to an Excel sheet
        /// </summary>
        /// <param name="columnValues"></param>
        private void AddVerticalTableRow(string[] columnValues, bool merge)
        {
            // get a range
            Excel.Range range = null;

            // set column values
            for (int i = 0; i < columnValues.Count(); i++)
            {
                range = GetSingleCellRange(this.worksheet, currentColumnOffset + i, currentRow);
                range.Value = columnValues[i];
                if (i == 0) { 
                    range.EntireColumn.ColumnWidth = 24;
                }
                if (i == 0 && currentRow != 0)
                {
                    // assume ID column, idk why it should be?
                    if (!columnValues[0].Equals("ID"))
                    {
                        range.Hyperlinks.Add(range, settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems#_a=edit&id="
                            + columnValues[0] + "&triage=true", Type.Missing, Type.Missing, columnValues[0]);
                        range.Font.Size = 10;
                        range.Font.Name = "Arial";
                        range.Font.Bold = 1;
                        range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    }
                }
            }

            // set row height
            range.EntireRow.RowHeight = 33;

            // merge if supplied
            if (merge)
            {
                range = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);
                range.Merge();
            }

            // increase the current row count
            this.currentRow++;
        }

        /// <summary>
        /// Splits the table
        /// </summary>
        private void AdvanceRow(int numExtraRows = 1)
        {
            currentRow++;
            for (int i = 0; i < numExtraRows; i++) { currentRow++; }
        }

        /// <summary>
        /// Autosizes the workbook
        /// </summary>
        private void AutoSizeWorksheet()
        {
            Excel.Range sized = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);
            sized.Columns.AutoFit();
        }

        /// <summary>
        /// Formats anything post document creation
        /// </summary>
        public override void CreateDocumentSpecificPostFormatting(bool wide = false)
        {
            this.worksheet.UsedRange.WrapText = true;

            for (int i = 1; i <= this.worksheet.UsedRange.Columns.Count + 1; i++)
            {
                Excel.Range columnCell = GetSingleCellRange(worksheet, i, 1);
                if (!wide)
                {
                    if (i >= currentColumnOffset)
                    {
                        columnCell.EntireColumn.ColumnWidth = 28;
                    }
                    else
                    {
                        columnCell.EntireColumn.ColumnWidth = 10;
                    }
                }
                else
                {
                    if (i >= currentColumnOffset)
                    {
                        columnCell.EntireColumn.ColumnWidth = 50;
                    }
                    else
                    {
                        columnCell.EntireColumn.ColumnWidth = 10;
                    }
                }
            }

            for (int i = 1; i <= this.worksheet.UsedRange.Rows.Count; i++)
            {
                Excel.Range rowCell = GetSingleCellRange(worksheet, 1, i);
                rowCell.EntireRow.RowHeight = 28;
            }			
        }

        /// <summary>
        /// Creates a new page with a given name (if Excel)
        /// </summary>
        /// <param name="worksheetName"></param>
        public override void CreateNewPage(string worksheetName)
        {
            // post format the previous document
            this.CreateDocumentSpecificPostFormatting();

            // worksheet information
            worksheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, worksheet);
            worksheet.Select();
            worksheet.Name = worksheetName;

            // excel positioning vars
            this.starterRow = 1;
            this.currentRow = 2;
            this.currentColumnCount = 4;
            this.currentColumnOffset = 2;

            // rem grid lines
            app.ActiveWindow.DisplayGridlines = false;
        }

        /// <summary>
        /// Sets the workbooks default theme
        /// </summary>
        private void SetDefaultTheme(bool header)
        {
            Excel.Range styled = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);
            Excel.XlYesNoGuess headerExists = Excel.XlYesNoGuess.xlNo;
            if (header)
                headerExists = Excel.XlYesNoGuess.xlYes;
            worksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, 
                styled, Type.Missing, headerExists, Type.Missing).Name = "TableStyle";
            worksheet.ListObjects.Item["TableStyle"].TableStyle = "TableStyleMedium2";
        }

        /// <summary>
        /// Sets a basic theme for the currently selected table range
        /// </summary>
        private void SetBasicTheme(bool bordersThemed)
        {
            Excel.Range styled = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);

            if (bordersThemed)
            {
                styled.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Excel.XlRgbColor.rgbBlack;
                styled.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = Excel.XlRgbColor.rgbBlack;
                styled.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Excel.XlRgbColor.rgbBlack;
                styled.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Excel.XlRgbColor.rgbBlack;
                styled.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Color = Excel.XlRgbColor.rgbBlack;
                styled.Borders[Excel.XlBordersIndex.xlInsideVertical].Color = Excel.XlRgbColor.rgbBlack;
            }
        }

        /// <summary>
        /// Gets a range block in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <returns></returns>
        private Excel.Range GetBlockedRange(Excel.Worksheet currentSheet, char firstCol, char lastCol, int firstRow, int lastRow)
        {
            return (Excel.Range)currentSheet.Range[firstCol.ToString() + firstRow.ToString(), lastCol.ToString() + lastRow.ToString()];
        }

        /// <summary>
        /// Gets a range block in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="firstRow"></param>
        /// <param name="lastRow"></param>
        /// <returns></returns>
        private Excel.Range GetBlockedRange(Excel.Worksheet currentSheet, int firstCol, int lastCol, int firstRow, int lastRow)
        {
            return (Excel.Range)currentSheet.Range[currentSheet.Cells[firstRow, firstCol], currentSheet.Cells[lastRow, lastCol]];
        }

        /// <summary>
        /// Gets a row in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private Excel.Range GetMultiCellRange(Excel.Worksheet currentSheet, char firstCol, char lastCol, int row)
        {
            return (Excel.Range)currentSheet.Range[firstCol.ToString() + row.ToString() + ":" + lastCol.ToString() + row.ToString(), Type.Missing];
        }

        /// <summary>
        /// Gets a row in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private Excel.Range GetMultiCellRange(Excel.Worksheet currentSheet, int firstCol, int lastCol, int row)
        {
            return (Excel.Range)currentSheet.Range[currentSheet.Cells[row, firstCol], currentSheet.Cells[row, lastCol]];
        }

        /// <summary>
        /// Gets a single cell in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private Excel.Range GetSingleCellRange(Excel.Worksheet currentSheet, char col, int row)
        {
            return (Excel.Range)currentSheet.Range[col.ToString() + row.ToString(), Type.Missing];
        }

        /// <summary>
        /// Gets a single cell in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private Excel.Range GetSingleCellRange(Excel.Worksheet currentSheet, int col, int row)
        {
            return (Excel.Range)currentSheet.Range[currentSheet.Cells[row, col], currentSheet.Cells[row, col]];
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~ExcelGenerator()
        {
            try
            {
                // remove user control
                app.UserControl = false;

                // save this workbook in the application directory
                workbook.SaveAs(Utilities.GetExecutingPath() + settings["Project Name"] + " " + settings["Iteration"] + " Release Notes.xlsx",
                    Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // quit
                app.Workbooks.Close();
                app.Quit();


                // unmarshall all COM objects
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);

                // set to null
                worksheet = null;
                workbook = null;
                app = null;
            }
            catch (COMException e)
            {
                // exception, up to system to free objects
                // once program is gone
                (new Logger())
                    .SetLoggingType(Logger.Type.Warning)
                    .SetLoggingSilenceState(this.silent)
                    .SetMessage(e.Message + "\n Excel may not have been freed from user control, \n" +
                                            "is waiting on user save, \nor cannot save (another open workbook?).")
                    .Display();
                throw;
            }

            // collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
