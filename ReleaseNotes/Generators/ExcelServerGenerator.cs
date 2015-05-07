using ReleaseNotesLibrary.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.Table;
using System.Runtime.InteropServices;
using System.Threading;

namespace ReleaseNotesLibrary.Generators
{
    public class ExcelServerGenerator : ReleaseNotesGenerator
    {
        // excel persistent objects
        private ExcelPackage app;
        private ExcelWorkbook workbook;
        private ExcelWorksheet worksheet;

        // excel positioning vars
        private int starterRow = 1;
        private int currentRow = 1;
        private int currentColumnCount = 4;
        private int currentColumnOffset = 2;
        private const int totalAllowedColumns = 24;

        private static int tableNumber = 1;

        // result
        private MemoryStream ms;

        // alphabet (for columns)
        private char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        /// <summary>
        /// Generates an excel instance
        /// </summary>
        /// <param name="settings"></param>
        private ExcelServerGenerator(NamedLookup settings, bool silent)
            : base(settings, silent)
        {
            ms = new MemoryStream();
            app = new ExcelPackage(ms);
            workbook = app.Workbook;
            app.Workbook.Worksheets.Add("Release Notes");
            app.Workbook.Worksheets.MoveToStart("Release Notes");
            worksheet = app.Workbook.Worksheets[1];
            worksheet.Name = "Release Notes";
            worksheet.View.ShowGridLines = false;
        }

        /// <summary>
        /// Generates an Excel instance
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static ExcelServerGenerator ExcelServerGeneratorFactory(NamedLookup settings, bool silent)
        {
            try
            {
                return new ExcelServerGenerator(settings, silent);
            }
            catch (Exception e)
            {
                (new Logger())
                    .SetLoggingType(Logger.Type.Error)
                    .SetMessage(e.Message + "Excel not initialized. \n Are you trying to run this server-side?...")
                    .Display();
                return null;
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
                    ExcelRange cellRange = GetSingleCellRange(worksheet, j + currentColumnOffset - 1, currentRow);

                    string currentKey = "";
                    if (counter != tableKeys.Count())
                        currentKey = tableKeys.ElementAt(counter);

                    worksheet.Row(i).Height = 18;
                    cellRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cellRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cellRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    cellRange.Style.Font.Bold = true;
                    cellRange.Style.Font.Size = 10;
                    cellRange.Style.Font.Name = "Arial";

                    if (j % 2 != 0)
                    {
                        cellRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        cellRange.Value = currentKey;
                    }
                    else
                    {
                        if (counter != tableKeys.Count())
                        {
                            cellRange.Style.Font.Color.SetColor(Color.FromArgb(0, 112, 192));
                            if (currentKey.Equals("Source"))
                            {
                                // hyperlink
                                cellRange.Style.Font.Name = "Arial";
                                cellRange.Style.Font.Size = 10;
                                cellRange.Style.Font.Bold = true;
                                cellRange.Value = settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems" + Environment.NewLine + data[currentKey];
                                string address = settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems";
                                cellRange.Hyperlink = new Uri(address);
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
                        cellRange.Style.Fill.BackgroundColor.SetColor(Color.White);
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
            ExcelRange titleRowRange = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);

            // merge
            titleRowRange.Merge = true;
            // side effects, blahhhh
            worksheet.Row(this.currentRow).Height = 30;

            // set the title
            titleRowRange.Style.Font.Name = "Times New Roman";
            titleRowRange.Style.Font.Size = 14;
            titleRowRange.Style.Font.Bold = true;
            titleRowRange.Style.Font.Color.SetColor(Color.Black);
            titleRowRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            titleRowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            titleRowRange.Value = titleText;

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
            ExcelRange titleRowRange = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);

            // merge
            titleRowRange.Merge = true;
            // side effects, blahhhh
            worksheet.Row(this.currentRow).Height = 15;

            // set the title
            titleRowRange.Style.Font.Name = "Calibri";
            titleRowRange.Style.Font.Size = 14;
            titleRowRange.Style.Font.Bold = true;
            titleRowRange.Style.Font.Color.SetColor(Color.Black);
            titleRowRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            titleRowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            titleRowRange.Value = headingText;

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
            int height = 70;
            int width = 125;
            try
            {
                if (!File.Exists(Global.appDataPath))
                {
                    Image i = Resources.Resources.ACAS;
                    // recommended before save
                    Thread.Sleep(30);
                    i.Save(Global.appDataPath + "ACAS.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                // add a picture to the worksheet
                var logo = Image.FromFile(new Uri(Global.appDataPath + "ACAS.jpg").LocalPath);
                var docPicture = worksheet.Drawings.AddPicture("Logo", logo);
                docPicture.SetPosition(0, 0);
                docPicture.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
                docPicture.SetSize(65, 75);
            }
            catch (ExternalException e)
            {
                (new Logger())
                    .SetLoggingType(Logger.Type.Warning)
                    .SetMessage(e.Message + "Image could not be saved server side")
                    .Display();
            }

            // resize the first row to avoid a border issue
            worksheet.Row(this.currentRow).Height = height + 1;

            // resize the first columnm
            worksheet.Column(currentColumnOffset).Width = width + 1;

            // go forward 
            AdvanceRow(0);
        }

        /// <summary>
        /// Creates a section to display one piece of information
        /// </summary>
        /// <param name="headername"></param>
        /// <param name="text"></param>
        /// <param name="hyperlink"></param>
        public override void CreateNamedSection(string headername, string text, string hyperlink)
        {
            NamedLookup namedSectionData = new NamedLookup(headername);
            namedSectionData[text] = hyperlink;
            CreateHorizontalTable(namedSectionData, 1, true);
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
            ExcelRange range = null;

            // set column values
            for (int i = 0; i < columnValues.Count(); i++)
            {
                range = GetSingleCellRange(this.worksheet, currentColumnOffset + i, currentRow);
                range.Value = columnValues[i];
                range.Style.WrapText = true;
                if (i == 0)
                {
                    worksheet.Column(i + 1).Width = 24;
                }
                if (i == 0 && currentRow != 0)
                {
                    // assume ID column, idk why it should be?
                    if (!columnValues[0].Equals("ID"))
                    {
                        string address = settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems#_a=edit&id="
                            + columnValues[0] + "&triage=true";
                        range.Hyperlink = new Uri(address);
                        range.Value = columnValues[i];
                        range.Style.Font.Size = 10;
                        range.Style.Font.Name = "Arial";
                        range.Style.Font.Bold = true;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }
                }
            }

            // set row height
            worksheet.Row(this.currentRow).Height = 33;

            // merge if supplied
            if (merge)
            {
                range = GetMultiCellRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, currentRow);
                range.Merge = true;
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
            ExcelRange sized = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);
            sized.AutoFitColumns();
        }

        /// <summary>
        /// Formats anything post document creation
        /// </summary>
        public override void CreateDocumentSpecificPostFormatting(bool wide = false)
        {
            GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow).Style.WrapText = true;

            for (int i = this.worksheet.Dimension.Start.Column + 1; i <= this.worksheet.Dimension.End.Column; i++)
            {
                if (!wide)
                {
                    if (i >= currentColumnOffset)
                    {
                        worksheet.Column(i).Width = 28;
                    }
                    else
                    {
                        worksheet.Column(i).Width = 10;
                    }
                }
                else
                {
                    if (i >= currentColumnOffset)
                    {
                        worksheet.Column(i).Width = 50;
                    }
                    else
                    {
                        worksheet.Column(i).Width = 10;
                    }
                }
            }

            for (int i = this.worksheet.Dimension.Start.Row 
                + 1; i <= this.worksheet.Dimension.End.Row; i++)
            {
                worksheet.Row(i).Height = 28;
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
            worksheet = (ExcelWorksheet) workbook.Worksheets.Add(worksheetName);
            worksheet.View.ShowGridLines = false;

            // excel positioning vars
            this.starterRow = 1;
            this.currentRow = 2;
            this.currentColumnCount = 4;
            this.currentColumnOffset = 2;
        }

        /// <summary>
        /// Sets the workbooks default theme
        /// </summary>
        private void SetDefaultTheme(bool header)
        {
            ExcelRange styled = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);
            ExcelTable table = worksheet.Tables.Add(styled, "table" + tableNumber.ToString());
            if (header)
                table.ShowHeader = true;
            else
                table.ShowHeader = false;
            table.TableStyle = TableStyles.Medium2;
            tableNumber++;
        }

        /// <summary>
        /// Sets a basic theme for the currently selected table range
        /// </summary>
        private void SetBasicTheme(bool bordersThemed)
        {
            ExcelRange styled = GetBlockedRange(worksheet, currentColumnOffset, currentColumnCount + currentColumnOffset - 1, starterRow, currentRow);
            if (bordersThemed)
            {
                styled.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                SetAllBorders(styled, ExcelBorderStyle.Thin);
            }
            styled.Style.WrapText = true;
        }

        /// <summary>
        /// Sets all borders in a given range (thanks EPPlus and Codeplex)
        /// </summary>
        /// <param name="this"></param>
        /// <param name="borderStyle"></param>
        private void SetAllBorders(ExcelRange @this, ExcelBorderStyle borderStyle)
        {
            var fromRow = @this.Start.Row;
            var fromCol = @this.Start.Column;
            var toRow = @this.End.Row;
            var toCol = @this.End.Column;

            var numRows = toRow - fromRow + 1;
            var numCols = toCol - fromCol + 1;

            @this.Style.Border.BorderAround(borderStyle);

            for (var rowOffset = 1; rowOffset < numRows; rowOffset += 2)
            {
                var row = @this.Offset(rowOffset, 0, 1, numCols);
                row.Style.Border.BorderAround(borderStyle);
            }

            for (var colOffset = 1; colOffset < numCols; colOffset += 2)
            {
                var col = @this.Offset(0, colOffset, numRows, 1);
                col.Style.Border.BorderAround(borderStyle);
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
        private ExcelRange GetBlockedRange(ExcelWorksheet currentSheet, char firstCol, char lastCol, int firstRow, int lastRow)
        {
            return currentSheet.Cells[firstCol.ToString() + firstRow.ToString() + ":" + lastCol.ToString() + lastRow.ToString()];
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
        private ExcelRange GetBlockedRange(ExcelWorksheet currentSheet, int firstCol, int lastCol, int firstRow, int lastRow)
        {
            return currentSheet.Cells[firstRow, firstCol, lastRow, lastCol];
        }

        /// <summary>
        /// Gets a row in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private ExcelRange GetMultiCellRange(ExcelWorksheet currentSheet, char firstCol, char lastCol, int row)
        {
            return currentSheet.Cells[firstCol.ToString() + row.ToString() + ":" + lastCol.ToString() + row.ToString()];
        }

        /// <summary>
        /// Gets a row in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="firstCol"></param>
        /// <param name="lastCol"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private ExcelRange GetMultiCellRange(ExcelWorksheet currentSheet, int firstCol, int lastCol, int row)
        {
            return currentSheet.Cells[row, firstCol, row, lastCol];
        }

        /// <summary>
        /// Gets a single cell in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private ExcelRange GetSingleCellRange(ExcelWorksheet currentSheet, char col, int row)
        {
            return currentSheet.Cells[col.ToString() + row.ToString()];
        }

        /// <summary>
        /// Gets a single cell in Excel
        /// </summary>
        /// <param name="currentSheet"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <returns></returns>
        private ExcelRange GetSingleCellRange(ExcelWorksheet currentSheet, int col, int row)
        {
            return currentSheet.Cells[row, col, row, col];
        }

        public override byte[] Save()
        {
            app.Save();
            // save this workbook in the application directory
            string filePath = Utilities.GetExecutingPath() + settings["Project Name"] + " " + settings["Iteration"] + " Release Notes.xlsx";
            FileInfo f = new FileInfo(new Uri(filePath).LocalPath);
            if (f.Exists) f.Delete();
            FileStream fs = f.Create();
            // stream to output
            fs.Position = 0;
            ms.WriteTo(fs);
            ms.Close();
            fs.Close();
            return ms.ToArray();
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~ExcelServerGenerator()
        {
            try
            {
                // set to null
                worksheet = null;
                workbook = null;
                app = null;
            }
            catch (Exception e)
            {
                // exception, up to system to free objects
                // once program is gone
                (new Logger())
                    .SetLoggingType(Logger.Type.Warning)
                    .SetLoggingSilenceState(this.silent)
                    .SetMessage(e.Message + "\n EPPlus encountered an error saving.")
                    .Display();
                throw;
            }
        }
    }
}
