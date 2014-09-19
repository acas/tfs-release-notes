using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using ReleaseNotes.Utility;

namespace ReleaseNotes
{
    class ExcelGenerator : ReleaseNotesGenerator
    {
        // excel persistent object
        private Excel.Application app;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private Excel.Range currentRange;
        private int starterRow = 1;
        private int currentRow = 1;
        private int currentColumnCount = 1;
        private int currentColumnOffset = 1;
        private const int totalAllowedRows = 10;

        // alphabet (for columns)
        private char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        /// <summary>
        /// Generates an excel instance
        /// </summary>
        /// <param name="settings"></param>
        private ExcelGenerator(NamedLookup settings) : base(settings)
        {
            app = new Excel.Application();
            app.Visible = !this.silent;
            app.UserControl = !this.silent;
            workbook = (Excel.Workbook)app.Workbooks.Add();
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = "Release Notes";
        }

        /// <summary>
        /// Generates an Excel instance
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static ExcelGenerator ExcelGeneratorFactory(NamedLookup settings)
        {
            try
            {
                return new ExcelGenerator(settings);
            }
            catch (COMException e)
            {
                (new Logger())
                    .setType(Logger.Type.Error)
                    .setMessage(e.Message + "Excel not initialized. \n Are you trying to run this server-side?...")
                    .display();
                return null;
            }
        }

        /// <summary>
        /// Creates a vertical style table
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="headerText"></param>
        /// <param name="header"></param>
        public override void createVerticalTable(DataTable dataTable, string headerText, bool header)
        {
            // set the current column count
            this.starterRow = currentRow;
            this.currentColumnCount = dataTable.Columns.Count;

            // add header row
            addVerticalTableRow(Utilities.tableColumnsToStringArray(dataTable), false);

            // add table information
            foreach (DataRow row in dataTable.Rows)
                addVerticalTableRow(Utilities.tableRowToStringArray(row), false);

            // set sizing and theming
            autoSize();
            setDefaultTheme(header);

            // move ahead a row
            this.currentRow++;
        }

        /// <summary>
        /// Creates a title for this Excel document
        /// </summary>
        /// <param name="titleText"></param>
        public override void createTitle(string titleText)
        {
  
        }

        /// <summary>
        /// Creates an Excel table heading
        /// </summary>
        /// <param name="headingText"></param>
        public override void createHeading(string headingText)
        {
            
        }

        /// <summary>
        /// Creates an error message in the table
        /// </summary>
        /// <param name="message"></param>
        public override void createErrorMessage(string message)
        {
            // autosize and theme, with error message
            addVerticalTableRow(new string[] { message }, true);

            // set sizing and theming
            autoSize();
            setDefaultTheme(false);
        }

        /// <summary>
        /// Add a row to an Excel sheet
        /// </summary>
        /// <param name="columnValues"></param>
        private void addVerticalTableRow(string[] columnValues, bool merge)
        {
            // set column values
            for (int i = 0; i < columnValues.Count(); i++)
            {
                currentRange = getSingleCellRange(this.worksheet, i + 1, currentRow);
                currentRange.Value = columnValues[i];
                if (i == 0) { currentRange.EntireColumn.ColumnWidth = 24; }
            }

            // set row height
            currentRange.EntireRow.RowHeight = 24;

            // merge if supplied
            if (merge)
            {
                currentRange = getMultiCellRange(worksheet, currentColumnOffset, currentColumnCount, currentRow);
                currentRange.Merge();
            }

            // increase the current row count
            this.currentRow++;
        }

        /// <summary>
        /// Autosizes the workbook
        /// </summary>
        private void autoSize()
        {
            Excel.Range sized = getBlockedRange(worksheet, currentColumnOffset, currentColumnCount, starterRow, currentRow);
            sized.Columns.AutoFit();
        }

        /// <summary>
        /// Sets the workbooks default theme
        /// </summary>
        private void setDefaultTheme(bool header)
        {
            Excel.Range styled = getBlockedRange(worksheet, currentColumnOffset, currentColumnCount, starterRow, currentRow);
            Excel.XlYesNoGuess headerExists = Excel.XlYesNoGuess.xlNo;
            if (header)
                headerExists = Excel.XlYesNoGuess.xlYes;
            worksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, 
                styled, Type.Missing, headerExists, Type.Missing).Name = "TableStyle";
            worksheet.ListObjects.Item["TableStyle"].TableStyle = "TableStyleMedium2";
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
        private Excel.Range getBlockedRange(Excel.Worksheet currentSheet, char firstCol, char lastCol, int firstRow, int lastRow)
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
        private Excel.Range getBlockedRange(Excel.Worksheet currentSheet, int firstCol, int lastCol, int firstRow, int lastRow)
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
        private Excel.Range getMultiCellRange(Excel.Worksheet currentSheet, char firstCol, char lastCol, int row)
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
        private Excel.Range getMultiCellRange(Excel.Worksheet currentSheet, int firstCol, int lastCol, int row)
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
        private Excel.Range getSingleCellRange(Excel.Worksheet currentSheet, char col, int row)
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
        private Excel.Range getSingleCellRange(Excel.Worksheet currentSheet, int col, int row)
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
                workbook.SaveAs(Utilities.getExecutingPath() + settings["Project Name"] + " " + settings["Iteration"] + " Release Notes.xlsx",
                    Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // quit
                app.Workbooks.Close();
                app.Quit();


                // unmarshall all COM objects
                Marshal.ReleaseComObject(currentRange);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(app);

                // set to null
                currentRange = null;
                worksheet = null;
                workbook = null;
                app = null;
            }
            catch (COMException e)
            {
                // exception, up to system to free objects
                // once program is gone
                (new Logger())
                    .setType(Logger.Type.Warning)
                    .setSilence(this.silent)
                    .setMessage(e.Message + "\n Excel may not have been freed from user control, \n" +
                                            "is waiting on user save, \nor cannot save (another open workbook?).")
                    .display();
            }

            // collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
