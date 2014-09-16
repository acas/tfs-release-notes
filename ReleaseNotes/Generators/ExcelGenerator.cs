using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace ReleaseNotes
{
    class ExcelGenerator : ReleaseNotesGenerator
    {
        // excel persistent object
        private Excel.Application app;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private Excel.Range currentRange;
        private int currentRow = 1;

        // alphabet (for columns)
        private char[] alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        /// <summary>
        /// Excel generator constructor
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="isVisible"></param>
        /// <param name="silent"></param>
        /// <param name="logger"></param>
        private ExcelGenerator(string worksheetName, bool silent, Logger logger)
        {
            app = new Excel.Application();
            this.silent = silent;
            app.Visible = !this.silent;
            workbook = (Excel.Workbook) app.Workbooks.Add();
            worksheet = (Excel.Worksheet) workbook.ActiveSheet;
            worksheet.Name = worksheetName;
            this.logger = logger.setSilence(this.silent);
        }

        /// <summary>
        /// Factory for generating Excel writers
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="isVisible"></param>
        /// <param name="silent"></param>
        /// <param name="logger"></param>
        /// <returns>A release notes generator for Excel</returns>
        public static ExcelGenerator ExcelGeneratorFactory(string worksheetName = "Unknown", bool silent = false, Logger logger = null)
        {
            try
            {
                if (logger == null) logger = new Logger();
                return new ExcelGenerator(worksheetName, silent, logger);
            }
            catch (COMException e)
            {
                (new Logger())
                    .setType(Logger.Type.Error)
                    .setSilence(silent)
                    .setMessage(e.Message + "\n Are you trying to run this server-side?...")
                    .display();
                return null;
            }
        }

        /// <summary>
        /// Documented in super
        /// </summary>
        public override void generateReleaseNotes()
        {
            // create excel writer
            logger.setMessage("Generating Excel release notes table...").setType(Logger.Type.Information).display();

            try
            {
                // add header row
                addRow("#", "ID", "Work Item Type", "Title", "Area Path", "Iteration", "Description");

                // get release notes work item collection
                WorkItemCollection c = TFSAccessor.TFSAccessorFactory().getReleaseNotesFromQuery(this.projectName, this.iterationPath);
                if (c == null) throw new Exception("Work items could not be retrieved.");
                
                // add table information
                int counter = 1;
                foreach (WorkItem i in c)
                {
                    addRow(counter.ToString(), i.Id.ToString(), i.Type.Name, i.Title.ToString(),
                        i.AreaPath, i.IterationPath, StripHtml.StripHtmlContrived(i.Description, true));
                    counter++;
                }

                // set sizing and theming
                autoSize();
                setDefaultTheme();
                
                // done!
                logger.setType(Logger.Type.Success)
                    .setMessage("Table generated.")
                    .display();
            }
            catch (Exception e)
            {
                // autosize and theme, with error message
                addRow(e.Message);
                
                // set sizing and theming
                autoSize();
                setDefaultTheme();
                logger.setType(Logger.Type.Error)
                    .setMessage("Table not generated. " + e.Message)
                    .display();
            }

            // give the user control
            giveUserControl();
        }

        /// <summary>
        /// Add a row to an Excel sheet
        /// </summary>
        /// <param name="columnNames"></param>
        public void addRow(params string[] columnNames)
        {
            for (int i = 0; i < columnNames.Count(); i++) {
                currentRange = getSingleCellRange(this.worksheet, i + 1, currentRow);
                currentRange.Value = columnNames[i];
                if (i == 0) { currentRange.EntireColumn.ColumnWidth = 24; }
            }
            currentRange.EntireRow.RowHeight = 24;
            this.currentRow++;
        }

        /// <summary>
        /// Gives the user control of the workbook
        /// </summary>
        public void giveUserControl(bool userControl = true)
        {
            app.UserControl = userControl;
        }

        /// <summary>
        /// Autosizes the workbook
        /// </summary>
        public void autoSize() {
            worksheet.UsedRange.Columns.AutoFit();
        }

        /// <summary>
        /// Sets the workbooks default theme
        /// </summary>
        public void setDefaultTheme()
        {
            Excel.Range styled = worksheet.UsedRange;
            worksheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, styled, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "TableStyle";
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
            return (Excel.Range) currentSheet.Range[firstCol.ToString() + firstRow.ToString(), lastCol.ToString() + lastRow.ToString()];
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
            return (Excel.Range) currentSheet.Range[currentSheet.Cells[firstRow, firstCol], currentSheet.Cells[lastRow, lastCol]];
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
            return (Excel.Range) currentSheet.Range[firstCol.ToString() + row.ToString() + ":" + lastCol.ToString() + row.ToString(), Type.Missing];
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
            return (Excel.Range) currentSheet.Range[currentSheet.Cells[row, firstCol], currentSheet.Cells[row, lastCol]];
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
            return (Excel.Range) currentSheet.Range[col.ToString() + row.ToString(), Type.Missing];
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
            return (Excel.Range) currentSheet.Range[currentSheet.Cells[row, col], currentSheet.Cells[row, col]];
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~ExcelGenerator()
        {
            if (app != null)
            {
                // remove user control
                app.UserControl = false;

                try
                {
                    // save this workbook in the application directory
                    workbook.SaveAs(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
                            + "/" + worksheet.Name + ".xlsx",
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
                        .setMessage(e.Message + "\n Excel may not have been freed from user control, " + 
                                                "is waiting on user save, or cannot save (another open workbook?).")
                        .display();   
                }
            }
            //collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
