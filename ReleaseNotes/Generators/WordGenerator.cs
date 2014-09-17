using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.IO;
using System.Threading;

namespace ReleaseNotes
{
    class WordGenerator : ReleaseNotesGenerator
    {
        private Word.Application app;
        private Word.Document document;

        /// <summary>
        /// Constructor for a word generator object
        /// </summary>
        /// <param name="documentName"></param>
        /// <param name="silent"></param>
        /// <param name="logger"></param>
        private WordGenerator(string documentName, bool silent, Logger logger)
        {
            app = new Word.Application();
            this.silent = silent;
            app.Visible = !this.silent;
            document = app.Documents.Add(Type.Missing, Type.Missing, Word.WdNewDocumentType.wdNewBlankDocument, !this.silent);
            document.UserControl = !this.silent;
            this.logger = logger.setSilence(this.silent);
        }

        /// <summary>
        /// Constructor (factory method) for a word generator object
        /// </summary>
        /// <param name="documentName"></param>
        /// <param name="isVisible"></param>
        /// <param name="silent"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static WordGenerator WordGeneratorFactory(string documentName = "Unknown", bool silent = false, Logger logger = null)
        {
            try
            {
                if (logger == null) logger = new Logger();
                return new WordGenerator(documentName, silent, logger);
            }
            catch (COMException e)
            {
                (new Logger())
                    .setSilence(silent)
                    .setType(Logger.Type.Error)
                    .setMessage(e.Message).display();
                return null;
            }
        }

        /// <summary>
        /// Documented in super
        /// </summary>
        public override void generateReleaseNotes()
        {
            // create excel writer
            logger.setMessage("Generating Word release notes document.")
                .setType(Logger.Type.Information)
                .display();

            // try to generate the document
            try
            {
                // set margins
                document.PageSetup.LeftMargin = app.InchesToPoints(0.25F);
                document.PageSetup.TopMargin = app.InchesToPoints(0.25F);
                document.PageSetup.BottomMargin = app.InchesToPoints(0.5F);
                document.PageSetup.RightMargin = app.InchesToPoints(0.25F);

                // set header distance (for top)
                document.PageSetup.HeaderDistance = app.InchesToPoints(0.25F);

                // thread sleep to allow COM interop to catch up
                Thread.Sleep(100);

                // add header graphics
                // save the graphic before its path can be referenced
                if (!File.Exists(Utilities.getExecutingPath() + "ACAS.jpg"))
                {
                    Resources.Resources.ACAS.Save(@"ACAS.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                // get the word document range of the header
                Word.Section firstSection = document.Sections.First;
                Word.Range headerSectionRange = firstSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                // get the shape back, put in the header and resize
                Word.InlineShape ACASLogo = headerSectionRange.InlineShapes.AddPicture(Utilities.getExecutingPath() + "ACAS.jpg", false, true);

                // thread sleep to allow COM interop to catch up
                Thread.Sleep(100);

                // scale height and width
                ACASLogo.ScaleHeight = 55.0F;
                ACASLogo.ScaleWidth = 55.0F;

                // get range of the first paragraph
                Word.Paragraph titleParagraph = document.Paragraphs.Add();
                Word.Range headerRange = titleParagraph.Range;

                // create the document title
                headerRange.Font.Name = "Times New Roman";
                headerRange.Font.Size = 12;
                headerRange.Text = "APPLICATION BUILD/RELEASE NOTES\n";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.Bold = 1; // true

                // connect to TFS
                TFSAccessor TFS = TFSAccessor.TFSAccessorFactory();

                // add another paragraph
                Word.Paragraph programInformationParagraph = document.Paragraphs.Add();

                // aggregate information
                Dictionary<string, string> informationTableData = new Dictionary<string, string>();
                informationTableData.Add("Application", this.projectName);
                informationTableData.Add("Release Date", DateTime.Now.ToShortDateString());
                informationTableData.Add("Release", this.projectName + " " + this.iterationPath);
                informationTableData.Add("Iteration (Sprint) #", "Unavailable");
                informationTableData.Add("Build #", "Unavailable");

                // create application information table
                Word.Table programInformationTable = createHorizontalStackedTable(programInformationParagraph.Range, 2, informationTableData, Word.WdRowAlignment.wdAlignRowCenter);

                // split
                insertTableSplit(programInformationParagraph);

                // new heading
                createHeading("Access", false);
                Word.Paragraph accessParagraph = document.Paragraphs.Add();
                accessParagraph.Range.Text = "Application is accessible at: https://" + projectName.ToLowerInvariant() + ".americancapital.com/";

                // several indents needed
                for (int i = 0; i < 3; i++) { accessParagraph.Indent(); }

                // split
                insertTableSplit(accessParagraph);

                // new heading
                createHeading("Details");

                // add another paragraph
                Word.Paragraph programServerParagraph = document.Paragraphs.Add();

                // aggregate information
                Dictionary<string, string> programServerData = new Dictionary<string, string>();
                programServerData.Add("Web Server", "pro00websv01.acas.corp.americancapital.com");
                programServerData.Add("Database Server", "SQLNONFinancialCluster02.americancapital.com");
                programServerData.Add("Database", projectName);
                programServerData.Add("Source", Settings.Settings.Default.TFSServer
                    + projectName + "/_versionControl\n"
                    + "(Changeset: " + "Unavailable" + ")");

                // create application information table
                Word.Table programServerTable = createHorizontalStackedTable(programServerParagraph.Range, 1, 
                    programServerData, Word.WdRowAlignment.wdAlignRowCenter);

                // split
                insertTableSplit(programServerParagraph);

                // create a heading
                createHeading("Included Requirements");

                // get release notes work item data table
                DataTable workItemsDataTable = TFS.getReleaseNotesAsDataTable(this.projectName, this.iterationPath);
                if (workItemsDataTable == null) throw new Exception("Work items table could not be retrieved.");

                // add another paragraph
                Word.Paragraph workItemsParagraph = document.Paragraphs.Add();

                // create work items data table
                Word.Table workItemsTable = createVerticalTable(workItemsParagraph.Range, workItemsDataTable, 
                    true, Word.WdRowAlignment.wdAlignRowCenter);

                // done!
                logger.setType(Logger.Type.Success)
                    .setMessage("Document generated.")
                    .display();
            }
            catch (Exception e)
            {
                // set sizing and theming
                logger.setType(Logger.Type.Error)
                    .setMessage("Document not generated. " + e.Message)
                    .display();
            }

            // give the user control
            giveUserControl();
        }

        /// <summary>
        /// Gives the user control over the document
        /// </summary>
        public void giveUserControl(bool userControl = true)
        {
            document.UserControl = userControl;
        }

        // utility methods
        /// <summary>
        /// Creates a horizontally stacked data table from the data
        /// </summary>
        /// <param name="range"></param>
        /// <param name="numberOfSplits"></param>
        /// <param name="tableKeyValuePairs"></param>
        /// <param name="tableAlignment"></param>
        /// <returns></returns>
        private Word.Table createHorizontalStackedTable(Word.Range range, int numberOfSplits, 
            Dictionary<string, string> tableKeyValuePairs, Word.WdRowAlignment tableAlignment)
        {
            // 2 splits = 4 columns
            int numberOfColumns = 2 * numberOfSplits;

            // get a list of the keys
            List<string> tableKeys = tableKeyValuePairs.Keys.ToList();

            // determine the optimal number of rows for the table
            int numberOfRows = (tableKeys.Count() / numberOfSplits) + (tableKeys.Count() % numberOfSplits);

            // create the entire table with styling
            Word.Table table = document.Tables.Add(range, numberOfRows, numberOfColumns,
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
            table.PreferredWidth = app.InchesToPoints(6.0F);
            table.Rows.Alignment = tableAlignment;

            // goto first cell
            Word.Cell tableCell = table.Cell(1, 1);

            // counter variable
            int counter = 0;

            // style the entire table
            table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
            table.Borders.InsideColor = Word.WdColor.wdColorGray45;
            table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
            table.Borders.OutsideColor = Word.WdColor.wdColorGray55;

            // set styling for horizontal columns
            for (int i = 1; i <= numberOfRows; i++)
            {
                for (int j = 1; j <= numberOfColumns; j++)
                {
                    tableCell = table.Cell(i, j);
                    string currentKey = "";
                    if (counter != tableKeys.Count())
                        currentKey = tableKeys.ElementAt(counter);

                    tableCell.Height = 18;
                    tableCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    tableCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    tableCell.Range.Bold = 1;
                    tableCell.Range.Font.Size = 8;
                    tableCell.Range.Font.Name = "Arial";

                    if (j % 2 != 0)
                    {
                        tableCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        tableCell.Range.Text = currentKey;
                    }
                    else
                    {
                        tableCell.Range.Font.ColorIndex = Word.WdColorIndex.wdTeal;
                        if (counter != tableKeys.Count())
                        {
                            tableCell.Range.Text = tableKeyValuePairs[currentKey];
                            counter++;
                        }
                        else
                        {
                            tableCell.Range.Text = "";
                        }
                    }
                }
            }
            return table;
        }

        /// <summary>
        /// Creates a vertical style data table from the data
        /// </summary>
        /// <param name="range"></param>
        /// <param name="dt"></param>
        /// <param name="header"></param>
        /// <param name="tableAlignment"></param>
        /// <returns>A vertical style data table</returns>
        private Word.Table createVerticalTable(Word.Range range, DataTable dt, bool header, Word.WdRowAlignment tableAlignment)
        {
            // need a counter to get word rows
            int counter = 1;

            // try to print data out
            try
            {
                // find errors with pulled data
                if (dt == null) throw new InvalidDataException("Data object was not initialized.");
                if (((dt.Rows.Count == 0 && !header) || (dt.Rows.Count < 1 && header)) || dt.Columns.Count == 0)
                    throw new InvalidDataException("Not enough data was pulled in");

                // create the entire table with styling
                Word.Table table = document.Tables.Add(range, dt.Rows.Count + 1, dt.Columns.Count,
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
                table.Rows.Alignment = tableAlignment;
                table.PreferredWidth = app.InchesToPoints(7.0F);
                table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleNone;
                table.Range.Font.Size = 8;

                if (header == true)
                {
                    // create header row
                    Word.Range headerRange = table.Rows[counter].Range;
                    headerRange.Bold = 1;
                    headerRange.Font.Name = "Arial";
                    headerRange.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    headerRange.Rows.Height = 18;
                    headerRange.Font.ColorIndex = Word.WdColorIndex.wdWhite;
                    headerRange.Shading.BackgroundPatternColor = (Word.WdColor) ColorTranslator.ToOle(Color.FromArgb(23, 64, 109));
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        headerRange.Cells[i + 1].Range.Text = Utilities.spaceCapitalizedNames(dt.Columns[i].ToString());
                        headerRange.Cells[i + 1].VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    }
                    // don't create more
                    header = false;
                    counter++;
                }

                foreach (DataRow row in dt.Rows)
                {
                    // apply data styling
                    Word.Range rowRange = table.Rows[counter].Range;
                    rowRange.Bold = 0;
                    rowRange.Font.Name = "Arial";
                    rowRange.Font.Size = 8;
                    rowRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;

                    // alternate row colors
                    if (counter % 2 == 0)
                        rowRange.Shading.BackgroundPatternColor = (Word.WdColor) ColorTranslator.ToOle(Color.WhiteSmoke);
                    else
                        rowRange.Shading.BackgroundPatternColor = (Word.WdColor) ColorTranslator.ToOle(Color.LightBlue);

                    // get all field values and apply to the data table
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        rowRange.Cells[i + 1].Range.Text = row[i].ToString();
                    }
                    counter++;
                }

                // give back table
                return table;
            }
            catch (Exception e)
            {
                // log
                this.logger.setType(Logger.Type.Error)
                    .setMessage("Table could not be created. " + e.Message)
                    .display();

                // insert a break
                Word.Paragraph errorParagraph = document.Paragraphs.Add();
                insertTableSplit(errorParagraph);

                // create the entire table with styling
                Word.Table table = document.Tables.Add(errorParagraph.Range, 1, 1,
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
                table.Rows.Alignment = tableAlignment;
                table.PreferredWidth = app.InchesToPoints(6.0F);
                table.Cell(1, 1).Range.Bold = 0;
                table.Cell(1, 1).Range.Font.Name = "Arial";
                table.Cell(1, 1).Range.Font.ColorIndex = Word.WdColorIndex.wdWhite;
                table.Cell(1, 1).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.DarkBlue);
                table.Cell(1, 1).Range.Text = "Table could not be created. " + e.Message;

                // done
                return table;
            }
        }

        /// <summary>
        /// Inserts a table split into the document
        /// </summary>
        /// <param name="paragraph"></param>
        private void insertTableSplit(Word.Paragraph paragraph)
        {
            // paragraph.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
        }

        /// <summary>
        /// Creates a word section heading
        /// </summary>
        /// <param name="headerText"></param>
        /// <param name="extraSpace">If extra formatting space needed (for tables)</param>
        private void createHeading(string headerText, bool extraSpace = true)
        {
            // add a details header
            Word.Paragraph heading = document.Paragraphs.Add();
            heading.Range.Text = headerText;
            heading.Range.set_Style("Heading 2");
            heading.Range.Font.Name = "Arial";
            heading.Range.Font.Size = 12;
            heading.Range.Font.Bold = 1;
            heading.Range.Font.ColorIndex = Word.WdColorIndex.wdBlack;
            heading.Indent();

            // split
            insertTableSplit(heading);

            // add paragraph
            if (extraSpace == true)
            {
                Word.Paragraph empty = document.Paragraphs.Add();
            }
        }

        /// <summary>
        /// Destructor
        /// </summary>
        ~WordGenerator()
        {
            try
            {
                // remove user control
                document.UserControl = false;

                // save this document
                document.SaveAs2(Utilities.getExecutingPath() + projectName + " " + iterationPath + " " + "Release Notes.docx", Word.WdSaveFormat.wdFormatDocument,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, true, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Word.WdLineEndingType.wdCRLF, Type.Missing, Type.Missing);

                // quit
                app.Documents.Close(false);
                app.Quit(Word.WdSaveOptions.wdSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, Type.Missing);

                // unmarshall all COM objects
                Marshal.ReleaseComObject(document);
                Marshal.ReleaseComObject(app);

                // set to null
                document = null;
                app = null;
            }
            catch (COMException e)
            {
                // exception, up to system to free objects
                // once program is gone
                (new Logger())
                    .setType(Logger.Type.Warning)
                    .setSilence(this.silent)
                    .setMessage(e.Message + "\n Word may not have been freed from user control, " +
                                            "is waiting on user save, or cannot save (another open document?).")
                    .display();
            }

            // collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
