using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            logger.setMessage("Generating Word release notes document...")
                .setType(Logger.Type.Information)
                .display();

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
                if (!File.Exists(Utilities.GetExecutingPath() + "ACAS.jpg"))
                {
                    Resources.Resources.ACAS.Save(@"ACAS.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                }

                // get the word document range of the header
                Word.Section firstSection = document.Sections.First;
                Word.Range headerSectionRange = firstSection.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                // get the shape back, put in the header and resize
                Word.InlineShape ACASLogo = headerSectionRange.InlineShapes.AddPicture(Utilities.GetExecutingPath() + "ACAS.jpg", false, true);

                // thread sleep to allow COM interop to catch up
                Thread.Sleep(100);

                // scale height and width
                ACASLogo.ScaleHeight = 55.0F;
                ACASLogo.ScaleWidth = 55.0F;

                // get range of the first paragraph
                document.Paragraphs.Add();
                Word.Range headerRange = document.Paragraphs[2].Range;

                // create the document title
                headerRange.Font.Name = "Times New Roman";
                headerRange.Font.Size = 12;
                headerRange.Text = "APPLICATION BUILD/RELEASE NOTES\n";
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.Bold = 1; // true

                // add another paragraph
                Word.Paragraph programInformationParagraph = document.Paragraphs.Add();

                // create application information table
                Word.Table programInformationTable = createBlankInformationTable(programInformationParagraph.Range, 3, 4);
                
                // aggregate information
                Dictionary<string, string> informationTableData = new Dictionary<string, string>();
                informationTableData.Add("Application", this.projectName);
                informationTableData.Add("Release", this.projectName + " " + this.iterationPath);
                informationTableData.Add("Build #", "Can't get yet from TFS");
                informationTableData.Add("Release Date", DateTime.Now.ToShortDateString());
                informationTableData.Add("Iteration (Sprint) #", "Can't get yet from TFS");
                informationTableData.Add("", "");

                // fill in table
                fillInformationTableValues(programInformationTable, informationTableData);

                // insert para post table
                programInformationParagraph.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                programInformationParagraph.Range.InsertParagraphAfter();
                programInformationParagraph.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // add a paragraph
                Word.Paragraph programServerParagraph = document.Paragraphs.Add();

                // create application information table
                Word.Table programServerTable = createBlankInformationTable(programServerParagraph.Range, 4, 2);

                // aggregate information
                Dictionary<string, string> programServerData = new Dictionary<string, string>();
                programServerData.Add("Web Server", "pro00websv01.acas.corp.americancapital.com");
                programServerData.Add("Database Server", "SQLNONFinancialCluster02.americancapital.com");
                programServerData.Add("Database", projectName);
                programServerData.Add("Source", Settings.Settings.Default.TFSServer 
                    + projectName + "/_versionControl\n" 
                    + "(Changeset: " + "Can't get from TFS at the moment" + ")");

                // fill in table
                fillInformationTableValues(programServerTable, programServerData);

                // get release notes work item collection
                TFSAccessor TFS = TFSAccessor.TFSAccessorFactory();
                WorkItemCollection c = TFS.getReleaseNotesFromQuery(this.projectName, this.iterationPath);
                if (c == null) throw new Exception("Work items could not be retrieved.");

                // add table information
                int counter = 1;
                foreach (WorkItem i in c)
                {
                    //addRow(counter.ToString(), i.Id.ToString(), i.Type.Name, i.Title.ToString(), 
                    // i.AreaPath, i.IterationPath, StripHtml.StripHtmlContrived(i.Description, true));
                    counter++;
                }

                // done!
                logger.setType(Logger.Type.Success)
                    .setMessage("Document generated.")
                    .display();
            }
            catch (Exception e)
            {
                // autosize and theme, with error message
                //addRow(e.Message);

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
        private Word.Table createBlankInformationTable(Word.Range range, int numberOfRows, int numberOfColumns)
        {
            // create the entire table with styling
            Word.Table informationTable = document.Tables.Add(range, numberOfRows, numberOfColumns, 
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
            informationTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            informationTable.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth100pt;
            informationTable.Borders.InsideColor = Word.WdColor.wdColorGray45;
            informationTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            informationTable.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth150pt;
            informationTable.Borders.OutsideColor = Word.WdColor.wdColorGray55;

            // set the current table range
            // this.range = currentTable.Range;

            // set styling for horizontal columns
            for (int i = 1; i <= numberOfRows; i++)
            {
                for (int j = 1; j <= numberOfColumns; j++)
                {
                    Word.Cell informationCell = informationTable.Cell(i, j);
                    informationCell.Height = 18;
                    if (j % 2 != 0)
                        informationCell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                    informationCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    informationCell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    informationCell.Range.Bold = 1;
                    informationCell.Range.Font.Size = 8;
                    informationCell.Range.Font.Name = "Arial";
                }
            }

            return informationTable;
        }

        private void fillInformationTableValues(Word.Table table, Dictionary<string, string> tableKVs)
        {
            // goto first cell
            Word.Cell tableCell = table.Cell(1, 1);

            // get a list of the keys
            List<string> keys = tableKVs.Keys.ToList();

            // counter variable
            int counter = 0;

            // get a key count and see if the argument list matches up
            if (tableKVs.Keys.Count() == ((table.Columns.Count * table.Rows.Count) / 2) && table != null)
            {
                for (int i = 1; i <= table.Rows.Count; i++)
                {
                    for (int j = 1; j <= table.Columns.Count; j++)
                    {
                        tableCell = table.Cell(i, j);
                        string currentKey = keys.ElementAt(counter);

                        if (j % 2 != 0)
                        {
                            tableCell.Range.Text = currentKey;
                        }
                        else
                        {
                            tableCell.Range.Font.ColorIndex = Word.WdColorIndex.wdTeal;
                            tableCell.Range.Text = tableKVs[currentKey];
                            counter++;
                        }
                    }
                }
            }
            else
            {
                throw new InvalidDataException("Incorrect number of arguments");
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
                // unimplemented

                // quit
                app.Documents.Close(false);
                app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges, Word.WdOriginalFormat.wdOriginalDocumentFormat, Type.Missing);

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
