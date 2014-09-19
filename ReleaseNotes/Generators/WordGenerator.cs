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
using ReleaseNotes.Utility;

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
        private WordGenerator(NamedLookup settings) : base(settings)
        {
            app = new Word.Application();
            app.Visible = !this.silent;
            document.UserControl = !this.silent;
            document = app.Documents.Add(Type.Missing, Type.Missing, Word.WdNewDocumentType.wdNewBlankDocument, !this.silent);
        }

        /// <summary>
        /// Constructor (factory method) for a word generator object
        /// </summary>
        /// <param name="documentName"></param>
        /// <param name="isVisible"></param>
        /// <param name="silent"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static WordGenerator WordGeneratorFactory(NamedLookup settings)
        {
            try
            {
                return new WordGenerator(settings);
            }
            catch (COMException e)
            {
                (new Logger())
                    .setType(Logger.Type.Error)
                    .setMessage(e.Message).display();
                return null;
            }
        }

        /// <summary>
        /// Formats document pre creation
        /// </summary>
        public override void createDocumentSpecificPreFormatting()
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
        }

        /// <summary>
        /// Formats document post creation
        /// </summary>
        public override void createDocumentSpecificPostFormatting()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Creates a header graphic
        /// </summary>
        public override void createCorporateHeaderGraphic()
        {
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
        }
         
        /// <summary>
        /// Creates a title with associated text
        /// </summary>
        /// <param name="text"></param>
        public override void createTitle(string text)
        {
            // get range of the first paragraph
            Word.Paragraph titleParagraph = document.Paragraphs.Add();
            Word.Range headerRange = titleParagraph.Range;

            // create the document title
            headerRange.Font.Name = "Times New Roman";
            headerRange.Font.Size = 12;
            headerRange.Text = text;
            headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            headerRange.Font.Bold = 1;
        }

        /// <summary>
        /// Create key value hyperlink section
        /// </summary>
        /// <param name="headername"></param>
        /// <param name="text"></param>
        /// <param name="hyperlink"></param>
        public override void createNamedSection(string headername, string text, string hyperlink)
        {
            // new heading
            createHeading(headername);
            Word.Paragraph accessParagraph = document.Paragraphs.Add();

            // create caption
            string accessParagraphText = text;
            string webLink = hyperlink;
            accessParagraph.Range.Text = Utilities.implicitMalloc(accessParagraphText, webLink.Length);

            // several indents needed
            for (int i = 0; i < 3; i++) { accessParagraph.Indent(); }

            document.Hyperlinks.Add(document.Range(accessParagraph.Range.Start + accessParagraphText.Length,
                accessParagraph.Range.Start + accessParagraphText.Length + webLink.Length),
                webLink, Type.Missing, settings["Project Name"], webLink, Type.Missing);

            // split
            insertTableSplit(accessParagraph);
        }

        /// <summary>
        /// Creates a gorizontal stacked table in Word
        /// </summary>
        /// <param name="data"></param>
        /// <param name="splits"></param>
        /// <param name="header"></param>
        public override void createHorizontalTable(NamedLookup data, int splits, bool header)
        {
            // if header needed
            if (header)
                createHeading(data.getName());

            // add another paragraph
            Word.Paragraph paragraph = document.Paragraphs.Add();
            Word.Range range = paragraph.Range;

            // 2 splits = 4 columns
            int numberOfColumns = 2 * splits;

            // get a list of the keys
            List<string> tableKeys = data.getLookup().Keys.ToList();

            // determine the optimal number of rows for the table
            int numberOfRows = (tableKeys.Count() / splits) + (tableKeys.Count() % splits);

            // create the entire table with styling
            Word.Table table = document.Tables.Add(range, numberOfRows, numberOfColumns,
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
            table.PreferredWidth = app.InchesToPoints(6.0F);
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

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
                        tableCell.Range.Font.TextColor.RGB = ColorTranslator.ToOle(Color.FromArgb(0, 112, 192));
                        if (counter != tableKeys.Count())
                        {
                            if (currentKey.Equals("Source"))
                            {
                                tableCell.Range.Hyperlinks.Add(document.Range(tableCell.Range.Start, tableCell.Range.End), data[currentKey], Type.Missing,
                                    "Source Control", data[currentKey], Type.Missing);
                                tableCell.Range.Font.Name = "Arial";
                                tableCell.Range.Font.Size = 8;
                            }
                            else
                            {
                                tableCell.Range.Text = data[currentKey];
                            }
                            counter++;
                        }
                        else
                        {
                            tableCell.Range.Text = "";
                        }
                    }
                }
            }

            // split
            insertTableSplit(paragraph);
        }

        /// <summary>
        /// Creates a vertical table in Word
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="headerText"></param>
        /// <param name="header"></param>
        public override void createVerticalTable(DataTable dt, string headerText, bool header)
        {
            // need a counter to get word rows
            int counter = 1;

            // try to print data out
            try
            {
                // add another paragraph
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;

                // create header
                if (header)
                    createHeading(headerText);

                // find errors with pulled data
                if (dt == null) throw new InvalidDataException("Data object was not initialized.");
                if (((dt.Rows.Count == 0 && !header) || (dt.Rows.Count < 1 && header)) || dt.Columns.Count == 0)
                    throw new InvalidDataException("Not enough data was pulled in");

                // create the entire table with styling
                Word.Table table = document.Tables.Add(range, dt.Rows.Count + 1, dt.Columns.Count,
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
                table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
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
                    headerRange.Rows.Height = 24;
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
                    table.Rows[counter].Height = 24;
                    rowRange.Bold = 0;
                    rowRange.Font.Name = "Arial";
                    rowRange.Font.Size = 8;
                    rowRange.Font.ColorIndex = Word.WdColorIndex.wdBlack;

                    // alternate row colors
                    if (counter % 2 == 0)
                        rowRange.Shading.BackgroundPatternColor = (Word.WdColor) ColorTranslator.ToOle(Color.WhiteSmoke);
                    else
                        rowRange.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.FromArgb(220, 233, 238));

                    // get all field values and apply to the data table
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (dt.Columns[i].ColumnName == "ID")
                        {
                            string hyperlinkText = settings["Team Project Path"] + settings["Project Name"] + "/_workitems#_a=edit&id=" + row[i].ToString() + "&triage=true";
                            rowRange.Cells[i + 1].Range.Text = Utilities.implicitMalloc(row[i].ToString(), 6);
                            rowRange.Cells[i + 1].Range.Bold = 1;
                            rowRange.Cells[i + 1].Range.Hyperlinks.Add(document.Range(rowRange.Cells[i+1].Range.Start, rowRange.Cells[i+1].Range.End), 
                                hyperlinkText, Type.Missing, "Work Item", row[i].ToString(), Type.Missing);
                        }
                        else
                        {
                            rowRange.Cells[i + 1].Range.Text = row[i].ToString();
                        }
                    }
                    counter++;
                }

                // split
                insertTableSplit(paragraph);
            }
            catch (Exception e)
            {
                // log
                this.logger.setType(Logger.Type.Error)
                    .setMessage("Table could not be created. " + e.Message)
                    .display();

                // create error message
                createErrorMessage(e.Message); 
            }
        }

        /// <summary>
        /// Creates an error table at the end of the current paragraph from a string
        /// </summary>
        /// <param name="message"></param>
        public override void createErrorMessage(string message)
        {
            // insert a break
            Word.Paragraph errorParagraph = document.Paragraphs.Add();

            // insert table split
            insertTableSplit(errorParagraph);

            // create the entire table with styling
            Word.Table table = document.Tables.Add(errorParagraph.Range, 1, 1,
                Word.WdDefaultTableBehavior.wdWord9TableBehavior, Word.WdAutoFitBehavior.wdAutoFitFixed);
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            table.PreferredWidth = app.InchesToPoints(6.0F);
            table.Cell(1, 1).Range.Bold = 0;
            table.Cell(1, 1).Range.Font.Name = "Arial";
            table.Cell(1, 1).Range.Font.ColorIndex = Word.WdColorIndex.wdWhite;
            table.Cell(1, 1).Range.Shading.BackgroundPatternColor = (Word.WdColor)ColorTranslator.ToOle(Color.DarkBlue);
            table.Cell(1, 1).Range.Text = "Table could not be created. " + message;
        }

        /// <summary>
        /// Creates a word section heading
        /// </summary>
        /// <param name="headerText"></param>
        /// <param name="extraSpace">If extra formatting space needed (for tables)</param>
        public override void createHeading(string headerText)
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
            Word.Paragraph empty = document.Paragraphs.Add();
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
        /// Destructor
        /// </summary>
        ~WordGenerator()
        {
            try
            {
                // remove user control
                document.UserControl = false;

                // save this document
                document.SaveAs2(Utilities.getExecutingPath() + settings["Project Name"] + " " + settings["Iteration"]
                    + " Release Notes.docx", Word.WdSaveFormat.wdFormatDocumentDefault,
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
                    .setMessage(e.Message + "\n Word may not have been freed from user control, \n" +
                                            "is waiting on user save, \n or cannot save (another open document?).")
                    .display();
            }

            // collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
