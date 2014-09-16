using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

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
            logger.setMessage("Generating Word release notes document...").setType(Logger.Type.Information).display();

            try
            {
                // get release notes work item collection
                WorkItemCollection c = TFSAccessor.TFSAccessorFactory().getReleaseNotesFromQuery(this.projectName, this.iterationPath);
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

        /// <summary>
        /// Destructor
        /// </summary>
        ~WordGenerator()
        {
            if (app != null)
            {
                // remove user control
                document.UserControl = false;
                try
                {
                    // save this document
                    
                    // quit
                    app.Documents.Close();
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
            }
            //collect the remaining garbage
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
