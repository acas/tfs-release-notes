using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using ReleaseNotes.Utility;
using System.Diagnostics.Contracts;

namespace ReleaseNotes
{
    class ReleaseNotesGenerator
    {
        protected bool silent;
        protected Logger logger;
        protected TFSAccessor TFS;
        protected NamedLookup settings;
        private List<NamedLookup> propertiesList = new List<NamedLookup>();

        // override all of these
        public virtual void createTitle(string titleText) { }
        public virtual void createHeader(string headingText) { }

        public virtual void createHorizontalTable(NamedLookup data, int splits, bool header) 
        {
            Contract.Requires<ArgumentNullException>(splits > 0, "At least 1 table split must be specified");
        }

        public virtual void createVerticalTable(DataTable dt, string headerText, bool header) 
        {
            Contract.Requires<ArgumentNullException>(dt != null, "Data cannot be null");
            Contract.Requires<ArgumentNullException>((dt.Rows.Count > 0 && !header) || (dt.Rows.Count > 1 && !header), "Row count must be greater than 0.");
            Contract.Requires<ArgumentNullException>(dt.Columns.Count > 0, "Column count must be greater than 0.");
            Contract.Requires<ArgumentNullException>(headerText != null, "Header text cannot be null");
        }

        public virtual void createDocumentSpecificPreFormatting() { }
        public virtual void createDocumentSpecificPostFormatting() { }
        public virtual void createNamedSection(string headername, string text, string hyperlink) { }
        public virtual void createErrorMessage(string message) { }
        public virtual void createHeaderGraphic(string path) { }
        public virtual void createNewPage(string worksheetName) { }
        public virtual void save() { }

        public ReleaseNotesGenerator(NamedLookup settings)
        {
            this.settings = settings;
            checkRequiredFields();
            this.TFS = TFSAccessor.TFSAccessorFactory(settings["Team Project Path"], settings["Project Name"], settings["Iteration"]);
            this.logger = new Logger();
            this.silent = false;
        }

        public void checkRequiredFields()
        {
            List<bool> keysAlright = new List<bool>();
            keysAlright.Add(settings.getLookup().ContainsKey("Team Project Path"));
            keysAlright.Add(settings.getLookup().ContainsKey("Project Name"));
            keysAlright.Add(settings.getLookup().ContainsKey("Iteration"));
            if (keysAlright.Where(a => a == false).ToList().Count() > 0)
                throw new Exception("Expected params not found.");
        }

        public void addPropertiesList(string name, Dictionary<string, string> lookup)
        {
            if (lookup != null && name != null)
                propertiesList.Add(new NamedLookup(name));
        }

        public void removePropertiesList(int index)
        {
            propertiesList.RemoveAt(index);
        }

        public void removePropertiesList(string name)
        {
            propertiesList = propertiesList.Where(a => !a.getName().Equals(name)).ToList();
        }

        public List<NamedLookup> getPropertiesList()
        {
            return this.propertiesList;
        }

        /// <summary>
        /// Creates some default executive summary details
        /// </summary>
        /// <returns></returns>
        public NamedLookup getDefaultExecutiveSummary()
        {
            NamedLookup executiveSummary = new NamedLookup("Executive Summary");
            executiveSummary["Application"] = settings["Project Name"];
            executiveSummary["Release Date"] = DateTime.Now.ToShortDateString();
            executiveSummary["Release"] = settings["Project Name"] + " " + settings["Iteration"];
            executiveSummary["Iteration (Sprint) #"] = settings["Iteration"];
            executiveSummary["Build #"] = TFS.getLatestBuildNumber();
            return executiveSummary;
        }

        /// <summary>
        /// Creates some default details you can choose 
        /// </summary>
        /// <returns></returns>
        public NamedLookup getDefaultDetails()
        {
            NamedLookup sourceServerInformation = new NamedLookup("Details");
            sourceServerInformation["Web Server"] = settings["Web Server"];
            sourceServerInformation["Database Server"] = settings["Database Server"];
            sourceServerInformation["Database"] = settings["Database"];
            sourceServerInformation["Source"] = "(Changeset: " + TFS.getLatestChangesetNumber() + ")";
            return sourceServerInformation;
        }

        /// <summary>
        /// Generate the release notes
        /// </summary>
        public void generateReleaseNotes()
        {
            // set silent to false
            silent = false;

            // create excel writer
            logger.setMessage("Generating Word release notes document.")
                .setType(Logger.Type.Information)
                .display();

            // try to generate the document
            try
            {
                // log generating document
                logger.setMessage("Preparing document, please wait...")
                    .setType(Logger.Type.Information)
                    .display();

                // pre formatting
                createDocumentSpecificPreFormatting();

                // create graphic
                createHeaderGraphic(null);

                // create heading
                createTitle(settings["Doc Type"]);

                // create horizontal table paragraph
                createHorizontalTable(getDefaultExecutiveSummary(), 2, true);

                // create access section
                createNamedSection("Access", "Application is accessible at: ", settings["Web Location"]);

                // create the details section
                createHorizontalTable(getDefaultDetails(), 1, true);

                // create a vertical table
                createVerticalTable(TFS.getReleaseNotesAsDataTable(), "Included Requirements", true);

                // create a new worksheet
                createNewPage("Test Cases");

                // create a vertical table for test cases/user stories here
                // unimplemented

                // post formatting
                createDocumentSpecificPostFormatting();

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
        }
    }
}
