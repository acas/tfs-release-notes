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
		public virtual void CreateTitle(string titleText) { }
		public virtual void CreateHeader(string headingText) { }

		public virtual void CreateHorizontalTable(NamedLookup data, int splits, bool header) 
		{
			// Contract.Requires<ArgumentNullException>(splits > 0, "At least 1 table split must be specified");
		}

		public virtual void CreateVerticalTable(DataTable dt, string headerText, bool header) 
		{
			// Contract.Requires<ArgumentNullException>(dt != null, "Data table cannot be null");
			// Contract.Requires<ArgumentNullException>((dt.Rows.Count > 0 && !header) || (dt.Rows.Count > 1 && !header), "Row count must be greater than 0.");
			// Contract.Requires<ArgumentNullException>(dt.Columns.Count > 0, "Column count must be greater than 0.");
			// Contract.Requires<ArgumentNullException>(headerText != null, "Header text cannot be null");
		}

		public virtual void CreateDocumentSpecificPreFormatting() { }
		public virtual void CreateDocumentSpecificPostFormatting(bool wide = false) { }
		public virtual void CreateNamedSection(string headername, string text, string hyperlink) { }
		public virtual void CreateErrorMessage(string message) { }
		public virtual void CreateHeaderGraphic(string path) { }
		public virtual void CreateNewPage(string worksheetName) { }
		public virtual void Save() { }

		public ReleaseNotesGenerator(NamedLookup settings, bool silent)
		{
			this.settings = settings;
			CheckRequiredFields();
			this.TFS = TFSAccessor.TFSAccessorFactory(settings["Team Project Path"], settings["Project Name"], settings["Iteration"], settings["Project Subpath"]);
			this.logger = new Logger();
			this.silent = silent;
		}

		public void CheckRequiredFields()
		{
			List<bool> keysAlright = new List<bool>();
			keysAlright.Add(settings.GetLookup().ContainsKey("Team Project Path"));
			keysAlright.Add(settings.GetLookup().ContainsKey("Project Name"));
			keysAlright.Add(settings.GetLookup().ContainsKey("Iteration"));
			if (keysAlright.Where(a => a == false).ToList().Count() > 0)
				throw new Exception("Expected params not found.");
		}

		public void AddPropertiesList(string name, Dictionary<string, string> lookup)
		{
			if (lookup != null && name != null)
				propertiesList.Add(new NamedLookup(name));
		}

		public void RemovePropertiesList(int index)
		{
			propertiesList.RemoveAt(index);
		}

		public void RemovePropertiesList(string name)
		{
			propertiesList = propertiesList.Where(a => !a.GetName().Equals(name)).ToList();
		}

		public List<NamedLookup> GetPropertiesList()
		{
			return this.propertiesList;
		}

		/// <summary>
		/// Creates some default executive summary details
		/// </summary>
		/// <returns></returns>
		public NamedLookup GetDefaultExecutiveSummary()
		{
			NamedLookup executiveSummary = new NamedLookup("Executive Summary");
			executiveSummary["Application"] = settings["Project Name"];
			executiveSummary["Release Date"] = DateTime.Now.ToShortDateString();            
			executiveSummary["Release (Sprint)"] = settings["Iteration"];
			string buildNumber = TFS.GetLatestBuildNumber();
			if (buildNumber != null)
			{
				executiveSummary["Build #"] = buildNumber;
			}            
			return executiveSummary;
		}

		/// <summary>
		/// Creates some default details you can choose 
		/// </summary>
		/// <returns></returns>
		public NamedLookup GetDefaultDetails()
		{
			NamedLookup sourceServerInformation = new NamedLookup("Details");
			sourceServerInformation["Web Server"] = settings["Web Server"];
			sourceServerInformation["Database Server"] = settings["Database Server"];
			sourceServerInformation["Database"] = settings["Database"];
			sourceServerInformation["Source"] = "(Changeset: " + TFS.GetLatestChangesetNumber() + ")";
			return sourceServerInformation;
		}

		/// <summary>
		/// Generate the release notes
		/// </summary>
		public void GenerateReleaseNotes()
		{
			// set silent to false
			silent = true;

			// create excel writer
			logger.SetMessage("Generating release notes document.")
				.SetLoggingType(Logger.Type.Information)
				.Display();

			// try to generate the document
			try
			{
				// log generating document
				logger.SetMessage("Preparing document, please wait...")
					.SetLoggingType(Logger.Type.Information)
					.Display();

				// pre formatting
				CreateDocumentSpecificPreFormatting();

				// create graphic
				CreateHeaderGraphic(null);

				// create heading
				CreateTitle(settings["Doc Type"]);

				// create horizontal table paragraph
				CreateHorizontalTable(GetDefaultExecutiveSummary(), 2, true);

				// create access section
				CreateNamedSection("Access", "Application is accessible at: ", settings["Web Location"]);

				// create the details section
				CreateHorizontalTable(GetDefaultDetails(), 1, true);

				// create a vertical table
				CreateVerticalTable(TFS.GetReleaseNotesAsDataTable(), "Included Requirements", true);

				// create a new worksheet
				CreateNewPage("Test Cases");

				// create a vertical table for test cases/user stories here
				// special post formatting
				CreateVerticalTable(TFS.GetTestCases(), "Test Cases", true);
				CreateDocumentSpecificPostFormatting(true);

                // save this!
                Save();

				// done!
				logger.SetLoggingType(Logger.Type.Success)
					.SetMessage("Document generated and saved in the current directory.")
					.Display();
			}
			catch (Exception e)
			{
				// set sizing and theming
				logger.SetLoggingType(Logger.Type.Error)
					.SetMessage("Document not generated. " + e.Message)
					.Display();
				throw;
			}
		}
	}
}
