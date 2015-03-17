using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Proxy;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Build.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using ReleaseNotes.Utility;

namespace ReleaseNotes
{
	class TFSAccessor
	{
		private TfsTeamProjectCollection projectCollection;
		private WorkItemStore workItems;
		private TfsClientCredentials credentials = new TfsClientCredentials();
		private string projectName;
		private string projectSubpath;
		private string iterationNumber;

		/// <summary>
		/// Constructor for TfsAccessor, creates a TFS interface
		/// </summary>
		/// <param name="TfsServerUri"></param>
		private TFSAccessor(string TfsServerUri, string projectName, string iterationNumber, string projectSubpath = "")
		{
			this.projectCollection = new TfsTeamProjectCollection(new Uri(TfsServerUri));
			this.projectCollection.EnsureAuthenticated();
			this.workItems = (WorkItemStore) projectCollection.GetService(typeof(WorkItemStore));
			this.projectName = projectName;
			this.projectSubpath = projectSubpath;
			this.iterationNumber = iterationNumber;
		}

		/// <summary>
		/// Gets project collection
		/// </summary>
		/// <returns>A project collection</returns>
		private TfsTeamProjectCollection GetProjectCollection() {
			return this.projectCollection;
		}

		/// <summary>
		/// Gets work items
		/// </summary>
		/// <returns>A list of work items</returns>
		private WorkItemStore GetWorkItems() {
			return this.workItems;
		}

		/// <summary>
		/// Gets work items from a query
		/// </summary>
		/// <param name="query"></param>
		/// <returns>A collection of work items from the designated query</returns>
		private WorkItemCollection GetWorkItemsFromQuery(string query)
		{
			try
			{
				return this.workItems.Query(query);
			}
			catch (Exception e)
			{
				(new Logger()).SetLoggingType(Logger.Type.Error)
					.SetMessage(e.Message)
					.Display();
				return null;
			}
		}

		/// <summary>
		/// Gets release notes data from a query
		/// </summary>
		/// <param name="projectName"></param>
		/// <param name="iterationNumber"></param>
		/// <returns>Release notes work item collection</returns>
		public WorkItemCollection GetReleaseNotesFromQuery()
		{
			(new Logger())
				.SetMessage("Querying work items.")
				.SetLoggingType(Logger.Type.Information)
				.Display();
			return GetWorkItemsFromQuery("SELECT * " +
				"FROM workitems " +
				"WHERE [System.TeamProject] = '" + this.projectName + "' " +
				"AND ([System.Tags] CONTAINS 'Service Now'" +
				"OR [System.Tags] CONTAINS 'UAT' " +
				"OR [System.Tags] CONTAINS 'PROD' " +
				"OR [System.WorkItemType] = 'Product Backlog Item')" +
				"AND [System.State] IN ('Committed', 'Done')" +
				"AND [System.IterationPath] = '" + this.projectName + "\\" + (this.projectSubpath != null && this.projectSubpath != "" ? this.projectSubpath + "\\" : "") + this.iterationNumber + "'");
		}

		/// <summary>
		/// Gets the release notes query as a datatable (it should do it this way in the first place)
		/// </summary>
		/// <param name="projectName"></param>
		/// <param name="iterationNumber"></param>
		/// <returns>A data table of the minimal release notes data</returns>
		public DataTable GetReleaseNotesAsDataTable()
		{
			DataTable releaseNotesTable = new DataTable();
			releaseNotesTable.Columns.Add("ID", typeof(int));
			releaseNotesTable.Columns.Add("WorkItem", typeof(string));
			releaseNotesTable.Columns.Add("Title", typeof(string));
			releaseNotesTable.Columns.Add("Area", typeof(string));
			releaseNotesTable.Columns.Add("Description", typeof(string));

			WorkItemCollection c = GetReleaseNotesFromQuery();
			if (c != null)
				foreach (WorkItem i in c)
					releaseNotesTable.Rows.Add(i.Id, i.Type.Name, i.Title, i.AreaPath, Utilities.StripHtmlContrived(i.Description, false) /*, i.Tags */);
			return releaseNotesTable;
		}

		/// <summary>
		/// Gets the latest changeset for this project
		/// </summary>
		/// <param name="projectName"></param>
		/// <param name="iterationNumber"></param>
		/// <returns></returns>
		public int GetLatestChangesetNumber()
		{
			(new Logger())
				.SetMessage("Querying changeset numbers.")
				.SetLoggingType(Logger.Type.Information)
				.Display();
			VersionControlServer versionControlServer = this.projectCollection.GetService<VersionControlServer>();
			var teamProjectServerPath = versionControlServer.GetTeamProject(projectName).ServerItem;
			var latestChangesetQuery = versionControlServer.QueryHistory(new QueryHistoryParameters(new ItemSpec(teamProjectServerPath, RecursionType.Full, 0)));
			var latestChangesets = latestChangesetQuery.Cast<Changeset>();
			return latestChangesets.First().ChangesetId;
		}

		/// <summary>
		/// Gets the latest build number
		/// </summary>
		/// <param name="projectName"></param>
		/// <returns></returns>
		public string GetLatestBuildNumber()
		{
			(new Logger())
				.SetMessage("Querying build definitions.")
				.SetLoggingType(Logger.Type.Information)
				.Display();

			IBuildServer buildServer = this.projectCollection.GetService<IBuildServer>();
			IBuildDetailSpec buildSpec = buildServer.CreateBuildDetailSpec(projectName);
			buildSpec.MaxBuildsPerDefinition = 1;
			buildSpec.QueryOrder = BuildQueryOrder.FinishTimeDescending;
			IBuildQueryResult query = buildServer.QueryBuilds(buildSpec);

			if (query.Builds.Count() == 0)
			{
				return null;
			}
			IBuildDetail detail = query.Builds[0];
			return detail.BuildNumber;
		}

		/// <summary>
		/// Gets test cases
		/// </summary>
		public DataTable GetTestCases()
		{
			(new Logger())
			.SetMessage("Querying test cases.")
			.SetLoggingType(Logger.Type.Information)
			.Display();

			// create steps data table
			DataTable testCasesTable = new DataTable();
			testCasesTable.Columns.Add("ID", typeof(int));
			testCasesTable.Columns.Add("Title", typeof(string));
			testCasesTable.Columns.Add("Steps", typeof(string));

			// get the TFS test management server and query
			ITestManagementService testManagementService = this.projectCollection.GetService<ITestManagementService>();
			ITestManagementTeamProject testProject = testManagementService.GetTeamProject(this.projectName);
			ITestCaseHelper testCaseHelper = testProject.TestCases;
			IEnumerable<ITestCase> tcc = testProject.TestCases.Query("SELECT * " +
				"FROM workitems " +
				"WHERE [System.TeamProject] = '" + this.projectName + "' " +
				"AND [System.WorkItemType] = 'Test Case'" +
				"AND [System.IterationPath] = '" + this.projectName + "\\" + (this.projectSubpath != null && this.projectSubpath != "" ? this.projectSubpath + "\\" : "") + this.iterationNumber + "'");
			
			// through all items that are test cases
			foreach(ITestCase testCase in tcc) {
				
				// get the work item data
				WorkItem workItem = testCase.Links.WorkItem;
				int id = workItem.Id;
				string title = workItem.Title;
				string desc = workItem.Description;
				string steps = "";

				// get collection of test actions
				TestActionCollection testActionCollection = testCase.Actions;
				int counter = 1;

				// iterate through all test actions
				for (int i = 0; i < testActionCollection.Count; i++)
				{
					// get an action
					var action = testCase.Actions[i];

					// if not test step, ignore
					if (!(action is ITestStep))
						continue;
					else
					{
						// get the test step
						var step = action as ITestStep;
						int stepNumber = counter;
						string stepTitle = Utilities.StripHtmlContrived(step.Title, true);
						string result = Utilities.StripHtmlContrived(step.ExpectedResult, true);
						steps += "Step #" + stepNumber + ": " + stepTitle + "\n";
						if (result.Trim().Count() != 0)
						{
							steps += "Expected Result: " + result + "\n";
						}
						steps += "\n";
						counter++;
					}
				}

				// add to data table
				testCasesTable.Rows.Add(id, title, steps);
			}
			return testCasesTable;
		}

		/// <summary>
		/// Factory pattern interface for creating a TfsAccessor.
		/// Catches errors relates to creating the interface (ie. authentication issues)
		/// </summary>
		/// <returns>A TfsAccessor</returns>
		public static TFSAccessor TFSAccessorFactory(string serverTeamProjectUrl, string projectName, string iterationNumber, string projectSubpath = "")
		{
			var errorLogger = (new Logger())
				.SetMessage("Connected to TFS")
				.SetLoggingType(Logger.Type.Information);
			try
			{
				TFSAccessor a = new TFSAccessor(serverTeamProjectUrl, projectName, iterationNumber, projectSubpath);
				if (a != null)
					errorLogger.Display();
				return a;
			}
			catch (Exception e)
			{
				errorLogger.SetMessage(e.Message).SetLoggingType(Logger.Type.Error).Display();
				return null;
			}
		}
	}
}
