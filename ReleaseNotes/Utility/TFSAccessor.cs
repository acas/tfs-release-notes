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
        private string iterationNumber;

        /// <summary>
        /// Constructor for TfsAccessor, creates a TFS interface
        /// </summary>
        /// <param name="TfsServerUri"></param>
        private TFSAccessor(string TfsServerUri, string projectName, string iterationNumber)
        {
            this.projectCollection = new TfsTeamProjectCollection(new Uri(TfsServerUri));
            this.projectCollection.EnsureAuthenticated();
            this.workItems = (WorkItemStore) projectCollection.GetService(typeof(WorkItemStore));
            this.projectName = projectName;
            this.iterationNumber = iterationNumber;
        }

        /// <summary>
        /// Gets project collection
        /// </summary>
        /// <returns>A project collection</returns>
        private TfsTeamProjectCollection getProjectCollection() {
            return this.projectCollection;
        }

        /// <summary>
        /// Gets work items
        /// </summary>
        /// <returns>A list of work items</returns>
        private WorkItemStore getWorkItems() {
            return this.workItems;
        }

        /// <summary>
        /// Gets work items from a query
        /// </summary>
        /// <param name="query"></param>
        /// <returns>A collection of work items from the designated query</returns>
        private WorkItemCollection getWorkItemsFromQuery(string query)
        {
            try
            {
                return this.workItems.Query(query);
            }
            catch (Exception e)
            {
                (new Logger()).setType(Logger.Type.Error)
                    .setMessage(e.Message)
                    .display();
                return null;
            }
        }

        /// <summary>
        /// Gets release notes data from a query
        /// </summary>
        /// <param name="projectName"></param>
        /// <param name="iterationNumber"></param>
        /// <returns>Release notes work item collection</returns>
        public WorkItemCollection getReleaseNotesFromQuery()
        {
            (new Logger())
                .setMessage("Querying work items.")
                .setType(Logger.Type.Information)
                .display();
            return getWorkItemsFromQuery("SELECT * " +
                "FROM workitems " +
                "WHERE [System.TeamProject] = '" + projectName + "' " +
                "AND ([System.Tags] CONTAINS 'Service Now'" +
                "OR [System.Tags] CONTAINS 'UAT' " +
                "OR [System.Tags] CONTAINS 'PROD' " +
                "OR [System.WorkItemType] = 'Product Backlog Item')" +
                "AND [System.State] IN ('Committed', 'Done')" +
                "AND [System.IterationPath] = '" + projectName + "\\Release " + iterationNumber + "'");
        }

        /// <summary>
        /// Gets the release notes query as a datatable (it should do it this way in the first place)
        /// </summary>
        /// <param name="projectName"></param>
        /// <param name="iterationNumber"></param>
        /// <returns>A data table of the minimal release notes data</returns>
        public DataTable getReleaseNotesAsDataTable()
        {
            DataTable releaseNotesTable = new DataTable();
            releaseNotesTable.Columns.Add("ID", typeof(int));
            releaseNotesTable.Columns.Add("WorkItem", typeof(string));
            releaseNotesTable.Columns.Add("Title", typeof(string));
            releaseNotesTable.Columns.Add("Area", typeof(string));
            releaseNotesTable.Columns.Add("Description", typeof(string));
            // releaseNotesTable.Columns.Add("Tag", typeof(string));

            WorkItemCollection c = getReleaseNotesFromQuery();
            if (c != null)
                foreach (WorkItem i in c)
                    releaseNotesTable.Rows.Add(i.Id, i.Type.Name, i.Title, i.AreaPath, Utilities.stripHtmlContrived(i.Description, false) /*, i.Tags */);
            return releaseNotesTable;
        }

        /// <summary>
        /// Gets the latest changeset for this project
        /// </summary>
        /// <param name="projectName"></param>
        /// <param name="iterationNumber"></param>
        /// <returns></returns>
        public int getLatestChangesetNumber()
        {
            (new Logger())
                .setMessage("Querying changeset numbers.")
                .setType(Logger.Type.Information)
                .display();
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
        public string getLatestBuildNumber()
        {
            (new Logger())
                .setMessage("Querying build definitions.")
                .setType(Logger.Type.Information)
                .display();

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
        public DataTable getTestCases()
        {
            (new Logger())
            .setMessage("Querying test cases.")
            .setType(Logger.Type.Information)
            .display();

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
                "WHERE [System.TeamProject] = '" + projectName + "' " +
                "AND [System.WorkItemType] = 'Test Case'" +
                "AND [System.IterationPath] = '" + projectName + "\\Release " + iterationNumber + "'");
            
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
                        string stepTitle = Utilities.stripHtmlContrived(step.Title, true);
                        string result = Utilities.stripHtmlContrived(step.ExpectedResult, true);
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
        public static TFSAccessor TFSAccessorFactory(string serverTeamProjectUrl, string projectName, string iterationNumber)
        {
            var errorLogger = (new Logger())
                .setMessage("Connected to TFS")
                .setType(Logger.Type.Information);
            try
            {
                TFSAccessor a = new TFSAccessor(serverTeamProjectUrl, projectName, iterationNumber);
                if (a != null)
                    errorLogger.display();
                return a;
            }
            catch (Exception e)
            {
                errorLogger.setMessage(e.Message).setType(Logger.Type.Error).display();
                return null;
            }
        }
    }
}
