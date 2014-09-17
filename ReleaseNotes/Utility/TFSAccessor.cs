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

namespace ReleaseNotes
{
    class TFSAccessor
    {
        private TfsTeamProjectCollection projectCollection;
        private WorkItemStore workItems;
        private TfsClientCredentials credentials = new TfsClientCredentials(); 

        /// <summary>
        /// Constructor for TfsAccessor, creates a TFS interface
        /// </summary>
        /// <param name="TfsServerUri"></param>
        private TFSAccessor(string TfsServerUri)
        {
            this.projectCollection = new TfsTeamProjectCollection(new Uri(TfsServerUri));
            this.projectCollection.EnsureAuthenticated();
            this.workItems = (WorkItemStore) projectCollection.GetService(typeof(WorkItemStore));
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
        public WorkItemCollection getReleaseNotesFromQuery(string projectName, string iterationNumber)
        {
            (new Logger())
                .setMessage("Querying work items.")
                .setType(Logger.Type.Information)
                .display();
            return getWorkItemsFromQuery("SELECT [System.ID], [System.Title], [System.AreaPath], [System.IterationPath], [System.Description] " +
                "FROM workitems " +
                "WHERE [System.TeamProject] = '" + projectName + "' " +
                "AND ([System.Tags] CONTAINS 'Service Now'" +
                "OR [System.Tags] CONTAINS 'UAT' " +
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
        public DataTable getReleaseNotesAsDataTable(string projectName, string iterationNumber)
        {
            DataTable releaseNotesTable = new DataTable();
            releaseNotesTable.Columns.Add("ID", typeof(int));
            releaseNotesTable.Columns.Add("WorkItem", typeof(string));
            releaseNotesTable.Columns.Add("Title", typeof(string));
            releaseNotesTable.Columns.Add("Area", typeof(string));
            releaseNotesTable.Columns.Add("Iteration", typeof(string));

            WorkItemCollection c = getReleaseNotesFromQuery(projectName, iterationNumber);
            if (c != null)
                foreach (WorkItem i in c)
                    releaseNotesTable.Rows.Add(i.Id, i.Type.Name, i.Title, i.AreaPath, i.IterationPath);
            return releaseNotesTable;
        }

        /// <summary>
        /// Gets the latest changeset for this project
        /// </summary>
        /// <param name="projectName"></param>
        /// <param name="iterationNumber"></param>
        /// <returns></returns>
        public int getLatestChangesetNumber(string projectName)
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
        public string getLatestBuildNumber(string projectName)
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
            IBuildDetail detail = query.Builds[0];
            return detail.BuildNumber;
        }

        /// <summary>
        /// Factory pattern interface for creating a TfsAccessor.
        /// Catches errors relates to creating the interface (ie. authentication issues)
        /// </summary>
        /// <returns>A TfsAccessor</returns>
        public static TFSAccessor TFSAccessorFactory()
        {
            var errorLogger = (new Logger())
                .setMessage("Connected to TFS")
                .setType(Logger.Type.Information);
            try
            {
                TFSAccessor a = new TFSAccessor(Settings.Settings.Default.TFSServer);
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
