using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotes
{
    abstract class ReleaseNotesGenerator
    {
        protected bool silent;
        protected Logger logger;
        protected string projectName = "";
        protected string iterationPath = "";
        protected string webServer = "";
        protected string databaseServer = "";
        protected string database = "";
        protected string webLink = "";

        /// <summary>
        /// Generates the release notes for the specified application
        /// and iteration path
        /// </summary>
        public abstract void generateReleaseNotes();

        /// <summary>
        /// Gets the project name
        /// </summary>
        /// <returns></returns>
        public string getProjectName()
        {
            return this.projectName;
        }

        /// <summary>
        /// Sets the project name
        /// </summary>
        /// <param name="projectName"></param>
        public void setProjectName(string projectName)
        {
            this.projectName = projectName;
        }

        /// <summary>
        /// Gets the iteration path
        /// </summary>
        /// <returns></returns>
        public string getIterationPath()
        {
            return this.iterationPath;
        }

        /// <summary>
        /// Sets the iteration path
        /// </summary>
        /// <param name="iterationPath"></param>
        public void setIterationPath(string iterationPath)
        {
            this.iterationPath = iterationPath;
        }

        /// <summary>
        /// Sets the database server
        /// </summary>
        /// <param name="databaseServer"></param>
        public void setDatabaseServer(string databaseServer)
        {
            this.databaseServer = databaseServer;
        }

        /// <summary>
        /// Sets the web server
        /// </summary>
        /// <param name="webServer"></param>
        public void setWebServer(string webServer)
        {
            this.webServer = webServer;
        }

        /// <summary>
        /// Sets the database name
        /// </summary>
        /// <param name="database"></param>
        public void setDatabase(string database)
        {
            this.database = database;
        }

        /// <summary>
        /// Project web link
        /// </summary>
        /// <param name="link"></param>
        public void setProjectWebLink(string webLink)
        {
            this.webLink = webLink;
        }
    }
}
