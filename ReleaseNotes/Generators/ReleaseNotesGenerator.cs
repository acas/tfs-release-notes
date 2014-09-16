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
    }
}
