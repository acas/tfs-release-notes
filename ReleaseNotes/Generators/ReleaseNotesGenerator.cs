using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotes
{
    interface IReleaseNotesGenerator
    {
        void generateReleaseNotes();
    }

    class BaseReleaseNotesGenerator
    {
        public struct NamedLookup
        {
            private string name;
            private Dictionary<string, string> lookup;
            
            public NamedLookup(string name) {
                this.name = name;
                this.lookup = new Dictionary<string, string>();
            }

            public NamedLookup(string name, Dictionary<string, string> predefinedLookup)
            {
                this.name = name;
                this.lookup = predefinedLookup;
            }

            public string getName()
            {
                return this.name;
            }

            public string this[string name] 
            {
                get { return this.lookup[name]; }
                set
                {
                    if (name != null && value != null)
                        lookup[name] =  value.ToString();
                }
            }

            public void removeProperty(string name)
            {
                lookup.Remove(name);
            }

            public Dictionary<string, string> getLookup() {
                return this.lookup;
            }
        }

        protected bool silent;
        protected Logger logger;
        protected TFSAccessor TFS;
        private List<NamedLookup> propertiesList = new List<NamedLookup>();
        protected NamedLookup settings;

        public BaseReleaseNotesGenerator(NamedLookup settings)
        {
            this.settings = settings;
            this.TFS = TFSAccessor.TFSAccessorFactory(settings["Team Project Path"], settings["Project Name"], settings["Iteration"]);
            this.logger = new Logger();
            this.silent = false;
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

        public NamedLookup getDefaultDetails()
        {
            NamedLookup sourceServerInformation = new NamedLookup("Details");
            sourceServerInformation["Web Server"] = settings["Web Server"];
            sourceServerInformation["Database Server"] = settings["Database Server"];
            sourceServerInformation["Database"] = settings["Database"];
            sourceServerInformation["Source"] = "(Changeset: " + TFS.getLatestChangesetNumber() + ")";
            return sourceServerInformation;
        }
    }
}
