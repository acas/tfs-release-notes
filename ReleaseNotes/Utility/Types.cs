using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReleaseNotes.Utility
{
    public struct NamedLookup
    {
        private string name;
        private Dictionary<string, string> lookup;
        private List<string> columnNames;

        public NamedLookup(string name)
        {
            this.name = name;
            this.lookup = new Dictionary<string, string>();
            this.columnNames = new List<string>();
        }

        public NamedLookup(string name, Dictionary<string, string> predefinedLookup)
        {
            this.name = name;
            this.lookup = predefinedLookup;
            this.columnNames = new List<string>();
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
                    lookup[name] = value.ToString();
            }
        }

        public void removeProperty(string name)
        {
            lookup.Remove(name);
        }

        public Dictionary<string, string> getLookup()
        {
            return this.lookup;
        }

        public void addColumnName(string columnName)
        {
            this.columnNames.Add(columnName);
        }

        public void addColumnNames(string[] columnNames)
        {
            foreach (string name in columnNames)
            {
                this.columnNames.Add(name);
            }
        }

        public List<string> getColumnNames()
        {
            return this.columnNames;
        }
    }
}
