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

        public string GetName()
        {
            return this.name;
        }

        public string this[string Name]
        {
            get { return this.lookup[Name]; }
            set
            {
                if (Name != null && value != null)
                    lookup[Name] = value.ToString();
            }
        }

        public void RemoveProperty(string name)
        {
            lookup.Remove(name);
        }

        public Dictionary<string, string> GetLookup()
        {
            return this.lookup;
        }

        public void AddColumnName(string columnName)
        {
            this.columnNames.Add(columnName);
        }

        public void AddColumnNames(string[] columnNames)
        {
            foreach (string name in columnNames)
            {
                this.columnNames.Add(name);
            }
        }

        public List<string> GetColumnNames()
        {
            return this.columnNames;
        }
    }
}
