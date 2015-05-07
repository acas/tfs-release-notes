using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using ReleaseNotesLibrary.Utility;

namespace ReleaseNotesLibrary.Generators
{
    public class HTMLGenerator : ReleaseNotesGenerator
    {
        private string htmlString = "";

        private HTMLGenerator(NamedLookup settings, bool silent)
            : base(settings, silent)
        {
            this.settings = settings;
            this.silent = silent;
        }

        /// <summary>
        /// HTML generator factory
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="silent"></param>
        /// <returns></returns>
        public static HTMLGenerator HTMLGeneratorFactory(NamedLookup settings, bool silent)
        {
            return new HTMLGenerator(settings, silent);
        }

        public override void CreateDocumentSpecificPreFormatting()
        {
            this.htmlString += "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" /></head>";
            this.htmlString += "<body style=\"font-family: Arial, sans-serif;\">";
            this.htmlString += "<style>";
            this.htmlString += @"table {
                                    border-collapse: collapse;
                                }
                                table, td, th {
                                    border: 1px solid black;
                                    padding: 5px 5px 5px 5px;
                                }";
            this.htmlString += "</style>";
        }

        public override void CreateDocumentSpecificPostFormatting(bool wide = false)
        {
            this.htmlString += "</body>";
        }

        public override void CreateTitle(string titleText)
        {
            this.htmlString += "<h2>" + titleText + "</h2>";
        }

        public override void CreateHeader(string headingText)
        {
            this.htmlString += "<h3>" + headingText + "</h3>";
        }

        public override void CreateHorizontalTable(NamedLookup data, int splits, bool header)
        {
            // create table start
            this.htmlString += "<table>";
            // 2 splits = 4 columns
            int currentColumnCount = 2 * splits;

            // if header needed
            if (header) CreateHeader(data.GetName());

            // get a list of the keys
            List<string> tableKeys = data.GetLookup().Keys.ToList();

            // determine the optimal number of rows for the table
            int optimalRowNumber = (tableKeys.Count() / splits) + (tableKeys.Count() % splits);

            int counter = 0;

            for (int i = 1; i <= optimalRowNumber; i++)
            {
                this.htmlString += "<tr>";
                for (int j = 1; j <= currentColumnCount; j++)
                {
                    this.htmlString += "<td>";
                    string currentKey = "";
                    if (counter != tableKeys.Count())
                        currentKey = tableKeys.ElementAt(counter);

                    if (j % 2 != 0)
                    {
                        this.htmlString += currentKey;
                    }
                    else
                    {
                        if (counter != tableKeys.Count())
                        {
                            if (currentKey.Equals("Source"))
                            {
                                this.htmlString += settings["Team Project Path"] + "/" + settings["Project Name"] + "/_workitems" + Environment.NewLine + data[currentKey];
                            }
                            else
                            {
                                this.htmlString += data[currentKey];
                            }
                            counter++;
                        }
                        else
                        {
                            this.htmlString += "";
                        }
                    }
                    this.htmlString += "</td>";
                }
                this.htmlString += "</tr>";
            }

            this.htmlString += "</table>";
        }

        public override void CreateVerticalTable(DataTable dataTable, string headerText, bool header)
        {
            this.htmlString += "<table>";

            // set the column count
            int currentColumnCount = dataTable.Columns.Count;

            // create header
            if (header)
            {
                CreateHeader(headerText);
            }

            // add header row
            string[] columnHeaderArray = Utilities.TableColumnsToStringArray(dataTable);
            this.htmlString += "<thead><tr>";
            foreach (string columnHeader in columnHeaderArray)
            {
                this.htmlString += "<th>" + columnHeader + "</th>";
            }
            this.htmlString += "</thead></tr>";

            // add table information
            foreach (DataRow row in dataTable.Rows)
            {
                this.htmlString += "<tr>";
                string[] rowArray = Utilities.TableRowToStringArray(row);
                foreach (string rowValue in rowArray)
                {
                    this.htmlString += "<td>" + rowValue.Replace("\n", "<br>") + "</td>";
                }
                this.htmlString += "</tr>";
            }

            this.htmlString += "</table>";
        }

        public override void CreateNamedSection(string headername, string text, string hyperlink)
        {
            NamedLookup namedSectionData = new NamedLookup(headername);
            namedSectionData[text] = hyperlink;
            CreateHorizontalTable(namedSectionData, 1, true);
        }

        public override void CreateErrorMessage(string message)
        {
            this.htmlString += "<h5>Error: " + message + "</h5>";
        }

        public override byte[] Save()
        {
            return System.Text.UTF8Encoding.UTF8.GetBytes(this.htmlString);
        }

        ~HTMLGenerator()
        {
            this.htmlString = "";
        }
    }
}
