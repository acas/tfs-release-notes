﻿using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ReleaseNotesLibrary
{
    public class Utilities
    {
        /// <summary>
        /// The app data path for saving temp files
        /// </summary>
        public static string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ReleaseNotes\\";

        /// <summary>
        /// Strips contrived HTML from a string
        /// Stackoverflow: http://stackoverflow.com/questions/19523913/remove-html-tags-from-string-including-nbsp-in-c-sharp
        /// Slightly modified for effect
        /// </summary>
        /// <param name="inputHTML"></param>
        /// <param name="removeWhitespace"></param>
        /// <returns>An HTML (almost) free string</returns>
        public static string StripHtmlContrived(string inputHTML, bool removeWhitespace)
        {
            string noHTML = Regex.Replace(Regex.Replace(Regex.Replace(Regex.Replace(Regex.Replace(inputHTML, "<br>", "\n\n"), "<li>", "- "), "</p>|</div>|</ul>|</li>|</ol>", "\n"), "&nbsp;", "\n"), @"<[^>]+>|&quot;", "").Trim();
            if (removeWhitespace)
                noHTML = Regex.Replace(noHTML, @"\s{2,}", " ");
            noHTML = noHTML.Replace("&gt;", ">").Replace("&amp;", "&");
            noHTML = Regex.Replace(noHTML, "(\\n){3,}|(\n){3,}", "\n");
            return noHTML.Trim();
        }

        /// <summary>
        /// Gets the path the program is currently executing in
        /// </summary>
        /// <returns>The path (with ending slash) </returns>
        public static string GetExecutingPath()
        {
            return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "/";
        }

        /// <summary>
        /// Finds the greatest common integer divisor of two numbers
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static int GreatestCommonDivisor(int a, int b)
        {
            while (b != 0)
            {
                int temp = b;
                b = a % b;
                a = temp;
            }
            return a;
        }

        /// <summary>
        /// Takes a Pascal case identifier and turns it into separate words
        /// </summary>
        /// <param name="name"></param>
        /// <returns>A string where capitalize letters are split into separate words</returns>
        public static string SpaceCapitalizedNames(string name)
        {
            return string.Join(
                    string.Empty,
                    name.Select((x, i) => (
                    char.IsUpper(x) && i > 0 &&
                    (char.IsLower(name[i - 1]) || 
                    (i < name.Count() - 1 && char.IsLower(name[i + 1])))) ? " " + x : x.ToString()));
        }

        /// <summary>
        /// Implicitly mallocs a fixed size for a hyperlink, to be calculated
        /// </summary>
        /// <param name="existingBuffer"></param>
        /// <param name="sizeofRemaining"></param>
        /// <returns>A string with right padding given a particular size</returns>
        public static string ImplicitMalloc(string existingBuffer, int sizeofRemaining)
        {
            for (int i = 0; i < sizeofRemaining; i++) { existingBuffer += " "; }
            return existingBuffer;
        }

        /// <summary>
        /// Data table columns to string array
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string[] TableColumnsToStringArray(DataTable dt)
        {
            List<String> tableColumns = new List<String>();
            foreach (DataColumn dc in dt.Columns)
                tableColumns.Add(dc.ColumnName);
            return tableColumns.ToArray<string>();
        }

        /// <summary>
        /// Turns a data row into an array of string
        /// </summary>
        /// <param name="dr"></param>
        /// <returns></returns>
        public static string[] TableRowToStringArray(DataRow dr)
        {
            List<string> tableRows = new List<string>();
            string[] tableColumns = TableColumnsToStringArray(dr.Table);
            foreach (string columnName in tableColumns)
                tableRows.Add(dr[columnName].ToString());
            return tableRows.ToArray<String>();
        }
    }
}
