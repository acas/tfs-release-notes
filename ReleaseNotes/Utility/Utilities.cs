using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ReleaseNotes
{
    class Utilities
    {
        /// <summary>
        /// Strips contrived HTML from a string
        /// Stackoverflow: http://stackoverflow.com/questions/19523913/remove-html-tags-from-string-including-nbsp-in-c-sharp
        /// Slightly modified for effect
        /// </summary>
        /// <param name="inputHTML"></param>
        /// <param name="removeWhitespace"></param>
        /// <returns>An HTML (almost) free string</returns>
        public static string stripHtmlContrived(string inputHTML, bool removeWhitespace)
        {
            string noHTML = Regex.Replace(inputHTML, @"<[^>]+>|&nbsp;|&quot;", "").Trim();
            if (removeWhitespace)
                noHTML = Regex.Replace(noHTML, @"\s{2,}", " ");
            return noHTML.Trim();
        }

        /// <summary>
        /// Gets the path the program is currently executing in
        /// </summary>
        /// <returns>The path (with ending slash) </returns>
        public static string getExecutingPath()
        {
            return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "/";
        }

        /// <summary>
        /// Finds the greatest common integer divisor of two numbers
        /// </summary>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <returns></returns>
        public static int greatestCommonDivisor(int a, int b)
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
        public static string spaceCapitalizedNames(string name)
        {
            return string.Join(
                    string.Empty,
                    name.Select((x, i) => (
                    char.IsUpper(x) && i > 0 &&
                    (char.IsLower(name[i - 1]) || 
                    (i < name.Count() - 1 && char.IsLower(name[i + 1])))) ? " " + x : x.ToString()));
        }
    }
}
