﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ReleaseNotes
{
    class StripHtml
    {
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
            string noHTML = Regex.Replace(inputHTML, @"<[^>]+>|&nbsp;|&quot;", "").Trim();
            if (removeWhitespace)
                noHTML = Regex.Replace(noHTML, @"\s{2,}", " ");
            return noHTML.Trim();
        }
    }
}
