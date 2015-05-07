using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotesLibrary
{
    public static class Global
    {
        /// <summary>
        /// The app data path for saving temp files
        /// </summary>
        public static string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ReleaseNotes\\";
    }
}
