using ReleaseNotes.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotes.Generators
{
    class ExcelServerGenerator : ReleaseNotesGenerator
    {
        // this generator will use EPPlus

        /// <summary>
        /// Generates an EPPlus instance
        /// </summary>
        /// <param name="settings"></param>
        private ExcelServerGenerator(NamedLookup settings, bool silent) : base(settings, silent)
        {
            
        }
    }
}
