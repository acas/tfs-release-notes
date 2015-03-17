using ReleaseNotes.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

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
            throw new NotImplementedException();
        }

        public override void CreateTitle(string titleText)
        {
            throw new NotImplementedException();
        }

        public override void CreateHeader(string headingText)
        {
            throw new NotImplementedException();
        }

        public override void CreateHorizontalTable(NamedLookup nl, int splits, bool header)
        {
            throw new NotImplementedException();
        }

        public override void CreateVerticalTable(DataTable dt, string headerText, bool header)
        {
            throw new NotImplementedException();
        }

        public override void CreateDocumentSpecificPreFormatting()
        {
            throw new NotImplementedException();
        }

        public override void CreateDocumentSpecificPostFormatting(bool wide = false)
        {
            throw new NotImplementedException();
        }

        public override void CreateNamedSection(string headername, string text, string hyperlink)
        {
            throw new NotImplementedException();
        }

        public override void CreateErrorMessage(string message)
        {
            throw new NotImplementedException();
        }

        public override void CreateHeaderGraphic(string path)
        {
            throw new NotImplementedException();
        }
    }
}
