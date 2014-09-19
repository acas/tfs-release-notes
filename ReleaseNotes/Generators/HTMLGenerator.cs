﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using ReleaseNotes.Utility;

namespace ReleaseNotes.Generators
{
    class HTMLGenerator : ReleaseNotesGenerator
    {
        public HTMLGenerator(NamedLookup settings)
            : base(settings)
        {
            throw new NotImplementedException();
        }

        public override void createTitle(string titleText)
        {
            throw new NotImplementedException();
        }

        public override void createHeading(string headingText)
        {
            throw new NotImplementedException();
        }

        public override void createHorizontalTable(NamedLookup nl, int splits, bool header)
        {
            throw new NotImplementedException();
        }

        public override void createVerticalTable(DataTable dt, string headerText, bool header)
        {
            throw new NotImplementedException();
        }

        public override void createDocumentSpecificPreFormatting()
        {
            throw new NotImplementedException();
        }

        public override void createDocumentSpecificPostFormatting()
        {
            throw new NotImplementedException();
        }

        public override void createNamedSection(string headername, string text, string hyperlink)
        {
            throw new NotImplementedException();
        }

        public override void createErrorMessage(string message)
        {
            throw new NotImplementedException();
        }

        public override void createHeaderGraphic(string path)
        {
            throw new NotImplementedException();
        }
    }
}
