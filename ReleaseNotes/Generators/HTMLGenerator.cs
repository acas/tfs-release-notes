﻿using System;
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
        public HTMLGenerator(NamedLookup settings, bool silent)
            : base(settings, silent)
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

        public override byte[] Save()
        {
            throw new NotImplementedException();
        }
    }
}
