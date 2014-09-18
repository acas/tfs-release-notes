using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReleaseNotes.Generators
{
    class HTMLGenerator : BaseReleaseNotesGenerator, IReleaseNotesGenerator
    {
        public HTMLGenerator(NamedLookup settings) : base(settings)
        {
            // nada yet
        }

        public void generateReleaseNotes()
        {
            throw new NotImplementedException();
        }
    }
}
