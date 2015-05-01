using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.IO;

namespace ReleaseNotes.Utility
{
    class SaveFile
    {
        private JObject saveFileConfiguration;

        public SaveFile(JObject saveFileConfiguration)
        {
            this.saveFileConfiguration = saveFileConfiguration;
        }

        public static SaveFile CreateSaveFileFromPath(string path)
        {
            try
            {
                return new SaveFile(JObject.Parse(File.ReadAllText(path)));
            }
            catch (Exception e)
            {
                throw;
            }
        }

        public JObject GetInternalObject()
        {
            return this.saveFileConfiguration;
        }
    }
}
