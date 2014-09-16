using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Client;

namespace ReleaseNotes
{
    class Program
    {
        static void Main(string[] args)
        {
            // create logger
            Logger logger = new Logger()
                .setMessage("Release Notes v.001")
                .setType(Logger.Type.Message)
                .display();
            try {

            // set vars from args (hardcoded until able to run with cmd line args)
            string projectName = "DealSpan";
            string iterationPath = "14.4";
            string generatorType = "EXCEL".ToLowerInvariant();
            string documentDescription = projectName + " " + iterationPath + " Release Notes";

            if (projectName == null || projectName == "") { throw new Exception("Project name invalid"); };
            if (iterationPath == null || iterationPath == "") { throw new Exception("Iteration path invalid."); }

            // create release notes generator
            ReleaseNotesGenerator generator = null;
                switch (generatorType)
                {
                    case "excel":
                        generator = ExcelGenerator.ExcelGeneratorFactory(documentDescription);
                        break;
                    case "word":
                        generator = WordGenerator.WordGeneratorFactory(documentDescription);
                        break;
                    case "html":
                        throw new NotImplementedException("Not implemented generator type");
                    default:
                        throw new NotImplementedException("Invalid generator type specified");
                }

                // set relevant vars and generate
                generator.setProjectName(projectName);
                generator.setIterationPath(iterationPath);
                generator.generateReleaseNotes();
            }
            catch (Exception e)
            {
                logger
                    .setType(Logger.Type.Error)
                    .setMessage(e.Message)
                    .display();
            }

            // wait for exit
            logger.setType(Logger.Type.Message).setMessage("Press any key to exit.").display();
            Console.ReadKey();
        }
    }
}
