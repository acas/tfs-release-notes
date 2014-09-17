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
            // print the program header
            printProgramHeader();

            // create logger
            Logger logger = new Logger()
                .setType(Logger.Type.Message);

            // silent mode (server side?) is false by default
            bool silent = false;

            // try to generate the notes
            try
            {
                // set vars from args (hardcoded until able to run with cmd line args)
                string projectName = "DealSpan";
                string iterationPath = "14.4";
                string generatorType = "WORD".ToLowerInvariant();
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
                // display error
                logger
                    .setType(Logger.Type.Error)
                    .setMessage(e.Message)
                    .display();
            }

            if (!silent)
            {
                // wait for exit
                logger.setType(Logger.Type.General)
                    .setMessage("Press any key to exit.")
                    .display();

                // wait for key
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Prints the program header
        /// </summary>
        static void printProgramHeader()
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Blue;
            Console.WriteLine(" *******************************************************");
            Console.WriteLine(" * ACAS Release Notes                                  *");
            Console.WriteLine(" * Author: Jon Fast                                    *");
            Console.WriteLine(" * License: MIT                                        *");
            Console.WriteLine(" *******************************************************");
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
        }
    }
}
