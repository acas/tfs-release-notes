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
            // silent mode (server side?) is false by default
            bool silent = false;
            
            if (!silent)
            {
                // print the program header
                printProgramHeader();
            }

            // create logger
            Logger logger = new Logger()
                .setType(Logger.Type.Message);

            // try to generate the notes
            try
            {
                // check cmd args length
                if (args.Length != 7) { throw new IndexOutOfRangeException("Too many/few command line arguments"); }

                // set vars from args (hardcoded until able to run with cmd line args)
                string projectName = args[0];
                string iterationPath = args[1];
                string generatorType = args[2].ToLowerInvariant();
                string documentDescription = projectName + " " + iterationPath + " Release Notes";
                string databaseServer = args[3];
                string webServer = args[4];
                string database = args[5];
                string webLink = args[6];

                if (projectName == null || projectName == "") { throw new Exception("Project name invalid."); };
                if (iterationPath == null || iterationPath == "") { throw new Exception("Iteration path invalid."); }
                if (databaseServer == null || databaseServer == "") { throw new Exception("Database server name invalid."); }
                if (webServer == null || webServer == "") { throw new Exception("Web server nmame invalid."); }
                if (database == null || database == "") { throw new Exception("Database invalid"); }
                if (webLink == null || webLink == "") { throw new Exception("Web link invalid"); }

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
                        throw new Exception("Invalid generator type specified");
                }

                // set relevant vars and generate
                generator.setProjectName(projectName);
                generator.setIterationPath(iterationPath);
                generator.setDatabase(database);
                generator.setDatabaseServer(databaseServer);
                generator.setWebServer(webServer);
                generator.setProjectWebLink(webLink);
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
