using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Client;
using ReleaseNotes.Utility;
using System.Threading;
namespace ReleaseNotes
{
    class Program
    {
        static void Main(string[] args)
        {
            // silent mode is true by default -
			// it's faster and quieter.
            bool silent = true;

			#if DEBUG
				silent = false;
			#endif
            
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
                // create release notes generator
                ReleaseNotesGenerator generator = null;

                // check cmd args length
                if (args.Length != 8) { throw new IndexOutOfRangeException("Too many/few command line arguments. Exactly eight must be supplied."); }

                // set vars from args (hardcoded until able to run with cmd line args)
                string generatorType = args[3].ToLowerInvariant();

                var settings = new NamedLookup("Settings");
                settings["Team Project Path"] = args[0];
                settings["Project Name"] = args[1];
                settings["Iteration"] = args[2];
                settings["Database"] = args[6];
                settings["Database Server"] = args[4];
                settings["Web Server"] = args[5];
                settings["Doc Type"] = "APPLICATION BUILD/RELEASE NOTES\n";
                settings["Web Location"] = args[7];

                switch (generatorType)
                {
                    case "excel":
                        generator = ExcelGenerator.ExcelGeneratorFactory(settings, silent);
                        break;
                    case "word":
                        generator = WordGenerator.WordGeneratorFactory(settings, silent);
                        break;
                    case "html":
                        throw new NotImplementedException("Not implemented generator type");
                    default:
                        throw new Exception("Invalid generator type specified");
                }

                // generate
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

            if (!silent) //if we're in silent mode, the program exits. The file has been saved.
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
            correctHeader(); Console.WriteLine("*******************************************************");
            correctHeader(); Console.WriteLine("* ACAS Release Notes                                  *");
            correctHeader(); Console.WriteLine("* Author: Jon Fast                                    *");
            correctHeader(); Console.WriteLine("* License: MIT                                        *");
            correctHeader(); Console.WriteLine("*******************************************************");
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
        }

        static void correctHeader()
        {
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(" ");
            Console.BackgroundColor = ConsoleColor.Blue;
        }
    }
}
