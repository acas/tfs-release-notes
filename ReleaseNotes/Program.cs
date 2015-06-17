using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.Framework.Client;
using ReleaseNotesLibrary.Utility;
using System.Threading;
using ReleaseNotesLibrary.Generators;
namespace ReleaseNotesLibrary
{
	public class Program
	{
        public static void programStart(string[] args)
        {
            // silent mode is true by default -
			// it's faster and quieter.
			bool silent = true;

			#if DEBUG
				silent = false;
			#endif
			
			// program header
			printProgramHeader();

			// create logger
			Logger logger = new Logger()
				.SetLoggingType(Logger.Type.Message);

			// try to generate the notes
			try
			{
				// arguments
				if (args.Length == 0) throw new Exception("Settings.json file path argument missing.");

				// create release notes generator
				ReleaseNotesGenerator generator = null;

				// set vars from args
				var configuration = (SaveFile.CreateSaveFileFromPath(args[0])).GetInternalObject();
				string generatorType = configuration.GetValue("Generator Type").ToString();

				var settings = new NamedLookup("Settings");
				settings["Team Project Path"] = configuration.GetValue("Team Project Path").ToString();
				settings["Project Name"] = configuration.GetValue("Project Name").ToString();
				settings["Project Subpath"] = configuration.GetValue("Project Subpath").ToString();
				settings["Iteration"] = configuration.GetValue("Iteration").ToString();
				settings["Database"] = configuration.GetValue("Database").ToString();
				settings["Database Server"] = configuration.GetValue("Database Server").ToString();
				settings["Web Server"] = configuration.GetValue("Web Server").ToString();
				settings["Doc Type"] = "APPLICATION BUILD/RELEASE NOTES\n";
				settings["Web Location"] = configuration.GetValue("Web Location").ToString();

				switch (generatorType.ToLowerInvariant())
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
				generator.GenerateReleaseNotes();
			}
			catch (Exception e)
			{
				// display error
				logger
					.SetLoggingType(Logger.Type.Error)
					.SetMessage(e.Message)
					.Display();
				Thread.Sleep(1000);
			}

			if (!silent) //if we're in silent mode, the program exits. The file has been saved.
			{
				// wait for exit
				logger.SetLoggingType(Logger.Type.General)
					.SetMessage("Press any key to exit.")
					.Display();

				// wait for key
				Console.ReadKey();
			}
        }

		/// <summary>
		/// Prints the program header
		/// </summary>
		public static void printProgramHeader()
		{
			Console.WriteLine();
			Console.ForegroundColor = ConsoleColor.White;
			Console.BackgroundColor = ConsoleColor.Blue;
			correctHeader(); Console.WriteLine("*******************************************************");
			correctHeader(); Console.WriteLine("* ACAS Release Notes                                  *");
			correctHeader(); Console.WriteLine("* Authors: Jon Fast / Aaron Greenwald                 *");
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
