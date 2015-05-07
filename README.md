TFS Release Notes
=================
TFS Release Notes automatically generates release notes for a project hosted on Team Foundation Server using Excel.
The release notes contain all PBIs in a given iteration plus any bugs in that iteration with specific tags. 
They also contain the test cases for those work items.

Prerequisites
------------------
* TFS Release Notes generates MS Excel documents using EPPlus. There are no installation prerequisites.
* The Windows user running the ReleaseNotesWeb application must have read access to the team project specified.
* The Windows user NETWORK SERVICE must be able to modify the contents of the folder: C:\Users\CurrentUser\AppDataFolder\Roaming\ReleaseNotes. Make sure it is not read-only.
* The ReleaseNotes library project must be built before ReleaseNotesWeb.

Usage
------------------
The ReleaseNotesWeb folder contains an ASP.NET Web application that you can run
on IIS to generate release notes. If you have Visual Studio, open the ReleaseNotes.sln
file, click on ReleaseNotesWeb in the Solution Explorer, then click on the play button
at the top of Visual Studio to run in your browser of choice, prefer Google Chrome!

If the build fails, you will need to build the ReleaseNotes project in the solution file. Change the build configuration to 'Release'
and click the Play button at the top of the screen. Then, build ReleaseNotesWeb. You must build with the 'Release' configuration selected
because the project file ReleaseNotesWeb points to the ReleaseNotesLibrary assembly in the solution's ReleaseNotes project folder: ...SolutionDir\ReleaseNotes\bin\Release\ReleaseNotesLibrary.dll

The default AppData folder for this application is C:\Users\CurrentUser\AppDataFolder\Roaming\ReleaseNotes,
which is required for both saving resource images and application persistence files.
The program will attempt to create the directory, though with IIS Express that may be impossible given permissions issues.
To resolve, simply give the local machine's 'NETWORK SERVICE' account access to the folder. 
You may need to create this folder by hand (though the application will attempt this for you).

The web form takes 8 parameters:

* TFS Project Collection Path (eg. https://mytfsserver.com/tfs/MyProjectCollection)
* TFS Project Name (eg. MyProject)
* Iteration (The full name of the sprint/iteration for which to generate release notes, eg. Release 0.00)
* TFS Project Subpath (any path between MyProjectCollection and Release 0.00, most likely will be not used)
* Database Server (arbitrary string - the name of the database server)
* Web Server (arbitrary string - the name of the web server)
* Database (arbitrary string - the name of the database)
* Web Location (arbitrary string - the location of the published application)
* Generator Type (either `excel` or `html`).