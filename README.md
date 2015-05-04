tfs-release-notes
=================
Generates customizable release notes for a project hosted on Team Foundation Server with Word and Excel.

Usage
------------------
The ReleaseNotesWeb folder contains an ASP.NET Web application that you can run
on IIS to generate release notes. If you have Visual Studio, open the ReleaseNotes.sln
file, click on ReleaseNotesWeb in the Solution Explorer, then click on the play button
at the top of Visual Studio to run in your browser of choice.

The web form takes 8 params:

* TFS Project Collection Path (eg https://mytfsserver.com/tfs/MyProjectCollection)
* TFS Project Name (eg MyProject)
* TFS Project Subpath (any path between MyProjectCollection and Release 0.00, most likely will be not used)
* Iteration (The full name of the sprint/iteration for which to generate release notes, ie. Release 0.00)
* Database Server (arbitrary string - the name of the database server)
* Web Server (arbitrary string - the name of the web server)
* Database (arbitrary string - the name of the database)
* Web Location (arbitrary string - the location of the published application)
* Generator Type (either `excel` or `html`, but unfortunately `html` is not yet supported).