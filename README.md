tfs-release-notes
=================

Generates customizeable release notes for a project hosted on Team Foundation Server with Word and Excel.


Usage
------------------
Run the executable (ReleaseNotes.exe) with exactly eight arguments, in the following order: 

* TFS Project Collection Path (eg https://mytfsserver.com/MyProjectCollection)
* TFS Project Name (eg MyProject)
* Iteration (The full name of the sprint/iteration for which to generate release notes)
* Release Notes Type (one of  `excel`, `word`, `html`)
* Database Server (arbitrary string - the name of the database server)
* Web Server (arbitrary string - the name of the web server)
* Database (arbitrary string - the name of the database)
* Web URL (arbitrary string - the location of the published application)